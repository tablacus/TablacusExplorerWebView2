//Tablacus Explorer

nTabMax = 0;
TabIndex = -1;
g_x = { Menu: null, Addons: null };
g_Chg = { Menus: false, Addons: false, Tab: false, Tree: false, View: false, Data: null };
g_arMenuTypes = ["Default", "Context", "Background", "Tabs", "Tree", "File", "Edit", "View", "Favorites", "Tools", "Help", "Systray", "System", "Alias"];
g_MenuType = "";
g_Id = "";
g_dlgAddons = null;
g_bDrag = false;
g_pt = { x: 0, y: 0 };
g_Gesture = null;
g_tidResize = null;
g_drag5 = false;
g_nResult = 0;
g_bChanged = true;
g_bClosed = false;
g_nSort = {
	"1_1" : 1,
	"1_3" : 1
};

g_ovPanel = null;

urlAddons = "https://tablacus.github.io/TablacusExplorerAddons/";
urlIcons = urlAddons + "te/iconpacks/";
xhr = null;
xmlAddons = null;

(async function () {
	arLangs = [await GetLangId()];
	var res = /(\w+)_/.exec(arLangs[0]);
	if (res && !/zh_cn/i.test(arLangs[0])) {
		arLangs.push(res[1]);
	}
	if (!/^en/.test(arLangs[0])) {
		arLangs.push("en");
	}
	arLangs.push("General");
	MainWindow.RunEvent1("BrowserCreated", document);
})();

async function SetDefaultLangID() {
	SetDefault(document.F.Conf_Lang, await GetLangId(true));
}

function SetDefault(o, v) {
	setTimeout(async function () {
		if (await confirmOk()) {
			SetValue(o, v);
		}
	}, 99);
}

function OpenGroup(id) {
	var o = document.getElementById(id);
	o.style.display = /block/i.test(o.style.display) ? "none" : "block";
}

async function LoadChecked(form) {
	for (var i = 0; i < form.length; ++i) {
		var o = form[i];
		var ar = o.id.split("=");
		if (ar.length > 1 && form[ar[0]].value == (await api.sscanf(ar[1], "0x%x") || ar[1])) {
			form[i].checked = true;
		}
		ar = o.id.split(":");
		if (ar.length > 1 && form[ar[0]].value & (await api.sscanf(ar[1], "0x%x") || ar[1])) {
			form[i].checked = true;
		}
	}
}

async function ResetForm() {
	var TV = await te.Ctrl(CTRL_TV);
	if (TV) {
		document.F.Tree_Align.value = await TV.Align;
		document.F.Tree_Style.value = await TV.Style;
		document.F.Tree_EnumFlags.value = await TV.EnumFlags;
		document.F.Tree_RootStyle.value = await TV.RootStyle;
		if (await TV.Align & 2) {
			document.F.Tree_Root.value = await TV.Root;
		}
		document.F.Tree_Width.value = await TV.Width;
	}
	var TC = await te.Ctrl(CTRL_TC);
	if (TC) {
		document.F.Tab_Style.value = await TC.Style;
		document.F.Tab_Align.value = await TC.Align;
		document.F.Tab_TabWidth.value = await TC.TabWidth;
		document.F.Tab_TabHeight.value = await TC.TabHeight;

		document.F.Tab_Left.value = await TC.Left;
		document.F.Tab_Top.value = await TC.Top;
		document.F.Tab_Width.value = await TC.Width;
		document.F.Tab_Height.value = await TC.Height;
	}
	var FV = await te.Ctrl(CTRL_FV);
	if (FV) {
		document.F.View_Type.value = await FV.Type;
		document.F.View_ViewMode.value = await FV.CurrentViewMode;
		document.F.View_fFlags.value = await FV.FolderFlags;
		document.F.View_Options.value = await FV.Options;
		document.F.View_ViewFlags.value = await FV.ViewFlags;
	}
	document.F.Conf_SizeFormat.value = await te.Data.Conf_SizeFormat || 0;
	for (i = 0; i < document.F.length; ++i) {
		o = document.F[i];
		if (SameText(o.type, 'checkbox')) {
			if (!/^!?Conf_/.test(o.id)) {
				o.checked = false;
			}
		}
	}
	await LoadChecked(document.F);
	var s = GetWebColor(document.F.Conf_TrailColor.value);
	document.F.Conf_TrailColor.value = s;
	document.F.Color_Conf_TrailColor.style.backgroundColor = s;
	document.getElementById("_EDIT").checked = true;
}

async function ResizeTabPanel() {
	var h = 4.8;
	if (await g_.Inline && !/\w/.test(document.getElementById('toolbar').innerHTML)) {
		h -= 2.4;
	}
	if (await g_.NoTab) {
		h -= 2.4;
	}
	CalcElementHeight(g_ovPanel, h);
}

async function ClickTab(o, nMode) {
	nMode = GetNum(nMode);
	if (o && o.id) {
		nTabIndex = o.id.replace(new RegExp('tab', 'g'), '') - 0;
	}
	var i = 0;
	var ovTab;
	while (ovTab = document.getElementById('tab' + i)) {
		var ovPanel = document.getElementById('panel' + i);
		if (i == nTabIndex) {
			try {
				ovTab.focus();
			} catch (e) { }
			ovTab.className = 'activetab';
			ovTab.style.zIndex = 11;
			ovPanel.style.display = 'block';
			g_ovPanel = ovPanel;
			await ResizeTabPanel();
			if (window.OnTabChanged) {
				OnTabChanged(i);
			}
		} else {
			ovTab.className = i < nTabIndex ? 'tab' : 'tab2';
			ovTab.style.zIndex = 10 - i;
			ovPanel.style.display = 'none';
		}
		++i;
	}
	nTabMax = i;
}

async function ClickTree(o, nMode, strChg, bForce) {
	if (strChg) {
		g_Chg[strChg] = true;
	}
	nMode = GetNum(nMode);
	var newTab = TabIndex != -1 ? TabIndex : 0;
	if (o && o.id) {
		var res = /tab([^_]+)(_?)(.*)/.exec(o.id);
		if (res) {
			newTab = res[1] + res[2] + res[3];
			ClickButton(res[1], true);
			if (res[3]) {
				if (res[1] == 1) {
					setTimeout(function (Id) {
						if (g_.elAddons[Id]) {
							if (!g_.elAddons[Id].contentWindow || !g_.elAddons[Id].contentWindow.document.body.innerHTML) {
								AddonOptions(Id, null, null, true);
							}
						}
					}, 999, res[3]);
				}
			}
			ShowButtons(/^1$|^1_1|^1_3$/.test(newTab), res[1] == 2, newTab);
			if (nMode == 0) {
				switch (res[1] - 0) {
					case 1:
						setTimeout(LoadAddons, 9);
						if (newTab == '1_1') {
							setTimeout(GetAddons, 9);
						}
						if (newTab == '1_3') {
							setTimeout(GetIconPacks, 9);
						}
						break;
					case 2:
						LoadMenus(res[3] - 0);
						break;
				}
			}
		}
	}
	if (newTab != TabIndex || bForce) {
		if (newTab == "0") {
			var o = document.getElementById("DefaultLangID");
			if (o && o.innerHTML == "") {
				var s = await GetLangId(1);
				var s2 = await GetLangId(2);
				if (s != s2) {
					s += ' (' + s2 + ')';
				}
				o.innerHTML = s;
			}
		}
		var ovTab = document.getElementById('tab' + TabIndex);
		if (ovTab) {
			var ovPanel = document.getElementById('panel' + TabIndex) || document.getElementById('panel' + TabIndex.replace(/_\w+/, ""));
			ovTab.className = 'button';
			ovPanel.style.display = 'none';
		}
		TabIndex = newTab;
		ovTab = document.getElementById('tab' + TabIndex);
		if (ovTab) {
			ovPanel = document.getElementById('panel' + TabIndex) || document.getElementById('panel' + TabIndex.replace(/_\w+/, ""));
			var res = /2_(.+)/.exec(TabIndex);
			if (res) {
				document.F.Menus.selectedIndex = res[1];
				SwitchMenus(document.F.Menus);
			}
			ovTab.className = 'hoverbutton';
			ovPanel.style.display = 'block';
			await CalcElementHeight(ovPanel, 3);
			var o = document.getElementById("tab_");
			await CalcElementHeight(o, 3);
			var i = ovTab.offsetTop - o.scrollTop;
			if (i < 0 || i >= o.offsetHeight - ovTab.offsetHeight) {
				ovTab.scrollIntoView(i < 0);
			}
		}
	}
}

function ClickButton(n, f) {
	var o = document.getElementById("tabbtn" + n);
	var op = document.getElementById("tab" + n + "_");
	if (f || (f === void 0 && /none/i.test(op.style.display))) {
		o.innerHTML = BUTTONS.opened;
		op.style.display = "block";
	} else {
		o.innerHTML = BUTTONS.closed;
		op.style.display = "none";
	}
}

function SetRadio(o) {
	var ar = o.id.split("=");
	o.form[ar[0]].value = ar[1];
	var res = /^(Tab|Tree|View|Conf)/.exec(ar[0]);
	if (res) {
		MainWindow.g_.OptionsWindow.g_Chg[res[1]] = true;
		MainWindow.g_.OptionsWindow.g_bChanged = true;
	}
}

async function SetCheckbox(o) {
	var ar = o.id.split(":");
	if (o.checked) {
		o.form[ar[0]].value |= (await api.sscanf(ar[1], "0x%x") || ar[1]);
	} else {
		o.form[ar[0]].value &= ~(await api.sscanf(ar[1], "0x%x") || ar[1]);
	}
	var res = /^(Tab|Tree|View|Conf)/.exec(ar[0]);
	if (res) {
		MainWindow.g_.OptionsWindow.g_Chg[res[1]] = true;
		MainWindow.g_.OptionsWindow.g_bChanged = true;
	}
}

function SetValue(n, v) {
	if (v.value != '!') {
		n.value = v.value.replace(/\\n/, "\n");
	}
}

function AddValue(name, i, min, max) {
	var o = document.F[name];
	o.value = Math.min(Math.max(i + GetNum(o.value), min), max);
}

function ChooseColor1(o) {
	setTimeout(async function () {
		var o2 = o.form[o.id.replace("Color_", "")];
		var c = await ChooseColor(o2.value || o2.placeholder);
		if (isFinite(c)) {
			o2.value = c;
			o.style.backgroundColor = await GetWebColor(c);
		}
	}, 99);
}

function ChooseColor2(o) {
	setTimeout(async function () {
		var o2 = o.form[o.id.replace("Color_", "")];
		var c = await ChooseColor(GetWinColor(o2.value || o2.placeholder));
		if (isFinite(c)) {
			c = await GetWebColor(c);
			o2.value = c;
			o.style.backgroundColor = c;
		}
	}, 99);
}

function ChooseFolder1(o) {
	setTimeout(async function () {
		var r = await BrowseForFolder(o.value);
		if (r) {
			if (await api.GetKeyState(VK_CONTROL) < 0 && /textarea/i.test(o.tagName)) {
				var ar = o.value.replace(/\s+$/g, "").split(/\n/);
				ar.push(r);
				o.value = ar.join("\n");
			} else {
				o.value = r;
			}
		}
	}, 99);
}

async function AddTabControl() {
	if (document.F.Tab_Width.value == 0) {
		MessageBox("Please enter the width.", TITLE, MB_ICONEXCLAMATION);
		return;
	}
	if (document.F.Tab_Height.value == 0) {
		MessageBox("Please enter the height.", TITLE, MB_ICONEXCLAMATION);
		return;
	}
	var TC = await te.CreateCtrl(CTRL_TC, document.F.Tab_Left.value, document.F.Tab_Top.value, document.F.Tab_Width.value, document.F.Tab_Height.value, document.F.Tab_Style.value, document.F.Tab_Align.value, document.F.Tab_TabWidth.value, document.F.Tab_TabHeight.value);
	TC.Selected.Navigate2("C:\\", SBSP_NEWBROWSER, document.F.View_Type.value, document.F.View_ViewMode.value, document.F.View_fFlags.value, 0, document.F.View_Options.value, document.F.View_ViewFlags.value);
}

async function DelTabControl() {
	var TC = await te.Ctrl(CTRL_TC);
	if (TC) {
		TC.Close();
	}
}

async function SetTabControls() {
	if (g_Chg.Tab) {
		if (document.getElementById("Conf_TabDefault").checked) {
			var cTC = await te.Ctrls(CTRL_TC);
			for (var i = 0; i < await cTC.Count; ++i) {
				SetTabControl(await cTC[i]);
			}
		} else {
			var TC = await te.Ctrl(CTRL_TC);
			if (TC) {
				SetTabControl(TC);
			}
		}
	}
}

async function SetFolderViews() {
	if (g_Chg.View) {
		var o = await api.CreateObject("Object");
		o.Layout = await te.Data.Conf_Layout;
		o.ViewMode = document.F.View_ViewMode.value;
		MainWindow.g_.TEData = o;
		o = await api.CreateObject("Object");
		o.All = document.getElementById("Conf_ListDefault").checked;
		o.FolderFlags = document.F.View_fFlags.value;
		o.Options = document.F.View_Options.value;
		o.ViewFlags = document.F.View_ViewFlags.value;
		o.Type = document.F.View_Type.value;
		MainWindow.g_.FVData = o;
	}
}

async function SetTreeControls() {
	if (g_Chg.Tree) {
		var o = await api.CreateObject("Object");
		o.All = document.getElementById("Conf_TreeDefault").checked;
		o.Align = document.F.Tree_Align.value;
		o.Style = document.F.Tree_Style.value;
		o.Width = document.F.Tree_Width.value;
		o.Root = document.F.Tree_Root.value;
		o.EnumFlags = document.F.Tree_EnumFlags.value;
		o.RootStyle = document.F.Tree_RootStyle.value;
		MainWindow.g_.TVData = o;
	}
}

async function SetTabControl(TC) {
	if (TC) {
		TC.Style = document.F.Tab_Style.value;
		TC.Align = document.F.Tab_Align.value;
		TC.TabWidth = document.F.Tab_TabWidth.value;
		TC.TabHeight = document.F.Tab_TabHeight.value;
	}
}

async function GetTabControl() {
	var TC = await te.Ctrl(CTRL_TC);
	if (TC) {
		document.F.Tab_Left.value = await TC.Left;
		document.F.Tab_Top.value = await TC.Top;
		document.F.Tab_Width.value = await TC.Width;
		document.F.Tab_Height.value = await TC.Height;
	}
}

async function MoveTabControl() {
	var TC = await te.Ctrl(CTRL_TC);
	if (TC) {
		if (document.F.Tab_Width.value != "" && document.F.Tab_Height.value != "") {
			TC.Left = document.F.Tab_Left.value;
			TC.Top = document.F.Tab_Top.value;
			TC.Width = document.F.Tab_Width.value;
			TC.Height = document.F.Tab_Height.value;
		}
	}
}

async function SwapTabControl() {
	var TC1 = await te.Ctrl(CTRL_TC);
	var cTC = await te.Ctrls(CTRL_TC, false);
	for (var list = api.CreateObject("Enum", cTC); !await list.atEnd(); await list.moveNext()) {
		var TC = await cTC[await list.item()];
		if (await TC.Visible && await TC.Id != await TC1.Id && await TC.Left == document.F.Tab_Left.value && await TC.Top == document.F.Tab_Top.value &&
			await TC.Width == document.F.Tab_Width.value && await TC.Height == document.F.Tab_Height.value) {
			TC.Left = await TC1.Left;
			TC.Top = await TC1.Top;
			TC.Width = await TC1.Width;
			TC.Height = await TC1.Height;
			TC1.Left = document.F.Tab_Left.value;
			TC1.Top = document.F.Tab_Top.value;
			TC1.Width = document.F.Tab_Width.value;
			TC1.Height = document.F.Tab_Height.value;
			break;
		}
	}
}

async function InitConfig(o) {
	var InstallPath = await fso.GetParentFolderName(api.GetModuleFileName(null));
	if (InstallPath == await te.Data.DataFolder) {
		return;
	}
	if (!await confirmOk()) {
		return;
	}
	await api.SHFileOperation(FO_MOVE, await fso.BuildPath(InstallPath, "layout"), await te.Data.DataFolder, 0, false);
	o.disabled = true;
}

function SelectPos(o, s) {
	var v = o[o.selectedIndex].value;
	if (v != "") {
		o.form[s].value = v;
	}
}

function SwitchMenus(o) {
	if (g_x.Menus) {
		g_x.Menus.style.display = "none";
		var o = o || document.F.elements.Menus;
		for (var i = o.length; i-- > 0;) {
			var a = o[i].value.split(",");
			if ("Menus_" + a[0] == g_x.Menus.name) {
				s = a[0] + "," + o.form["Menus_Base"].selectedIndex + "," + o.form["Menus_Pos"].value;
				if (s != o[i].value) {
					g_Chg.Menus = true;
					o[i].value = s;
				}
				break;
			}
		}
	}
	if (o) {
		var a = o.value.split(",");
		g_x.Menus = o.form["Menus_" + a[0]];
		g_x.Menus.style.display = "inline";
		o.form["Menus_Base"].selectedIndex = a[1];
		o.form["Menus_Pos"].value = GetNum(a[2]);
		CancelX("Menus");
	}
}

function SwitchX(mode, o, form) {
	g_x[mode].style.display = "none";
	g_x[mode] = (form || o.form)[mode + o.value];
	g_x[mode].style.display = "inline";
	CancelX(mode);
}

function ClearX(mode) {
	g_Chg.Data = null;
}

function CancelX(mode) {
	g_x[mode].selectedIndex = -1;
	EnableSelectTag(g_x[mode]);
}

ChangeX = function (mode) {
	g_Chg.Data = mode;
	g_bChanged = true;
}

function ConfirmX(bCancel, fn) {
	g_bChanged |= g_Chg.Data;
	return SetOptions(function () {
		SetChanged(fn);
	},
		function () {
			g_bChanged = false;
			ClearX(g_Chg.Data);
		}, !bCancel, true);
}

async function SetOptions(fnYes, fnNo) {
	if (g_nResult == 2 || document.activeElement && SameText(document.activeElement.value, await GetText("Cancel"))) {
		if (fnNo) {
			fnNo();
		}
		return false;
	}
	if (document.activeElement) {
		document.activeElement.blur();
	}
	if (g_nResult == 1 && !g_Chg.Data) {
		if (fnYes) {
			fnYes();
		}
		return true;
	}
	if (g_bChanged || g_Chg.Data) {
		switch (await MessageBox("Do you want to replace?", TITLE, MB_ICONQUESTION | MB_YESNO)) {
			case IDYES:
				if (fnYes) {
					fnYes();
				}
				return true;
			case IDNO:
				if (fnNo) {
					fnNo();
				}
				if (g_nResult == 1) {
					ClearX();
					if (fnYes) {
						fnYes();
					}
					return true;
				}
				return false;
			default:
				return false;
		}
	}
	if (fnNo) {
		fnNo();
	}
	return false;
}

async function EditMenus() {
	if (g_x.Menus.selectedIndex < 0) {
		return;
	}
	ClearX("Menus");
	var a = g_x.Menus[g_x.Menus.selectedIndex].value.split(g_sep);
	var a2 = a[0].split(/\\t/);
	if (!a[5]) {
		a2[0] = GetText(a2[0]);
	}
	document.F.Menus_Key.value = a2.length > 1 ? GetKeyName(a2.pop()) : "";
	document.F.Menus_Name.value = a2.join("\\t");
	document.F.Menus_Filter.value = a[1];
	var p = await api.CreateObject("Object");
	p.s = a[2];
	await MainWindow.OptionDecode(a[3], p);
	document.F.Menus_Path.value = await p.s;
	SetType(document.F.Menus_Type, a[3]);
	document.F.Icon.value = a[4] || "";
	document.F.IconSize.value = a[6] || "";
	SetImage();
}

function EditXEx(o, s, form) {
	if (document.getElementById("_EDIT").checked) {
		o(s, form);
	}
}

EditX = async function (mode, form) {
	if (g_x[mode].selectedIndex < 0) {
		return;
	}
	if (!form) {
		form = document.F;
	}
	ClearX(mode);
	var a = g_x[mode][g_x[mode].selectedIndex].value.split(g_sep);
	form[mode + mode].value = a[0];
	var p = await api.CreateObject("Object");
	p.s = a[1];
	await MainWindow.OptionDecode(a[2], p);
	form[mode + "Path"].value = await p.s;
	SetType(form[mode + "Type"], a[2]);
	if (api.StrCmpI(mode, "key") == 0) {
		SetKeyShift();
	}
	var o = form[mode + "Name"];
	if (o) {
		o.value = GetText(a[3] || "");
	}
}

function SetType(o, value) {
	value = value.replace(/&|\.\.\.$/g, "");
	var i = o.length;
	while (--i >= 0) {
		if (SameText(o[i].value, value)) {
			o.selectedIndex = i;
			break;
		}
	}
	if (i < 0) {
		i = o.length++;
		o[i].value = value;
		o[i].innerText = value;
		o.selectedIndex = i;
	}
}

function InsertX(sel) {
	if (!sel) {
		return;
	}
	++sel.length;
	if (sel.selectedIndex < 0) {
		sel.selectedIndex = sel.length - 1;
	} else {
		++sel.selectedIndex;
		for (var i = sel.length; i-- > sel.selectedIndex;) {
			sel[i].text = sel[i - 1].text;
			sel[i].value = sel[i - 1].value;
		}
	}
}

function AddX(mode, fn, form) {
	InsertX(g_x[mode]);
	(fn || ReplaceX)(mode, form);
	EnableSelectTag(g_x[mode]);
}

async function ReplaceMenus() {
	ClearX("Menus");
	if (g_x.Menus.selectedIndex < 0) {
		InsertX(g_x.Menus);
	}
	var sel = g_x.Menus[g_x.Menus.selectedIndex];
	var o = document.F.Menus_Type;
	var s = await GetSourceText(document.F.Menus_Name.value.replace(/\\/g, "/"));
	var org = (SameText(s, document.F.Menus_Name.value) && api.GetKeyState(VK_SHIFT) >= 0) ? "1" : ""
	if (document.F.Menus_Key.value.length) {
		s += "\\t" + GetKeyKeyG(document.F.Menus_Key.value);
	}
	var p = await api.CreateObject("Object");
	p.s = document.F.Menus_Path.value;
	await MainWindow.OptionEncode(o[o.selectedIndex].value, p);
	SetMenus(sel, [s, document.F.Menus_Filter.value, await p.s, o[o.selectedIndex].value, document.F.Icon.value, org, document.F.IconSize.value]);
	g_Chg.Menus = true;
}

async function ReplaceX(mode, form) {
	if (!g_x[mode]) {
		return;
	}
	if (!form) {
		form = document.F;
	}
	ClearX(mode);
	if (g_x[mode].selectedIndex < 0) {
		InsertX(g_x[mode]);
		EnableSelectTag(g_x[mode]);
	}
	var sel = g_x[mode][g_x[mode].selectedIndex];
	var o = form[mode + "Type"];
	var p = { s: form[mode + "Path"].value };
	MainWindow.OptionEncode(o[o.selectedIndex].value, p);
	var o2 = form[mode + "Name"];
	SetData(sel, [form[mode + mode].value, p.s, o[o.selectedIndex].value, o2 ? await GetSourceText(o2.value) : ""]);
	g_Chg[mode] = true;
	g_bChanged = true;
}

function RemoveX(mode) {
	ClearX(mode);
	if (g_x[mode].selectedIndex < 0 || !confirmOk()) {
		return;
	}
	var i = g_x[mode].selectedIndex;
	var j = i;
	while (j >= 0 && g_x[mode][j]) {
		g_x[mode][j] = null;
		j = g_x[mode].selectedIndex;
	}
	g_Chg[mode] = true;
	g_bChanged = true;
	if (i >= g_x[mode].length) {
		i = g_x[mode].length - 1;
	}
	if (i >= 0) {
		g_x[mode].selectedIndex = i;
		FireEvent(g_x[mode][i], "change");
	}
}

function MoveX(mode, n) {
	if (n < 0) {
		for (var i = 0; i < g_x[mode].length + n; ++i) {
			if (!g_x[mode][i].selected && g_x[mode][i + 1].selected) {
				var ar = [g_x[mode][i].text, g_x[mode][i].value];
				g_x[mode][i].text = g_x[mode][i + 1].text;
				g_x[mode][i].value = g_x[mode][i + 1].value;
				g_x[mode][i + 1].text = ar[0];
				g_x[mode][i + 1].value = ar[1];
				g_x[mode][i + 1].selected = false;
				g_x[mode][i].selected = true;
			}
		}
	} else {
		for (var i = g_x[mode].length; i-- > n;) {
			if (!g_x[mode][i].selected && g_x[mode][i - 1].selected) {
				var ar = [g_x[mode][i].text, g_x[mode][i].value];
				g_x[mode][i].text = g_x[mode][i - 1].text;
				g_x[mode][i].value = g_x[mode][i - 1].value;
				g_x[mode][i - 1].text = ar[0];
				g_x[mode][i - 1].value = ar[1];
				g_x[mode][i - 1].selected = false;
				g_x[mode][i].selected = true;
			}
		}
	}
	g_Chg[mode] = true;
	g_bChanged = true;
	EnableSelectTag(g_x[mode]);
}

function SetMenus(sel, a) {
	sel.value = PackData(a);
	var a2 = a[0].split(/\\t/);
	sel.text = [a[5] ? a2[0] : GetText(a2[0]), a[1]].join(" ").replace(/[\r\n].*/, "");
	if (!sel.text.length) {
		sel.text = "********";
	}
}

async function LoadMenus(nSelected) {
	var oa = document.F.Menus_Type;
	if (!g_x.Menus) {
		var arFunc = await api.CreateObject("Array");
		await MainWindow.RunEvent1("AddType", arFunc);
		for (var i = 0; i < await GetLength(arFunc); ++i) {
			var o = oa[++oa.length - 1];
			o.value = arFunc[i];
			o.innerText = await GetText(await arFunc[i]);
		}

		oa = document.F.Menus;
		oa.length = 0;

		for (var j in g_arMenuTypes) {
			var s = g_arMenuTypes[j];
			document.getElementById("Menus_List").insertAdjacentHTML("BeforeEnd", ['<select name="Menus_', s, '" size="17" style="width: 12em; height: 32em; height: calc(100vh - 6em); min-height: 20em; display: none" onchange="EditXEx(EditMenus)" ondblclick="EditMenus()" oncontextmenu="CancelX(\'Menus\')" multiple></select>'].join(""));
			var menus = await teMenuGetElementsByTagName(s);
			if (menus && await GetLength(menus)) {
				oa[++oa.length - 1].value = s + "," + await menus[0].getAttribute("Base") + "," + await menus[0].getAttribute("Pos");
				var o = document.F["Menus_" + s];
				var items = await menus[0].getElementsByTagName("Item");
				if (items) {
					var i = await GetLength(items);
					o.length = i;
					while (--i >= 0) {
						var item = await items[i];
						SetMenus(o[i], [await item.getAttribute("Name"), await item.getAttribute("Filter"), await item.text, await item.getAttribute("Type"), await item.getAttribute("Icon"), await item.getAttribute("Org"), await item.getAttribute("Height")]);
					}
				}
			} else {
				oa[++oa.length - 1].value = s;
			}
			oa[oa.length - 1].text = await GetText(s);
		}
		SwitchMenus(oa[nSelected]);
	}
	for (var j in g_arMenuTypes) {
		var ar = String(g_MenuType).split(",");
		if (SameText(ar[0], g_arMenuTypes[j])) {
			nSelected = oa.length - 1;
			oa[nSelected].selected = true;
			g_MenuType = j;
			setTimeout(async function (o, v) {
				await ClickTree(document.getElementById("tab2_" + g_MenuType));
				await EditMenus();
				g_MenuType = "";
				if (isFinite(v)) {
					o.selectedIndex = v;
					EnableSelectTag(o);
					FireEvent(o, "change");
				}
			}, 99, document.F["Menus_" + ar[0]], ar[1]);
		}
	}
}

async function LoadX(mode, fn, form) {
	if (!g_x[mode]) {
		if (!form) {
			form = document.F;
		}
		var arFunc = await api.CreateObject("Array");
		await MainWindow.RunEvent1("AddType", arFunc);
		var oa = form[mode + "Type"] || form.Type;
		while (oa.length) {
			oa.removeChild(oa[0]);
		}
		for (var i = 0; i < await GetLength(arFunc); ++i) {
			var o = oa[++oa.length - 1];
			o.value = arFunc[i];
			o.innerText = await GetText(arFunc[i]);
		}
		g_x[mode] = form[mode + "All"];
		if (g_x[mode]) {
			oa = form[mode];
			oa.length = 0;
			var xml = await OpenXml(mode + ".xml", false, true);
			for (var j in g_Types[mode]) {
				var o = oa[++oa.length - 1];
				o.text = await GetTextEx(g_Types[mode][j]);
				o.value = g_Types[mode][j];
				o = form[mode + g_Types[mode][j]];
				var items = await xml.getElementsByTagName(g_Types[mode][j]);
				var i = await GetLength(items);
				if (i == 0 && g_Types[mode][j] == "List") {
					items = await xml.getElementsByTagName("Folder");
					i = await GetLength(items);
				}
				o.length = i;
				while (--i >= 0) {
					var item = await items[i];
					var s = await item.getAttribute(mode);
					if (SameText(mode, "Key")) {
						var ar = /,$/.test(s) ? [s] : s.split(",");
						for (var k = ar.length; k--;) {
							ar[k] = await GetKeyNameG(ar[k]);
						}
						s = ar.join(",");
					}
					SetData(o[i], [s, await item.text, await item.getAttribute("Type"), await item.getAttribute("Name")]);
				}
			}
		} else {
			g_x[mode] = form.List;
			g_x[mode].length = 0;
			var path = await fso.GetParentFolderName(await api.GetModuleFileName(null));
			var xml = await te.Data["xml" + AddonName];
			if (!xml) {
				xml = await api.CreateObject("Msxml2.DOMDocument");
				xml.async = false;
				await xml.load(await fso.BuildPath(path, "config\\" + AddonName + ".xml"));
				te.Data["xml" + AddonName] = xml;
			}

			var items = await xml.getElementsByTagName("Item");
			var i = await GetLength(items);
			g_x[mode].length = i;
			while (--i >= 0) {
				var item = await items[i];
				SetData(g_x[mode][i], [await item.getAttribute("Name"), await item.text, await item.getAttribute("Type"), await item.getAttribute("Icon"), await item.getAttribute("Height")]);
			}
			xml = null;
		}
		EnableSelectTag(g_x[mode]);
		fn && fn();
	}
}

async function SaveMenus() {
	if (g_Chg.Menus) {
		var xml = await CreateXml();

		var root = await xml.createElement("TablacusExplorer");
		for (var j in g_arMenuTypes) {
			var o = document.F["Menus_" + g_arMenuTypes[j]];
			var items = await xml.createElement(g_arMenuTypes[j]);
			var a = document.F.elements.Menus[j].value.split(",");
			items.setAttribute("Base", GetNum(a[1]));
			items.setAttribute("Pos", GetNum(a[2]));
			for (var i = 0; i < o.length; ++i) {
				var item = await xml.createElement("Item");
				var a = o[i].value.split(g_sep);
				item.setAttribute("Name", a[0]);
				item.setAttribute("Filter", a[1]);
				item.text = a[2];
				item.setAttribute("Type", a[3]);
				item.setAttribute("Icon", a[4]);
				if (a[5]) {
					item.setAttribute("Org", 1);
				}
				if (a[6]) {
					item.setAttribute("Height", a[6]);
				}
				await items.appendChild(item);
			}
			await root.appendChild(items);
		}
		await xml.appendChild(root);
		te.Data.xmlMenus = xml;
		await MainWindow.RunEvent1("ConfigChanged", "Menus");
	}
}

async function GetKeyKeyEx(s) {
	var n = await GetKeyKey(s);
	return n & 0xff ? await api.sprintf(9, "$%x", n) : s;
}

async function GetKeyKeyG(s) {
	var n = await GetKeyKeyEx(s);
	var s = await GetKeyName(n, true);
	return /^\w$|\+\w$/i.test(s) ? s : n;
}

async function GetKeyNameG(s) {
	return await GetKeyName(/^$/.test(s) ? s : await GetKeyKeyEx(s));
}

async function SaveX(mode, form) {
	if (g_Chg[mode]) {
		var xml = await CreateXml();
		var root = await xml.createElement("TablacusExplorer");
		for (var j in g_Types[mode]) {
			var o = (form || document.F)[mode + g_Types[mode][j]];
			for (var i = 0; i < o.length; ++i) {
				var item = await xml.createElement(g_Types[mode][j]);
				var a = o[i].value.split(g_sep);
				var s = a[0];
				if (SameText(mode, "key")) {
					var ar = /,$/.test(s) ? [s] : s.split(",");
					for (var k = ar.length; k--;) {
						ar[k] = await GetKeyKeyG(ar[k]);
					}
					s = ar.join(",");
				} else {
					item.setAttribute("Name", a[3]);
				}
				item.setAttribute(mode, s);
				item.text = a[1];
				item.setAttribute("Type", a[2]);
				await root.appendChild(item);
			}
		}
		await xml.appendChild(root);
		await SaveXmlEx(mode.toLowerCase() + ".xml", xml);
	}
}

async function SaveAddons() {
	try {
		if (g_Chg.Addons || await te.Data.bErrorAddons) {
			te.Data.bErrorAddons = false;
			var xml = await CreateXml();
			var root = await xml.createElement("TablacusExplorer");
			var table = document.getElementById("Addons");
			for (var j = 0; j < table.rows.length; ++j) {
				var div = table.rows(j).cells(0).firstChild;
				var Id = div.id.replace("Addons_", "").toLowerCase();
				var item = null;
				var items = await te.Data.Addons.getElementsByTagName(Id);
				if (GetLength(items)) {
					item = await items[0].cloneNode(true);
				}
				if (!item) {
					item = await xml.createElement(Id);
				}
				var Enabled = document.getElementById("enable_" + Id).checked;
				if (Enabled) {
					var AddonFolder = fso.BuildPath(fso.GetParentFolderName(api.GetModuleFileName(null)), "addons\\" + Id);
					Enabled = fso.FolderExists(AddonFolder + "\\lang") ? 2 : 0;
					if (fso.FileExists(AddonFolder + "\\script.vbs")) {
						Enabled |= 8;
					}
					if (fso.FileExists(AddonFolder + "\\script.js")) {
						Enabled |= 1;
					}
					Enabled = (Enabled & 9) ? Enabled : 4;
				} else {
					await MainWindow.AddonDisabled(Id);
				}
				item.setAttribute("Enabled", Enabled);
				await root.appendChild(item);
			}
			await xml.appendChild(root);
			te.Data.Addons = xml;
			await MainWindow.RunEvent1("ConfigChanged", "Addons");
		}
	} catch (e) { }
}

function SetChanged(fn, form) {
	if (g_Chg.Data) {
		g_bChanged = true;
		if (g_x[g_Chg.Data] && g_x[g_Chg.Data].selectedIndex >= 0) {
			(fn || ReplaceX)(g_Chg.Data, form);
		} else {
			AddX(g_Chg.Data, fn, form);
		}
	}
}

function SetData(sel, a, t) {
	sel.value = PackData(a);
	if ("boolean" === typeof t) {
		t = (t ? String.fromCharCode(9745, 32) : String.fromCharCode(9744, 32)) + a[0];
	}
	sel.text = t || a[0];
}

function PackData(a) {
	for (var i = a.length; i-- > 0;) {
		a[i] = String(a[i] || "").replace(g_sep, "`  ~");
	}
	return a.join(g_sep);
}

async function LoadAddons() {
	if (g_x.Addons) {
		OpenAddonsOptions();
		return;
	}
	g_x.Addons = true;

	var AddonId = [];
	var wfd = await api.Memory("WIN32_FIND_DATA");
	var path = await fso.BuildPath(await fso.GetParentFolderName(await api.GetModuleFileName(null)), "addons\\");
	var hFind = await api.FindFirstFile(path + "*", wfd);
	for (var bFind = hFind != INVALID_HANDLE_VALUE; await bFind; bFind = await api.FindNextFile(hFind, wfd)) {
		var Id = await wfd.cFileName;
		if (Id != "." && Id != ".." && !AddonId[Id]) {
			AddonId[Id] = 1;
		}
	}
	api.FindClose(hFind);

	var table = document.getElementById("Addons");
	table.ondragover = Over5;
	table.ondrop = Drop5;
	table.deleteRow(0);
	var root = await te.Data.Addons.documentElement;
	if (root) {
		var items = await root.childNodes;
		if (items) {
			for (var i = 0; i < await GetLength(items); ++i) {
				var item = await items[i];
				var Id = await item.nodeName;
				if (AddonId[Id]) {
					AddAddon(table, Id, GetNum(await item.getAttribute("Enabled")));
					delete AddonId[Id];
				}
			}
		}
	}
	for (var Id in AddonId) {
		if (await fso.FileExists(path + Id + "\\config.xml")) {
			AddAddon(table, Id, false);
		}
	}
	OpenAddonsOptions();
}

function OpenAddonsOptions() {
	if (g_Id) {
		if (document.getElementById("opt_" + g_Id)) {
			AddonOptions(g_Id);
		}
		g_Id = "";
	}
}

function AddAddon(table, Id, bEnable, Alt) {
	SetAddon(Id, bEnable, table.insertRow().insertCell(), Alt);
	if (!Alt) {
		var sorted = document.getElementById("SortedAddons");
		if (sorted.rows.length) {
			SetAddon(Id, bEnable, sorted.insertRow().insertCell(), "Sorted_");
		}
	}
}

async function SetAddon(Id, bEnable, td, Alt) {
	if (!td) {
		td = document.getElementById(Alt || "Addons_" + Id).parentNode;
	}
	var info = await GetAddonInfo(Id);
	var bMinVer = await info.MinVersion && await AboutTE(0) < await CalcVersion(await info.MinVersion);
	if (bMinVer) {
		bEnable = false;
	}
	var s = ['<div ', (Alt ? '' : 'draggable="true" ondragstart="Start5(this)" ondragend="End5(this)"'), ' title="', Id, '" Id="', Alt || "Addons_", Id, '"', bEnable ? "" : ' class="disabled"', '>'];
	s.push('<table><tr style="border-top: 1px solid buttonshadow"><td>', (Alt ? '&nbsp;' : '<input type="radio" name="AddonId" id="_' + Id + '">'), '</td><td style="width: 100%"><label for="_', Id, '">', await info.Name, "&nbsp;", await info.Version, '<br><a href="#" onclick="return AddonInfo(\'', Id, '\', this)"  class="link" style="font-size: .9em">', GetText('Details'), ' (', Id, ')</a>');
	s.push(' <a href="#" onclick="AddonRemove(\'', Id, '\'); return false;" style="color: red; font-size: .9em; white-space: nowrap; margin-left: 2em">', GetText('Delete'), "</a>");
	if (bMinVer) {
		s.push('</td><td class="danger" style="align: right; white-space: nowrap; vertical-align: middle">', await info.MinVersion.replace(/^20/, (await api.LoadString(hShell32, 60) || "%").replace(/%.*/, "")).replace(/\.0/g, '.'), ' ', await GetText("is required."), '</td>');
	} else if (await info.Options) {
		s.push('</td><td style="white-space: nowrap; vertical-align: middle; padding-right: 1em"><a href="#" onclick="AddonOptions(\'', Id, '\'); return false;" class="link" id="opt_', Id, '">', await GetText('Options'), '</td>');
	}
	s.push('<td style="vertical-align: middle', bMinVer ? ';display: none"' : "", '"><input type="checkbox" ', (Alt ? "" : 'id="enable_' + Id + '"'), ' onclick="AddonEnable(this, \'', Id, '\')" ', bEnable ? " checked" : "", '></td>');
	s.push('<td style="vertical-align: middle', bMinVer ? ';display: none"' : "", '"><label for="enable_', Id, '" style="display: block; width: 6em; white-space: nowrap">', GetText(bEnable ? "Enabled" : "Enable"), '</label></td>');
	s.push('</tr></table></label></div>');
	td.innerHTML = s.join("");
	ApplyLang(td);
	if (!Alt) {
		var div = document.getElementById("Sorted_" + Id);
		if (div) {
			SetAddon(Id, bEnable, div.parentNode, "Sorted_");
		}
	}
}

async function Start5(o) {
	if (await api.GetKeyState(VK_LBUTTON) < 0) {
		event.dataTransfer.effectAllowed = 'move';
		g_drag5 = o.id;
		return true;
	}
	return false;
}

function End5() {
	g_drag5 = false;
}

function Over5() {
	if (g_drag5) {
		if (event.preventDefault) {
			event.preventDefault();
		} else {
			event.returnValue = false;
		}
	}
}

function Drop5(e) {
	if (g_drag5) {
		var o = document.elementFromPoint(e.clientX, e.clientY);
		do {
			if (/Addons_/i.test(o.id)) {
				setTimeout(function (src, dest) {
					AddonMoveEx(src, dest);
				}, 99, GetRowIndexById(g_drag5), GetRowIndexById(o.id));
				break;
			}
		} while (o = o.parentNode);
	}
}

function GetRowIndexById(id) {
	try {
		var o = document.getElementById(id);
		if (o) {
			while (o = o.parentNode) {
				if (o.rowIndex !== void 0) {
					return o.rowIndex;
				}
			}
		}
	} catch (e) { }
}

async function AddonInfo(Id, o) {
	o.style.textDecoration = "none";
	o.style.color = GetWebColor(MainWindow.GetSysColor(COLOR_WINDOWTEXT));
	o.style.cursor = "default";
	o.onclick = null;
	var info = await GetAddonInfo(Id);
	o.innerHTML = await info.Description;
	return false;
}

async function AddonWebsite(Id) {
	var info = await GetAddonInfo(Id);
	wsh.run(await info.URL);
}

function AddonEnable(o, Id) {
	var div = document.getElementById("Addons_" + Id);
	SetAddon(Id, o.checked);
	document.getElementById("enable_" + Id).checked = o.checked;
	g_Chg.Addons = true;
}

async function OptionMove(dir) {
	if (/^1/.test(TabIndex)) {
		var r = document.F.AddonId;
		for (i = 0; i < r.length; ++i) {
			if (r[i].checked) {
				if (await api.GetKeyState(VK_CONTROL) < 0) {
					if (dir < 0) {
						dir = -i;
					} else {
						dir = document.getElementById("Addons").rows.length - i - 1;
					}
				}
				try {
					AddonMoveEx(i, i + dir);
				} catch (e) { }
				break;
			}
		}
	} else if (/^2/.test(TabIndex)) {
		if (g_x.Menus.selectedIndex < 0 || g_x.Menus.selectedIndex + dir < 0 || g_x.Menus.selectedIndex + dir >= g_x.Menus.length) {
			return;
		}
		MoveX("Menus", dir);
	}
}

function AddonMoveEx(src, dest) {
	var table = document.getElementById("Addons");
	if (dest < 0 || dest >= table.rows.length || src == dest) {
		return false;
	}
	var tr = table.rows(src);
	var td = tr.cells(0);

	var s = td.innerHTML
	var md = td.onmousedown;
	var mu = td.onmouseup;
	var mm = td.onmousemove;

	table.deleteRow(src);

	tr = table.insertRow(dest);
	td = tr.insertCell();
	td.innerHTML = s;
	td.onmousedown = md;
	td.onmouseup = mu;
	td.onmousemove = mm;
	var o = document.getElementById('panel1');
	var i = td.offsetTop - o.scrollTop;
	if (i < 0 || i >= o.offsetHeight - td.offsetHeight) {
		td.scrollIntoView(i < 0);
	}
	document.F.AddonId[dest].checked = true;
	g_Chg.Addons = true;
	return false;
}

async function AddonRemove(Id) {
	if (!await confirmOk()) {
		return;
	}
	MainWindow.AddonDisabled(Id);
	if (await AddonBeforeRemove(Id) < 0) {
		return;
	}
	var sf = await api.Memory("SHFILEOPSTRUCT");
	sf.hwnd = await api.GetForegroundWindow();
	sf.wFunc = FO_DELETE;
	sf.fFlags = 0;
	sf.pFrom = await fso.BuildPath(await fso.GetParentFolderName(await api.GetModuleFileName(null)), "addons\\" + Id) + "\0";
	if (await api.SHFileOperation(sf) == 0) {
		if (!await sf.fAnyOperationsAborted) {
			var table = document.getElementById("Addons");
			table.deleteRow(GetRowIndexById("Addons_" + Id));
			var table = document.getElementById("SortedAddons");
			if (table.rows.length) {
				table.deleteRow(GetRowIndexById("Sorted_" + Id));
			}
			g_Chg.Addons = true;
		}
	}
}

async function ApplyOptions() {
	var hwnd = await GetBrowserWindow(document);
	var hwnd1 = hwnd;
	while (hwnd1 = await api.GetParent(hwnd)) {
		hwnd = hwnd1;
	}
	if (await api.GetKeyState(VK_SHIFT) >= 0 && !await api.IsZoomed(hwnd) && !await api.IsIconic(hwnd)) {
		var r = 12 / (Math.abs(await MainWindow.DefaultFont.lfHeight) || 12);
		te.Data.Conf_OptWidth = GetNum(document.documentElement.offsetWidth || document.body.offsetWidth) * r;
		te.Data.Conf_OptHeight = GetNum(document.documentElement.offsetHeight || document.body.offsetHeight) * r;
	}
	SetChanged(ReplaceMenus);
	for (var i in document.F.elements) {
		if (!/=|:/.test(i)) {
			var res = /^(!?)(Tab_.+|Tree_.+|View_.+|Conf_.+)/.exec(i);
			if (res && !/_$/.test(i)) {
				var v = GetElementValue(document.F[i]);
				te.Data[res[2]] = res[1] ? !v : v;
			}
		}
	}
	await SaveAddons();
	await SaveMenus();
	await SetTabControls();
	await SetTreeControls();
	await SetFolderViews();
	await MainWindow.RunEvent1("ConfigChanged", "Config");

	try {
		for (var esw = await new Enumerator(sha.Windows()); !await esw.atEnd(); await esw.moveNext()) {
			var x = await esw.item();
			if (x && await x.Document) {
				var w = await x.Document.parentWindow;
				if (w && await w.te && await w.te.Data && await te.Data.DataFolder == await w.te.Data.DataFolder) {
					w.te.Data.bReload = true;
				}
			}
		}
	} catch (e) { }
}

async function CancelOptions() {
	if (await te.Data.bErrorAddons) {
		await SaveAddons();
		te.Data.bReload = true;
	}
}

InitOptions = await function () {
	ApplyLang(document);
	document.getElementById("tab1_3").innerHTML = api.sprintf(99, GetText("Get %s"), GetText("Icon"));
	document.title = GetText("Options") + " - " + TITLE;
	MainWindow.g_.OptionsWindow = window;
	var InstallPath = fso.GetParentFolderName(api.GetModuleFileName(null));
	document.F.ButtonInitConfig.disabled = (InstallPath == te.Data.DataFolder) | !fso.FolderExists(fso.BuildPath(InstallPath, "layout"));
	for (i in document.F.elements) {
		if (!/=|:/.test(i)) {
			var res = /^(!?)(Tab_.+|Tree_.+|View_.+|Conf_.+)/.exec(i);
			if (res) {
				var v = te.Data[res[2]];
				if (v !== void 0 || res[1]) {
					SetElementValue(document.F[i], res[1] ? !v : v);
				}
			}
		}
	}

	await ResetForm();
	var s = [];
	for (var i in g_arMenuTypes) {
		var j = g_arMenuTypes[i];
		s.push('<label id="tab2_' + i + '" class="button" style="width: 100%" onmousedown="ClickTree(this, null, \'Menus\');">' + GetText(j) + '</label><br>');
	}
	document.getElementById("tab2_").innerHTML = s.join("");

	AddEventEx(window, "load", function () {
		for (var i = 6; i--;) {
			ClickButton(i, false);
		}
		SetTab(dialogArguments.Data);
	});

	AddEventEx(window, "resize", function () {
		clearTimeout(g_tidResize);
		g_tidResize = setTimeout(function () {
			ClickTree(null, null, null, true);
		}, 500);
	});

	SetOnChangeHandler();
	AddEventEx(window, "beforeunload", function () {
		alert("unload");
		for (var i in g_.elAddons) {
			try {
				var w = g_.elAddons[i].contentWindow;
				w.g_nResult = g_nResult;
				if (g_nResult == 1) {
					w.TEOk();
				}
				w.close();
			} catch (e) { }
		}
		g_.elAddons = {};
		g_bChanged |= g_Chg.Addons || g_Chg.Menus || g_Chg.Tab || g_Chg.Tree || g_Chg.View;
		SetOptions(ApplyOptions, CancelOptions);
	});
}

OpenIcon = function (o) {
	setTimeout(async function () {
		var data = [];
		var a = o.id.split(/,/);
		if (a[0] == "b") {
			var dllpath = await fso.BuildPath(system32, "ieframe.dll");
			a[0] = await fso.GetFileName(dllpath);
			var a1 = a[1];
			var hModule = await LoadImgDll(a, 0);
			if (hModule) {
				var lpbmp = isFinite(a[1]) ? a[1] - 0 : a[1];
				var himl = await api.ImageList_LoadImage(hModule, lpbmp, a[2], CLR_NONE, CLR_NONE, IMAGE_BITMAP, LR_CREATEDIBSECTION | LR_LOADTRANSPARENT);
				a[1] = a1;
				var nCount = himl ? await api.ImageList_GetImageCount(himl) : 0;
				if (nCount == 0) {
					if (lpbmp == 206 || lpbmp == 204) {
						nCount = 20;
					} else if (lpbmp == 216 || lpbmp == 214) {
						nCount = 32;
					}
				}
				a[0] = await fso.GetFileName(dllpath);
				for (a[3] = 0; a[3] < nCount; a[3]++) {
					var s = "bitmap:" + a.join(",");
					var src = await MakeImgSrc(s, 0, false, a[2]);
					data.push('<img src="' + src + '" class="button" onclick="SelectIcon(this)" onmouseover="MouseOver(this)" onmouseout="MouseOut()" title="' + s + '"> ');
				}
				if (himl) {
					api.ImageList_Destroy(himl);
				}
				api.FreeLibrary(hModule);
			}
		} else {
			dllPath = await ExtractMacro(te, a[1]);
			if (!/^[A-Z]:\\|^\\\\/i.test(dllPath)) {
				dllPath = await fso.BuildPath(system32, a[1]);
			}
			var nCount = await api.ExtractIconEx(dllPath, -1, null, null, 0);
			for (var i = 0; i < nCount; ++i) {
				var s = ["icon:" + a[1], i].join(",");
				var src = await MakeImgSrc(s, 0, false, 32);
				data.push('<img src="' + src + '" class="button" onclick="SelectIcon(this)" onmouseover="MouseOver(this)" onmouseout="MouseOut()" title="' + s + '" style="max-height: 24pt"> ');
			}
		}
		o.innerHTML = data.join("");
		o.cursor = "";
		o.onclick = null;
		document.body.style.cursor = "auto";
	}, 1);
	document.body.style.cursor = "wait";
}

async function SearchIcon(o) {
	o.onclick = null;
	var s = [];
	var wfd = await api.Memory("WIN32_FIND_DATA");
	var hFind = api.FindFirstFile(fso.BuildPath(system32, "*"), wfd);
	for (var bFind = hFind != INVALID_HANDLE_VALUE; await bFind; bFind = await api.FindNextFile(hFind, wfd)) {
		var nCount = await api.ExtractIconEx(await fso.BuildPath(system32, await wfd.cFileName), -1, null, null, 0);
		if (nCount) {
			var id = "i," + await wfd.cFileName.toLowerCase();
			if (!document.getElementById(id)) {
				s.push('<div id="', id, '" onclick="OpenIcon(this)" style="cursor: pointer"><span class="tab">', await wfd.cFileName, ' : ', nCount, '</span></div>');
			}
		}
	}
	api.FindClose(hFind);
	o.innerHTML = s.join("");
}

InitDialog = async function () {
	document.title = TITLE;
	var Query = String(await dialogArguments.Query || location.search.replace(/\?/, "")).toLowerCase();
	var res = /^icon(.*)/.exec(Query);
	if (res) {
		var a =
		{
			"16px ieframe,216": "b,216,16",
			"24px ieframe,214": "b,214,24",
			"16px ieframe,206": "b,206,16",
			"24px ieframe,204": "b,204,24",
			"16px ieframe,699": "b,699,16",
			"24px ieframe,697": "b,697,24",

			"shell32": "i,shell32.dll",
			"imageres": "i,imageres.dll",
			"wmploc": "i,wmploc.dll",
			"setupapi": "i,setupapi.dll",
			"dsuiext": "i,dsuiext.dll",
			"inetcpl": "i,inetcpl.cpl",
		};
		var s = [];
		var wfd = await api.Memory("WIN32_FIND_DATA");
		var path = await fso.BuildPath(await te.Data.DataFolder, "icons");
		var hFind = await api.FindFirstFile(path + "\\*", wfd);
		for (var bFind = hFind != INVALID_HANDLE_VALUE; bFind; bFind = await api.FindNextFile(hFind, wfd)) {
			if ((await wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && /^[a-z]/i.test(await wfd.cFileName)) {
				var arfn = [];
				var path2 = await fso.BuildPath(path, await wfd.cFileName);
				var hFind2 = await api.FindFirstFile(path2 + "\\*.png", wfd);
				for (var bFind2 = hFind != INVALID_HANDLE_VALUE; bFind2; bFind2 = await api.FindNextFile(hFind2, wfd)) {
					arfn.push(await wfd.cFileName);
				}
				arfn.sort(async function (a, b) {
					return await api.StrCmpLogical(a, b);
				});
				for (var i = 0; i < arfn.length; ++i) {
					var src = ["icon:" + await fso.GetFileName(path2), arfn[i].replace(/\.png$/i, "")].join(",");
					s.push('<img src="', await fso.BuildPath(path2, arfn[i]), '" class="button" onclick="SelectIcon(this)" onmouseover="MouseOver(this)" onmouseout="MouseOut()" title="', src, '" style="max-height: 24pt"> ');
				}
				s.push("<br>");
			}
		}
		for (var i in a) {
			if (a[i].charAt(0) == "i" || res[1] != "2") {
				s.push('<div id="', a[i], '" onclick="OpenIcon(this)" style="cursor: pointer"><span class="tab">', i.replace(/ieframe,21\d/, await GetText("General")).replace(/ieframe,20\d/, await GetText("Browser")), '</span></div>');
			}
		}
		s.push('<div onclick="SearchIcon(this)" style="cursor: pointer"><span class="tab">', await GetText("Search"), '</span></div>');
		document.getElementById("Content").innerHTML = s.join("");
	}
	if (Query == "mouse") {
		document.getElementById("Content").innerHTML = '<div id="Gesture" style="width: 100%; height: 100%; text-align: center" onmousedown="return MouseDown()" onmouseup="return MouseUp()" onmousemove="return MouseMove()" ondblclick="MouseDbl()" onmousewheel="return MouseWheel()"></div>';
		document.getElementById("Selected").innerHTML = '<input type="text" name="q" style="width: 100%" autocomplete="off" onkeydown="setTimeout(\'returnValue=document.F.q.value\',100)">';
	}
	if (Query == "key") {
		returnValue = false;
		document.getElementById("Content").innerHTML = '<div style="padding: 8px;" style="display: block;"><label>Key</label><br><input type="text" name="q" autocomplete="off" style="width: 100%; ime-mode: disabled" onfocus="this.blur()"></div>';
		var fn = async function (e) {
			var key = (e || event).keyCode;
			var k = await api.MapVirtualKey(key, 0) | ((key >= 33 && key <= 46 || key >= 91 && key <= 93 || key == 111 || key == 144) ? 256 : 0);
			if (k == 42 || k == 29 || k == 56 || k == 347) {
				return false;
			}
			var s = await api.sprintf(10, "$%x", k | await GetKeyShift());
			returnValue = await GetKeyName(s);
			if (/^\$\w02a$|^\$\w01d$|^\$\w038$/i.test(returnValue)) {
				returnValue = await GetKeyName(s.slice(0, 3) + "1e").replace(/\+A$/, "");
			}
			document.F.q.value = returnValue;
			document.F.q.title = s;
			document.F.ButtonOk.disabled = false;
			return false;
		}
		AddEventEx(document.body, "keydown", fn);
		AddEventEx(document.body, "keyup", fn);
	}
	if (Query == "new") {
		returnValue = false;
		var s = [];
		s.push('<div style="padding: 8px;" style="display: block;"><label><input type="radio" name="mode" id="folder" onclick="document.F.path.focus()">New Folder</label> <label><input type="radio" name="mode" id="file" onclick="document.F.path.focus()">New File</label><br>', await dialogArguments.path, '<br><input type="text" name="path" style="width: 100%"></div>');
		document.getElementById("Content").innerHTML = s.join("");
		AddEventEx(document.body, "keydown", function (e) {
			setTimeout(function () {
				document.F.ButtonOk.disabled = !document.F.path.value;
			}, 99);
			var key = (e || event).keyCode;
			if (key == VK_RETURN && document.F.path.value) {
				SetResult(1);
			}
			if (key == VK_ESCAPE) {
				SetResult(2);
			}
			return true;
		});

		AddEventEx(document.body, "paste", function () {
			setTimeout(function () {
				document.F.ButtonOk.disabled = !document.F.path.value;
			}, 99);
		});

		setTimeout(function () {
			document.F[dialogArguments.Mode].checked = true;
			document.F.path.focus();
		}, 99);

		AddEventEx(window, "beforeunload", async function () {
			if (g_nResult == 1) {
				path = document.F.path.value;
				if (path) {
					if (!/^[A-Z]:\\|^\\/i.test(path)) {
						path = await fso.BuildPath(await dialogArguments.path, path.replace(/^\s+/, ""));
					}
					if (GetElement("folder").checked) {
						await CreateFolder(path);
					} else if (GetElement("file").checked) {
						await CreateFile(path);
					}
				}
			}
		});
	}
	if (Query == "fileicon" && await dialogArguments.element) {
		var s = await api.PathUnquoteSpaces(await dialogArguments.element.value);
		document.title = s + " - " + TITLE;
		GetElement("Content").innerHTML = '<div id="i,' + s + '" style="cursor: pointer"></div>';
		await OpenIcon(GetElement("i," + s));
	}
	if (Query == "about") {
		returnValue = false;
		var s = ['<table style="border-spacing: 2em; border-collapse: separate; width: 100%"><tr><td>'];
		var src = await MakeImgSrc(await api.GetModuleFileName(null), 0, true, 48);
		s.push('<img src="', src, '"></td><td><span style="font-weight: bold; font-size: 120%">', await AboutTE(2), '</span> (', await GetTextR((await api.sizeof("HANDLE") * 8) + '-bit'), ')<br>');
		s.push('<br><a href="javascript:Run(0)" class="link">', await api.GetModuleFileName(null), '</a> (', await fso.GetFileVersion(await api.GetModuleFileName(null)), ')<br>');
		s.push('<br><a href="javascript:Run(1)" class="link">', await fso.BuildPath(await te.Data.DataFolder, "config"), '</a><br>');
		s.push('<br><label>Information</label><input type="text" value="', await AboutTE(3), '" style="width: 100%" onclick="this.select()" readonly><br>');
		var root = await te.Data.Addons.documentElement;
		if (root) {
			var ar = [];
			var items = await root.childNodes;
			if (items) {
				for (var i = 0; i < await GetLength(items); ++i) {
					if (GetNum(await items[i].getAttribute("Enabled"))) {
						ar.push(await items[i].nodeName);
					}
				}
			}
			s.push('<br><label>Add-ons</label><input type="text" value="', ar.join(","), '" style="width: 100%" onclick="this.select()"><br>');
		}
		s.push('<br><input type="button" value="Visit website" onclick="Run(2)">');
		s.push('&nbsp;<input type="button" value="Check for updates" onclick="Run(3)">');
		s.push('</td></tr></table>');
		document.getElementById("Content").innerHTML = s.join("");
		document.F.ButtonOk.disabled = false;
		document.getElementById("buttonCancel").style.display = "none";

		Run = function (n) {
			var ar = [
				'Navigate(api.GetModuleFileName(null) + "\\\\..", SBSP_NEWBROWSER)',
				'Navigate(fso.BuildPath(te.Data.DataFolder, "config"), SBSP_NEWBROWSER)',
				'wsh.Run("https://tablacus.github.io/explorer_en.html")',
				'CheckUpdate()'
			];
			MainWindow.setTimeout(ar[n], 500);
			window.close();
		}
	}
	if (await dialogArguments.element) {
		AddEventEx(window, "beforeunload", async function () {
			try {
				if (g_nResult == 1) {
					dialogArguments.element.value = returnValue;
					var cb = await dialogArguments.element.onchange;
					if (cb) {
						await cb();
					}
				}
			} catch (e) { }
		});
	}
	DialogResize = function () {
		CalcElementHeight(document.getElementById("panel0"), 3);
	};
	AddEventEx(window, "resize", function () {
		clearTimeout(g_tidResize);
		g_tidResize = setTimeout(DialogResize, 500);
	});
	ApplyLang(document);
	DialogResize();
}

MouseDown = async function () {
	if (g_Gesture) {
		var n = 1;
		var ar = [0, 0, 2, 1];
		for (i = 1; i < 6; ++i) {
			if (g_Gesture.indexOf(i + "") < 0) {
				if ((event.buttons && event.buttons & n) || event.button == ar[i]) {
					returnValue += i + "";
				}
			}
			n *= 2;
		}
	} else {
		returnValue = await GetGestureKey() + await GetGestureButton();
		api.RedrawWindow(await GetBrowserWindow(document), null, 0, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN);
	}
	document.F.q.value = returnValue;
	g_Gesture = returnValue;
	g_pt.x = event.clientX;
	g_pt.y = event.clientY;
	document.F.ButtonOk.disabled = false;
	var o = document.getElementById("Gesture");
	var s = o.style.height;
	o.style.height = "1px";
	o.style.height = s;
	return false;
}

MouseUp = async function () {
	g_Gesture = await GetGestureButton();
	return false;
}

MouseMove = async function () {
	if (await api.GetKeyState(VK_XBUTTON1) < 0 || await api.GetKeyState(VK_XBUTTON2) < 0) {
		returnValue = await GetGestureKey() + await GetGestureButton();
		document.F.q.value = returnValue;
	}
	if (document.F.q.value.length && (await api.GetKeyState(VK_RBUTTON) < 0 || (await te.Data.Conf_Gestures && (await api.GetKeyState(VK_MBUTTON) < 0)))) {
		var pt = { x: event.clientX, y: event.clientY };
		var x = (pt.x - g_pt.x);
		var y = (pt.y - g_pt.y);
		if (Math.abs(x) + Math.abs(y) >= 20) {
			if (te.Data.Conf_TrailSize) {
				var hwnd = await GetBrowserWindow(document);
				var hdc = await api.GetWindowDC(hwnd);
				if (hdc) {
					await api.MoveToEx(hdc, g_pt.x, g_pt.y, null);
					var pen1 = await api.CreatePen(PS_SOLID, await te.Data.Conf_TrailSize, await te.Data.Conf_TrailColor);
					var hOld = await api.SelectObject(hdc, pen1);
					await api.LineTo(hdc, pt.x, pt.y);
					await api.SelectObject(hdc, hOld);
					api.DeleteObject(pen1);
					api.ReleaseDC(hwnd, hdc);
				}
			}
			g_pt = pt;
			var s = (Math.abs(x) >= Math.abs(y)) ? ((x < 0) ? "L" : "R") : ((y < 0) ? "U" : "D");
			if (s != document.F.q.value.charAt(document.F.q.value.length - 1)) {
				returnValue += s;
				document.F.q.value = returnValue;
			}
		}
	}
	return false;
}

MouseDbl = function () {
	returnValue += returnValue.replace(/\D/g, "");
	document.F.q.value = returnValue;
	return false;
}

MouseWheel = async function () {
	returnValue = await GetGestureKey() + await GetGestureButton() + (event.wheelDelta > 0 ? "8" : "9");
	document.F.q.value = returnValue;
	document.F.ButtonOk.disabled = false;
	return false;
}

InitLocation = async function () {
	var ar = await api.CreateObject("Array"), param = await api.CreateObject("Array");
	Addon_Id = await dialogArguments.Data.id;
	for (var i = 10; i--;) {
		var o = document.getElementById('tab' + i);
		o.className = "tab";
		o.hidefocus = true;
		o.style.display = "none";
		(function (o) {
			o.onmousedown = function () {
				ClickTab(o, 1);
			};
			o.onfocus = function () {
				o.blur()
			};
		})(o);
	}
	await LoadAddon("js", Addon_Id, await ar, param);
	if (await GetLength(ar)) {
		setTimeout(function (ar) {
			MessageBox(ar.join("\n\n"), TITLE, MB_OK);
		}, 500, ar);
	}
	ar = [];
	var s = "CSA";
	for (var i = 0; i < s.length; ++i) {
		ar.push('<input type="button" value="', await MainWindow.g_.KeyState[i][0], '" title="', s.charAt(i), '" onclick="AddMouse(this)">');
	}
	document.getElementById("__MOUSEDATA").innerHTML = ar.join("");
	ApplyLang(document);
	var info = await GetAddonInfo(await dialogArguments.Data.id);
	document.title = await info.Name;
	var items = await te.Data.Addons.getElementsByTagName(await dialogArguments.Data.id);
	var item = null;
	if (GetLength(items)) {
		item = await items[0];
		var Location = await item.getAttribute("Location");
		if (!Location) {
			Location = await param.Default;
		}
		for (var i = document.L.elements.length; i--;) {
			if (SameText(Location, document.L[i].value)) {
				document.L[i].checked = true;
			}
		}
	}
	var locs = {};
	items = await MainWindow.g_.Locations;
	for (var list = await api.CreateObject("Enum", items); !await list.atEnd(); await list.moveNext()) {
		var i = await list.item();
		locs[i] = [];
		var item1 = await items[i];
		for (var j = await GetLength(item1); j--;) {
			var ar = await item1[j].split("\0");
			locs[i].unshift(await GetImgTag({ src: ar[1], title: await GetAddonInfo(ar[0]).Name, "class": ar[1] ? "" : "text1" }, 16) + '<span style="font-size: 1px"> </span>');
		}
	}
	for (var i in locs) {
		var s = locs[i].join("");
		try {
			var o = document.getElementById('_' + i);
			ApplyLang(o);
			o.parentNode.title = o.innerHTML.replace(/<[^>]*>|[\r\n]|\s\s+/g, "");
			o.innerHTML = s;
		} catch (e) { }
	}
	var oa = document.F.Menu;
	oa.length = 0;
	var o = oa[++oa.length - 1];
	o.value = "";
	o.text = GetText("Select");
	for (var j in g_arMenuTypes) {
		var s = g_arMenuTypes[j];
		if (!/Default|Alias/.test(s)) {
			o = oa[++oa.length - 1];
			o.value = s;
			o.text = GetText(s);
		}
	}
	var ar = ["Key", "Mouse"];
	for (i in ar) {
		var mode = ar[i];
		var oa = document.F[mode + "On"];
		oa.length = 0;
		o = oa[++oa.length - 1];
		o.value = "";
		o.text = GetText("Select");
		for (var list = await api.CreateObject("Enum", MainWindow.eventTE[mode]); !await list.atEnd(); await list.moveNext()) {
			var j = list.item();
			o = oa[++oa.length - 1];
			o.text = await GetTextEx(j);
			o.value = j;
		}
	}
	if (item) {
		var ele = document.F.elements;
		for (var i = ele.length; i--;) {
			var n = ele[i].id || ele[i].name;
			if (n && !/=/.test(n)) {
				s = (/^!/.test(n) ? !await item.getAttribute(n.slice(1)) : await item.getAttribute(n)) || "";
				if (/Name$/.test(n)) {
					s = await GetText(s);
				}
				if (n == "Key") {
					s = await GetKeyNameG(s);
				}
				if (s || s === 0) {
					SetElementValue(ele[n], s);
				}
			}
		}
	}
	await LoadChecked(document.F);

	if (!await dialogArguments.Data.show) {
		dialogArguments.Data.show = "6";
		dialogArguments.Data.index = 6;
	}
	if (!/,/.test(await dialogArguments.Data.show)) {
		g_.NoTab = true;
		setTimeout(function () {
			document.getElementById("tabs").style.display = "none";
		}, 99);
	}
	if (/[8]/.test(await dialogArguments.Data.show)) {
		await MakeKeySelect();
		await SetKeyShift();
	}
	var a = document.F.MenuName.value.split(/\t/);
	document.F._MenuName.value = GetText(a[0]);

	try {
		var ar = await dialogArguments.Data.show.split(/,/);
		for (var i in ar) {
			document.getElementById("tab" + ar[i]).style.display = "inline";
		}
		nTabIndex = await dialogArguments.Data.index;
	} catch (e) { }

	await SetOnChangeHandler();
	AddEventEx(window, "load", function () {
		ApplyLang(document);
		ClickTab(null, 1);
	});
	AddEventEx(window, "resize", function () {
		clearTimeout(g_tidResize);
		g_tidResize = setTimeout(ResizeTabPanel, 500);
	});

	AddEventEx(window, "beforeunload", function () {
		SetOptions(async function () {
			if (window.SaveLocation) {
				SaveLocation();
			}
			MainWindow.g_.OptionsWindow = void 0;
			var items = await te.Data.Addons.getElementsByTagName(await dialogArguments.Data.id);
			if (GetLength(items)) {
				var item = await items[0];
				item.removeAttribute("Location");
				for (var i = document.L.elements.length; i--;) {
					if (document.L[i].checked) {
						item.setAttribute("Location", document.L[i].value);
						te.Data.bReload = true;
						await MainWindow.RunEvent1("ConfigChanged", "Addons");
						break;
					}
				}
				var ele = document.F.elements;
				ele.MenuName.value = await GetSourceText(ele._MenuName.value);
				if (dialogArguments.Data.show == "6") {
					ele.Set.value = "";
				}
				for (var i = ele.length; i--;) {
					var n = ele[i].id || ele[i].name;
					if (n && n.charAt(0) != "_") {
						if (n == "Key") {
							document.F[n].value = await GetKeyKeyG(document.F[n].value);
						}
						if (SetAttribEx(item, document.F, n)) {
							te.Data.bReload = true;
							MainWindow.RunEvent1("ConfigChanged", "Addons");
						}
					}
				}
			}
		});
	});
	if (item) {
		InitColor1(item);
	}
}

function SetAttrib(item, n, s) {
	if (/^!/.test(n)) {
		n = n.slice(1);
		s = !s;
	}
	if (s) {
		item.setAttribute(n, s);
	} else {
		item.removeAttribute(n);
	}
}

function GetElementValue(o) {
	if (o.type) {
		if (/checkbox/i.test(o.type)) {
			return o.checked ? 1 : 0;
		}
		if (/hidden|text|number|url|password|range|color|date|time/i.test(o.type)) {
			return o.value;
		}
		if (/select/i.test(o.type)) {
			return o[o.selectedIndex].value;
		}
	}
}

function SetElementValue(o, s) {
	if (o && o.type) {
		if (/checkbox/i.test(o.type)) {
			o.checked = GetNum(s);
			return;
		}
		if (/select/i.test(o.type)) {
			for (var i = o.options.length; i-- > 0;) {
				if (o.options[i].value == s) {
					o.selectedIndex = i;
					break;
				}
			}
			return;
		}
		o.value = s;
	}
}

async function SetAttribEx(item, f, n) {
	var s = GetElementValue(f[n]);
	if (s != await GetAttribEx(item, f, n)) {
		SetAttrib(item, n, s);
		return true;
	}
	return false;
}

async function GetAttribEx(item, f, n) {
	var s;
	var res = /([^=]*)=(.*)/.exec(n);
	if (res) {
		s = await item.getAttribute(res[1]);
		if (s == res[2]) {
			document.getElementById(n).checked = true;
		}
		return;
	}
	s = /^!/.test(n) ? !item.getAttribute(n.slice(1)) : item.getAttribute(n);
	if (s || s === 0) {
		if (n == "Key") {
			s = GetKeyName(s);
		}
		SetElementValue(f[n], s);
	}
}

function RefX(Id, bMultiLine, oButton, bFilesOnly, Filter, f) {
	setTimeout(async function () {
		var o = GetElement(Id, f);
		if (/Path/.test(Id)) {
			var s = Id.replace("Path", "Type");
			if (o) {
				var pt = await api.CreateObject("Object");
				if (oButton) {
					var pt1 = GetPos(oButton, 1);
					pt.x = await pt1.x;
					pt.y = await pt1.y + oButton.offsetHeight
					pt.width = oButton.offsetWidth;
				} else {
					await api.GetCursorPos(pt);
				}
				var oType = o.form[s];
				var r = MainWindow.OptionRef(oType[oType.selectedIndex].value, o.value, pt);
				if ("string" === typeof r) {
					var p = await api.CreateObject("Object");
					p.s = r;
					await MainWindow.OptionDecode(oType[oType.selectedIndex].value, p);
					if (bMultiLine && await api.GetKeyState(VK_CONTROL) < 0 && await api.ILCreateFromPath(await p.s)) {
						AddPath(Id, await p.s, f);
					} else {
						SetValue(o, await p.s);
					}
				}
				return;
			}
		}

		var path = o.value || o.getAttribute("placeholder") || "";
		var res = /^icon:([^,]*)|^bitmap:([^,]*)/i.exec(path) || [];
		path = await OpenDialogEx(res[1] || res[2] || path, Filter, GetNum(bFilesOnly));
		if (path) {
			if (bMultiLine) {
				AddPath(Id, path);
				return;
			}
			SetValue(o, path);
			if (/Icon|Large|Small/i.test(Id)) {
				var s = await api.PathUnquoteSpaces(ExtractMacro(te, path));
				if (await api.ExtractIconEx(s, -1, null, null, 0) > 1) {
					ShowDialogEx("fileicon", 640, 480, o);
					return;
				}
			}
		}
	}, 99);
}

async function PortableX(Id) {
	if (!await confirmOk()) {
		return;
	}
	var o = GetElement(Id);
	var s = await fso.GetDriveName(await api.GetModuleFileName(null));
	SetValue(o, o.value.replace(await wsh.ExpandEnvironmentStrings("%UserProfile%"), "%UserProfile%").replace(new RegExp('^("?)' + s, "igm"), "$1%Installed%").replace(new RegExp('( "?)' + s, "igm"), "$1%Installed%").replace(new RegExp('(:)' + s, "igm"), "$1%Installed%"));
}

function GetElement(Id, o) {
	return (o && o[Id]) || (document.F && document.F[Id]) || document.getElementById(Id) || (document.E && document.E[Id]);
}

function AddPath(Id, strValue, f) {
	var o = GetElement(Id, f);
	if (o) {
		var s = o.value;
		if (/\n$/.test(s) || s == "") {
			s += strValue;
		} else {
			s += "\n" + strValue;
		}
		SetValue(o, s);
	}
}

function SetValue(o, s) {
	if (o.value != s) {
		o.value = s;
		FireEvent(o, "change");
	}
}

async function GetCurrentSetting(s) {
	var FV = await te.Ctrl(CTRL_FV);

	if (await confirmOk()) {
		AddPath(s, await api.PathQuoteSpaces(await api.GetDisplayNameOf(await FV.FolderItem, SHGDN_FORPARSINGEX | SHGDN_FORPARSING)));
	}
}

function SetTab(s) {
	var o = null;
	var arg = String(s).split(/&/);
	for (var i in arg) {
		var ar = arg[i].split(/=/);
		var n = ar[0].toLowerCase();
		if (n == "tab") {
			if (SameText(ar[1], "Get Addons")) {
				o = document.getElementById('tab1_1');
			}
			if (SameText(ar[1], "Get Icons")) {
				o = document.getElementById('tab1_3');
			}
			var s = await GetText(ar[1]).toLowerCase();
			var ovTab;
			for (var j = 0; ovTab = document.getElementById('tab' + j); ++j) {
				if (s == ovTab.innerText.toLowerCase()) {
					o = ovTab;
					break;
				}
			}
		} else if (n == "menus") {
			g_MenuType = ar[1];
		} else if (n == "id") {
			g_Id = ar[1];
		}
	}
	ClickTree(o);
}

function AddMouse(o) {
	(document.F["MouseMouse"] || document.F["Mouse"]).value += o.title;
}

async function InitAddonOptions(bFlag) {
	returnValue = false;
	await LoadLang2(await fso.BuildPath(await fso.GetParentFolderName(await api.GetModuleFileName(null)), "addons\\" + Addon_Id + "\\lang\\" + await GetLangId() + ".xml"));

	ApplyLang(document);
	info = await GetAddonInfo(Addon_Id);
	document.title = await info.Name;
	SetOnChangeHandler();
	AddEventEx(window, "beforeunload", function () {
		SetOptions(SetAddonOptions);
	});
	var items = await te.Data.Addons.getElementsByTagName(Addon_Id);
	if (GetLength(items)) {
		InitColor1(await items[0]);
	}
}

function SetOnChangeHandler() {
	g_nResult = 3;
	g_bChanged = false;
	var ar = ["input", "select", "textarea"];
	for (var j in ar) {
		var o = document.getElementsByTagName(ar[j]);
		if (o) {
			for (var i = o.length; i--;) {
				if ((o[i].name || o[i].id) && o[i].name != "List" && !/^_/.test(o[i].id)) {
					AddEventEx(o[i], "change", function () {
						g_bChanged = true;
					});
				}
			}
		}
	}
}
async function SetAddonOptions() {
	if (!g_bClosed) {
		var items = await te.Data.Addons.getElementsByTagName(Addon_Id);
		if (GetLength(items)) {
			var item = await items[0];
			var ele = document.F.elements;
			for (var i = ele.length; i--;) {
				var n = ele[i].id || ele[i].name;
				if (n) {
					if (await SetAttribEx(item, document.F, n)) {
						returnValue = true;
					}
				}
			}
		}
		g_bClosed = true;
		if (returnValue) {
			TEOk();
		}
		window.close();
	}
}

function SelectIcon(o) {
	returnValue = o.title;
	document.F.ButtonOk.disabled = false;
	document.getElementById("Selected").innerHTML = o.outerHTML;
	DialogResize();
}

TestX = async function (id, f) {
	if (await confirmOk()) {
		if (!f) {
			f = document.F;
		}
		var o = f[id + "Type"];
		var p = await api.CreateObject("Object");
		p.s = f[id + "Path"].value;
		await MainWindow.OptionEncode(o[o.selectedIndex].value, p);
		await MainWindow.focus();
		await MainWindow.Exec(await te.Ctrl(CTRL_FV), await p.s, o[o.selectedIndex].value);
		focus();
	}
}

SetImage = async function (f, n) {
	var o = document.getElementById(n || "_Icon");
	if (o) {
		if (!f) {
			f = document.F;
		}
		var h = GetNum((f.IconSize || f.Height || { value: window.IconSize || 24 }).value);
		var src = await MakeImgSrc(api.PathUnquoteSpaces(f.Icon.value), 0, true, h);
		o.innerHTML = src ? '<img src="' + src + '" ' + (h ? 'height="' + h + 'px"' : "") + ' style="max-width: 32pt; max-height: 32pt">' : "";
	}
}

ShowIcon = ShowIconEx;

async function SelectLangID(o) {
	var i = 0;
	var Langs = [];
	var wfd = await api.Memory("WIN32_FIND_DATA");
	var hFind = await api.FindFirstFile(await fso.BuildPath(await fso.GetParentFolderName(await api.GetModuleFileName(null)), "lang\\*.xml"), wfd);
	for (var bFind = hFind != INVALID_HANDLE_VALUE; await bFind; bFind = await api.FindNextFile(hFind, wfd)) {
		Langs.push(await wfd.cFileName.replace(/\..*$/, ""));
	}
	api.FindClose(hFind);
	Langs.sort();
	var path = await fso.BuildPath(await fso.GetParentFolderName(await api.GetModuleFileName(null)), "lang\\");
	var hMenu = await api.CreatePopupMenu();
	for (i in Langs) {
		var xml = await api.CreateObject("Msxml2.DOMDocument");
		xml.async = false;
		var title = Langs[i];
		await xml.load(path + title + '.xml');
		var items = await xml.getElementsByTagName('lang');
		if (items && await GetLength(items)) {
			var item = await items[0];
			var en = await item.getAttribute("en");
			en = (en && SameText(await item.text, en)) ? "" : ' / ' + en;
			title = await item.text + en + " (" + title + ")\t" + await item.getAttribute("author");
		}
		api.InsertMenu(hMenu, i, MF_BYPOSITION | MF_STRING, GetNum(i) + 1, title);
	}
	var pt = GetPos(o, 1);
	var nVerb = await api.TrackPopupMenuEx(hMenu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y + o.offsetHeight, te.hwnd, null, null);
	if (nVerb) {
		document.F.Conf_Lang.value = Langs[nVerb - 1];
		g_bChanged = true;
	}
	api.DestroyMenu(hMenu);
}

async function GetTextEx(s) {
	var ar = s.split(/_/);
	var s = await GetText(ar.shift());
	if (ar && ar.length) {
		s += "(" + await GetText(ar.join(" ")) + ")";
	}
	return s;
}

async function GetAddons() {
	await MainWindow.OpenHttpRequest(urlAddons + "index.xml", "http", AddonsList);
}

async function GetIconPacks() {
	await MainWindow.OpenHttpRequest(urlIcons + "index.json", "http", IconPacksList);
}

function UpdateAddon(Id, o) {
	if (!o) {
		AddAddon(document.getElementById("Addons"), Id, "Disable");
		g_Chg.Addons = true;
	}
}

async function CheckAddon(Id) {
	return await fso.FileExists(await fso.BuildPath(await fso.GetParentFolderName(await api.GetModuleFileName(null)), "addons\\" + Id + "\\config.xml"));
}

function AddonsSearch() {
	if (xmlAddons) {
		AddonsAppend()
	} else {
		GetAddons();
	}
	return true;
}

function AddonsKeyDown() {
	if (event.keyCode == VK_RETURN) {
		AddonsSearch();
	}
	return true;
}

async function AddonsList(xhr2) {
	if (xmlAddons) {
		return;
	}
	if (xhr2) {
		xhr = xhr2;
	}
	var xml = await xhr.get_responseXML ? xhr.get_responseXML() : xhr.responseXML;
	if (xml) {
		xmlAddons = await xml.getElementsByTagName("Item");
		AddonsAppend()
	}
}

function SetTable(table, td) {
	if (table) {
		while (table.rows.length > 0) {
			table.deleteRow(0);
		}
		for (var i = 0; i < td.length; ++i) {
			var tr = table.insertRow(i);
			var td1 = tr.insertCell(0);
			td[i].shift();
			td1.innerHTML = td[i].join("");
			ApplyLang(td1);
			td1.className = (i & 1) ? "oddline" : "";
			td1.style.borderTop = "1px solid buttonshadow";
			td1.style.padding = "3px";
		}
	}
}

async function AddonsAppend() {
	var Progress = api.CreateObject("ProgressDialog");
	var i = 0, td = [];
	Progress.StartProgressDialog(te.hwnd, null, 2);
	try {
		Progress.SetAnimation(hShell32, 150);
		Progress.SetLine(1, api.LoadString(hShell32, 13585) || api.LoadString(hShell32, 6478), true);
		while (!await Progress.HasUserCancelled() && await xmlAddons[i]) {
			ArrangeAddon(xmlAddons[i++], td, Progress);
			var nLen = await GetLength(xmlAddons);
			Progress.SetTitle(Math.floor(100 * i / nLen) + "%");
			Progress.SetProgress(i, nLen);
		}
		if (g_nSort["1_1"] == 1) {
			td.sort(function (a, b) {
				return b[0] - a[0];
			});
		} else {
			td.sort();
		}
		SetTable(document.getElementById("Addons1"), td);
	} catch (e) {
		ShowError(e);
	}
	Progress.StopProgressDialog();
}

async function ArrangeAddon(xml, td, Progress) {
	var Id = await xml.getAttribute("Id");
	var s = [];
	var strUpdate = "";
	if (await Search(xml)) {
		var info = api.CreateObject("Object");
		for (var i = arLangs.length; i--;) {
			await GetAddonInfo2(xml, info, arLangs[i]);
		}
		var pubDate = "";
		var dt = new Date(await info.pubDate);
		Progress.SetLine(2, await info.Name, true);
		if (await info.pubDate) {
			pubDate = await api.GetDateFormat(LOCALE_USER_DEFAULT, 0, dt, await api.GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE)) + " ";
		}
		s.push('<table width="100%"><tr><td width="100%"><b style="font-size: 1.3em">', await info.Name, "</b>&nbsp;", await info.Version, "&nbsp;", await info.Creator, "<br>", await info.Description, "<br>");
		if (await info.Details) {
			s.push('<a href="#" title="', await info.Details, '" class="link" onclick="wsh.run(this.title); return false;">', await GetText('Details'), '</a><br>');
		}
		s.push(pubDate, '</td><td align="right">');
		var filename = await info.filename;
		if (!filename) {
			filename = Id + '_' + await info.Version.replace(/\D/, '') + '.zip';
		}
		var dt2 = (dt.getTime() / (24 * 60 * 60 * 1000)) - await info.Version;
		var bUpdate = false;
		if (CheckAddon(Id)) {
			var installed = await GetAddonInfo(Id);
			if (await installed.Version >= await info.Version) {
				try {
					if (!await installed.DllVersion) {
						return;
					}
					var path = await fso.BuildPath(await fso.GetParentFolderName(await api.GetModuleFileName(null)), "addons\\" + Id);
					var wfd = await api.Memory("WIN32_FIND_DATA");
					var hFind = await api.FindFirstFile(fso.BuildPath(path, "*" + (await api.sizeof("HANDLE") * 8) + ".dll"), wfd);
					api.FindClose(hFind);
					if (hFind == INVALID_HANDLE_VALUE) {
						return;
					}
					if (CalcVersion(await installed.DllVersion) <= CalcVersion(await fso.GetFileVersion(await fso.BuildPath(path, await wfd.cFileName)))) {
						return;
					}
				} catch (e) {
					return;
				}
			}
			strUpdate = '<br><b id="_Addons_' + Id + '"  class="danger" style="white-space: nowrap;">' + await GetText('Update available') + '</b>';
			dt2 += MAXINT * 2;
			bUpdate = true;
		} else {
			dt2 += MAXINT;
		}
		if (await info.MinVersion && await AboutTE(0) >= CalcVersion(await info.MinVersion)) {
			s.push('<input type="button" onclick="Install(this,', bUpdate, ')" title="', Id, '_', await info.Version, '" value="', await GetText("Install"), '">');
		} else {
			s.push('<input type="button"  class="danger" onclick="MainWindow.CheckUpdate()" value="', await info.MinVersion.replace(/^20/, (await api.LoadString(hShell32, 60) || "%").replace(/%.*/, "")).replace(/\.0/g, '.'), ' ', await GetText("is required."), '">');
		}
		s.push(strUpdate, '</td></tr></table>');
		s.unshift(g_nSort["1_1"] == 1 ? dt2 : g_nSort["1_1"] ? Id : await info.Name);
		td.push(s);
	}
}

async function Search(xml) {
	var q = document.getElementById('_GetAddonsQ');
	if (!q) {
		return true;
	}
	q = q.value.toUpperCase();
	if (q == "") {
		return true;
	}
	for (var k = arLangs.length; k-- > 0;) {
		var items = xml.getElementsByTagName(arLangs[k]);
		if (await GetLength(items)) {
			var item = await items[0].childNodes;
			for (var i = await GetLength(item); i-- > 0;) {
				var item1 = await item[i];
				if (await item1.tagName) {
					if ((await item1.textContent || await item1.text).toUpperCase().indexOf(q) >= 0) {
						return true;
					}
				}
			}
		}
	}
	return false;
}

async function Install(o, bUpdate) {
	if (!bUpdate && !await confirmOk("Do you want to install it now?")) {
		return;
	}
	var Id = o.title.replace(/_.*$/, "");
	await MainWindow.AddonDisabled(Id);
	if (await AddonBeforeRemove(Id) < 0) {
		return;
	}
	var file = o.title.replace(/\./, "") + '.zip';
	await MainWindow.OpenHttpRequest(urlAddons + Id + '/' + file, "http", Install2, o);
}

async function Install2(xhr, url, o) {
	var Id = o.title.replace(/_.*$/, "");
	var file = o.title.replace(/\./, "") + '.zip';
	var temp = await fso.BuildPath(await wsh.ExpandEnvironmentStrings("%TEMP%"), "tablacus");
	await CreateFolder(temp);
	var dest = await fso.BuildPath(temp, Id);
	await DeleteItem(dest);
	var hr = await MainWindow.Extract(fso.BuildPath(temp, file), temp, xhr);
	if (hr) {
		MessageBox([api.LoadString(hShell32, 4228).replace(/^\t/, "").replace("%d", await api.sprintf(99, "0x%08x", hr)), await GetText("Extract"), file].join("\n\n"), TITLE, MB_OK | MB_ICONSTOP);
		return;
	}
	var configxml = dest + "\\config.xml";
	var nDog = 300;
	while (!await fso.FileExists(configxml)) {
		if (await wsh.Popup(await GetText("Please wait."), 1, TITLE, MB_ICONINFORMATION | MB_OKCANCEL) == IDCANCEL || nDog-- == 0) {
			return;
		}
	}
	await api.SHFileOperation(FO_MOVE, dest, await fso.BuildPath(await fso.GetParentFolderName(await api.GetModuleFileName(null)), "addons"), FOF_NOCONFIRMATION | FOF_NOCONFIRMMKDIR, false);
	o.disabled = true;
	o.value = await GetText("Installed");
	o = document.getElementById('_Addons_' + Id);
	if (o) {
		o.style.display = "none";
	}
	UpdateAddon(Id, o);
}

function IconsKeyDown() {
	if (event.keyCode == VK_RETURN) {
		IconPacksList();
	}
	return true;
}

async function InstallIcon(o) {
	if (!await confirmOk("Do you want to install it now?")) {
		return;
	}
	var Id = o.title.replace(/_[^_]*$/, "");
	await MainWindow.OpenHttpRequest(urlIcons + Id + '/' + o.title.replace(/\./g, "") + '.zip', "http", InstallIcon2, o);
}

async function InstallIcon2(xhr, url, o) {
	var file = o.title.replace(/\./, "") + '.zip';
	var temp = await fso.BuildPath(await wsh.ExpandEnvironmentStrings("%TEMP%"), "tablacus");
	await CreateFolder(temp);
	var dest = await fso.BuildPath(await te.Data.DataFolder, "icons");
	await CreateFolder(dest);
	var hr = await MainWindow.Extract(await fso.BuildPath(temp, file), dest, xhr);
	if (hr) {
		MessageBox([api.LoadString(hShell32, 4228).replace(/^\t/, "").replace("%d", await api.sprintf(99, "0x%08x", hr)), await GetText("Extract"), file].join("\n\n"), TITLE, MB_OK | MB_ICONSTOP);
		return;
	}
	IconPacksList();
}

async function IconPacksList1(s, Id, info, json) {
	var q = document.getElementById('_GetIconsQ').value;
	if (q && !json) {
		if (!await api.PathMatchSpec(JSON.stringify(info), "*" + q + "*")) {
			return false;
		}
	}
	var langId = await GetLangId();
	s.push('<img src="', urlIcons, Id, '/preview.png" align="left" style="margin-right: 8px"><b style="font-size: 1.3em">', info.name[langId] || info.name.en, '</b> ');
	s.push(info.version, " ");
	if (info.URL) {
		s.push('<a href="#" class="link" onclick="wsh.run(\'', info.URL, '\'); return false;">');
	}
	s.push(info.creator[langId] || info.creator.en);
	if (info.URL) {
		s.push('</a>');
	}
	s.push("<br>", info.descprition[langId] || info.descprition.en, "<br>");
	if (json) {
		s.push(GetText("Installed"));
		s.push('<input type="button" onclick="DeleteIconPacks()" value="Delete" style="float: right;">');
		if (json[Id] && Number(json[Id].info.version) > Number(info.version)) {
			s.push('<hr><b class="danger" style="white-space: nowrap;">', GetText('Update available'), '</b> ', json[Id].info.version);
			info = json[Id].info;
			json = false;
		}
	}
	if (!json) {
		s.push('<input type="button" onclick="InstallIcon(this)" title="', Id, '_', info.version, '" value="', GetText("Install"), '" style="float: right;">');
	}
	s.push("<br>", await api.GetDateFormat(LOCALE_USER_DEFAULT, 0, new Date(info.pubDate), api.GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE)));
	return true;
}

async function IconPacksList(xhr) {
	if (xhr) {
		g_xhrIcons = xhr;
	} else {
		xhr = window.g_xhrIcons;
	}
	if (!xhr) {
		return;
	}
	var s;
	var ado = await OpenAdodbFromTextFile(fso.BuildPath(te.Data.DataFolder, "icons\\config.json"), "utf-8");
	if (ado) {
		s = ado.ReadText();
		ado.Close();
	}
	var json1 = JSON.parse(s || '{}');
	var text = xhr.get_responseText ? xhr.get_responseText() : xhr.responseText;
	var json = JSON.parse(text);
	var td = [];
	var Installed = "";
	if (json1.info) {
		Installed = json1.info.id || json1.info.name.en.replace(/\W/g, "_");
	}
	for (var n in json) {
		if (n != Installed) {
			var s1;
			var s = [];
			info = json[n].info;
			if (IconPacksList1(s, n, info)) {
				if (g_nSort["1_3"] == 0) {
					s1 = info.name[await GetLangId()] || info.name.en;
				} else if (g_nSort["1_3"] == 1) {
					if (json.pubDate) {
						s1 = await api.GetDateFormat(LOCALE_USER_DEFAULT, 0, new Date(info.pubDate), "yyyyMMdd");
					}
				} else {
					s1 = n;
				}
				td.push([s1 + "\t" + n, s.join("")]);
			}
		}
	}
	if (g_nSort["1_3"] == 1) {
		td.sort(function (a, b) {
			return b[0] - a[0];
		});
	} else {
		td.sort();
	}
	if (json1.info) {
		s = [];
		if (IconPacksList1(s, Installed, json1.info, json)) {
			td.unshift(["", s.join("")]);
		}
	}
	SetTable(document.getElementById("IconPacks1"), td);
}

async function DeleteIconPacks() {
	if (await confirmOk()) {
		DeleteItem(await fso.BuildPath(await te.Data.DataFolder, "icons"));
		IconPacksList();
	}
}

async function EnableSelectTag(o) {
	if (o) {
		var hwnd = await GetBrowserWindow(document);
		api.SendMessage(hwnd, WM_SETREDRAW, false, 0);
		o.style.visibility = "hidden";
		setTimeout(async function () {
			o.style.visibility = "visible";
			await api.SendMessage(hwnd, WM_SETREDRAW, true, 0);
			api.RedrawWindow(hwnd, null, 0, RDW_NOERASE | RDW_INVALIDATE | RDW_ALLCHILDREN);
		}, 99);
	}
}

SetResult = function (i) {
	g_nResult = i;
	window.close();
}

function InitColor1(item) {
	var ele = document.F.elements;
	for (var i = ele.length; i--;) {
		var n = ele[i].id || ele[i].name;
		if (n) {
			GetAttribEx(item, document.F, n);
		}
	}
	for (var i = ele.length; i--;) {
		var n = ele[i].id || ele[i].name;
		if (n) {
			var res = /^Color_(.*)/.exec(n);
			if (res) {
				var o = document.F[res[1]];
				if (o) {
					ele[i].style.backgroundColor = GetWebColor(o.value || o.placeholder);
				}
			}
		}
	}
}

async function ChangeColor1(ele) {
	var o = document.getElementById("Color_" + (ele.id || ele.name));
	if (o) {
		o.style.backgroundColor = await GetWebColor(ele.value || ele.placeholder);
	}
}

function EnableInner() {
	ChangeForm([["__Inner", "disabled", false]]);
}

function ChangeForm(ar) {
	AddEventEx(window, "load", function () {
		for (var i in ar) {
			var o = document.getElementById(ar[i][0]);
			for (var s = ar[i][1].split("/"); s.length > 1;) {
				o = o[s.shift()];
			}
			o[s[0]] = ar[i][2];
		}
	});
}

async function SetTabContents(id, name, value) {
	var oPanel = document.getElementById("panel" + id);
	if (name) {
		document.getElementById("tab" + id).innerHTML = await GetText(name);
	}
	oPanel.innerHTML = value.join ? value.join('') : value;
}

function ShowButtons(b1, b2, SortMode) {
	if (SortMode) {
		g_SortMode = SortMode;
	}
	var o = document.getElementById("SortButton");
	o.style.display = b1 ? "inline-block" : "none";
	if (g_SortMode == 1) {
		var table = document.getElementById("Addons");
		var bSorted = /none/i.test(table.style.display);
		document.getElementById("MoveButton").style.display = (b1 || b2) && !bSorted ? "inline-block" : "none";
		for (var i = 3; i--;) {
			o = document.getElementById("SortButton_" + i);
			o.style.border = bSorted && g_nSort[1] == i ? "1px solid highlight" : "";
			o.style.padding = bSorted && g_nSort[1] == i ? "0" : "";
		}
	} else {
		document.getElementById("MoveButton").style.display = b2 ? "inline-block" : "none";
		for (var i = 3; i--;) {
			o = document.getElementById("SortButton_" + i);
			o.style.border = g_nSort[g_SortMode] == i ? "1px solid highlight" : "";
			o.style.padding = g_nSort[g_SortMode] == i ? "0" : "";
		}
	}
}

async function SortAddons(n) {
	if (g_SortMode == 1) {
		var table = document.getElementById("Addons");
		if (table.rows.length < 2) {
			return;
		}
		var sorted = document.getElementById("SortedAddons");
		while (sorted.rows.length > 0) {
			sorted.deleteRow(0);
		}
		if (/none/i.test(table.style.display) && g_nSort[1] == n) {
			table.style.display = "";
			sorted.style.display = "none";
		} else {
			g_nSort[1] = n;
			var s, ar = [];
			for (var j = table.rows.length; j--;) {
				var div = table.rows(j).cells(0).firstChild;
				var Id = div.id.replace("Addons_", "").toLowerCase();
				if (g_nSort[1] == 0) {
					s = table.rows(j).cells(0).innerText;
				} else if (g_nSort[1] == 1) {
					s = "";
					var info = await GetAddonInfo(Id);
					var pubDate = await info.pubDate;
					if (pubDate) {
						s = await api.GetDateFormat(LOCALE_USER_DEFAULT, 0, new Date(pubDate), "yyyyMMdd");
					}
				} else {
					s = Id;
				}
				ar.push(s + "\t" + Id);
			}
			ar.sort();
			if (g_nSort[1] == 1) {
				ar = ar.reverse();
			}
			var i = 0, td = [];
			var Progress = await api.CreateObject("ProgressDialog");
			Progress.SetAnimation(hShell32, 150);
			Progress.StartProgressDialog(te.hwnd, null, 2);
			try {
				for (var i in ar) {
					if (Progress.HasUserCancelled()) {
						break;
					}
					Progress.SetTitle(Math.floor(100 * i / ar.length) + "%");
					Progress.SetProgress(i, ar.length);
					var data = ar[i].split("\t");
					var Id = data[data.length - 1];
					Progress.SetLine(1, Id, true);
					AddAddon(sorted, Id, document.getElementById("enable_" + Id).checked, "Sorted_");
				}
			} catch (e) {
				ShowError(e);
			}
			Progress.StopProgressDialog();
			var bCancelled = await Progress.HasUserCancelled();
			table.style.display = bCancelled ? "" : "none";
			sorted.style.display = bCancelled ? "none" : "";
		}
	} else if (g_SortMode == "1_1") {
		if (n != g_nSort["1_1"]) {
			g_nSort["1_1"] = n;
			AddonsSearch();
		}
	} else if (g_SortMode == "1_3") {
		if (n != g_nSort["1_3"]) {
			g_nSort["1_3"] = n;
			IconPacksList();
		}
	}
	ShowButtons(true);
}
