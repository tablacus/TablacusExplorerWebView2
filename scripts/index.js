// Tablacus Explorer

async function Resize2() {
	if (ui_.tidResize) {
		clearTimeout(ui_.tidResize);
		ui_.tidResize = void 0;
	}
	ResetScroll();
	var o = document.getElementById("toolbar");
	var offsetTop = o ? o.offsetHeight : 0;

	var h = 0;
	o = document.getElementById("bottombar");
	var offsetBottom = o.offsetHeight;
	o = document.getElementById("client");
	var ode = document.documentElement || document.body;
	if (o) {
		h = ode.offsetHeight - offsetBottom - offsetTop;
		if (h < 0) {
			h = 0;
		}
		o.style.height = h + "px";
	}
	await ResizeSizeBar("Left", h);
	await ResizeSizeBar("Right", h);
	o = document.getElementById("Background");
	pt = GetPos(o);
	te.offsetLeft = pt.x;
	te.offsetRight = ode.offsetWidth - o.offsetWidth - te.offsetLeft;
	te.offsetTop = pt.y;
	pt = GetPos(document.getElementById("bottombar"));
	te.offsetBottom = ode.offsetHeight - pt.y;
	RunEvent1("Resize");
	api.PostMessage(await te.hwnd, WM_SIZE, 0, 0);
}

async function ResizeSizeBar(z, h) {
	var o = await g_.Locations;
	var w = (await o[z + "Bar1"] || await o[z + "Bar2"] || await o[z + "Bar3"]) ? await te.Data["Conf_" + z + "BarWidth"] : 0;
	o = document.getElementById(z.toLowerCase() + "bar");
	if (w > 0) {
		o.style.display = "";
		if (w != o.offsetWidth) {
			o.style.width = w + "px";
			for (var i = 1; i <= 3; ++i) {
				document.getElementById(z + "Bar" + i).style.width = w + "px";
			}
			document.getElementById(z.toLowerCase() + "barT").style.width = w + "px";
		}
	} else {
		o.style.display = "none";
	}
	document.getElementById(z.toLowerCase() + "splitter").style.display = w ? "" : "none";

	o = document.getElementById(z.toLowerCase() + "barT");
	var th = Math.round(Math.max(h, 0));
	o.style.height = th + "px";

	var h2 = o.clientHeight - document.getElementById(z + "Bar1").offsetHeight - document.getElementById(z + "Bar3").offsetHeight;
	document.getElementById(z + "Bar2").style.height = Math.abs(h2 - o.clientHeight - th) + "px";
}

function ResetScroll() {
	if (document.documentElement && document.documentElement.scrollLeft) {
		document.documentElement.scrollLeft = 0;
	}
}

function PanelCreated(Ctrl) {
	UI.RunEvent1("PanelCreated", Ctrl);
}

GetAddonLocation = async function (strName) {
	var items = await te.Data.Addons.getElementsByTagName(strName);
	return (await GetLength(items) ? await items[0].getAttribute("Location") : null);
}

SetAddon = async function (strName, Location, Tag, strVAlign) {
	if (strName) {
		var s = await GetAddonLocation(strName);
		if (s) {
			Location = s;
		}
	}
	if (Tag) {
		if (Tag.join) {
			Tag = Tag.join("");
		}
		var o = document.getElementById(Location);
		if (o) {
			if ("string" === typeof Tag) {
				o.insertAdjacentHTML("BeforeEnd", Tag);
			} else {
				o.appendChild(Tag);
			}
			o.style.display = (ui_.IEVer >= 8 && SameText(o.tagName, "td")) ? "table-cell" : "block";
			if (strVAlign && !o.style.verticalAlign) {
				o.style.verticalAlign = strVAlign;
			}
		} else if (Location == "Inner") {
			AddEvent("PanelCreated", function (Ctrl) {
				SetAddon(null, "Inner1Left_" + Ctrl.Id, Tag.replace(/\$/g, Ctrl.Id));
			});
		}
		if (strName) {
			if (!await g_.Locations[Location]) {
				g_.Locations[Location] = await api.CreateObject("Array");
			}
			var res = /<img.*?src=["'](.*?)["']/i.exec(String(Tag));
			if (res) {
				strName += "\0" + res[1];
			}
			await g_.Locations[Location].push(strName);
		}
	}
	return Location;
}

UI.Resize = function () {
	if (!ui_.tidResize) {
		clearTimeout(ui_.tidResize);
	}
	ui_.tidResize = setTimeout(Resize2, 500);
}

Resize = UI.Resize;

UI.OpenInExplorer = function (Path) {
	setTimeout(async function (Path) {
		await $.OpenInExplorer(Path);
	}, 99, Path);
}

UI.StartGestureTimer = async function () {
	var i = await te.Data.Conf_GestureTimeout;
	if (i) {
		clearTimeout(await g_.mouse.tidGesture);
		g_.mouse.tidGesture = setTimeout(function () {
			g_.mouse.EndGesture(true);
		}, i);
	}
}

UI.Focus = function (o, tm) {
	if (o) {
		setTimeout(function () {
			o.focus();
		}, tm)
	}
}

UI.FocusFV = function (tm) {
	setTimeout(async function () {
		var hFocus = await api.GetFocus();
		if (!hFocus || hFocus == await te.hwnd) {
			var FV = te.Ctrl(CTRL_FV);
			if (await FV) {
				FV.Focus();
			}
		}
	}, tm || ui_.DoubleClickTime);
}

UI.FocusWB = function () {
	var o = document.activeElement;
	if (o) {
		if (!/input|textarea/i.test(o.tagName)) {
			setTimeout(function () {
				if (o === document.activeElement) {
					GetFolderView().Focus();
				}
			}, ui_.DoubleClickTime, o);
		}
	}
}

UI.SelectItem = function (FV, path, wFlags, tm) {
	setTimeout(function () {
		FV.SelectItem(path, wFlags);
	}, tm);
}

UI.SelectNewItem = function () {
	setTimeout(async function () {
		var FV = te.Ctrl(CTRL_FV);
		if (FV) {
			if (!await api.StrCmpI(await FV.FolderItem.Path, await fso.GetParentFolderName(await g_.NewItemPath))) {
				FV.SelectItem(await g_.NewItemPath, SVSI_FOCUSED | SVSI_ENSUREVISIBLE | SVSI_DESELECTOTHERS | SVSI_SELECTIONMARK | SVSI_SELECT);
			}
		}
	}, 800);
}

UI.SelectNext = function (FV) {
	setTimeout(async function () {
		if (!await api.SendMessage(await FV.hwndList, LVM_GETEDITCONTROL, 0, 0) || WINVER < 0x600) {
			FV.SelectItem(FV.Item(await FV.GetFocusedItem + (await api.GetKeyState(VK_SHIFT) < 0 ? -1 : 1)) || await api.GetKeyState(VK_SHIFT) < 0 ? FV.ItemCount(SVGIO_ALLVIEW) - 1 : 0, SVSI_EDIT | SVSI_FOCUSED | SVSI_SELECT | SVSI_DESELECTOTHERS);
		}
	}, 99);
}

UI.EndMenu = function () {
	setTimeout(async function () {
		await api.EndMenu();
	}, 200);
}

UI.ShowStatusText = async function (Ctrl, Text, iPart, tm) {
	if (ui_.Status && ui_.Status[5]) {
		clearTimeout(ui_.Status[5]);
		delete ui_.Status;
	}
	ui_.Status = [Ctrl, Text, iPart, tm, new Date().getTime(), setTimeout(function () {
		if (ui_.Status) {
			if (new Date().getTime() - ui_.Status[4] > ui_.Status[3] / 2) {
				$.ShowStatusText(ui_.Status[0], ui_.Status[1], ui_.Status[2]);
				delete ui_.Status;
			}
		}
	}, tm)];
}

UI.CancelWindowRegistered = function () {
	clearTimeout(ui_.tidWindowRegistered);
	ui_.bWindowRegistered = false;
	ui_.tidWindowRegistered = setTimeout(function () {
		ui_.bWindowRegistered = true;
	}, 9999);
}

UI.ExecGesture = function (Ctrl, hwnd, pt, str) {
	setTimeout(function () {
		g_.mouse.Exec(Ctrl, hwnd, pt, str);
	}, 99);
}

UI.InitWindow = function (cb, cb2) {
	setTimeout(async function (cb) {
		Resize();
		(await cb)(cb);
		setTimeout(async function (cb2) {
			(await cb2)();
		}, 500, cb2);
	}, 99, cb);
}

UI.ExitFullscreen = function () {
	if (document.msExitFullscreen) {
		document.msExitFullscreen();
	}
}

te.OnArrange = async function (Ctrl, rc, cb) {
	var Type = await Ctrl.Type;
	if (Type == CTRL_TE) {
		ui_.TCPos = {};
	}
	await RunEvent1("Arrange", Ctrl, rc, cb);
	if (Type == CTRL_TC) {
		var Id = await Ctrl.Id;
		var o = ui_.Panels[Id];
		if (!o) {
			o = document.createElement("table");
			o.id = "Panel_" + Id;
			o.className = "layout";
			o.style.position = "absolute";
			o.style.zIndex = 1;
			var s = '<tr><td id="InnerLeft_$" class="sidebar" style="width: 0; display: none; overflow: auto"></td><td style="width: 100%"><div id="InnerTop_$" style="display: none"></div>';
			s += '<table id="InnerTop2_$" class="layout"><tr><td id="Inner1Left_$" class="toolbar1"></td><td id="Inner1Center_$" class="toolbar2" style="white-space: nowrap"></td><td id="Inner1Right_$" class="toolbar3"></td></tr></table>';
			s += '<table id="InnerView_$" class="layout" style="width: 100%"><tr><td id="Inner2Left_$" style="width: 0"></td><td id="Inner2Center_$"></td><td id="Inner2Right_$" style="width: 0; overflow: auto"></td></tr></table>';
			s += '<div id="InnerBottom_$"></div></td><td id="InnerRight_$" class="sidebar" style="width: 0; display: none"></td></tr>';
			o.innerHTML = s.replace(/\$/g, Id);
			document.getElementById("Panel").appendChild(o);
			PanelCreated(await Ctrl);
			ui_.Panels[Id] = o;
			ApplyLang(o);
			ChangeView(await Ctrl.Selected);
		}
		o.style.left = await rc.left + "px";
		o.style.top = await rc.top + "px";
		if (await Ctrl.Visible) {
			var s = [await Ctrl.Left, await Ctrl.Top, await Ctrl.Width, await Ctrl.Height].join(",");
			if (ui_.TCPos[s] && ui_.TCPos[s] != Id) {
				Ctrl.Close();
				return;
			} else {
				ui_.TCPos[s] = Id;
			}
			o.style.display = (ui_.IEVer >= 8 && SameText(o.tagName, "td")) ? "table-cell" : "block";
		} else {
			o.style.display = "none";
		}
		o.style.width = Math.max(await rc.right - await rc.left, 0) + "px";
		o.style.height = Math.max(await rc.bottom - await rc.top, 0) + "px";
		rc.top = await rc.top + document.getElementById("InnerTop_" + Id).offsetHeight + document.getElementById("InnerTop2_" + Id).offsetHeight;
		var w1 = 0, w2 = 0, x = '';
		for (var i = 0; i <= 1; ++i) {
			w1 += Number(document.getElementById("Inner" + x + "Left_" + Id).style.width.replace(/\D/g, "")) || 0;
			w2 += Number(document.getElementById("Inner" + x + "Right_" + Id).style.width.replace(/\D/g, "")) || 0;
			x = '2';
		}
		rc.left = await rc.left + w1;
		rc.right = await rc.right - w2;
		rc.bottom = await rc.bottom - document.getElementById("InnerBottom_" + Id).offsetHeight;
		o = document.getElementById("Inner2Center_" + Id).style;
		o.width = Math.max(await rc.right - await rc.left, 0) + "px";
		o.height = Math.max(await rc.bottom - await rc.top, 0) + "px";
		(await cb)(await Ctrl, await rc);
	}
}

g_.event.windowregistered = function (Ctrl) {
	if (ui_.bWindowRegistered) {
		RunEvent1("WindowRegistered", Ctrl);
	}
	ui_.bWindowRegistered = true;
}

ArrangeAddons = async function () {
	g_.Locations = await api.CreateObject("Object");
	$.IconSize = await te.Data.Conf_IconSize || screen.logicalYDPI / 4;
	var xml = await OpenXml("addons.xml", false, true);
	te.Data.Addons = xml;
	if (await api.GetKeyState(VK_SHIFT) < 0 && await api.GetKeyState(VK_CONTROL) < 0) {
		IsSavePath = function (path) {
			return false;
		}
		return;
	}
	var AddonId = [];
	var root = await xml.documentElement;
	if (root) {
		var items = await root.childNodes;
		if (items) {
			var arError = await api.CreateObject("Array");
			for (var i = 0; i < await GetLength(items); ++i) {
				var item = await items[i];
				var Id = await item.nodeName;
				g_.Error_source = Id;
				if (!AddonId[Id]) {
					var Enabled = GetNum(await item.getAttribute("Enabled"));
					if (Enabled & 6) {
						LoadLang2(await fso.BuildPath(await te.Data.Installed, "addons\\" + Id + "\\lang\\" + await GetLangId() + ".xml"));
					}
					if (Enabled & 8) {
						LoadAddon("vbs", Id, arError);
					}
					if (Enabled & 1) {
						LoadAddon("js", Id, arError);
					}
					AddonId[Id] = true;
				}
				g_.Error_source = "";
			}
			if (await arError.length || await arError.Count) {
				setTimeout(async function (arError) {
					if (await MessageBox(await arError.join("\n\n"), TITLE, MB_OKCANCEL) != IDCANCEL) {
						te.Data.bErrorAddons = true;
						ShowOptions("Tab=Add-ons");
					}
				}, 500, arError);
			}
		}
	}
	UI.RunEvent1("BrowserCreated", document);
	var cl = GetWinColor(window.getComputedStyle ? getComputedStyle(document.body).getPropertyValue('background-color') : document.body.currentStyle.backgroundColor);
	ArrangeAddons1(cl);
}

// Events

AddEvent("VisibleChanged", async function (Ctrl) {
	if (await Ctrl.Type == CTRL_TC) {
		var o = ui_.Panels[Ctrl.Id];
		if (o) {
			if (await Ctrl.Visible) {
				o.style.display = (ui_.IEVer >= 8 && SameText(o.tagName, "td")) ? "table-cell" : "block";
				ChangeView(await Ctrl.Selected);
			} else {
				o.style.display = "none";
			}
		}
	}
});

AddEvent("SystemMessage", async function (Ctrl, hwnd, msg, wParam, lParam) {
	if (await Ctrl.Type == CTRL_WB) {
		if (msg == WM_KILLFOCUS) {
			var o = document.activeElement;
			if (o) {
				var s = o.style.visibility;
				o.style.visibility = "hidden";
				o.style.visibility = s;
				FireEvent(o, "blur");
			}
		}
	}
});

// Browser Events

AddEventEx(window, "load", async function () {
	ApplyLang(document);
	if (await api.GetKeyState(VK_SHIFT) < 0 && await api.GetKeyState(VK_CONTROL) < 0) {
		ShowOptions("Tab=Add-ons");
	}
});

AddEventEx(window, "resize", Resize);

AddEventEx(window, "beforeunload", Finalize);

AddEventEx(window, "blur", ResetScroll);

AddEventEx(document, "MSFullscreenChange", function () {
	FullscreenChanged(document.msFullscreenElement != void 0);
});
	
(async function () {
	UI.OnLoad();
	await InitCode();
	DefaultFont = await $.DefaultFont;
	HOME_PATH = await $.HOME_PATH;
	await InitMouse();
	OpenMode = await $.OpenMode;
	await InitMenus();
	await LoadLang();
	ArrangeAddons();
	setTimeout(await InitWindow, 9);
})();
