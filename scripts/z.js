function ReadFile(fn) {
	var ado = new ActiveXObject("ADODB.Stream");
	ado.CharSet = "utf-8";
	ado.Open();
	ado.LoadFromFile(fn);
	var s = ado.ReadText();
	ado.Close();
	return s;
}

function WriteFile(fn, s) {
	if (s == ReadFile(fn)) {
		return false;
	}
	var ado = new ActiveXObject("ADODB.Stream");
	ado.CharSet = "utf-8";
	ado.Open();
	ado.WriteText(s);
	ado.SaveToFile(fn, 2);
	ado.Close();
	//Delete BOM
	ado = new ActiveXObject("ADODB.Stream");
	ado.Type = 1;
	ado.Open();
	ado.LoadFromFile(fn);
	ado.Position = 3;
	s = ado.Read(-1);
	ado.Close();
	ado = new ActiveXObject("ADODB.Stream");
	ado.Type = 1;
	ado.Open();
	ado.Write(s);
	ado.SaveToFile(fn, 2);
	ado.Close();
	return true;
}
var ar = ["consts.js", "index.js", "ui.js", "common.js", "options.js"];
ar.push("index.html", "options.html", "dialog.html");
for (var i in ar) {
	var fn = ar[i];
	var src = ReadFile(fn);
	if (/\.js$/.test(fn)) {
		src = src.replace(/([^\.\w])(async |await |debugger;)/g, "$1");
	}
	if (WriteFile('..\\script\\' + fn, src)) {
		WScript.Echo(fn);
	}
}
