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
for (;;) {
	var ar = ["index.js", "ui.js", "common.js", "options.js"];
	for (var i in ar) {
		var fn = ar[i];
		//WScript.Echo(fn);
		var src = ReadFile('..\\script\\' + fn);
		src = src.replace(/([^\.\w])(async |await |debugger;)/g, "$1");
		if (WriteFile(fn, src)) {
			WScript.Echo(new Date().toLocaleString());
			WScript.Echo(fn);
		}
	}
/*	var ar= ["consts.js", "index.html", "options.html", "dialog.html", "location.html", "index.css", "options.css"];
	for (var i in ar) {
		var fn = ar[i];
		//WScript.Echo(fn);
		var src = ReadFile(fn);
		if (WriteFile('..\\script\\' + fn, src)) {
			WScript.Echo(new Date().toLocaleString());
			WScript.Echo(fn);
		}
	}*/
	WScript.Sleep(3000);
}
