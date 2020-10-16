// Tablacus Explorer

if (window.UI) {
	BlurId = function (Id) {
		api.Invoke(UI.BlurId, Id);
	}

	clearTimeout = function (tid) {
		api.Invoke(UI.clearTimeout, tid);
	}

	OpenHttpRequest = function () {
		api.Invoke(UI.OpenHttpRequest, arguments);
	}

	ReloadCustomize = function () {
		api.Invoke(UI.ReloadCustomize);
	}

	Resize = function () {
		api.Invoke(UI.Resize);
	}
	
	setTimeout = function (fn) {
		api.OutputDebugString(fn + "\n");
		api.Invoke(UI.setTimeoutAsync, arguments);
	}

	ShowDialog = function () {
		api.Invoke(UI.ShowDialog, arguments);
	}
}
