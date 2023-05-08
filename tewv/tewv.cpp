// Tablacus WebView2 (C)2020 Gaku
// MIT Lisence
// Visual Studio Express 2017 for Windows Desktop
// https://tablacus.github.io/

#include "tewv.h"

// Global Variables:
const TCHAR g_szProgid[] = TEXT("Tablacus.WebView2");
const TCHAR g_szClsid[] = TEXT("{55BBF1B8-0D30-4908-BE0C-D576612A0F48}");
HINSTANCE	g_hinstDll = NULL;
LPWSTR g_versionInfo = NULL;
LONG		g_lLocks = 0;
#ifdef _DEBUG
HMODULE		g_hWebView2Loader = NULL;
LPFNCreateCoreWebView2EnvironmentWithOptions _CreateCoreWebView2EnvironmentWithOptions = NULL;
#endif

std::unordered_map<std::string, DISPID> g_umSW = {
	{ "name", TE_PROPERTY + 1 },
	{ "fullname", TE_PROPERTY + 2 },
	{ "path", TE_PROPERTY + 3 },
	{ "visible", TE_PROPERTY + 4 },
	{ "document", TE_PROPERTY + 5 },
	{ "window", TE_PROPERTY + 6 },
};

std::unordered_map<std::wstring, DISPID> g_umArray = {
	{ L"Item", DISPID_TE_ITEM },
	{ L"Count", DISPID_TE_COUNT },
	{ L"length", DISPID_TE_COUNT },
	{ L"push", TE_METHOD + 1 },
	{ L"pop", TE_METHOD + 2 },
	{ L"shift", TE_METHOD + 3 },
	{ L"unshift", TE_METHOD + 4 },
	{ L"join", TE_METHOD + 5 },
	{ L"slice", TE_METHOD + 6 },
	{ L"splice", TE_METHOD + 7 }
};

// Unit
VOID SafeRelease(PVOID ppObj)
{
	try {
		IUnknown **ppunk = static_cast<IUnknown **>(ppObj);
		if (*ppunk) {
			(*ppunk)->Release();
			*ppunk = NULL;
		}
	} catch (...) {}
}

VOID LockModule()
{
	::InterlockedIncrement(&g_lLocks);
}

VOID UnlockModule()
{
	::InterlockedDecrement(&g_lLocks);
}

BSTR GetLPWSTRFromVariant(VARIANT *pv)
{
	switch (pv->vt) {
		case VT_VARIANT | VT_BYREF:
			return GetLPWSTRFromVariant(pv->pvarVal);
		case VT_BSTR:
		case VT_LPWSTR:
			return pv->bstrVal;
		default:
			return NULL;
	}//end_switch
}

int GetIntFromVariant(VARIANT *pv)
{
	if (pv) {
		if (pv->vt == (VT_VARIANT | VT_BYREF)) {
			return GetIntFromVariant(pv->pvarVal);
		}
		if (pv->vt == VT_I4) {
			return pv->lVal;
		}
		if (pv->vt == VT_UI4) {
			return pv->ulVal;
		}
		if (pv->vt == VT_R8) {
			return (int)(LONGLONG)pv->dblVal;
		}
		VARIANT vo;
		VariantInit(&vo);
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_I4)) {
			return vo.lVal;
		}
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_UI4)) {
			return vo.ulVal;
		}
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_I8)) {
			return (int)vo.llVal;
		}
	}
	return 0;
}

int GetIntFromVariantClear(VARIANT *pv)
{
	int i = GetIntFromVariant(pv);
	VariantClear(pv);
	return i;
}

VOID teSetBool(VARIANT *pv, BOOL b)
{
	if (pv) {
		pv->boolVal = b ? VARIANT_TRUE : VARIANT_FALSE;
		pv->vt = VT_BOOL;
	}
}

VOID teSysFreeString(BSTR *pbs)
{
	if (*pbs) {
		::SysFreeString(*pbs);
		*pbs = NULL;
	}
}

VOID teSetLong(VARIANT *pv, LONG i)
{
	if (pv) {
		pv->lVal = i;
		pv->vt = VT_I4;
	}
}

VOID teSetLL(VARIANT *pv, LONGLONG ll)
{
	if (pv) {
		pv->lVal = static_cast<int>(ll);
		if (ll == static_cast<LONGLONG>(pv->lVal)) {
			pv->vt = VT_I4;
			return;
		}
		pv->dblVal = static_cast<DOUBLE>(ll);
		if (ll == static_cast<LONGLONG>(pv->dblVal)) {
			pv->vt = VT_R8;
			return;
		}
		SAFEARRAY *psa;
		psa = SafeArrayCreateVector(VT_I4, 0, sizeof(LONGLONG) / sizeof(int));
		if (psa) {
			PVOID pvData;
			if (::SafeArrayAccessData(psa, &pvData) == S_OK) {
				::CopyMemory(pvData, &ll, sizeof(LONGLONG));
				::SafeArrayUnaccessData(psa);
				pv->vt = VT_ARRAY | VT_I4;
				pv->parray = psa;
			}
		}
	}
}

BOOL teSetObject(VARIANT *pv, PVOID pObj)
{
	if (pObj) {
		try {
			IUnknown *punk = static_cast<IUnknown *>(pObj);
			if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pv->pdispVal))) {
				pv->vt = VT_DISPATCH;
				return true;
			}
			if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pv->punkVal))) {
				pv->vt = VT_UNKNOWN;
				return true;
			}
		} catch (...) {}
	}
	return false;
}

BOOL teSetObjectRelease(VARIANT *pv, PVOID pObj)
{
	if (pObj) {
		try {
			IUnknown *punk = static_cast<IUnknown *>(pObj);
			if (pv) {
				if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pv->pdispVal))) {
					pv->vt = VT_DISPATCH;
					SafeRelease(&punk);
					return true;
				}
				if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pv->punkVal))) {
					pv->vt = VT_UNKNOWN;
					SafeRelease(&punk);
					return true;
				}
			}
			SafeRelease(&punk);
		} catch (...) {}
	}
	return false;
}

VOID teSetSZ(VARIANT *pv, LPCWSTR lpstr)
{
	if (pv) {
		pv->bstrVal = ::SysAllocString(lpstr);
		pv->vt = VT_BSTR;
	}
}

VOID teSetBSTR(VARIANT *pv, BSTR bs, int nLen)
{
	if (pv) {
		pv->vt = VT_BSTR;
		if (bs) {
			if (nLen < 0) {
				nLen = lstrlen(bs);
			}
			if (::SysStringLen(bs) == nLen) {
				pv->bstrVal = bs;
				return;
			}
		}
		pv->bstrVal = SysAllocStringLen(bs, nLen);
		teSysFreeString(&bs);
	}
}

HRESULT teGetDispIdNum(LPOLESTR lpszName, int nMax, DISPID *pid)
{
	VARIANT v, vo;
	teSetSZ(&v, lpszName);
	VariantInit(&vo);
	if (SUCCEEDED(VariantChangeType(&vo, &v, 0, VT_I4))) {
		*pid = vo.lVal + DISPID_COLLECTION_MIN;
		VariantClear(&vo);
	}
	VariantClear(&v);
	if (*pid < DISPID_COLLECTION_MIN || *pid >= nMax + DISPID_COLLECTION_MIN) {
		*pid = DISPID_UNKNOWN;
	}
	return S_OK;
}

BOOL FindUnknown(VARIANT *pv, IUnknown **ppunk)
{
	if (pv) {
		if (pv->vt == VT_DISPATCH || pv->vt == VT_UNKNOWN) {
			*ppunk = pv->punkVal;
			return *ppunk != NULL;
		}
		if (pv->vt == (VT_VARIANT | VT_BYREF)) {
			return FindUnknown(pv->pvarVal, ppunk);
		}
		if (pv->vt == (VT_DISPATCH | VT_BYREF) || pv->vt == (VT_UNKNOWN | VT_BYREF)) {
			*ppunk = *pv->ppunkVal;
			return *ppunk != NULL;
		}
	}
	*ppunk = NULL;
	return FALSE;
}

HRESULT tePutProperty0(IUnknown *punk, LPOLESTR sz, VARIANT *pv, DWORD grfdex)
{
	HRESULT hr = E_FAIL;
	DISPID dispid, putid;
	DISPPARAMS dispparams;
	IDispatchEx *pdex;
	if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pdex))) {
		BSTR bs = ::SysAllocString(sz);
		hr = pdex->GetDispID(bs, grfdex, &dispid);
		if SUCCEEDED(hr) {
			putid = DISPID_PROPERTYPUT;
			dispparams.rgvarg = pv;
			dispparams.rgdispidNamedArgs = &putid;
			dispparams.cArgs = 1;
			dispparams.cNamedArgs = 1;
			hr = pdex->InvokeEx(dispid, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUTREF, &dispparams, NULL, NULL, NULL);
		}
		::SysFreeString(bs);
		SafeRelease(&pdex);
	}
	return hr;
}

HRESULT tePutProperty(IUnknown *punk, LPOLESTR sz, VARIANT *pv)
{
	return tePutProperty0(punk, sz, pv, fdexNameEnsure);
}

// VARIANT Clean-up of an array
VOID teClearVariantArgs(int nArgs, VARIANTARG *pvArgs)
{
	if (pvArgs && nArgs > 0) {
		for (int i = nArgs ; i-- >  0;){
			VariantClear(&pvArgs[i]);
		}
		delete[] pvArgs;
		pvArgs = NULL;
	}
}

HRESULT Invoke5(IDispatch *pdisp, DISPID dispid, WORD wFlags, VARIANT *pvResult, int nArgs, VARIANTARG *pvArgs)
{
	HRESULT hr;
	// DISPPARAMS
	DISPPARAMS dispParams;
	dispParams.rgvarg = pvArgs;
	dispParams.cArgs = abs(nArgs);
	DISPID dispidName = DISPID_PROPERTYPUT;
	if (wFlags & DISPATCH_PROPERTYPUT) {
		dispParams.cNamedArgs = 1;
		dispParams.rgdispidNamedArgs = &dispidName;
	} else {
		dispParams.rgdispidNamedArgs = NULL;
		dispParams.cNamedArgs = 0;
	}
	try {
		hr = pdisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
			wFlags, &dispParams, pvResult, NULL, NULL);
	} catch (...) {
		hr = E_FAIL;
	}
	teClearVariantArgs(nArgs, pvArgs);
	return hr;
}

HRESULT Invoke4(IDispatch *pdisp, VARIANT *pvResult, int nArgs, VARIANTARG *pvArgs)
{
	return Invoke5(pdisp, DISPID_VALUE, DISPATCH_METHOD, pvResult, nArgs, pvArgs);
}

VARIANTARG* GetNewVARIANT(int n)
{
	VARIANT *pv = new VARIANTARG[n];
	while (n--) {
		VariantInit(&pv[n]);
	}
	return pv;
}

BOOL GetDispatch(VARIANT *pv, IDispatch **ppdisp)
{
	IUnknown *punk;
	if (FindUnknown(pv, &punk)) {
		return SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(ppdisp)));
	}
	return FALSE;
}

VOID teVariantChangeType(__out VARIANTARG * pvargDest,
	__in const VARIANTARG * pvarSrc, __in VARTYPE vt)
{
	VariantInit(pvargDest);
	if FAILED(VariantChangeType(pvargDest, pvarSrc, 0, vt)) {
		pvargDest->llVal = 0;
	}
}

// Initialize & Finalize
BOOL WINAPI DllMain(HINSTANCE hinstDll, DWORD dwReason, LPVOID lpReserved)
{
	switch (dwReason) {
	case DLL_PROCESS_ATTACH:
#ifdef _DEBUG
		WCHAR pszPath[MAX_PATH * 2];
		g_hinstDll = hinstDll;
		GetModuleFileName(hinstDll, pszPath, MAX_PATH * 2);
		lstrcpy(PathFindFileName(pszPath), L"WebView2Loader.dll");
		g_hWebView2Loader = LoadLibrary(pszPath);
		if (g_hWebView2Loader) {
			LPFNGetAvailableCoreWebView2BrowserVersionString _GetAvailableCoreWebView2BrowserVersionString = NULL;
			*(FARPROC *)&_GetAvailableCoreWebView2BrowserVersionString = GetProcAddress(g_hWebView2Loader, "GetAvailableCoreWebView2BrowserVersionString");
			if (_GetAvailableCoreWebView2BrowserVersionString) {
				if (_GetAvailableCoreWebView2BrowserVersionString(NULL, &g_versionInfo) == S_OK) {
					*(FARPROC *)&_CreateCoreWebView2EnvironmentWithOptions = GetProcAddress(g_hWebView2Loader, "CreateCoreWebView2EnvironmentWithOptions");
				}
			}
		}
#endif
		break;
	case DLL_PROCESS_DETACH:
		if (g_versionInfo) {
			::CoTaskMemFree(g_versionInfo);
		}
#ifdef _DEBUG
		if (g_hWebView2Loader) {
			::FreeLibrary(g_hWebView2Loader);
		}
#endif
		break;
	}
	return TRUE;
}

// DLL Export

STDAPI DllCanUnloadNow(void)
{
	return g_lLocks == 0 ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
	static CteClassFactory serverFactory;
	CLSID clsid;
	HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;

	*ppv = NULL;
	CLSIDFromString(g_szClsid, &clsid);
	if (IsEqualCLSID(rclsid, clsid)) {
		hr = serverFactory.QueryInterface(riid, ppv);
	}
	return hr;
}

STDAPI DllRegisterServer(void)
{
	return E_NOTIMPL;
}

STDAPI DllUnregisterServer(void)
{
	return E_NOTIMPL;
}

//CteBase

CteBase::CteBase()
{
	m_cRef = 1;
	m_pOleClientSite = NULL;
	m_pdisp = NULL;
	m_bstrPath = NULL;
	m_pDocument = NULL;
	m_webviewController = NULL;
	m_webviewWindow = NULL;
}

CteBase::~CteBase()
{
	HWND hwnd;
	GetWindow(&hwnd);
	RevokeDragDrop(hwnd);
	teSysFreeString(&m_bstrPath);
	SafeRelease(&m_pOleClientSite);
	SafeRelease(&m_pdisp);
	SafeRelease(&m_pDocument);
	SafeRelease(&m_webviewController);
	SafeRelease(&m_webviewWindow);
}

STDMETHODIMP CteBase::QueryInterface(REFIID riid, void **ppvObject)
{
	if (IsEqualIID(riid, IID_IOleWindow)) {
		*ppvObject = static_cast<IOleInPlaceObject *>(this);
		AddRef();
		return S_OK;
	}
	static const QITAB qit[] =
	{
		QITABENT(CteBase, IDispatch),
		QITABENT(CteBase, IWebBrowser),
		QITABENT(CteBase, IWebBrowserApp),
		QITABENT(CteBase, IWebBrowser2),
		QITABENT(CteBase, IOleObject),
		QITABENT(CteBase, IOleInPlaceObject),
		QITABENT(CteBase, IDropTarget),
		QITABENT(CteBase, IShellBrowser),
		QITABENT(CteBase, IServiceProvider),
		QITABENT(CteBase, ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler),
		QITABENT(CteBase, ICoreWebView2CreateCoreWebView2ControllerCompletedHandler),
		QITABENT(CteBase, ICoreWebView2DocumentTitleChangedEventHandler),
		QITABENT(CteBase, ICoreWebView2NavigationCompletedEventHandler),
		{ 0 }
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteBase::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteBase::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}
	return m_cRef;
}

STDMETHODIMP CteBase::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteBase::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	CHAR pszName[9];
	for (int i = 0;; ++i) {
		WCHAR wc = rgszNames[0][i];
		if (i == _countof(pszName) - 1 || wc < '0' || wc > 'z') {
			pszName[i] = NULL;
			break;
		}
		pszName[i] = tolower(wc);
	}
	auto itr = g_umSW.find(pszName);
	if (itr != g_umSW.end()) {
		*rgDispId = itr->second;
		return S_OK;
	}
	OutputDebugStringA("GetIDsOfNames:");
	OutputDebugString(rgszNames[0]);
	OutputDebugStringA("\n");
	return DISP_E_UNKNOWNNAME;
}

STDMETHODIMP CteBase::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	HRESULT hr = S_OK;
	try {
		switch (dispIdMember) {

		case TE_PROPERTY + 1://name
			if (pVarResult) {
				get_Name(&pVarResult->bstrVal);
				pVarResult->vt = VT_BSTR;
			}
			return S_OK;

		case TE_PROPERTY + 2://fullname
			if (pVarResult) {
				get_FullName(&pVarResult->bstrVal);
				pVarResult->vt = VT_BSTR;
			}
			return S_OK;

		case TE_PROPERTY + 3://path
			if (pVarResult) {
				get_Path(&pVarResult->bstrVal);
				pVarResult->vt = VT_BSTR;
			}
			return S_OK;
	
		case TE_PROPERTY + 4://visible
			if (nArg >= 0) {
				put_Visible(GetIntFromVariant(&pDispParams->rgvarg[nArg]));
			}
			if (pVarResult) {
				get_Visible(&pVarResult->boolVal);
				pVarResult->vt = VT_BOOL;
			}
			return S_OK;

		case TE_PROPERTY + 5://document
			if (pVarResult) {
				if SUCCEEDED(get_Document(&pVarResult->pdispVal)) {
					pVarResult->vt = VT_DISPATCH;
				}
			}
			return S_OK;

		case TE_PROPERTY + 6://window
			if (pVarResult) {
				if SUCCEEDED(get_Document(&pVarResult->pdispVal)) {
					pVarResult->vt = VT_DISPATCH;
				}
			}
			return S_OK;

		case DISPID_VALUE://this
			teSetObject(pVarResult, this);
			return S_OK;
		}//end_switch
	} catch (...) {
		teSetLong(pVarResult, E_UNEXPECTED);
		return S_OK;
	}
	return DISP_E_MEMBERNOTFOUND;
}

//IWebBrowser
STDMETHODIMP CteBase::GoBack(void)
{
	return m_webviewWindow->GoBack();
}

STDMETHODIMP CteBase::GoForward(void)
{
	return m_webviewWindow->GoForward();
}

STDMETHODIMP CteBase::GoHome(void)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::GoSearch(void)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::Navigate(BSTR URL, VARIANT *Flags, VARIANT *TargetFrameName, VARIANT *PostData, VARIANT *Headers)
{
	if (GetIntFromVariant(Flags) == 1) {
		if (m_webviewWindow) {
			return m_webviewWindow->NavigateToString(URL);
		}
		return E_FAIL;
	}
	teSysFreeString(&m_bstrPath);
	m_bstrPath = ::SysAllocString(URL);
	if (m_webviewWindow) {
		return m_webviewWindow->Navigate(URL);
	}
	return S_OK;
}

STDMETHODIMP CteBase::Refresh(void)
{
	if (m_webviewWindow) {
		return m_webviewWindow->Reload();
	}
	return E_FAIL;
}

STDMETHODIMP CteBase::Refresh2(VARIANT *Level)
{
	return Refresh();
}

STDMETHODIMP CteBase::Stop(void)
{
	if (m_webviewWindow) {
		return m_webviewWindow->Stop();
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::get_Application(IDispatch **ppDisp)
{
	return QueryInterface(IID_PPV_ARGS(ppDisp));
}

STDMETHODIMP CteBase::get_Parent(IDispatch **ppDisp)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::get_Container(IDispatch **ppDisp)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::get_Document(IDispatch **ppDisp)
{
	if (m_pDocument) {
		return m_pDocument->QueryInterface(IID_PPV_ARGS(ppDisp));
	}
	return E_NOINTERFACE;
}

STDMETHODIMP CteBase::get_TopLevelContainer(VARIANT_BOOL *pBool)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::get_Type(BSTR *Type)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::get_Left(long *pl)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::put_Left(long Left)
{
	return S_OK;
}

STDMETHODIMP CteBase::get_Top(long *pl)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::put_Top(long Top)
{
	if (m_webviewController) {
		m_webviewController->NotifyParentWindowPositionChanged();
	}
	return S_OK;
}

STDMETHODIMP CteBase::get_Width(long *pl)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::put_Width(long Width)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::get_Height(long *pl)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::put_Height(long Height)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::get_LocationName(BSTR *LocationName)
{
	*LocationName = ::SysAllocString(m_bstrPath);
	return S_OK;
}

STDMETHODIMP CteBase::get_LocationURL(BSTR *LocationURL)
{
	*LocationURL = ::SysAllocString(m_bstrPath);
	return S_OK;
}

STDMETHODIMP CteBase::get_Busy(VARIANT_BOOL *pBool)
{
	*pBool = VARIANT_FALSE;
	return S_OK;
}

//IWebBrowserApp
STDMETHODIMP CteBase::Quit(void)
{
	return m_webviewWindow->Stop();
}

STDMETHODIMP CteBase::ClientToWindow(int *pcx, int *pcy)
{
	HWND hwnd;
	GetWindow(&hwnd);
	POINT pt = { *pcx, *pcy };
	ClientToScreen(hwnd, &pt);
	*pcx = pt.x;
	*pcy = pt.y;
	return S_OK;
}

STDMETHODIMP CteBase::PutProperty(BSTR Property, VARIANT vtValue)
{
	if (lstrcmpi(Property, L"document") == 0) {
		SafeRelease(&m_pDocument);
		GetDispatch(&vtValue, &m_pDocument);
	}
	return S_OK;
}

STDMETHODIMP CteBase::GetProperty(BSTR Property, VARIANT *pvtValue)
{
	if (lstrcmpi(Property, L"InvokeMethod") == 0) {
		BSTR bs = ::SysAllocString(L"_InvokeMethod();");
		m_webviewWindow->ExecuteScript(bs, this);
		teSysFreeString(&bs);
		return S_OK;
	}
	if (lstrcmpi(Property, L"version") == 0) {
		teSetLong(pvtValue, VER_Y * 1000000 + VER_M * 10000 + VER_D * 100 + VER_Z);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP CteBase::get_Name(BSTR *Name)
{
	*Name = ::SysAllocStringLen(L"WebView2/", lstrlen(g_versionInfo) + 9);
	lstrcat(*Name, g_versionInfo);
	return S_OK;
}

STDMETHODIMP CteBase::get_HWND(SHANDLE_PTR *pHWND)
{
	HWND hwnd;
	HRESULT hr = GetWindow(&hwnd);
	*pHWND = (HANDLE_PTR)hwnd;
	return S_OK;
}

STDMETHODIMP CteBase::get_FullName(BSTR *FullName)
{
	WCHAR pszPath[MAX_PATH];
	GetModuleFileName(NULL, pszPath, MAX_PATH);
	*FullName = ::SysAllocString(pszPath);
	return S_OK;
}

STDMETHODIMP CteBase::get_Path(BSTR *Path)
{
	WCHAR pszPath[MAX_PATH];
	GetModuleFileName(NULL, pszPath, MAX_PATH);
	LPWSTR lp = PathFindFileName(pszPath);
	if (lp) {
		lp[0] = NULL;
	}
	*Path = ::SysAllocString(pszPath);
	return S_OK;
}

STDMETHODIMP CteBase::get_Visible(VARIANT_BOOL *pBool)
{
	HRESULT hr = E_FAIL;
	BOOL bVisible;
	if (m_webviewController) {
		hr = m_webviewController->get_IsVisible(&bVisible);
		*pBool = bVisible ? VARIANT_TRUE : VARIANT_FALSE;
	}
	return hr;
}

STDMETHODIMP CteBase::put_Visible(VARIANT_BOOL Value)
{
	HRESULT hr = E_FAIL;
	if (m_webviewController) {
		hr = m_webviewController->put_IsVisible(Value);
	}
	return hr;
}

STDMETHODIMP CteBase::get_StatusBar(VARIANT_BOOL *pBool)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::put_StatusBar(VARIANT_BOOL Value)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::get_StatusText(BSTR *StatusText)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::put_StatusText(BSTR StatusText)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::get_ToolBar(int *Value)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::put_ToolBar(int Value)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::get_MenuBar(VARIANT_BOOL *Value)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::put_MenuBar(VARIANT_BOOL Value)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::get_FullScreen(VARIANT_BOOL *pbFullScreen)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::put_FullScreen(VARIANT_BOOL bFullScreen)
{
	return E_NOTIMPL;
}
//IWebBrowser2
STDMETHODIMP CteBase::Navigate2(VARIANT *URL, VARIANT *Flags, VARIANT *TargetFrameName, VARIANT *PostData, VARIANT *Headers)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::QueryStatusWB(OLECMDID cmdID, OLECMDF *pcmdf)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::ExecWB(OLECMDID cmdID, OLECMDEXECOPT cmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::ShowBrowserBar(VARIANT *pvaClsid, VARIANT *pvarShow, VARIANT *pvarSize)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::get_ReadyState(READYSTATE *plReadyState)
{
	*plReadyState = READYSTATE_COMPLETE;
	return S_OK;
}

STDMETHODIMP CteBase::get_Offline(VARIANT_BOOL *pbOffline)
{
	*pbOffline = VARIANT_FALSE;
	return S_OK;
}

STDMETHODIMP CteBase::put_Offline(VARIANT_BOOL bOffline)
{
	return S_OK;
}

STDMETHODIMP CteBase::get_Silent(VARIANT_BOOL *pbSilent)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::put_Silent(VARIANT_BOOL bSilent)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::get_RegisterAsBrowser(VARIANT_BOOL *pbRegister)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::put_RegisterAsBrowser(VARIANT_BOOL bRegister)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::get_RegisterAsDropTarget(VARIANT_BOOL *pbRegister)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::put_RegisterAsDropTarget(VARIANT_BOOL bRegister)
{
	HWND hwnd;
	GetWindow(&hwnd);
	RevokeDragDrop(hwnd);
	if (bRegister) {
		IDocHostUIHandler *pDocHostUIHandler;
		if SUCCEEDED(m_pOleClientSite->QueryInterface(IID_PPV_ARGS(&pDocHostUIHandler))) {
			IDropTarget	*pDropTarget;
			pDocHostUIHandler->GetDropTarget(this, &pDropTarget);
			RegisterDragDrop(hwnd, pDropTarget);
			pDropTarget->Release();
		}
	}
	return S_OK;
}

STDMETHODIMP CteBase::get_TheaterMode(VARIANT_BOOL *pbRegister)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::put_TheaterMode(VARIANT_BOOL bRegister)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::get_AddressBar(VARIANT_BOOL *Value)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::put_AddressBar(VARIANT_BOOL Value)
{
	return E_NOTIMPL;
}
STDMETHODIMP CteBase::get_Resizable(VARIANT_BOOL *Value)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::put_Resizable(VARIANT_BOOL Value)
{
	return E_NOTIMPL;
}

//IOleObject
STDMETHODIMP CteBase::SetClientSite(IOleClientSite *pClientSite)
{
	SafeRelease(&m_pOleClientSite);
	SafeRelease(&m_pdisp);
	if (pClientSite) {
		pClientSite->QueryInterface(IID_PPV_ARGS(&m_pOleClientSite));
		pClientSite->QueryInterface(IID_PPV_ARGS(&m_pdisp));
	}
	return S_OK;
}

STDMETHODIMP CteBase::GetClientSite(IOleClientSite **ppClientSite)
{
	return m_pOleClientSite ? m_pOleClientSite->QueryInterface(IID_PPV_ARGS(ppClientSite)) : E_NOINTERFACE;
}

STDMETHODIMP CteBase::SetHostNames(LPCOLESTR szContainerApp, LPCOLESTR szContainerObj)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::Close(DWORD dwSaveOption)
{
	return S_OK;
}

STDMETHODIMP CteBase::SetMoniker(DWORD dwWhichMoniker, IMoniker *pmk)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::InitFromData(IDataObject *pDataObject, BOOL fCreation, DWORD dwReserved)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::GetClipboardData(DWORD dwReserved, IDataObject **ppDataObject)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::DoVerb(LONG iVerb, LPMSG lpmsg, IOleClientSite *pActiveSite, LONG lindex, HWND hwndParent, LPCRECT lprcPosRect)
{
	if (iVerb == OLEIVERB_INPLACEACTIVATE) {
		m_hwndParent = hwndParent;
		WCHAR pszDataPath[MAX_PATH], pszSetting[MAX_PATH + 64], pszProxyServer[MAX_PATH];
		lstrcpy(pszSetting, L"--disable-web-security");
		GetTempPath(MAX_PATH, pszDataPath);
		PathAppend(pszDataPath, L"tablacus");
		pszProxyServer[0] = NULL;
		HKEY hKey;
		if (RegOpenKeyExA(HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings", 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
			DWORD dwSize = sizeof(DWORD);
			DWORD dwProxyEnable = 0;
			RegQueryValueEx(hKey, L"ProxyEnable", NULL, NULL, (LPBYTE)&dwProxyEnable, &dwSize);
			if (dwProxyEnable) {
				dwSize = MAX_PATH;
				RegQueryValueEx(hKey, L"ProxyServer", NULL, NULL, (LPBYTE)&pszProxyServer, &dwSize);
			}
			RegCloseKey(hKey);
		}
		auto options = Microsoft::WRL::Make<CoreWebView2EnvironmentOptions>();
		if (pszProxyServer[0]) {
			lstrcat(pszSetting, L" --proxy-server=");
			lstrcat(pszSetting, pszProxyServer);
		}
		options->put_AdditionalBrowserArguments(pszSetting);
#ifdef _DEBUG
		_CreateCoreWebView2EnvironmentWithOptions(NULL, pszDataPath, options.Get(), this);
#else
		CreateCoreWebView2EnvironmentWithOptions(NULL, pszDataPath, options.Get(), this);
#endif
		return S_OK;
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::EnumVerbs(IEnumOLEVERB **ppEnumOleVerb)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::Update(void)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::IsUpToDate(void)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::GetUserClassID(CLSID *pClsid)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::GetUserType(DWORD dwFormOfType, LPOLESTR *pszUserType)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::SetExtent(DWORD dwDrawAspect, SIZEL *psizel)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::GetExtent(DWORD dwDrawAspect, SIZEL *psizel)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::Advise(IAdviseSink *pAdvSink, DWORD *pdwConnection)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::Unadvise(DWORD dwConnection)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::EnumAdvise(IEnumSTATDATA **ppenumAdvise)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::GetMiscStatus(DWORD dwAspect, DWORD *pdwStatus)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::SetColorScheme(LOGPALETTE *pLogpal)
{
	return E_NOTIMPL;
}

//IOleWindow
STDMETHODIMP CteBase::GetWindow(HWND *phwnd)
{
	HRESULT hr = E_FAIL;
	if (m_webviewController) {
		HWND hwnd1, hwnd = NULL;
		hr = m_webviewController->get_ParentWindow(&hwnd1);
		if (hr == S_OK) {
			hwnd = hwnd1;
			hwnd1 = FindWindowEx(hwnd1, NULL, L"Chrome_WidgetWin_0", NULL);
			if (hwnd1) {
				hwnd = hwnd1;
				hwnd1 = FindWindowEx(hwnd1, NULL, L"Chrome_WidgetWin_1", NULL);
				if (hwnd1) {
					hwnd = hwnd1;
					hwnd1 = FindWindowEx(hwnd1, NULL, L"Chrome_RenderWidgetHostHWND", NULL);
					if (hwnd1) {
						hwnd = hwnd1;
					}
				}
			}
		}
		*phwnd = hwnd;
	}
	return hr;
}

STDMETHODIMP CteBase::ContextSensitiveHelp(BOOL fEnterMode)
{
	return E_NOTIMPL;
}

//IOleInPlaceObject
STDMETHODIMP CteBase::InPlaceDeactivate(void)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::UIDeactivate(void)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::SetObjectRects(LPCRECT lprcPosRect, LPCRECT lprcClipRect)
{
	RECT bounds;
	if (!lprcPosRect) {
		GetClientRect(m_hwndParent, &bounds);
		lprcClipRect = &bounds;
	}
	HDC hdc = GetDC(m_hwndParent);
	int deviceYDPI = GetDeviceCaps(hdc, LOGPIXELSY);
	ReleaseDC(m_hwndParent, hdc);
	m_webviewController->SetBoundsAndZoomFactor(*lprcClipRect, 96.0 / deviceYDPI);
	return S_OK;
}

STDMETHODIMP CteBase::ReactivateAndUndo(void)
{
	if (m_webviewController) {
		m_webviewController->MoveFocus(COREWEBVIEW2_MOVE_FOCUS_REASON_PROGRAMMATIC);
	}
	return S_OK;
}

//IDropTarget
STDMETHODIMP CteBase::DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::DragLeave()
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::Drop(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
{
	return E_NOTIMPL;
}

//IShellBrowser
STDMETHODIMP CteBase::InsertMenusSB(HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::SetMenuSB(HMENU hmenuShared, HOLEMENU holemenuRes, HWND hwndActiveObject)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::RemoveMenusSB(HMENU hmenuShared)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::SetStatusTextSB(LPCWSTR lpszStatusText)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::EnableModelessSB(BOOL fEnable)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::TranslateAcceleratorSB(LPMSG lpmsg, WORD wID)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::BrowseObject(PCUIDLIST_RELATIVE pidl, UINT wFlags)
{
	WCHAR pszPath[MAX_PATH];
	if (SHGetPathFromIDList(pidl, pszPath)) {
		VARIANTARG *pv = GetNewVARIANT(7);
		VARIANT_BOOL bCancel = VARIANT_FALSE;
		teSetObject(&pv[6], this);
		pv[5].bstrVal = ::SysAllocString(pszPath);
		pv[5].vt = VT_BSTR;
		pv[0].vt = VT_BYREF | VT_BOOL;
		pv[0].pboolVal = &bCancel;
		Invoke5(m_pdisp, DISPID_BEFORENAVIGATE2, DISPATCH_METHOD, NULL, 7, pv);
		if (bCancel) {
			return S_OK;
		}
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::GetViewStateStream(DWORD grfMode, IStream **ppStrm)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::GetControlWindow(UINT id, HWND *lphwnd)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::SendControlMsg(UINT id, UINT uMsg, WPARAM wParam, LPARAM lParam, LRESULT *pret){
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::QueryActiveShellView(IShellView **ppshv){
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::OnViewWindowActive(IShellView *ppshv){
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::SetToolbarItems(LPTBBUTTONSB lpButtons, UINT nButtons, UINT uFlags)
{
	return E_NOTIMPL;
}

//IServiceProvider
STDMETHODIMP CteBase::QueryService(REFGUID guidService, REFIID riid, void **ppv)
{
	if (IsEqualGUID(guidService, SID_TablacusObject)) {
		*ppv = new CteObjectEx();
		return S_OK;
	}
	if (IsEqualGUID(guidService, SID_TablacusArray)) {
		*ppv = new CteArray();
		return S_OK;
	}
	if (IsEqualIID(riid, IID_IShellBrowser)) {
		return QueryInterface(riid, ppv);
	}
	return E_NOINTERFACE;
}

//ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
STDMETHODIMP CteBase::Invoke(HRESULT result, ICoreWebView2Environment *created_environment)
{
	created_environment->CreateCoreWebView2Controller(m_hwndParent, this);
	return S_OK;
}

//ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
STDMETHODIMP CteBase::Invoke(HRESULT result, ICoreWebView2Controller *createdController)
{
	createdController->QueryInterface(IID_PPV_ARGS(&m_webviewController));
	m_webviewController->get_CoreWebView2(&m_webviewWindow);
	ICoreWebView2Settings* Settings;
	m_webviewWindow->get_Settings(&Settings);
	Settings->put_IsScriptEnabled(TRUE);
	Settings->put_AreDefaultScriptDialogsEnabled(TRUE);
	Settings->put_IsWebMessageEnabled(TRUE);
	Settings->put_IsStatusBarEnabled(FALSE);
	if (m_pOleClientSite) {
		IDocHostUIHandler *pDocHostUIHandler;
		if SUCCEEDED(m_pOleClientSite->QueryInterface(IID_PPV_ARGS(&pDocHostUIHandler))) {
			VARIANT v;
			if SUCCEEDED(pDocHostUIHandler->GetExternal(&v.pdispVal)) {
				v.vt = VT_DISPATCH;
				m_webviewWindow->AddHostObjectToScript(L"te", &v);
				VariantClear(&v);
			}
			pDocHostUIHandler->Release();
		}
		m_webviewWindow->add_DocumentTitleChanged(this, &m_documentTitleChangedToken);
		m_webviewWindow->add_NavigationStarting(this, &m_navigationStartingToken);
		m_webviewWindow->add_NavigationCompleted(this, &m_navigationCompletedToken);
	}
	SetObjectRects(NULL, NULL);
	if (m_bstrPath) {
		m_webviewWindow->Navigate(m_bstrPath);
	}
	m_webviewController->put_IsVisible(TRUE);
	return S_OK;
}

//ICoreWebView2ExecuteScriptCompletedHandler
STDMETHODIMP CteBase::Invoke(HRESULT result, LPCWSTR resultObjectAsJson)
{
	return S_OK;
}

//ICoreWebView2DocumentTitleChangedEventHandler
STDMETHODIMP CteBase::Invoke(ICoreWebView2* sender, IUnknown* args) {
	if (!m_pdisp) {
		return E_NOTIMPL;
	}
	LPWSTR title = NULL;
	if (sender->get_DocumentTitle(&title) == S_OK) {
		VARIANT v;
		v.bstrVal = ::SysAllocString(title);
		::CoTaskMemFree(title);
		v.vt = VT_BSTR;
		Invoke5(m_pdisp, DISPID_TITLECHANGE, DISPATCH_METHOD, NULL, -1, &v);
		VariantClear(&v);
	}
	return S_OK;
}

//ICoreWebView2DocumentTitleChangedEventHandler
STDMETHODIMP CteBase::Invoke(ICoreWebView2* sender, ICoreWebView2NavigationStartingEventArgs* args) {
	VARIANTARG *pv = GetNewVARIANT(7);
	VARIANT_BOOL bCancel = VARIANT_FALSE;
	teSetObject(&pv[6], this);
	LPWSTR lpPath;
	if SUCCEEDED(args->get_Uri(&lpPath)) {
		pv[5].bstrVal = ::SysAllocString(lpPath);
		pv[5].vt = VT_BSTR;
		CoTaskMemFree(lpPath);

	}
	pv[0].vt = VT_BYREF | VT_BOOL;
	pv[0].pboolVal = &bCancel;
	Invoke5(m_pdisp, DISPID_BEFORENAVIGATE2, DISPATCH_METHOD, NULL, 7, pv);
	if (bCancel) {
		args->put_Cancel(true);
	}
	return S_OK;
}

//ICoreWebView2NavigationCompletedEventHandler
STDMETHODIMP CteBase::Invoke(ICoreWebView2* sender, ICoreWebView2NavigationCompletedEventArgs* args)
{
	if (m_webviewController) {
		m_webviewController->NotifyParentWindowPositionChanged();
	}
	SetObjectRects(NULL, NULL);
	return Invoke5(m_pdisp, DISPID_DOCUMENTCOMPLETE, DISPATCH_METHOD, NULL, 0, NULL);
}
// CteClassFactory

STDMETHODIMP CteClassFactory::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteClassFactory, IClassFactory),
		{ 0 }
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteClassFactory::AddRef()
{
	LockModule();
	return 2;
}

STDMETHODIMP_(ULONG) CteClassFactory::Release()
{
	UnlockModule();
	return 1;
}

STDMETHODIMP CteClassFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject)
{
	*ppvObject = NULL;

	if (pUnkOuter != NULL) {
		return CLASS_E_NOAGGREGATION;
	}
#ifdef _DEBUG
	if (!_CreateCoreWebView2EnvironmentWithOptions) {
		return CLASS_E_CLASSNOTAVAILABLE;
	}
#else
	if (GetAvailableCoreWebView2BrowserVersionString(NULL, &g_versionInfo) != S_OK) {
		return CLASS_E_CLASSNOTAVAILABLE;
	}
#endif
	*ppvObject = new CteBase();
	return S_OK;
}

STDMETHODIMP CteClassFactory::LockServer(BOOL fLock)
{
	if (fLock) {
		LockModule();
		return S_OK;
	}
	UnlockModule();
	return S_OK;
}

///CteArray
CteArray::CteArray()
{
	m_cRef = 1;
	SAFEARRAYBOUND sab = { 0, 0 };
	m_psa = SafeArrayCreate(VT_VARIANT, 1, &sab);
}

CteArray::~CteArray()
{
	SafeArrayDestroy(m_psa);
}

STDMETHODIMP CteArray::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteArray, IDispatch),
		QITABENT(CteArray, IDispatchEx),
		{ 0 }
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteArray::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteArray::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteArray::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteArray::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteArray::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	auto itr = g_umArray.find(*rgszNames);
	if (itr != g_umArray.end()) {
		*rgDispId = itr->second;
		return S_OK;
	}
	return teGetDispIdNum(*rgszNames, TE_PROPERTY - DISPID_COLLECTION_MIN, rgDispId);
}

STDMETHODIMP CteArray::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	LONG n, lIndex;
	SAFEARRAYBOUND sab;
	try {
		if (wFlags == DISPATCH_PROPERTYGET && dispIdMember >= TE_METHOD) {
			teSetObjectRelease(pVarResult, new CteDispatch(this, 0, dispIdMember));
			return S_OK;
		}
		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		switch(dispIdMember) {

		case DISPID_TE_ITEM://Item
			if (nArg >= 0) {
				ItemEx(GetIntFromVariant(&pDispParams->rgvarg[nArg]), pVarResult, nArg >= 1 ? &pDispParams->rgvarg[nArg - 1] : NULL);
			}
			return S_OK;

		case DISPID_TE_COUNT://Count
			if (nArg >= 0 && (wFlags & DISPATCH_PROPERTYPUT)) {
				VARIANT v;
				VariantInit(&v);
				sab = { (ULONG)GetIntFromVariant(&pDispParams->rgvarg[nArg]), 0 };
				::SafeArrayRedim(m_psa, &sab);
			}
			teSetLong(pVarResult, GetCount());
			return S_OK;

		case TE_METHOD + 1://push
			n = GetCount();
			sab = { (ULONG)n + nArg + 1, 0 };
			::SafeArrayRedim(m_psa, &sab);
			for (int i = 0; i <= nArg; ++i) {
				lIndex = n + i;
				::SafeArrayPutElement(m_psa, &lIndex, &pDispParams->rgvarg[nArg - i]);
			}
			return S_OK;

		case TE_METHOD + 2://pop
			n = GetCount();
			if (--n >= 0) {
				if (pVarResult) {
					::SafeArrayGetElement(m_psa, &n, pVarResult);
				}
				sab = { (ULONG)n, 0 };
				::SafeArrayRedim(m_psa, &sab);
			}
			return S_OK;

		case TE_METHOD + 3://shift
			n = GetCount();
			if (n) {
				if (pVarResult) {
					LONG l = 0;
					::SafeArrayGetElement(m_psa, &l, pVarResult);
				}
				VARIANT *pv;
				if (::SafeArrayAccessData(m_psa, (LPVOID *)&pv) == S_OK) {
					::VariantClear(pv);
					::CopyMemory(pv, &pv[1], sizeof(VARIANT) * --n);
					::VariantInit(&pv[n]);
					::SafeArrayUnaccessData(m_psa);
					SAFEARRAYBOUND sab = { (ULONG)n, 0 };
					::SafeArrayRedim(m_psa, &sab);
				}
			}
			return S_OK;

		case TE_METHOD + 4://unshift
			n = GetCount();
			VARIANT *pv;
			sab = { (ULONG)n + nArg + 1, 0 };
			::SafeArrayRedim(m_psa, &sab);
			if (::SafeArrayAccessData(m_psa, (LPVOID *)&pv) == S_OK) {
				::CopyMemory(&pv[nArg + 1], pv, sizeof(VARIANT) * n);
				for (int i = nArg; i >= 0; --i) {
					::VariantInit(&pv[i]);
				}
				::SafeArrayUnaccessData(m_psa);
				for (LONG i = nArg; i >= 0; --i) {
					::SafeArrayPutElement(m_psa, &i, &pDispParams->rgvarg[nArg - i]);
				}
			}
			return S_OK;

		case TE_METHOD + 5://join
			if (pVarResult) {
				UINT n = 0;
				VARIANT vSeparator, v, vs;
				if (nArg >= 0) {
					teVariantChangeType(&vSeparator, &pDispParams->rgvarg[nArg], VT_BSTR);
				} else {
					teSetSZ(&vSeparator, L",");
				}
				UINT nSeparator = ::SysStringByteLen(vSeparator.bstrVal);
				for (LONG i = 0; i < GetCount(); ++i) {
					if (i) {
						n += nSeparator;
					}
					VariantInit(&vs);
					::SafeArrayGetElement(m_psa, &i, &vs);
					teVariantChangeType(&v, &vs, VT_BSTR);
					VariantClear(&vs);
					n += ::SysStringByteLen(v.bstrVal);
					VariantClear(&v);
				}
				pVarResult->vt = VT_BSTR;
				pVarResult->bstrVal = ::SysAllocStringByteLen(NULL, n);
				BYTE *p = (BYTE *)pVarResult->bstrVal;
				for (LONG i = 0; i < GetCount(); ++i) {
					if (i) {
						CopyMemory(p, vSeparator.bstrVal, nSeparator);
						p += nSeparator;
					}
					VariantInit(&vs);
					::SafeArrayGetElement(m_psa, &i, &vs);
					teVariantChangeType(&v, &vs, VT_BSTR);
					VariantClear(&vs);
					UINT nSize = ::SysStringByteLen(v.bstrVal);
					CopyMemory(p, v.bstrVal, nSize);
					p += nSize;
					VariantClear(&v);
				}
			}
			return S_OK;

		case TE_METHOD + 6://Slice
		case TE_METHOD + 7://Splice
			CteArray *pNewArray;
			if (pVarResult) {
				pNewArray = new CteArray();
			}
			if (nArg >= 0) {
				LONG nStart = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				LONG nLen = nArg >= 1 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) : MAXINT;
				if (nStart > GetCount() - nLen) {
					nLen = GetCount() - nStart;
				}
				if (nLen < 0) {
					nLen = 0;
				}
				VARIANT v;
				VariantInit(&v);
				if (pVarResult) {
					for (LONG i = nStart; i < nLen; ++i) {
						ItemEx(i, &v, NULL);
						pNewArray->ItemEx(-1, NULL, &v);
						VariantClear(&v);
					}
					teSetObjectRelease(pVarResult, pNewArray);
				}
				if (dispIdMember == TE_METHOD + 7 && nLen) {//Splice
					int n = GetCount();
					if (::SafeArrayAccessData(m_psa, (LPVOID *)&pv) == S_OK) {
						for (int i = nLen; --i >= 0;) {
							VariantClear(&pv[nStart + i]);
						}
						::CopyMemory(&pv[nStart], &pv[nStart + nLen], sizeof(VARIANT) * (n - (nStart + nLen)));
						for (int i = nLen; i > 0; --i) {
							::VariantInit(&pv[n - i]);
						}
						::SafeArrayUnaccessData(m_psa);
					}
					sab = { (ULONG)n - nLen, 0 };
					::SafeArrayRedim(m_psa, &sab);
				}
			}
			return S_OK;

		case DISPID_PROPERTYPUT:
			if ((wFlags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) && nArg >= 0) {
				if (pDispParams->rgvarg[nArg].vt == (VT_ARRAY | VT_VARIANT)) {
					SafeArrayDestroy(m_psa);
					SafeArrayCopy(pDispParams->rgvarg[nArg].parray, &m_psa);
				}
			}
			return S_OK;

		case DISPID_VALUE://Value
			teSetObject(pVarResult, this);
		case DISPID_UNKNOWN:
			return S_OK;

		default:
			if (dispIdMember >= DISPID_COLLECTION_MIN && dispIdMember <= DISPID_TE_MAX) {
				if (wFlags & DISPATCH_METHOD) {
					VARIANT v;
					VariantInit(&v);
					ItemEx(dispIdMember - DISPID_COLLECTION_MIN, &v, NULL);
					if (v.vt == VT_DISPATCH) {
						Invoke5(v.pdispVal, DISPID_VALUE, wFlags, pVarResult, -(int)pDispParams->cArgs, pDispParams->rgvarg);
					}
					VariantClear(&v);
					return S_OK;
				}
				ItemEx(dispIdMember - DISPID_COLLECTION_MIN, pVarResult, nArg >= 0 ? &pDispParams->rgvarg[nArg] : NULL);
				return S_OK;
			}
		}
	} catch (...) {
	}
	return DISP_E_MEMBERNOTFOUND;
}

//IDispatchEx
STDMETHODIMP CteArray::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
	return GetIDsOfNames(IID_NULL, &bstrName, 1, LOCALE_USER_DEFAULT, pid);
}

STDMETHODIMP CteArray::InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, EXCEPINFO *pei, IServiceProvider *pspCaller)
{
	return Invoke(id, IID_NULL, lcid, wFlags, pdp, pvarRes, pei, NULL);
}

STDMETHODIMP CteArray::DeleteMemberByName(BSTR bstrName, DWORD grfdex)
{
	DISPID id;
	HRESULT hr = teGetDispIdNum(bstrName, GetCount(), &id);
	if SUCCEEDED(hr) {
		return DeleteMemberByDispID(id);
	}
	return hr;
}

STDMETHODIMP CteArray::DeleteMemberByDispID(DISPID id)
{
	id -= DISPID_COLLECTION_MIN;
	if (id >= 0 && id < GetCount()) {
		VARIANT v;
		VariantInit(&v);
		return ::SafeArrayPutElement(m_psa, &id, &v);
	}
	return E_FAIL;
}

STDMETHODIMP CteArray::GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteArray::GetMemberName(DISPID id, BSTR *pbstrName)
{
	if (id >= DISPID_COLLECTION_MIN && id < DISPID_COLLECTION_MIN + GetCount()) {
		wchar_t pszName[8];
		swprintf_s(pszName, 8, L"%d", id - DISPID_COLLECTION_MIN);
		*pbstrName = ::SysAllocString(pszName);
		return S_OK;
	}
	for (auto itr = g_umArray.begin(); itr != g_umArray.end(); ++itr) {
		if (id == itr->second) {
			*pbstrName = ::SysAllocString(itr->first.data());
		}
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteArray::GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid)
{
	*pid = (id < DISPID_COLLECTION_MIN) ? DISPID_COLLECTION_MIN : id + 1;
	return *pid < GetCount() + DISPID_COLLECTION_MIN ? S_OK : S_FALSE;
}

STDMETHODIMP CteArray::GetNameSpaceParent(IUnknown **ppunk)
{
	return E_NOTIMPL;
}

LONG CteArray::GetCount()
{
	LONG lUBound, lLBound;
	SafeArrayGetUBound(m_psa, 1, &lUBound);
	SafeArrayGetLBound(m_psa, 1, &lLBound);
	return lUBound - lLBound + 1;
}

VOID CteArray::ItemEx(LONG nIndex, VARIANT *pVarResult, VARIANT *pVarNew)
{
	if (pVarNew) {
		if (nIndex < 0) {
			nIndex = GetCount();
		}
		if (nIndex >= GetCount()) {
			SAFEARRAYBOUND sab = { (ULONG)nIndex + 1, 0 };
			::SafeArrayRedim(m_psa, &sab);
		}
		::SafeArrayPutElement(m_psa, (LONG *)&nIndex, pVarNew);
	}
	if (pVarResult) {
		if (nIndex >= 0 && nIndex < GetCount()) {
			::SafeArrayGetElement(m_psa, (LONG *)&nIndex, pVarResult);
		}
	}
}

CteObjectEx::CteObjectEx()
{
	m_cRef = 1;
	m_dispId = DISPID_COLLECTION_MIN;
}

CteObjectEx::~CteObjectEx()
{
	for(auto itr = m_mData.begin(); itr != m_mData.end(); ++itr) {
		VariantClear(&itr->second);
	}
	m_umIndex.clear();
	m_mData.clear();
}

STDMETHODIMP CteObjectEx::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteObjectEx, IDispatch),
		QITABENT(CteObjectEx, IDispatchEx),
		{ 0 }
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteObjectEx::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteObjectEx::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteObjectEx::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteObjectEx::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteObjectEx::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	auto itr = m_umIndex.find(*rgszNames);
	if (itr == m_umIndex.end()) {
		*rgDispId = m_dispId;
		m_umIndex[*rgszNames] = m_dispId++;
		return  (m_dispId < TE_PROPERTY) ? S_OK : DISP_E_UNKNOWNNAME;
	}
	*rgDispId = itr->second;
	return S_OK;
}

STDMETHODIMP CteObjectEx::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	if (pVarResult) {
		VariantInit(pVarResult);
	}
	if (dispIdMember == DISPID_VALUE) {
		if (pDispParams->cArgs == 1 && pDispParams->rgvarg[0].vt == VT_BSTR) {
			auto itr2 = m_umIndex.find(pDispParams->rgvarg[0].bstrVal);
			if (itr2 != m_umIndex.end()) {
				auto itr = m_mData.find(itr2->second);
				if (itr != m_mData.end()) {
					VariantCopy(pVarResult, &itr->second);
				}
			}
			return S_OK;
		}
		teSetObject(pVarResult, this);
		return S_OK;
	}
	if (dispIdMember >= DISPID_COLLECTION_MIN && dispIdMember < TE_PROPERTY) {
		if (wFlags & DISPATCH_METHOD) {
			auto itr = m_mData.find(dispIdMember);
			if (itr != m_mData.end()) {
				if (itr->second.vt == VT_DISPATCH) {
					Invoke5(itr->second.pdispVal, DISPID_VALUE, wFlags, pVarResult, -(int)pDispParams->cArgs, pDispParams->rgvarg);
				}
			}
		} else if (wFlags & DISPATCH_PROPERTYGET) {
			auto itr = m_mData.find(dispIdMember);
			if (itr != m_mData.end()) {
				VariantCopy(pVarResult, &itr->second);
			}
		} else if (wFlags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) {
			int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
			if (nArg >= 0) {
				if (pDispParams->rgvarg[nArg].vt != VT_EMPTY) {
					VARIANT *pv = &m_mData[dispIdMember];
					VariantClear(pv);
					VariantCopy(pv, &pDispParams->rgvarg[nArg]);
				} else {
					DeleteMemberByDispID(dispIdMember);
				}
			}
		}
		return S_OK;
	}
	return DISP_E_MEMBERNOTFOUND;
}

//IDispatchEx
STDMETHODIMP CteObjectEx::GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid)
{
	return GetIDsOfNames(IID_NULL, &bstrName, 1, LOCALE_USER_DEFAULT, pid);
}

STDMETHODIMP CteObjectEx::InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, EXCEPINFO *pei, IServiceProvider *pspCaller)
{
	return Invoke(id, IID_NULL, lcid, wFlags, pdp, pvarRes, pei, NULL);
}

STDMETHODIMP CteObjectEx::DeleteMemberByName(BSTR bstrName, DWORD grfdex)
{
	auto itr = m_umIndex.find(bstrName);
	if (itr != m_umIndex.end()) {
		auto itr2 = m_mData.find(itr->second);
		if (itr2 != m_mData.end()) {
			VariantClear(&itr2->second);
		}
	}
	return S_OK;
}

STDMETHODIMP CteObjectEx::DeleteMemberByDispID(DISPID id)
{
	BSTR bstrName;
	if SUCCEEDED(GetMemberName(id, &bstrName)) {
		DeleteMemberByName(bstrName, fdexNameCaseSensitive);
		teSysFreeString(&bstrName);
	}
	return S_OK;
}

STDMETHODIMP CteObjectEx::GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteObjectEx::GetMemberName(DISPID id, BSTR *pbstrName)
{
	if (id >= DISPID_COLLECTION_MIN && id < TE_METHOD) {
		for(auto itr = m_umIndex.begin(); itr != m_umIndex.end(); ++itr) {
			if (id == itr->second) {
				*pbstrName = ::SysAllocString(itr->first.data());
				return S_OK;
			}
		}
	}
	return E_FAIL;
}

STDMETHODIMP CteObjectEx::GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid)
{
	auto itr = m_mData.find(id);
	if (itr == m_mData.end()) {
		if (id >= DISPID_COLLECTION_MIN) {
			return S_FALSE;
		}
		itr = m_mData.begin();
	} else {
		++itr;
	}
	if (itr != m_mData.end()) {
		while (itr->second.vt == VT_EMPTY) {
			if (++itr == m_mData.end()) {
				return S_FALSE;
			}
		}
		*pid = itr->first;
		return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CteObjectEx::GetNameSpaceParent(IUnknown **ppunk)
{
	return E_NOTIMPL;
}

//CteDispatch

CteDispatch::CteDispatch(IDispatch *pDispatch, int nMode, DISPID dispId)
{
	m_cRef = 1;
	pDispatch->QueryInterface(IID_PPV_ARGS(&m_pDispatch));
	m_dispIdMember = dispId;
}

CteDispatch::~CteDispatch()
{
	Clear();
}

VOID CteDispatch::Clear()
{
	SafeRelease(&m_pDispatch);
}

STDMETHODIMP CteDispatch::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteDispatch, IDispatch),
		{ 0 }
	};
	return QISearch(this, qit, riid, ppvObject);
}

STDMETHODIMP_(ULONG) CteDispatch::AddRef()
{
	return ::InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CteDispatch::Release()
{
	if (::InterlockedDecrement(&m_cRef) == 0) {
		delete this;
		return 0;
	}

	return m_cRef;
}

STDMETHODIMP CteDispatch::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

STDMETHODIMP CteDispatch::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteDispatch::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	return DISP_E_UNKNOWNNAME;
}

STDMETHODIMP CteDispatch::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	try {
		if (pVarResult) {
			VariantInit(pVarResult);
		}
		if (wFlags & DISPATCH_METHOD) {
			return m_pDispatch->Invoke(m_dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
		}
		teSetObject(pVarResult, this);
		return S_OK;
	} catch (...) {}
	return DISP_E_MEMBERNOTFOUND;
}
