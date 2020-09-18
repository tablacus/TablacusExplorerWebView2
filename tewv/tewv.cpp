// Tablacus WebView2 (C)2020 Gaku
// MIT Lisence
// Visual Studio Express 2017 for Windows Desktop
// https://tablacus.github.io/

#include "tewv.h"

// Global Variables:
const TCHAR g_szProgid[] = TEXT("Tablacus.WebView2");
const TCHAR g_szClsid[] = TEXT("{55BBF1B8-0D30-4908-BE0C-D576612A0F48}");
std::vector<DWORD>	g_pIconOverlayHandlers;
HINSTANCE	g_hinstDll = NULL;
HMODULE		g_hWebView2Loader = NULL;
LONG		g_lLocks = 0;
LPFNCreateCoreWebView2EnvironmentWithOptions _CreateCoreWebView2EnvironmentWithOptions = NULL;

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

#ifdef _WIN64
LONGLONG GetLLFromVariant(VARIANT *pv)
{
	if (pv) {
		if (pv->vt == (VT_VARIANT | VT_BYREF)) {
			return GetLLFromVariant(pv->pvarVal);
		}
		if (pv->vt == VT_I4) {
			return pv->lVal;
		}
		if (pv->vt == VT_R8) {
			return (LONGLONG)pv->dblVal;
		}
		if (pv->vt == (VT_ARRAY | VT_I4)) {
			LONGLONG ll = 0;
			PVOID pvData;
			if (::SafeArrayAccessData(pv->parray, &pvData) == S_OK) {
				::CopyMemory(&ll, pvData, sizeof(LONGLONG));
				::SafeArrayUnaccessData(pv->parray);
				return ll;
			}
		}
		VARIANT vo;
		VariantInit(&vo);
		if SUCCEEDED(VariantChangeType(&vo, pv, 0, VT_I8)) {
			return vo.llVal;
		}
	}
	return 0;
}
#endif

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

HRESULT teGetProperty(IDispatch *pdisp, LPOLESTR sz, VARIANT *pv)
{
	DISPID dispid;
	HRESULT hr = pdisp->GetIDsOfNames(IID_NULL, &sz, 1, LOCALE_USER_DEFAULT, &dispid);
	if (hr == S_OK) {
		hr = Invoke5(pdisp, dispid, DISPATCH_PROPERTYGET, pv, 0, NULL);
	}
	return hr;
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
		g_hWebView2Loader = LoadLibrary(L"WebView2Loader.dll");
		if (!g_hWebView2Loader) {
			WCHAR pszPath[MAX_PATH * 2];
			g_hinstDll = hinstDll;
			GetModuleFileName(hinstDll, pszPath, MAX_PATH * 2);
			lstrcpy(PathFindFileName(pszPath), L"WebView2Loader.dll");
			g_hWebView2Loader = LoadLibrary(pszPath);
		}
		if (g_hWebView2Loader) {
			LPFNGetAvailableCoreWebView2BrowserVersionString _GetAvailableCoreWebView2BrowserVersionString = NULL;
			*(FARPROC *)&_GetAvailableCoreWebView2BrowserVersionString = GetProcAddress(g_hWebView2Loader, "GetAvailableCoreWebView2BrowserVersionString");
			if (_GetAvailableCoreWebView2BrowserVersionString) {
				LPWSTR versionInfo;
				if (_GetAvailableCoreWebView2BrowserVersionString(NULL, &versionInfo) == S_OK) {
					CoTaskMemFree(versionInfo);
					*(FARPROC *)&_CreateCoreWebView2EnvironmentWithOptions = GetProcAddress(g_hWebView2Loader, "CreateCoreWebView2EnvironmentWithOptions");
				}
			}
		}
		break;
	case DLL_PROCESS_DETACH:
		if (g_hWebView2Loader) {
			FreeLibrary(g_hWebView2Loader);
		}
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
	CoCreateInstance(CLSID_WebBrowser, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&m_pWebBrowser));
	m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&m_pOleObject));
	m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&m_pOleInPlaceObject));
	m_pWebBrowser->QueryInterface(IID_PPV_ARGS(&m_pDispatch));

	m_pOleClientSite = NULL;
	m_bstrPath = NULL;
}

CteBase::~CteBase()
{
	teSysFreeString(&m_bstrPath);
	SafeRelease(&m_pOleClientSite);


	SafeRelease(&m_pDispatch);
	SafeRelease(&m_pOleInPlaceObject);
	SafeRelease(&m_pOleObject);
	SafeRelease(&m_pWebBrowser);
}

STDMETHODIMP CteBase::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteBase, IDispatch),
		QITABENT(CteBase, IWebBrowser),
		QITABENT(CteBase, IWebBrowserApp),
		QITABENT(CteBase, IWebBrowser2),
		QITABENT(CteBase, IOleObject),
		QITABENT(CteBase, IOleWindow),
		QITABENT(CteBase, IOleInPlaceObject),
		QITABENT(CteBase, IServiceProvider),
		QITABENT(CteBase, ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler),
		QITABENT(CteBase, ICoreWebView2CreateCoreWebView2ControllerCompletedHandler),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppvObject);
/*	HRESULT hr = QISearch(this, qit, riid, ppvObject);
	HRESULT hr = E_FAIL;
	return SUCCEEDED(hr) ? hr : m_pWebBrowser->QueryInterface(riid, ppvObject);*/
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

	OutputDebugStringA("GetIDsOfNames:");
	OutputDebugString(rgszNames[0]);
	OutputDebugStringA("\n");
	return m_pDispatch->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
}

STDMETHODIMP CteBase::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	return m_pDispatch->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);

	int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
	HRESULT hr = S_OK;
	try {
		switch (dispIdMember) {

		case 0x60010000://Init
			return S_OK;

		case 0x60010001://Finalize
			return S_OK;

		//this
		case DISPID_VALUE:
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
	return m_pWebBrowser->GoBack();
}

STDMETHODIMP CteBase::GoForward(void)
{
	return m_pWebBrowser->GoForward();
}

STDMETHODIMP CteBase::GoHome(void)
{
	return m_pWebBrowser->GoHome();
}

STDMETHODIMP CteBase::GoSearch(void)
{
	return m_pWebBrowser->GoSearch();
}

STDMETHODIMP CteBase::Navigate(BSTR URL, VARIANT *Flags, VARIANT *TargetFrameName, VARIANT *PostData, VARIANT *Headers)
{
	if (m_webviewWindow) {
		return E_NOTIMPL;
		return m_webviewWindow->Navigate(URL);
	} 
	teSysFreeString(&m_bstrPath);
//	m_bstrPath = ::SysAllocString(URL);
	m_bstrPath = ::SysAllocString(L"C:\\cpp\\TE\\Debug\\script\\index.html");
	return S_OK;
}

STDMETHODIMP CteBase::Refresh(void)
{
	return m_pWebBrowser->Refresh();
}

STDMETHODIMP CteBase::Refresh2(VARIANT *Level)
{
	return m_pWebBrowser->Refresh2(Level);
}

STDMETHODIMP CteBase::Stop(void)
{
	return m_pWebBrowser->Stop();
}

STDMETHODIMP CteBase::get_Application(IDispatch **ppDisp)
{
	return m_pWebBrowser->get_Application(ppDisp);
}

STDMETHODIMP CteBase::get_Parent(IDispatch **ppDisp)
{
	return m_pWebBrowser->get_Parent(ppDisp);
}

STDMETHODIMP CteBase::get_Container(IDispatch **ppDisp)
{
	return m_pWebBrowser->get_Container(ppDisp);
}

STDMETHODIMP CteBase::get_Document(IDispatch **ppDisp)
{
	return m_pWebBrowser->get_Document(ppDisp);
}

STDMETHODIMP CteBase::get_TopLevelContainer(VARIANT_BOOL *pBool)
{
	return m_pWebBrowser->get_TopLevelContainer(pBool);
}

STDMETHODIMP CteBase::get_Type(BSTR *Type)
{
	return m_pWebBrowser->get_Type(Type);
}

STDMETHODIMP CteBase::get_Left(long *pl)
{
	return m_pWebBrowser->get_Left(pl);
}

STDMETHODIMP CteBase::put_Left(long Left)
{
	return m_pWebBrowser->put_Left(Left);
}

STDMETHODIMP CteBase::get_Top(long *pl)
{
	return m_pWebBrowser->get_Top(pl);
}

STDMETHODIMP CteBase::put_Top(long Top)
{
	return m_pWebBrowser->put_Top(Top);
}

STDMETHODIMP CteBase::get_Width(long *pl)
{
	return m_pWebBrowser->get_Width(pl);
}

STDMETHODIMP CteBase::put_Width(long Width)
{
	return m_pWebBrowser->put_Width(Width);
}

STDMETHODIMP CteBase::get_Height(long *pl)
{
	return m_pWebBrowser->get_Height(pl);
}

STDMETHODIMP CteBase::put_Height(long Height)
{
	return m_pWebBrowser->put_Height(Height);
}

STDMETHODIMP CteBase::get_LocationName(BSTR *LocationName)
{
	return m_pWebBrowser->get_LocationName(LocationName);
}

STDMETHODIMP CteBase::get_LocationURL(BSTR *LocationURL)
{
	return m_pWebBrowser->get_LocationURL(LocationURL);
}

STDMETHODIMP CteBase::get_Busy(VARIANT_BOOL *pBool)
{
	return m_pWebBrowser->get_Busy(pBool);
}

//IWebBrowserApp
STDMETHODIMP CteBase::Quit(void)
{
	return m_pWebBrowser->Quit();
}

STDMETHODIMP CteBase::ClientToWindow(int *pcx, int *pcy)
{
	return m_pWebBrowser->ClientToWindow(pcx, pcy);
}

STDMETHODIMP CteBase::PutProperty(BSTR Property, VARIANT vtValue)
{
	return m_webviewWindow->AddHostObjectToScript(Property, &vtValue);
}

STDMETHODIMP CteBase::GetProperty(BSTR Property, VARIANT *pvtValue)
{
	HRESULT hr = m_webviewWindow->ExecuteScript(Property, this);
	teSetLong(pvtValue, hr);
	return hr;
}

STDMETHODIMP CteBase::get_Name(BSTR *Name)
{
	return m_pWebBrowser->get_Name(Name);
}

STDMETHODIMP CteBase::get_HWND(SHANDLE_PTR *pHWND)
{
	return m_pWebBrowser->get_HWND(pHWND);
}

STDMETHODIMP CteBase::get_FullName(BSTR *FullName)
{
	return m_pWebBrowser->get_FullName(FullName);
}

STDMETHODIMP CteBase::get_Path(BSTR *Path)
{
	return m_pWebBrowser->get_Path(Path);
}

STDMETHODIMP CteBase::get_Visible(VARIANT_BOOL *pBool)
{
	return m_pWebBrowser->get_Visible(pBool);
}

STDMETHODIMP CteBase::put_Visible(VARIANT_BOOL Value)
{
	return m_pWebBrowser->put_Visible(Value);
}

STDMETHODIMP CteBase::get_StatusBar(VARIANT_BOOL *pBool)
{
	return m_pWebBrowser->get_StatusBar(pBool);
}

STDMETHODIMP CteBase::put_StatusBar(VARIANT_BOOL Value)
{
	return m_pWebBrowser->put_StatusBar(Value);
}

STDMETHODIMP CteBase::get_StatusText(BSTR *StatusText)
{
	return m_pWebBrowser->get_StatusText(StatusText);
}

STDMETHODIMP CteBase::put_StatusText(BSTR StatusText)
{
	return m_pWebBrowser->put_StatusText(StatusText);
}

STDMETHODIMP CteBase::get_ToolBar(int *Value)
{
	return m_pWebBrowser->get_ToolBar(Value);
}

STDMETHODIMP CteBase::put_ToolBar(int Value)
{
	return m_pWebBrowser->put_ToolBar(Value);
}

STDMETHODIMP CteBase::get_MenuBar(VARIANT_BOOL *Value)
{
	return m_pWebBrowser->get_MenuBar(Value);
}

STDMETHODIMP CteBase::put_MenuBar(VARIANT_BOOL Value)
{
	return m_pWebBrowser->put_MenuBar(Value);
}

STDMETHODIMP CteBase::get_FullScreen(VARIANT_BOOL *pbFullScreen)
{
	return m_pWebBrowser->get_FullScreen(pbFullScreen);
}

STDMETHODIMP CteBase::put_FullScreen(VARIANT_BOOL bFullScreen)
{
	return m_pWebBrowser->put_FullScreen(bFullScreen);
}
//IWebBrowser2
STDMETHODIMP CteBase::Navigate2(VARIANT *URL, VARIANT *Flags, VARIANT *TargetFrameName, VARIANT *PostData, VARIANT *Headers)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteBase::QueryStatusWB(OLECMDID cmdID, OLECMDF *pcmdf)
{
	return m_pWebBrowser->QueryStatusWB(cmdID, pcmdf);
}

STDMETHODIMP CteBase::ExecWB(OLECMDID cmdID, OLECMDEXECOPT cmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut)
{
	return m_pWebBrowser->ExecWB(cmdID, cmdexecopt, pvaIn, pvaOut);
}

STDMETHODIMP CteBase::ShowBrowserBar(VARIANT *pvaClsid, VARIANT *pvarShow, VARIANT *pvarSize)
{
	return m_pWebBrowser->ShowBrowserBar(pvaClsid, pvarShow, pvarSize);
}

STDMETHODIMP CteBase::get_ReadyState(READYSTATE *plReadyState)
{
	return m_pWebBrowser->get_ReadyState(plReadyState);
}

STDMETHODIMP CteBase::get_Offline(VARIANT_BOOL *pbOffline)
{
	return m_pWebBrowser->get_Offline(pbOffline);
}

STDMETHODIMP CteBase::put_Offline(VARIANT_BOOL bOffline)
{
	return m_pWebBrowser->put_Offline(bOffline);
}

STDMETHODIMP CteBase::get_Silent(VARIANT_BOOL *pbSilent)
{
	return m_pWebBrowser->get_Silent(pbSilent);
}

STDMETHODIMP CteBase::put_Silent(VARIANT_BOOL bSilent)
{
	return m_pWebBrowser->put_Silent(bSilent);
}

STDMETHODIMP CteBase::get_RegisterAsBrowser(VARIANT_BOOL *pbRegister)
{
	return m_pWebBrowser->get_RegisterAsBrowser(pbRegister);
}

STDMETHODIMP CteBase::put_RegisterAsBrowser(VARIANT_BOOL bRegister)
{
	return m_pWebBrowser->put_RegisterAsBrowser(bRegister);
}

STDMETHODIMP CteBase::get_RegisterAsDropTarget(VARIANT_BOOL *pbRegister)
{
	return m_pWebBrowser->get_RegisterAsDropTarget(pbRegister);
}

STDMETHODIMP CteBase::put_RegisterAsDropTarget(VARIANT_BOOL bRegister)
{
	return m_pWebBrowser->put_RegisterAsDropTarget(bRegister);
}

STDMETHODIMP CteBase::get_TheaterMode(VARIANT_BOOL *pbRegister)
{
	return m_pWebBrowser->get_TheaterMode(pbRegister);
}

STDMETHODIMP CteBase::put_TheaterMode(VARIANT_BOOL bRegister)
{
	return m_pWebBrowser->put_TheaterMode(bRegister);
}

STDMETHODIMP CteBase::get_AddressBar(VARIANT_BOOL *Value)
{
	return m_pWebBrowser->get_AddressBar(Value);
}

STDMETHODIMP CteBase::put_AddressBar(VARIANT_BOOL Value)
{
	return m_pWebBrowser->put_AddressBar(Value);
}
STDMETHODIMP CteBase::get_Resizable(VARIANT_BOOL *Value)
{
	return m_pWebBrowser->get_Resizable(Value);
}

STDMETHODIMP CteBase::put_Resizable(VARIANT_BOOL Value)
{
	return m_pWebBrowser->put_Resizable(Value);
}

//IOleObject
STDMETHODIMP CteBase::SetClientSite(IOleClientSite *pClientSite)
{
	SafeRelease(&m_pOleClientSite);
	if (pClientSite) {
		pClientSite->QueryInterface(IID_PPV_ARGS(&m_pOleClientSite));
	}
	return m_pOleObject->SetClientSite(pClientSite);
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
		if (_CreateCoreWebView2EnvironmentWithOptions) {
			_CreateCoreWebView2EnvironmentWithOptions(NULL, NULL, NULL, this);
			ShowWindow(hwndParent, SW_SHOWNORMAL);
			return S_OK;
		}
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
	return m_pOleObject->Advise(pAdvSink, pdwConnection);
}

STDMETHODIMP CteBase::Unadvise(DWORD dwConnection)
{
	return m_pOleObject->Unadvise(dwConnection);
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
	*phwnd = FindWindowA("Chrome_WidgetWin_0", NULL);
	return *phwnd ? S_OK : E_FAIL;
}

STDMETHODIMP CteBase::ContextSensitiveHelp(BOOL fEnterMode)
{
	return m_pOleInPlaceObject->ContextSensitiveHelp(fEnterMode);
}

//IOleInPlaceObject
STDMETHODIMP CteBase::InPlaceDeactivate(void)
{
	return m_pOleInPlaceObject->InPlaceDeactivate();
}

STDMETHODIMP CteBase::UIDeactivate(void)
{
	return m_pOleInPlaceObject->UIDeactivate();
}

STDMETHODIMP CteBase::SetObjectRects(LPCRECT lprcPosRect, LPCRECT lprcClipRect)
{
	return m_webviewController->put_Bounds(*lprcClipRect);
}

STDMETHODIMP CteBase::ReactivateAndUndo(void)
{
	return m_pOleInPlaceObject->ReactivateAndUndo();
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
	m_webviewController = createdController;
	m_webviewController->get_CoreWebView2(&m_webviewWindow);
	ICoreWebView2Settings* Settings;
	m_webviewWindow->get_Settings(&Settings);
	Settings->put_IsScriptEnabled(TRUE);
	Settings->put_AreDefaultScriptDialogsEnabled(TRUE);
	Settings->put_IsWebMessageEnabled(TRUE);
	if (m_pOleClientSite) {
		IDocHostUIHandler *pDocHostUIHandler;
		if SUCCEEDED(m_pOleClientSite->QueryInterface(IID_PPV_ARGS(&pDocHostUIHandler))) {
			VARIANT v;
			if SUCCEEDED(pDocHostUIHandler->GetExternal(&v.pdispVal)) {
				v.vt = VT_DISPATCH;
				HRESULT hr = m_webviewWindow->AddHostObjectToScript(L"te", &v);
				Sleep(hr);
			}
			pDocHostUIHandler->Release();
		}
	}
	RECT bounds;
	GetClientRect(m_hwndParent, &bounds);
	m_webviewController->put_Bounds(bounds);
	if (m_bstrPath) {
		m_webviewWindow->Navigate(m_bstrPath);
	}
	return S_OK;
}

//ICoreWebView2ExecuteScriptCompletedHandler
STDMETHODIMP CteBase::Invoke(HRESULT result, LPCWSTR resultObjectAsJson)
{
	return S_OK;
}


// CteClassFactory

STDMETHODIMP CteClassFactory::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteClassFactory, IClassFactory),
		{ 0 },
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
	if (!_CreateCoreWebView2EnvironmentWithOptions) {
		return CLASS_E_CLASSNOTAVAILABLE;
	}
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
	m_bDestroy = TRUE;
}

CteArray::~CteArray()
{
	if (m_bDestroy) {
		SafeArrayDestroy(m_psa);
	}
}

STDMETHODIMP CteArray::QueryInterface(REFIID riid, void **ppvObject)
{
	static const QITAB qit[] =
	{
		QITABENT(CteArray, IDispatch),
		QITABENT(CteArray, IDispatchEx),
		{ 0 },
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
/*
		case TE_METHOD + 6://Slice
		case TE_METHOD + 7://Splice
			CteArray *pNewArray;
			if (pVarResult) {
				pNewArray = new CteArray();
			}
			if (nArg >= 0) {
				UINT nStart = GetIntFromVariant(&pDispParams->rgvarg[nArg]);
				size_t nLen = nArg >= 1 ? GetIntFromVariant(&pDispParams->rgvarg[nArg - 1]) : MAXINT;
				if (nStart + nLen > m_pArray.size()) {
					nLen = m_pArray.size() - nStart;
				}
				VARIANT v;
				VariantInit(&v);
				if (pVarResult) {
					for (UINT i = nStart; i < nLen; ++i) {
						ItemEx(i, &v, NULL);
						pNewArray->ItemEx(-1, NULL, &v);
					}
					teSetObjectRelease(pVarResult, pNewArray);
				}
				if (dispIdMember == TE_METHOD + 7) {//Splice
					m_pArray.erase(m_pArray.begin() + nStart, m_pArray.begin() + nStart + nLen);
				}
			}
			return S_OK;
*/
		case DISPID_PROPERTYPUT:
			if ((wFlags & DISPATCH_PROPERTYPUT) && nArg >= 0) {
				if (pDispParams->rgvarg[nArg].vt == (VT_ARRAY | VT_VARIANT)) {
					if (m_bDestroy) {
						m_bDestroy = FALSE;
						SafeArrayDestroy(m_psa);
						m_psa = pDispParams->rgvarg[nArg].parray;
					}
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
		{ 0 },
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
		} else if (wFlags & DISPATCH_PROPERTYPUT) {
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
		{ 0 },
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