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
CteBase		*g_pBase = NULL;
LONG		g_lLocks = 0;

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

VOID teSetSZA(VARIANT *pv, LPCSTR lpstr, int nCP)
{
	if (pv) {
		int nLenW = MultiByteToWideChar(nCP, 0, lpstr, -1, NULL, NULL);
		if (nLenW) {
			pv->bstrVal = ::SysAllocStringLen(NULL, nLenW - 1);
			pv->bstrVal[0] = NULL;
			MultiByteToWideChar(nCP, 0, (LPCSTR)lpstr, -1, pv->bstrVal, nLenW);
		} else {
			pv->bstrVal = NULL;
		}
		pv->vt = VT_BSTR;
	}
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

BSTR GetMemoryFromVariant(VARIANT *pv, BOOL *pbDelete, LONG_PTR *pLen)
{
	if (pv->vt == (VT_VARIANT | VT_BYREF)) {
		return GetMemoryFromVariant(pv->pvarVal, pbDelete, pLen);
	}
	BSTR pMemory = NULL;
	*pbDelete = FALSE;
	if (pLen) {
		if (pv->vt == VT_BSTR || pv->vt == VT_LPWSTR) {
			return pv->bstrVal;
		}
	}
	IUnknown *punk;
	if (FindUnknown(pv, &punk)) {
		IStream *pStream;
		if SUCCEEDED(punk->QueryInterface(IID_PPV_ARGS(&pStream))) {
			ULARGE_INTEGER uliSize;
			if (pLen) {
				LARGE_INTEGER liOffset;
				liOffset.QuadPart = 0;
				pStream->Seek(liOffset, STREAM_SEEK_END, &uliSize);
				pStream->Seek(liOffset, STREAM_SEEK_SET, NULL);
			} else {
				uliSize.QuadPart = 2048;
			}
			pMemory = ::SysAllocStringByteLen(NULL, uliSize.LowPart > 2048 ? uliSize.LowPart : 2048);
			if (pMemory) {
				if (uliSize.LowPart < 2048) {
					::ZeroMemory(pMemory, 2048);
				}
				*pbDelete = TRUE;
				ULONG cbRead;
				pStream->Read(pMemory, uliSize.LowPart, &cbRead);
				if (pLen) {
					*pLen = cbRead;
				}
			}
			pStream->Release();
		}
	} else if (pv->vt == (VT_ARRAY | VT_I1) || pv->vt == (VT_ARRAY | VT_UI1) || pv->vt == (VT_ARRAY | VT_I1 | VT_BYREF) || pv->vt == (VT_ARRAY | VT_UI1 | VT_BYREF)) {
		LONG lUBound, lLBound, nSize;
		SAFEARRAY *psa = (pv->vt & VT_BYREF) ? pv->pparray[0] : pv->parray;
		PVOID pvData;
		if (::SafeArrayAccessData(psa, &pvData) == S_OK) {
			SafeArrayGetUBound(psa, 1, &lUBound);
			SafeArrayGetLBound(psa, 1, &lLBound);
			nSize = lUBound - lLBound + 1;
			pMemory = ::SysAllocStringByteLen(NULL, nSize > 2048 ? nSize : 2048);
			if (pMemory) {
				if (nSize < 2048) {
					::ZeroMemory(pMemory, 2048);
				}
				::CopyMemory(pMemory, pvData, nSize);
				if (pLen) {
					*pLen = nSize;
				}
				*pbDelete = TRUE;
			}
			::SafeArrayUnaccessData(psa);
		}
		return pMemory;
	} else if (!pLen) {
		return (BSTR)GetPtrFromVariant(pv);
	}
	return pMemory;
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
		g_pBase = new CteBase();
		g_hinstDll = hinstDll;
		break;
	case DLL_PROCESS_DETACH:
		SafeRelease(&g_pBase);
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
	m_bstrPath = ::SysAllocString(L"C:\\cpp\\TE\\Debug\\script\\blink.html");
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
	return m_pWebBrowser->PutProperty(Property, vtValue);
}

STDMETHODIMP CteBase::GetProperty(BSTR Property, VARIANT *pvtValue)
{
	return m_pWebBrowser->GetProperty(Property, pvtValue);
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
		CreateCoreWebView2EnvironmentWithOptions(NULL, NULL, NULL, this);
 		ShowWindow(hwndParent, SW_SHOWNORMAL);
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
	return g_pBase->QueryInterface(riid, ppvObject);
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
}

CteArray::~CteArray()
{
	while (!m_pArray.empty()) {
		VariantClear(&m_pArray.back());
		m_pArray.pop_back();
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
	return teGetDispIdNum(*rgszNames, m_pArray.size(), rgDispId);
}

STDMETHODIMP CteArray::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	VARIANT v;
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
			if (nArg >= 0 && wFlags == DISPATCH_PROPERTYPUT) {
				VARIANT v;
				VariantInit(&v);
				m_pArray.resize(GetIntFromVariant(&pDispParams->rgvarg[nArg]), v);
			}
			teSetLong(pVarResult, m_pArray.size());
			return S_OK;

		case TE_METHOD + 1://push
			for (int i = 0; i <= nArg; ++i) {
				VariantInit(&v);
				VariantCopy(&v, &pDispParams->rgvarg[nArg - i]);
				m_pArray.push_back(v);
			}
			return S_OK;

		case TE_METHOD + 2://pop
			if (!m_pArray.empty()) {
				if (pVarResult) {
					VariantCopy(pVarResult, &m_pArray.back());
				}
				m_pArray.pop_back();
			}
			return S_OK;

		case TE_METHOD + 3://shift
			if (!m_pArray.empty()) {
				if (pVarResult) {
					VariantCopy(pVarResult, &m_pArray.front());
				}
				m_pArray.erase(m_pArray.begin());
			}
			return S_OK;

		case TE_METHOD + 4://unshift
			for (int i = nArg; i >= 0; --i) {
				VariantInit(&v);
				VariantCopy(&v, &pDispParams->rgvarg[nArg - i]);
				m_pArray.insert(m_pArray.begin(), v);
			}
			return S_OK;

		case TE_METHOD + 5://join
			if (pVarResult) {
				UINT n = 0;
				VARIANT vSeparator, v;
				if (nArg >= 0) {
					teVariantChangeType(&vSeparator, &pDispParams->rgvarg[nArg], VT_BSTR);
				} else {
					teSetSZ(&vSeparator, L",");
				}
				UINT nSeparator = ::SysStringByteLen(vSeparator.bstrVal);
				for (size_t i = 0; i < m_pArray.size(); ++i) {
					if (i) {
						n += nSeparator;
					}
					teVariantChangeType(&v, &m_pArray[i], VT_BSTR);
					n += ::SysStringByteLen(v.bstrVal);
					VariantClear(&v);
				}
				pVarResult->vt = VT_BSTR;
				pVarResult->bstrVal = ::SysAllocStringByteLen(NULL, n);
				BYTE *p = (BYTE *)pVarResult->bstrVal;
				for (size_t i = 0; i < m_pArray.size(); ++i) {
					if (i) {
						CopyMemory(p, vSeparator.bstrVal, nSeparator);
						p += nSeparator;
					}
					teVariantChangeType(&v, &m_pArray[i], VT_BSTR);
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
					IDispatch *pdisp;
					HRESULT hr = E_FAIL;
					if (GetDispatch(&v, &pdisp)) {
						hr = Invoke5(pdisp, DISPID_VALUE, wFlags, pVarResult, - nArg - 1, pDispParams->rgvarg);
						pdisp->Release();
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
	HRESULT hr = teGetDispIdNum(bstrName, m_pArray.size(), &id);
	if SUCCEEDED(hr) {
		return DeleteMemberByDispID(id);
	}
	return hr;
}

STDMETHODIMP CteArray::DeleteMemberByDispID(DISPID id)
{
	id -= DISPID_COLLECTION_MIN;
	if (id >= 0 && id < (DISPID)m_pArray.size()) {
		VariantClear(&m_pArray[id]);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP CteArray::GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteArray::GetMemberName(DISPID id, BSTR *pbstrName)
{
	if (id >= DISPID_COLLECTION_MIN && id < DISPID_COLLECTION_MIN + (DISPID)m_pArray.size()) {
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
	return *pid < (DISPID)m_pArray.size() + DISPID_COLLECTION_MIN ? S_OK : S_FALSE;
}

STDMETHODIMP CteArray::GetNameSpaceParent(IUnknown **ppunk)
{
	return E_NOTIMPL;
}

VOID CteArray::ItemEx(int nIndex, VARIANT *pVarResult, VARIANT *pVarNew)
{
	if (pVarNew) {
		if (nIndex >= 0) {
			if (nIndex < (int)m_pArray.size()) {
				if (m_pArray[nIndex].vt != VT_EMPTY) {
					VariantClear(&m_pArray[nIndex]);
				}
			} else {
				VARIANT v;
				VariantInit(&v);
				m_pArray.resize(nIndex + 1, v);
			}
			VariantCopy(&m_pArray[nIndex], pVarNew);
		} else {
			VARIANT v;
			VariantCopy(&v, pVarNew);
			m_pArray.push_back(v);
		}
	}
	if (pVarResult) {
		if (nIndex >= 0 && nIndex < (int)m_pArray.size()) {
			VariantCopy(pVarResult, &m_pArray[nIndex]);
		}
	}
}

CteObjectEx::CteObjectEx()
{
	m_cRef = 1;
}

CteObjectEx::~CteObjectEx()
{
	for(std::unordered_map<std::wstring, VARIANT>::iterator itr = m_umObject.begin(); itr != m_umObject.end(); ++itr) {
		VariantClear(&itr->second);
	}
	m_umObject.clear();
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
	auto itr = m_umObject.find(*rgszNames);
	if (itr == m_umObject.end()) {
		VARIANT v;
		VariantInit(&v);
		m_umObject[*rgszNames] = v;
		itr = m_umObject.find(*rgszNames);
	}
	*rgDispId = std::distance(m_umObject.begin(), itr) + DISPID_COLLECTION_MIN;
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
	if (dispIdMember >= DISPID_COLLECTION_MIN && dispIdMember < DISPID_COLLECTION_MIN + (DISPID)m_umObject.size()) {
		auto itr = m_umObject.begin();
		std::advance(itr, dispIdMember - DISPID_COLLECTION_MIN);
		int nArg = pDispParams ? pDispParams->cArgs - 1 : -1;
		if (nArg >= 0) {
			VariantClear(&itr->second);
			if (pDispParams->rgvarg[nArg].vt != VT_EMPTY) {
				VariantCopy(&itr->second, &pDispParams->rgvarg[nArg]);
			} else {
				m_umObject.erase(itr);
				return S_OK;
			}
		} else if (itr->second.vt == VT_EMPTY) {
			m_umObject.erase(itr);
			return S_OK;
		}
		if (pVarResult) {
			if (itr != m_umObject.end()) {
				VariantCopy(pVarResult, &itr->second);
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
	auto itr = m_umObject.find(bstrName);
	if (itr != m_umObject.end()) {
		VariantClear(&itr->second);
		m_umObject.erase(itr);
	}
	return S_OK;
}

STDMETHODIMP CteObjectEx::DeleteMemberByDispID(DISPID id)
{
	if (id >= DISPID_COLLECTION_MIN && id < DISPID_COLLECTION_MIN + (DISPID)m_umObject.size()) {
		auto itr = m_umObject.begin();
		std::advance(itr, id - DISPID_COLLECTION_MIN);
		VariantClear(&itr->second);
		m_umObject.erase(itr);
	}
	return S_OK;
}

STDMETHODIMP CteObjectEx::GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex)
{
	return E_NOTIMPL;
}

STDMETHODIMP CteObjectEx::GetMemberName(DISPID id, BSTR *pbstrName)
{
	if (id >= DISPID_COLLECTION_MIN && id < DISPID_COLLECTION_MIN + (DISPID)m_umObject.size()) {
		auto itr = m_umObject.begin();
		std::advance(itr, id - DISPID_COLLECTION_MIN);
		*pbstrName = ::SysAllocString(itr->first.data());
		return S_OK;
	}
	return E_NOTIMPL;
}

STDMETHODIMP CteObjectEx::GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid)
{
	*pid = (id < DISPID_COLLECTION_MIN) ? DISPID_COLLECTION_MIN : id + 1;
	return *pid < (DISPID)m_umObject.size() + DISPID_COLLECTION_MIN ? S_OK : S_FALSE;
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
