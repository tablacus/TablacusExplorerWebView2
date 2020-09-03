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

TEmethod methodBASE[] = {
	{ 0x60010000, L"Init" },
	{ 0x60010001, L"Finalize" },
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

/*
int teBSearch(TEmethod *method, int nSize, int* pMap, LPOLESTR bs)
{
	int nMin = 0;
	int nMax = nSize - 1;
	int nIndex, nCC;

	while (nMin <= nMax) {
		nIndex = (nMin + nMax) / 2;
		nCC = lstrcmpi(bs, method[pMap[nIndex]].name);
		if (nCC < 0) {
			nMax = nIndex - 1;
			continue;
		}
		if (nCC > 0) {
			nMin = nIndex + 1;
			continue;
		}
		return pMap[nIndex];
	}
	return -1;
}
*/
HRESULT teGetDispId(TEmethod *method, int nCount, int* pMap, LPOLESTR bs, DISPID *rgDispId)
{
/*	if (pMap) {
		int nIndex = teBSearch(method, nCount, pMap, bs);
		if (nIndex >= 0) {
			*rgDispId = method[nIndex].id;
			return S_OK;
		}
	} else {*/
		for (int i = 0; method[i].name; i++) {
			if (lstrcmpi(bs, method[i].name) == 0) {
				*rgDispId = method[i].id;
				return S_OK;
			}
		}
//	}
	return DISP_E_UNKNOWNNAME;
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
	return false;
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
}

CteBase::~CteBase()
{
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
//	return teGetDispId(methodBASE, _countof(methodBASE), NULL, *rgszNames, rgDispId);
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
	return m_pWebBrowser->Navigate(URL, Flags, TargetFrameName, PostData, Headers);
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
	return m_pWebBrowser->Navigate2(URL, Flags, TargetFrameName, PostData, Headers);
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
	return m_pOleObject->SetClientSite(pClientSite);
}

STDMETHODIMP CteBase::GetClientSite(IOleClientSite **ppClientSite)
{
	return m_pOleObject->GetClientSite(ppClientSite);
}

STDMETHODIMP CteBase::SetHostNames(LPCOLESTR szContainerApp, LPCOLESTR szContainerObj)
{
	return m_pOleObject->SetHostNames(szContainerApp, szContainerObj);
}

STDMETHODIMP CteBase::Close(DWORD dwSaveOption)
{
	return m_pOleObject->Close(dwSaveOption);
}

STDMETHODIMP CteBase::SetMoniker(DWORD dwWhichMoniker, IMoniker *pmk)
{
	return m_pOleObject->SetMoniker(dwWhichMoniker, pmk);
}

STDMETHODIMP CteBase::GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk)
{
	return m_pOleObject->GetMoniker(dwAssign, dwWhichMoniker, ppmk);
}

STDMETHODIMP CteBase::InitFromData(IDataObject *pDataObject, BOOL fCreation, DWORD dwReserved)
{
	return m_pOleObject->InitFromData(pDataObject, fCreation, dwReserved);
}

STDMETHODIMP CteBase::GetClipboardData(DWORD dwReserved, IDataObject **ppDataObject)
{
	return m_pOleObject->GetClipboardData(dwReserved, ppDataObject);
}

STDMETHODIMP CteBase::DoVerb(LONG iVerb, LPMSG lpmsg, IOleClientSite *pActiveSite, LONG lindex, HWND hwndParent, LPCRECT lprcPosRect)
{
	return m_pOleObject->DoVerb(iVerb, lpmsg, pActiveSite, lindex, hwndParent, lprcPosRect);
}

STDMETHODIMP CteBase::EnumVerbs(IEnumOLEVERB **ppEnumOleVerb)
{
	return m_pOleObject->EnumVerbs(ppEnumOleVerb);
}

STDMETHODIMP CteBase::Update(void)
{
	return m_pOleObject->Update();
}

STDMETHODIMP CteBase::IsUpToDate(void)
{
	return m_pOleObject->IsUpToDate();
}

STDMETHODIMP CteBase::GetUserClassID(CLSID *pClsid)
{
	return m_pOleObject->GetUserClassID(pClsid);
}

STDMETHODIMP CteBase::GetUserType(DWORD dwFormOfType, LPOLESTR *pszUserType)
{
	return m_pOleObject->GetUserType(dwFormOfType, pszUserType);
}

STDMETHODIMP CteBase::SetExtent(DWORD dwDrawAspect, SIZEL *psizel)
{
	return m_pOleObject->SetExtent(dwDrawAspect, psizel);
}

STDMETHODIMP CteBase::GetExtent(DWORD dwDrawAspect, SIZEL *psizel)
{
	return m_pOleObject->GetExtent(dwDrawAspect, psizel);
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
	return m_pOleObject->EnumAdvise(ppenumAdvise);
}

STDMETHODIMP CteBase::GetMiscStatus(DWORD dwAspect, DWORD *pdwStatus)
{
	return m_pOleObject->GetMiscStatus(dwAspect, pdwStatus);
}

STDMETHODIMP CteBase::SetColorScheme(LOGPALETTE *pLogpal)
{
	return m_pOleObject->SetColorScheme(pLogpal);
}

//IOleWindow
STDMETHODIMP CteBase::GetWindow(HWND *phwnd)
{
	return m_pOleInPlaceObject->GetWindow(phwnd);
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
	return m_pOleInPlaceObject->SetObjectRects(lprcPosRect, lprcClipRect);
}

STDMETHODIMP CteBase::ReactivateAndUndo(void)
{
	return m_pOleInPlaceObject->ReactivateAndUndo();
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
