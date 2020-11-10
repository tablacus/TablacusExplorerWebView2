#include "resource.h"
#include <windows.h>
#include <dispex.h>
#include <shlwapi.h>
#include <shlobj.h>
#include <process.h>
#include <vector>
#include <map>
#include <unordered_map>
#include <mshtmdid.h>
#include <exdispid.h>
//#include <wrl.h>
#include <Mshtml.h>
#include <mshtmhst.h>
#include <wil/com.h>
#include "WebView2.h"
#pragma comment (lib, "shlwapi.lib")
//using namespace Microsoft::WRL;

#define TE_VT 24
#define TE_VI 0xffffff
#define TE_METHOD		0x60010000
#define TE_METHOD_MAX	0x6001ffff
#define TE_METHOD_MASK	0x0000ffff
#define TE_PROPERTY		0x40010000
#define START_OnFunc	0x4001fc00
#define TE_OFFSET		0x4001ff00
#define DISPID_TE_ITEM  0x6001ffff
#define DISPID_TE_COUNT 0x4001ffff
#define DISPID_TE_INDEX 0x4001fffe
#define DISPID_TE_MAX TE_VI
#define MAX_PATHEX				32768

//Tablacus Explorer (Edge)
const CLSID CLSID_WebBrowserExt             = {0x55bbf1b8, 0x0d30, 0x4908, { 0xbe, 0x0c, 0xd5, 0x76, 0x61, 0x2a, 0x0f, 0x48}};
// {BD34E79B-963F-4AFB-B03E-C5BD289B5080}
const IID SID_TablacusObject                = {0xbd34e79b, 0x963f, 0x4afb, { 0xb0, 0x3e, 0xc5, 0xbd, 0x28, 0x9b, 0x50, 0x80}};
// {A7A52B88-B449-47BB-BD92-ABCCD8A6FED7}
const IID SID_TablacusArray                 = {0xa7a52b88, 0xb449, 0x47bb, { 0xbd, 0x92, 0xab, 0xcc, 0xd8, 0xa6, 0xfe, 0xd7 }};

#ifdef _WIN64
#define teSetPtr(pVar, nData)	teSetLL(pVar, (LONGLONG)nData)
#define GetPtrFromVariant(pv)	GetLLFromVariant(pv)
#else
#define teSetPtr(pVar, nData)	teSetLong(pVar, (LONG)nData)
#define GetPtrFromVariant(pv)	GetIntFromVariant(pv)
#endif

typedef HRESULT (WINAPI * LPFNCreateCoreWebView2EnvironmentWithOptions)(PCWSTR browserExecutableFolder, PCWSTR userDataFolder, ICoreWebView2EnvironmentOptions* environmentOptions, ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler* environment_created_handler);
typedef HRESULT (WINAPI * LPFNGetAvailableCoreWebView2BrowserVersionString)(PCWSTR browserExecutableFolder, LPWSTR* versionInfo);

// Base Object
class CteBase : public IWebBrowser2, public IOleObject, public IOleInPlaceObject, public IServiceProvider,
	public IDropTarget,
	public ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler,
	public ICoreWebView2CreateCoreWebView2ControllerCompletedHandler,
	public ICoreWebView2ExecuteScriptCompletedHandler,
	public ICoreWebView2DocumentTitleChangedEventHandler,
	public ICoreWebView2NavigationStartingEventHandler,
	public ICoreWebView2NavigationCompletedEventHandler
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	//IWebBrowser
	STDMETHODIMP GoBack(void);
	STDMETHODIMP GoForward(void);
	STDMETHODIMP GoHome(void);
	STDMETHODIMP GoSearch(void);
	STDMETHODIMP Navigate(BSTR URL, VARIANT *Flags, VARIANT *TargetFrameName, VARIANT *PostData, VARIANT *Headers);
	STDMETHODIMP Refresh(void);
	STDMETHODIMP Refresh2(VARIANT *Level);
	STDMETHODIMP Stop(void);
	STDMETHODIMP get_Application(IDispatch **ppDisp);
	STDMETHODIMP get_Parent(IDispatch **ppDisp);
	STDMETHODIMP get_Container(IDispatch **ppDisp);
	STDMETHODIMP get_Document(IDispatch **ppDisp);
	STDMETHODIMP get_TopLevelContainer(VARIANT_BOOL *pBool);
	STDMETHODIMP get_Type(BSTR *Type);
	STDMETHODIMP get_Left(long *pl);
	STDMETHODIMP put_Left(long Left);
	STDMETHODIMP get_Top(long *pl);
	STDMETHODIMP put_Top(long Top);
	STDMETHODIMP get_Width(long *pl);
	STDMETHODIMP put_Width(long Width);
	STDMETHODIMP get_Height(long *pl);
	STDMETHODIMP put_Height(long Height);
	STDMETHODIMP get_LocationName(BSTR *LocationName);
	STDMETHODIMP get_LocationURL(BSTR *LocationURL);
	STDMETHODIMP get_Busy(VARIANT_BOOL *pBool);
	//IWebBrowserApp
	STDMETHODIMP Quit(void);
	STDMETHODIMP ClientToWindow(int *pcx, int *pcy);
	STDMETHODIMP PutProperty(BSTR Property, VARIANT vtValue);
	STDMETHODIMP GetProperty(BSTR Property, VARIANT *pvtValue);
	STDMETHODIMP get_Name(BSTR *Name);
	STDMETHODIMP get_HWND(SHANDLE_PTR *pHWND);
	STDMETHODIMP get_FullName(BSTR *FullName);
	STDMETHODIMP get_Path(BSTR *Path);
	STDMETHODIMP get_Visible(VARIANT_BOOL *pBool);
	STDMETHODIMP put_Visible(VARIANT_BOOL Value);
	STDMETHODIMP get_StatusBar(VARIANT_BOOL *pBool);
	STDMETHODIMP put_StatusBar(VARIANT_BOOL Value);
	STDMETHODIMP get_StatusText(BSTR *StatusText);
	STDMETHODIMP put_StatusText(BSTR StatusText);
	STDMETHODIMP get_ToolBar(int *Value);
	STDMETHODIMP put_ToolBar(int Value);
	STDMETHODIMP get_MenuBar(VARIANT_BOOL *Value);
	STDMETHODIMP put_MenuBar(VARIANT_BOOL Value);
	STDMETHODIMP get_FullScreen(VARIANT_BOOL *pbFullScreen);
	STDMETHODIMP put_FullScreen(VARIANT_BOOL bFullScreen);
	//IWebBrowser2
	STDMETHODIMP Navigate2(VARIANT *URL, VARIANT *Flags, VARIANT *TargetFrameName, VARIANT *PostData, VARIANT *Headers);
	STDMETHODIMP QueryStatusWB(OLECMDID cmdID, OLECMDF *pcmdf);
	STDMETHODIMP ExecWB(OLECMDID cmdID, OLECMDEXECOPT cmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut);
	STDMETHODIMP ShowBrowserBar(VARIANT *pvaClsid, VARIANT *pvarShow, VARIANT *pvarSize);
	STDMETHODIMP get_ReadyState(READYSTATE *plReadyState);
	STDMETHODIMP get_Offline(VARIANT_BOOL *pbOffline);
	STDMETHODIMP put_Offline(VARIANT_BOOL bOffline);
	STDMETHODIMP get_Silent(VARIANT_BOOL *pbSilent);
	STDMETHODIMP put_Silent(VARIANT_BOOL bSilent);
	STDMETHODIMP get_RegisterAsBrowser(VARIANT_BOOL *pbRegister);
	STDMETHODIMP put_RegisterAsBrowser(VARIANT_BOOL bRegister);
	STDMETHODIMP get_RegisterAsDropTarget(VARIANT_BOOL *pbRegister);
	STDMETHODIMP put_RegisterAsDropTarget(VARIANT_BOOL bRegister);
	STDMETHODIMP get_TheaterMode(VARIANT_BOOL *pbRegister);
	STDMETHODIMP put_TheaterMode(VARIANT_BOOL bRegister);
	STDMETHODIMP get_AddressBar(VARIANT_BOOL *Value);
	STDMETHODIMP put_AddressBar(VARIANT_BOOL Value);
	STDMETHODIMP get_Resizable(VARIANT_BOOL *Value);
	STDMETHODIMP put_Resizable(VARIANT_BOOL Value);
	//IOleObject
	STDMETHODIMP SetClientSite(IOleClientSite *pClientSite);
	STDMETHODIMP GetClientSite(IOleClientSite **ppClientSite);
	STDMETHODIMP SetHostNames(LPCOLESTR szContainerApp, LPCOLESTR szContainerObj);
	STDMETHODIMP Close(DWORD dwSaveOption);
	STDMETHODIMP SetMoniker(DWORD dwWhichMoniker, IMoniker *pmk);
	STDMETHODIMP GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk);
	STDMETHODIMP InitFromData(IDataObject *pDataObject, BOOL fCreation, DWORD dwReserved);
	STDMETHODIMP GetClipboardData(DWORD dwReserved, IDataObject **ppDataObject);
	STDMETHODIMP DoVerb(LONG iVerb, LPMSG lpmsg, IOleClientSite *pActiveSite, LONG lindex, HWND hwndParent, LPCRECT lprcPosRect);
	STDMETHODIMP EnumVerbs(IEnumOLEVERB **ppEnumOleVerb);
	STDMETHODIMP Update(void);
	STDMETHODIMP IsUpToDate(void);
	STDMETHODIMP GetUserClassID(CLSID *pClsid);
	STDMETHODIMP GetUserType(DWORD dwFormOfType, LPOLESTR *pszUserType);
	STDMETHODIMP SetExtent(DWORD dwDrawAspect, SIZEL *psizel);
	STDMETHODIMP GetExtent(DWORD dwDrawAspect, SIZEL *psizel);
	STDMETHODIMP Advise(IAdviseSink *pAdvSink, DWORD *pdwConnection);
	STDMETHODIMP Unadvise(DWORD dwConnection);
	STDMETHODIMP EnumAdvise(IEnumSTATDATA **ppenumAdvise);
	STDMETHODIMP GetMiscStatus(DWORD dwAspect, DWORD *pdwStatus);
	STDMETHODIMP SetColorScheme(LOGPALETTE *pLogpal);
	//IOleWindow
	STDMETHODIMP GetWindow(HWND *phwnd);
	STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);
	//IOleInPlaceObject
	STDMETHODIMP InPlaceDeactivate(void);
	STDMETHODIMP UIDeactivate(void);
	STDMETHODIMP SetObjectRects(LPCRECT lprcPosRect, LPCRECT lprcClipRect);
	STDMETHODIMP ReactivateAndUndo(void);
	//IDropTarget
	STDMETHODIMP DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);
	STDMETHODIMP DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);
	STDMETHODIMP DragLeave();
	STDMETHODIMP Drop(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);
	//IServiceProvider
	STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void **ppv);
	//ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
	STDMETHODIMP Invoke(HRESULT result, ICoreWebView2Environment *created_environment);
	//ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
	STDMETHODIMP Invoke(HRESULT result, ICoreWebView2Controller *createdController);
	//ICoreWebView2ExecuteScriptCompletedHandler
	STDMETHODIMP Invoke(HRESULT result, LPCWSTR resultObjectAsJson);
	//ICoreWebView2DocumentTitleChangedEventHandler
	STDMETHODIMP Invoke(ICoreWebView2* sender, IUnknown* args);
	//ICoreWebView2NavigationStartingEventHandler
	STDMETHODIMP Invoke(ICoreWebView2* sender, ICoreWebView2NavigationStartingEventArgs* args);
	//ICoreWebView2NavigationCompletedEventHandler
	STDMETHODIMP Invoke(ICoreWebView2* sender, ICoreWebView2NavigationCompletedEventArgs* args);

	CteBase();
	~CteBase();
private:
	IOleClientSite *m_pOleClientSite;
	IDispatch *m_pdisp;
	HWND m_hwndParent;
	wil::com_ptr<ICoreWebView2Controller> m_webviewController;
	wil::com_ptr<ICoreWebView2> m_webviewWindow;
	EventRegistrationToken m_documentTitleChangedToken;
	EventRegistrationToken m_navigationStartingToken;
	EventRegistrationToken m_navigationCompletedToken;
	BSTR m_bstrPath;

	IDispatch	*m_pDocument;
	LONG		m_cRef;
};

// Class Factory
class CteClassFactory : public IClassFactory
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	STDMETHODIMP CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject);
	STDMETHODIMP LockServer(BOOL fLock);
};

class CteArray : public IDispatchEx
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	//IDispatchEx
	STDMETHODIMP GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid);
	STDMETHODIMP InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, EXCEPINFO *pei, IServiceProvider *pspCaller);
	STDMETHODIMP DeleteMemberByName(BSTR bstrName, DWORD grfdex);
	STDMETHODIMP DeleteMemberByDispID(DISPID id);
	STDMETHODIMP GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex);
	STDMETHODIMP GetMemberName(DISPID id, BSTR *pbstrName);
	STDMETHODIMP GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid);
	STDMETHODIMP GetNameSpaceParent(IUnknown **ppunk);

	CteArray();
	~CteArray();

	VOID ItemEx(LONG nIndex, VARIANT *pVarResult, VARIANT *pVarNew);
	LONG GetCount();
private:
	SAFEARRAY *m_psa;
	LONG	m_cRef;
	BOOL	m_bDestroy;
};

class CteObjectEx : public IDispatchEx
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
	//IDispatchEx
	STDMETHODIMP GetDispID(BSTR bstrName, DWORD grfdex, DISPID *pid);
	STDMETHODIMP InvokeEx(DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pvarRes, EXCEPINFO *pei, IServiceProvider *pspCaller);
	STDMETHODIMP DeleteMemberByName(BSTR bstrName, DWORD grfdex);
	STDMETHODIMP DeleteMemberByDispID(DISPID id);
	STDMETHODIMP GetMemberProperties(DISPID id, DWORD grfdexFetch, DWORD *pgrfdex);
	STDMETHODIMP GetMemberName(DISPID id, BSTR *pbstrName);
	STDMETHODIMP GetNextDispID(DWORD grfdex, DISPID id, DISPID *pid);
	STDMETHODIMP GetNameSpaceParent(IUnknown **ppunk);

	CteObjectEx();
	~CteObjectEx();
private:
	std::unordered_map<std::wstring, DISPID>	m_umIndex;
	std::map<DISPID, VARIANT>	m_mData;
	LONG	m_cRef;
	DISPID m_dispId;
};

class CteDispatch : public IDispatch
{
public:
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	//IDispatch
	STDMETHODIMP GetTypeInfoCount(UINT *pctinfo);
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

	CteDispatch(IDispatch *pDispatch, int nMode, DISPID dispId);
	~CteDispatch();

	VOID Clear();
public:
	DISPID		m_dispIdMember;
private:
	IDispatch	*m_pDispatch;
	LONG		m_cRef;
};
