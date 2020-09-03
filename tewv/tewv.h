#include "resource.h"
#include <windows.h>
#include <dispex.h>
#include <shlwapi.h>
#include <shlobj.h>
#include <process.h>
#include <vector>
#pragma comment (lib, "shlwapi.lib")

struct TEmethod
{
	LONG   id;
	LPWSTR name;
};

#ifdef _WIN64
#define teSetPtr(pVar, nData)	teSetLL(pVar, (LONGLONG)nData)
#define GetPtrFromVariant(pv)	GetLLFromVariant(pv)
#else
#define teSetPtr(pVar, nData)	teSetLong(pVar, (LONG)nData)
#define GetPtrFromVariant(pv)	GetIntFromVariant(pv)
#endif


// Base Object
class CteBase : public IWebBrowser2, public IOleObject, public IOleInPlaceObject
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

	CteBase();
	~CteBase();
private:
	LONG		m_cRef;
	IWebBrowser2 *m_pWebBrowser;
	IOleObject *m_pOleObject;
	IOleInPlaceObject *m_pOleInPlaceObject;
	IDispatch *m_pDispatch;
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
