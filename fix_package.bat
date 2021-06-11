rem xcopy /e /i .\packages\Microsoft.Web.WebView2.1.0.864.35\build\native .\packages\Microsoft.Web.WebView2.1.0.864.35\build
mkdir .\x64
mkdir .\x64\Debug
mkdir .\Debug
copy .\packages\Microsoft.Web.WebView2.1.0.864.35\build\x64\WebView2Loader.dll .\x64\Debug
copy .\packages\Microsoft.Web.WebView2.1.0.864.35\build\x86\WebView2Loader.dll .\Debug
