rem xcopy /e /i .\packages\Microsoft.Web.WebView2.1.0.2210.55\build\native .\packages\Microsoft.Web.WebView2.1.0.2210.55\build
mkdir .\x64
mkdir .\x64\Debug
mkdir .\Debug
copy .\packages\Microsoft.Web.WebView2.1.0.2210.55\build\native\x64\WebView2Loader.dll .\x64\Debug
copy .\packages\Microsoft.Web.WebView2.1.0.2210.55\build\native\x86\WebView2Loader.dll .\Debug
