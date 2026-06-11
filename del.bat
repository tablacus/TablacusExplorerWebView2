del tewv.ncb
attrib -h tewv.suo
del tewv.suo
del /q tewv\tewv.vcproj.*.user
del /q Debug\tewv*.*
del /q tewv\Debug\*
rmdir tewv\Debug
del /q Release\*
rmdir Release
del /q tewv\Release\*
rmdir tewv\Release
del /q /s x64\Release\*
rmdir x64\Release
del /q /s x64\Debug\*
rmdir x64\Debug
rmdir x64
del /q /s tewv\x64\*
rmdir tewv\x64\Release
rmdir tewv\x64\Debug
rmdir tewv\x64
del /q /s ARM64\Release\*
rmdir ARM64\Release
del /q /s ARM64\Debug\*
rmdir ARM64\Debug
rmdir ARM64
del /q /s tewv\ARM64\*
rmdir tewv\ARM64\Release
rmdir tewv\ARM64\Debug
rmdir tewv\ARM64
del /q /s *.sdf
rmdir /q /s ipch
rmdir /q /s Debug
rmdir /q /s .vs
rmdir /q /s packages
rmdir /q /s tewv\Debug
rmdir /q /s tewv\Release
rmdir /q /s tewv\x64
rmdir /q /s tewv\ARM64
