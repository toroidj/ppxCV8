@echo off
set RETAIL=1
rem *** set value ***
set arcname=ppxCV8R04.zip
set readme=PPxCV8.TXT
set srcname=source.7z
set exename=PPxCV8

rem *** main ***
rd /s /q archive
md archive
md archive\x64
md archive\x86
md archive\licenses
md archive\samples

rem nuget.exe install Microsoft.ClearScript.V8
rem nuget.exe install Microsoft.ClearScript.V8.Native.win-x86
rem nuget.exe install Microsoft.ClearScript.V8.Native.win-x64

msbuild PPxCV8.vcxproj -property:Platform=Win32 -property:Configuration=Release -t:Rebuild
xcopy .\Win32\Release\*.dll .\archive\x86
tfilesign sp archive\x86\%exename%.dll archive\x86\%exename%.dll
CT %readme% archive\x86\%exename%.dll

msbuild PPxCV8.vcxproj -property:Platform=x64 -property:Configuration=Release -t:Rebuild
xcopy .\x64\Release\*.dll .\archive\x64
tfilesign sp archive\x64\%exename%64.dll archive\x64\%exename%64.dll
CT %readme% archive\x64\%exename%64.dll

xcopy %readme% .\archive\
xcopy /e licenses\* archive\licenses\

rem *** Source Archive ***

if %RETAIL%==0 goto :skipsource

xcopy samples\* archive\samples\
for %%i in (archive\samples\*) do CT %readme% %%i

for %%i in (*) do CT %readme% %%i
ppb /c %%u/7-zip32.dll,a archive\%srcname% -hide -mx=9 archive.BAT *.C *.CPP *.DEF *.H *.RC *.RH *.sln *.vcxproj
CT %readme% archive\%srcname%
if %RETAIL%==1 goto :archive

:skipsource
del archive\x64\ClearScript*.dll
del archive\x64\PPlib*.dll
del archive\x86\ClearScript*.dll
del archive\x86\PPlib*.dll

:archive
cd archive
ppb /c %%u/7-ZIP32.DLL,a -tzip -hide -mx=7 ..\%arcname% *
cd ..
tfilesign s %arcname% %arcname%
CT %readme% %arcname%
