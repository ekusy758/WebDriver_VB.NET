@Rem ====== vbc.exe compile ======
Set VBC="C:\Windows\Microsoft.NET\Framework64\v4.0.30319\vbc.exe"
%VBC% /nologo VersionController.vb /t:library
%VBC% /nologo WebController.vb /t:library /r:VersionController.dll,WebDriver.dll
%VBC% /nologo SubWebAuto.vb /t:library /r:WebController.dll
%VBC% /nologo Main.vb /r:SubWebAuto.dll
dir | findstr "VersionController.dll WebController.dll SubWebAuto.dll Main.exe"
