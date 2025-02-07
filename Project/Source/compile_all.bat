@Rem ====== vbc.exe compile ======
Set VBC="C:\Windows\Microsoft.NET\Framework64\v4.0.30319\vbc.exe"
cd %~dp0%
%VBC% /nologo VersionController.vb /t:library /out:..\VersionController
%VBC% /nologo WebController.vb /t:library /r:..\VersionController.dll,..\WebDriver.dll /out:..\WebController
%VBC% /nologo SubWebAuto.vb /t:library /r:..\WebController.dll /out:..\SubWebAuto
%VBC% /nologo Main.vb /r:..\SubWebAuto.dll /out:..\WebAutoMain
dir ..\ | findstr "VersionController.dll WebController.dll SubWebAuto.dll WebAutoMain.exe"
