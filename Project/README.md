+--------------------------------------------------+
  はじめに
+--------------------------------------------------+
vbc.exeで作るWeb操作自動化プログラムです。
コンパイル方法や環境の準備手順は下記を参照してください。

+--------------------------------------------------+
  準備手順
+--------------------------------------------------+
# chromedriverをプロジェクトに配置する
chrome for developersサイトからダウンロード
(ブラウザと同じバージョンのドライバを選択する)

# Selenium WebDriverをプロジェクトに配置する
nugetから「.NET Framework」対応のパッケージをダウンロード
パッケージ内の「.\lib\net48\WebDriver.dll」を配置

# コンパイラのパス(vbc.exe)を環境変数に登録しておく
C:\Windows\Microsoft.NET\Framework64\v4.0.30319

# .NET Frameworkランタイムバージョン確認
Environment.Versionを表示する
→Ver:4.0.30319.42000
(4.6以降は、4.0.30319.42000 形式になる)

+--------------------------------------------------+
  通常作業 
+--------------------------------------------------+
# vbcコンパイル
vbc.exe VersionController.vb /t:library
vbc.exe WebController.vb /t:library /r:VersionController.dll,WebDriver.dll
vbc.exe Main.vb /r:WebController.dll

# ドライバーのバージョン確認
.\driver\chrome\win64\chromedriver.exe --version

# ドライバー残留プロセスチェック
tasklist | find "driver.exe"

# ドライバー残留プロセスKILL
taskkill /im chromedriver.exe /f
