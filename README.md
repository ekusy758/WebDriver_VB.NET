# WebDriver_VB.NET

## はじめに
vbc.exeで作るWeb操作自動化プログラムです。<br>
コンパイル方法や環境の準備手順は下記を参照してください。<br>

## 環境準備
  - chromedriverをプロジェクトに配置する<br>
    chrome for developersからブラウザと同バージョンをダウンロード<br>

  - Selenium WebDriverをプロジェクトに配置する<br>
    nugetから「.NET Framework」対応のパッケージをダウンロード<br>
    nupkgパッケージ内の「.\lib\net48\WebDriver.dll」を配置<br>

  - Chromeのパスを確認する(環境ごとに異なる)<br>
    C:\Program Files (x86)\Google\Chrome\Application\chrome.exe<br>
    C:\Program Files\Google\Chrome\Application\chrome.exe<br>

  - コンパイラのパス(vbc.exe)を環境変数に登録しておく<br>
    C:\Windows\Microsoft.NET\Framework64\v4.0.30319<br>

  - .NET Frameworkランタイムバージョン確認<br>
    Environment.Versionを表示する<br>
    →Ver:4.0.30319.42000<br>
    (4.6以降は、4.0.30319.42000 形式になる)<br>

## 通常作業
  - ドライバーのバージョン確認<br>
    chromedriver.exe --version<br>

  - ドライバー残留プロセスチェック<br>
    tasklist | find "driver.exe"<br>

  - ドライバー残留プロセスKILL<br>
    taskkill /im chromedriver.exe /f<br>
