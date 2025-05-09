Option Explicit On
Imports System
Imports System.IO
Imports System.Net
Imports System.Text

Namespace WebTool.Common

  '****************************************************************************
  '＜ChromeDriverダウンロードクラス＞
  Public Class DownloadChromeDriver

    ' プロパティ定義
    Public Property TargetMajorVersion As String ' 対象メジャーバージョン
    Public Property TargetFullVersion As String  ' 対象フルバージョン
    Public Property TargetZipFileUrl As String   ' ダウンロードするZIPファイルのURL
    Public Property DownloadPath As String       ' ダウンロード先パス
    Public Property DeployExePath As String      ' 解凍したChromeDriver.exeの配置先

    ' フィールド定義
    'Private zipName_v114 As String = "chromedriver_win32.zip"
    Private zipName_v115 As String = "chromedriver-win64.zip" ' Web上のZIPファイル名
    Private exeName As String = "chromedriver.exe"

    ' コンストラクタ定義
    Public Sub New(Byval browserMajorVersion As String)
      InitCtor(browserMajorVersion, ".\", ".\")
    End Sub

    ' コンストラクタ定義 (オーバーロード)
    Public Sub New(Byval browserMajorVersion As String, _ 
                   ByVal downloadDir As String)
      InitCtor(browserMajorVersion, downloadDir, ".\")
    End Sub

    ' コンストラクタ定義 (オーバーロード)
    Public Sub New(Byval browserMajorVersion As String, _ 
                   ByVal downloadDir As String _
                   ByVal deployDir As String)
      InitCtor(browserMajorVersion, downloadDir, deployDir)
    End Sub

    '**************************************************************************
    ' 概要：コンストラクタの初期化処理を行う
    ' 引数：String型 ダウンロード対象メジャーバージョン
    ' 　　　String型 ZIPファイルのダウンロード先フォルダ
    ' 　　　String型 ZIPファイルの解凍先フォルダ
    ' 戻値：なし
    ' 参照：なし
    ' 備考：なし
    '**************************************************************************
    Private Sub InitCtor(ByVal majorVersion As String, ByVal downloadDir As String, ByVal deployDir As String)
      Try
        ' プロパティに値をセット
        TargetMajorVersion = majorVersion
        TargetFullVersion = GetFullVersion()
        TargetZipFileUrl = CreateZipFileUrl()
        DownloadPath = Path.Combine(downloadDir, "chromedriver.zip")
        DeployExePath = deployDir
      Catch ex As Exception
        Throw
      End Try
    End Sub

    '**************************************************************************
    ' 概要：バージョン確認用のAPIからフルバージョンを取得する
    ' 引数：なし
    ' 戻値：String型 フルバージョン
    ' 参照：なし
    ' 備考：バージョン114以下は実装しない
    '**************************************************************************
    Private Function GetFullVersion() As String
      'Dim baseUrl_114 As String = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
      Dim baseUrl_115 As String = "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_"
      Dim fullUrl As String = ""
      Dim fullVer As String = ""

      If (Integer.Parse(TargetMajorVersion) <= 114) Then
        Throw New NotImplementedException("[Err:] Version 114以下の処理は未実装です。")
      Else
        ' Webサイトからフルバージョンを取得
        fullUrl = baseUrl_115 & TargetMajorVersion
        fullVer = GetHtmlSource()
      End If

      Console.WriteLine("[Msg:] Full Version: " & fullVer)

      Return fullVer

    End Function

  End Class

End Namespace
