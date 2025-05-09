Option Explicit On
Imports System
Imports System.IO
Imports OpenQA.Selenium.Chrome

Namespace WebTool.Common
  '****************************************************************************
  '＜Web操作制御クラス＞
  '  クラス概要　　：WebDriverによるWebブラウザ自動操作を制御する。
  '  　　　　　　　　バージョン管理クラスインスタンスを生成する。
  '  　　　　　　　　個別の自動操作ロジックの読み込みを行う。（※予定。開発中）
  '
  '  フィールド  　：browserExe  ブラウザのexeファイルパス（絶対パス）
  '  　　　　　　　　driverPath  ドライバの格納フォルダ
  '  　　　　　　　　driverExe   ドライバのexeファイルパス（絶対パス）
  '  　　　　　　　　verCon      バージョン管理クラスインスタンス
  '
  '  プロパティ　　：なし
  '
  '  コンストラクタ：バージョン管理クラスインスタンスを生成し、チェック結果を取得する。
  '  　　　　　　　　エラー時は例外をスローする。
  '  
  '  メソッド一覧　：InitCtor    コンストラクタの初期化処理を行う
  '  　　　　　　　　DriverTest  WebDriver起動テストを行う (Public)
  '  　　　　　　　　GetBrowserPath ブラウザのパスを取得する (Private)
  '
  '  最終更新日　　：2025.01.03 Ver1.0.0 新規作成
  '  　　　　　　　　2025.05.09 Ver1.0.1 ブラウザのパス取得メソッドを追加
  '****************************************************************************
  Public Class WebController

    ' フィールド定義
    Private driverPath As String = ".\Driver\Chrome\"
    Private driverExe As String = driverPath & "chromedriver.exe"
    Private verCon As VersionController

    ' コンストラクタ定義
    Public Sub New()
      InitCtor()
    End Sub

    '**************************************************************************
    ' 概要：コンストラクタの初期化処理を行う
    ' 引数：なし
    ' 戻値：なし
    ' 参照：なし
    ' 備考：なし
    '**************************************************************************
    Private Sub InitCtor()
      Try
        ' インスタンスを生成し、バージョン情報を取得する
        Console.WriteLine("[Msg:] VersionControllerインスタンスを生成します。")
        verCon = New VersionController(GetBrowserPath(), driverExe)
        
        ' バージョンチェック結果を確認する
        If Not (verCon.CheckStatus) Then
          ' [DEBUG] ここでMSGBOX＆最新ドライバ取得処理を書いていく
          ' ......開発中......
          Throw New ApplicationException("[Err:] バージョンチェック結果が不正です。")
        End If
      Catch ex As Exception
        Throw
      End Try
    End Sub

    '**************************************************************************
    ' 概要：Webブラウザ起動テストを行う
    ' 引数：なし
    ' 戻値：なし
    ' 参照：OpenQA.Selenium.Chrome.ChromeDriver
    ' 備考：なし
    '**************************************************************************
    Public Sub DriverTest()
      Console.WriteLine("[Msg:] WebDriver起動テストを開始します。")
      Console.WriteLine("+--------------------------------------------------+")

      Using driver = New ChromeDriver(driverPath)
        driver.Url = "https://www.google.co.jp/"
        driver.Quit()
      End Using

      Console.WriteLine("+--------------------------------------------------+")
      Console.WriteLine("[Msg:] WebDriver起動テストを終了します。")
    End Sub

    '**************************************************************************
    ' 概要：ブラウザのパスを候補から取得する
    ' 引数：なし
    ' 戻値：ブラウザのパス（絶対パス）
    ' 参照：System.IO
    ' 備考：なし
    '**************************************************************************
    Private Function GetBrowserPath() As String
      Dim possiblePaths As String() = {
        "C:\Program Files\Google\Chrome\Application\chrome.exe",
        "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
      }
      For Each path In possiblePaths
        If File.Exists(path) Then
          Return path
        End If
      Next
      Return Nothing ' 見つからない場合
    End Function

  End Class

End Namespace