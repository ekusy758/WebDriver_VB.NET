Option Explicit On
Imports System
Imports System.IO
Imports System.Diagnostics
Imports System.Environment
Imports System.Text.Encoding

Namespace WebTool.Common
  '****************************************************************************
  '＜バージョン管理クラス＞
  '  クラス概要　　：WebブラウザとWebDriverのバージョン管理を行う。
  '  　　　　　　　　バージョン差異があるとWebDriverインスタンスを起動できない。
  '
  '  プロパティ　　：ブラウザとドライバのバージョン情報を保持し、公開する。
  '
  '  コンストラクタ：バージョン情報を取得し、プロパティに格納する。
  '  　　　　　　　　エラー時は例外をスローする。
  '  　　　　　　　　引数にブラウザとドライバのexeパスを指定する。
  '
  '  メソッド一覧　：InitCtor            コンストラクタの初期化処理を行う
  '  　　　　　　　　IsCheckNormal       バージョンチェック結果を返す(Public)
  '  　　　　　　　　ExtractMajor        メジャーバージョンを抽出する
  '  　　　　　　　　GetBrowserVersion   ブラウザのバージョンを取得する
  '  　　　　　　　　GetProductVersinon  ファイルの製品バージョンを取得する
  '  　　　　　　　　GetDriverVersion    ドライバのバージョンを取得する
  '  　　　　　　　　StartConsoleProcess コンソールアプリケーションを起動する
  '
  '  最終更新日　　：2025.01.02 Ver1.0.0 新規作成
  '****************************************************************************   　　　 
  Public Class VersionController

    ' プロパティ定義
    Public Property BrowserFullVer As String
    Public Property DriverFullVer As String
    Public Property BrowserMajorVer As String
    Public Property DriverMajorVer As String
    Public Property CheckStatus As Boolean

    ' コンストラクタ定義
    Public Sub New(ByVal browserExe As String, ByVal driverExe As String)
      InitCtor(browserExe, driverExe)
    End Sub

    '**************************************************************************
    ' 概要：コンストラクタの初期化処理を行う
    ' 引数：String ブラウザexeパス, String ドライバexeパス
    ' 戻値：なし
    ' 参照：なし
    ' 備考：なし
    '**************************************************************************
    Private Sub InitCtor(ByVal browserExe As String, ByVal driverExe As String)
      Try
        BrowserFullVer = GetBrowserVersion(browserExe)
        DriverFullVer = GetDriverVersion(driverExe)
        BrowserMajorVer = ExtractMajor(BrowserFullVer)
        DriverMajorVer = ExtractMajor(DriverFullVer)
        CheckStatus = IsCheckNormal()
      Catch ex As Exception
        Throw
      End Try
    End Sub

    '**************************************************************************
    ' 概要：WebDriver実行環境が正常か判定する
    ' 引数：なし
    ' 戻値：Boolean 
    ' 参照：なし
    ' 備考：ブラウザとドライバのメジャーバージョンが一致していればTrueを返す
    '**************************************************************************
    Public Function IsCheckNormal() As Boolean
      Dim result As Boolean = False
      If (String.Equals(BrowserMajorVer, DriverMajorVer)) Then
        result = True
      End If
      Return result
    End Function

    '**************************************************************************
    ' 概要：メジャーバージョンを抽出する
    ' 引数：String バージョン情報（x.x.x.x）
    ' 戻値：String メジャーバージョン情報（x）
    ' 参照：なし
    ' 備考：Ex.) 131.0.6778.204の場合、131を返す
    '**************************************************************************
    Private Function ExtractMajor(ByVal version As String) As String
      Dim words() As String = version.Split(".")
      Dim majorVer As String = words(0)
      Return majorVer
    End Function

    '**************************************************************************
    ' 概要：ブラウザのバージョンを取得する
    ' 引数：String ブラウザexeパス
    ' 戻値：String バージョン情報（x.x.x.x）
    ' 参照：System.IO.File
    ' 備考：バージョン情報はファイルプロパティの製品バージョンから取得する
    '**************************************************************************
    Private Function GetBrowserVersion(ByVal browserExe As String) As String

      Dim version As String = "0.0.0.0"

      Console.WriteLine("[Msg:] ブラウザのバージョン情報取得を開始します。")

      If Not (File.Exists(browserExe)) Then
        Throw New ApplicationException("[Err:] ブラウザが見つかりません。Path:" & browserExe)
      End If

      ' ブラウザの製品バージョンを取得する
      version = GetProductVersinon(browserExe)

      Console.WriteLine("[Msg:] ブラウザのバージョンは「" &  version & "」です。")
      Console.WriteLine("[Msg:] ブラウザのバージョン情報取得が完了しました。")

      Return version 

    End Function

    '**************************************************************************
    ' 概要：ファイルの製品バージョンを取得する
    ' 引数：String ファイルパス
    ' 戻値：String ファイルプロパティの製品バージョン
    ' 参照：System.Diagnostics.FileVersionInfo
    ' 備考：なし
    '**************************************************************************
    Private Function GetProductVersinon(ByVal filePath As String) As String
      Dim verInfo = FileVersionInfo.GetVersionInfo(filePath)
      Dim prodVer As String = "0.0.0.0"
      prodVer = verInfo.ProductVersion
      Return prodVer
    End Function

    '**************************************************************************
    ' 概要：ドライバのバージョンを取得する（chromedriver.exeから取得）
    ' 引数：String ドライバexeパス
    ' 戻値：String バージョン情報（x.x.x.x）
    ' 参照：System.IO.File
    ' 備考：バージョン情報は「chromedriver.exe --version」の実行結果から取得する
    '**************************************************************************
    Private Function GetDriverVersion(ByVal driverExe As String) As String

      Dim command As String = driverExe & " --version"
      Dim version As String = "0.0.0.0"
      Dim result As String = ""

      Console.WriteLine("[Msg:] ドライバのバージョン情報取得を開始します。")

      If Not (File.Exists(driverExe)) Then
        Throw New ApplicationException("[Err:] ドライバが見つかりません。Path:" & driverExe)
      End If

      ' コンソールアプリケーションを起動して、コマンドからバージョン情報を取得する
      Console.WriteLine("[Msg:] コマンドからバージョンを取得します。")
      result = StartConsoleProcess(command)
      version = result.Split(" ")(1)

      Console.WriteLine("[Msg:] ドライバのバージョンは「" &  version & "」です。")
      Console.WriteLine("[Msg:] ドライバのバージョン情報取得が完了しました。")

      return version

    End Function

    '**************************************************************************
    ' 概要：コンソールプロセスを起動してDOSコマンドを実行する
    ' 引数：String コマンド
    ' 戻値：String コマンド実行結果（標準出力）
    ' 参照：System.Diagnostics.Process, System.Text.Encoding.GetEncoding,
    ' 　　　System.Environment.GetEnvironmentVariable
    ' 備考：標準出力はShift_JISで取得する
    '**************************************************************************
    Private Function StartConsoleProcess(ByVal cmd As String) As String

      Dim procCon = New Process()
      Dim cmdPath As String = GetEnvironmentVariable("ComSpec")
      Dim results As String = ""

      ' DOSコンソールのウィンドウを開かずに起動するための設定
      With procCon.StartInfo
        .FileName = cmdPath            ' コマンドプロンプトのパスを指定
        .Arguments = "/c " & cmd       ' /cはコマンド実行後にウィンドウを閉じる
        .CreateNoWindow = True         ' DOSコンソールのウィンドウを開かない
        .UseShellExecute = False       ' シェル機能を使用しない
        .RedirectStandardOutput = True ' 標準出力をリダイレクトする
        .RedirectStandardError = True  ' 標準エラー出力をリダイレクトする
        .StandardOutputEncoding = GetEncoding("Shift_JIS") ' 標準出力の文字コードを指定
        .StandardErrorEncoding = GetEncoding("Shift_JIS")  ' 標準エラー出力の文字コードを指定
      End With

      ' コンソールプロセスを起動してバージョン情報を取得する
      procCon.Start()

      ' 標準出力を取得する
      results = procCon.StandardOutput.ReadToEnd()

      ' プロセスの終了を待つ
      procCon.WaitForExit()
      procCon.Close()

      Return results

    End Function

  End Class

End Namespace
