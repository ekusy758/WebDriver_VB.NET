Option Explicit On
Imports System

Namespace WebTool.Common

  '****************************************************************************
  '＜Web自動操作起動クラス＞
  '  クラス概要　　：Webブラウザ自動操作を開始する。（staticクラス）
  '
  '  フィールド  　：なし
  '
  '  プロパティ　　：なし
  '
  '  コンストラクタ：staticクラス用コンストラクタ（インスタンス作成不可）
  '
  '  メソッド一覧　：StartWebAuto Web操作自動化処理を開始する(Public)
  '
  '  最終更新日　　：2025.01.05 Ver1.0.0 新規作成
  '****************************************************************************
  Public Class SubWebAuto

    ' staticクラス用コンストラクタ（インスタンス作成不可）
    Private Sub New()
    End Sub

    '**************************************************************************
    ' 概要：Webブラウザ自動操作を開始する
    ' 引数：args() コマンドライン引数
    ' 戻値：なし
    ' 参照：なし
    ' 備考：なし
    '**************************************************************************
    Public Shared Sub StartProcess(ByVal args() As String)
      Try
        ' WebControllerインスタンス初期化
        ' 実行環境の確認と自動操作ロジックの読み込みを行う
        Console.WriteLine("[Msg:] WebControllerインスタンスを生成します。")
        Dim webCon = New WebController()

        ' 引数の取得(第1引数を自動化ロジック起動パラメータとする)
        Console.WriteLine("[Msg:] 第1引数を取得します。")
        Dim choice As String = args(1)

        ' 引数による処理分岐
        Console.WriteLine("[Msg:] 引数による処理分岐を行います。Param=" & choice)
        Select choice
          Case "1"
            Console.WriteLine("1: 個別自動化ロジック未登録")
          Case "2"
            Console.WriteLine("2: 個別自動化ロジック未登録")
          Case "3"
            Console.WriteLine("3: 個別自動化ロジック未登録")
          Case "9"
            Console.WriteLine("9: Webブラウザ起動テスト")
            webCon.DriverTest()
          Case "-help", "/help", "?"
            ShowUsage()
          Case Else
            Console.WriteLine("[Err:] 引数が正しくないようです。")
            ShowUsage()
        End Select

      Catch ex As IndexOutOfRangeException

        ShowExceptionInfo(ex)
        ShowUsage()

      Catch ex As Exception

        ShowExceptionInfo(ex)

      End Try

    End Sub

    '**************************************************************************
    ' 概要：使用方法を表示する
    ' 引数：なし
    ' 戻値：なし
    ' 参照：なし
    ' 備考：なし
    '**************************************************************************
    Private Shared Sub ShowUsage()
      Console.WriteLine("+--------------------------------------------------+")
      Console.WriteLine("Usage: 引数に以下の数字を指定して起動ください。")
      Console.WriteLine("       1: 個別自動化ロジック未登録")
      Console.WriteLine("       2: 個別自動化ロジック未登録")
      Console.WriteLine("       3: 個別自動化ロジック未登録")
      Console.WriteLine("       9: Webブラウザ起動テスト")
      Console.WriteLine("+--------------------------------------------------+")
    End Sub

    '**************************************************************************
    ' 概要：例外情報を表示する
    ' 引数：ex 例外オブジェクト
    ' 戻値：なし
    ' 参照：なし
    ' 備考：なし
    '**************************************************************************
    Private Shared Sub ShowExceptionInfo(ByVal ex As Exception)
      Dim type As String = ex.GetType().ToString()
      Console.WriteLine("====== エラーが発生しました。======")
      Console.WriteLine("[キャッチした例外]" & vbCrLf & type & vbCrLf)
      Console.WriteLine("[エラーメッセージ]" & vbCrLf & ex.Message & vbCrLf)
      Console.WriteLine("[スタックトレース]" & vbCrLf & ex.StackTrace)
      Console.WriteLine("====== エラー出力を終了します。======")
    End Sub

  End Class

End Namespace