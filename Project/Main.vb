Option Explicit On
Imports System

Namespace WebTool.Common
  '****************************************************************************
  '＜Mainクラス＞
  '  クラス概要　　：エントリポイントを定義する。
  '  最終更新日　　：2025.01.05 Ver1.0.0 新規作成
  '****************************************************************************
  Class Main
    '**************************************************************************
    ' 概要：Mainプロシージャ
    ' 備考：VB.NETではMainプロシージャをクラスで宣言する場合はSharedが必要
    '**************************************************************************
    Shared Sub Main()
      Dim runtimeVersion As String = Environment.Version.ToString()
      Dim args() As String = Environment.GetCommandLineArgs()
      Console.WriteLine("+--------------------------------------------------+")
      Console.WriteLine("+ <VB.NET Main>")
      Console.WriteLine("+ Runtime Version: " & runtimeVersion)
      Console.WriteLine("+--------------------------------------------------+")
      console.WriteLine("====== プログラムを開始します。======")
      SubWebAuto.StartProcess(args)
      Console.WriteLine("====== プログラムを終了します。======")
    End Sub

  End Class

End Namespace