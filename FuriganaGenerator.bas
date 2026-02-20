Attribute VB_Name = "FuriganaGen.bas"
'/////////////////////////////////////////////////////
'// FuriganaGen.bas
'// 指定された Word の文書からルビを振った文書を作成する
'// API を使用する作成と Word の機能を使用する作成
'//
'// このプログラムで使用している Yahoo の API
'//
'// このプログラムは以下の記事を参考にした
'//
'// 関数:
'// FuriganaGen()
'// FuriganaGenByRuby()
'// FuriganaToWord()
'// GetFile()
'// GetNewFileName()
'//
'// 履歴:
'// 2026/02/20 作成開始
Option Explicit

'/////////////////////////////////////////////////////
'// FuriganaGen()
'// 指定された Word の文書からテキストを取り出し、
'// Yahoo のふりがな Api に投げてルビ付文書を作成する
'//
Public Sub FuriganaGen()

   On Error GoTo FuriganaGen_Error
   
   Dim fn As String
   Dim target As Document
   Dim allText As String

   fn = GetFile()
   If Len(fn) <= 0 Then Exit Sub 
   target = Documents.Open(fn)
   target.Select
   allText = target.Selection.Text
   fn = GetNewFileName(fn)
   Call target.SaveAs2(fn)
   target.Close

   Exit Sub
FuriganaGen_Error:
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "FuriganaGen: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
    Err.Clear
   
End Sub
