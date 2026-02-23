Attribute VB_Name = "FuriganaGenerator"
'/////////////////////////////////////////////////////
'// FuriganaGen.bas
'// 指定された Word の文書からルビを振った文書を作成する
'// API を使用する作成
'// 生成AI でルビテキスト(青空文庫形式)を作成し、それを利用して作成
'// Word の機能を使用する作成
'//
'// 青空文庫形式のルビ
'// https://www.aozora.gr.jp/aozora-manual/index-input.html
'// ルビは、ルビの付く文字列のあとに、「《》」でくくって入力します。（学術記号の「≪≫」と混同しやすいので注意してください。）
'// ルビの付く文字列がはじまる前には、「｜」を入れます。
'//
'// このプログラムで使用している Yahoo の API
'// Web Services by Yahoo! JAPAN （https://developer.yahoo.co.jp/sitemap/）
'//
'// このプログラムは以下の記事を参考にした
'// Wordの文章にルビ（ふりがな）を自動で振れるマクロ
'// https://kagakucafe.com/2020092311700.html
'// Wordの文章にルビ（ふりがな）を自動で振れるマクロ の補足
'// https://kagakucafe.com/2021071316030.html
'// Wordでのルビ振りを一括でできるようにした話
'// https://qiita.com/enoz_jp/items/0a746cd1c0c021599a1d
'// Wordでのルビ振りを一括でできるようにした話2
'// https://qiita.com/enoz_jp/items/915ee4db96ae2097fe02
'//
'// 下記のモジュールが必要
'// JsonConverter.bas (https://github.com/VBA-tools/VBA-JSON)
'// Dictionary.cls    (https://github.com/VBA-tools/VBA-Dictionary)
'// ReadFileToSJISTextModule.bas
'//
'// 関数:
'// FuriganaGen()
'// GetApiResult()
'// IsKanji()
'// SetRuby()
'// FuriganaByAozora()
'// GetWordFile()
'// GetNewFileName()
'// FuriganaGenByRuby()
'// IsContainKanji()
'//
'// 履歴:
'// 2026/02/20 作成開始
'// 2026/02/23 Ver.0.1
Option Explicit

Const G_API_GRADE = 1
Const G_API_URL = "https://jlp.yahooapis.jp/FuriganaService/V2/furigana"

'/////////////////////////////////////////////////////
'// FuriganaGen()
'// 指定された Word の文書からテキストを取り出し、
'// Yahoo のふりがな Api に投げてルビ付文書を作成する
'//
Public Sub FuriganaGen()

   On Error GoTo FuriganaGen_Error

   Dim token As String 
   Dim fn As String
   Dim target As Document
   Dim allText As String
   Dim json As Object

   fn = GetWordFile()
   If Len(fn) <= 0 Then Exit Sub 
   Set target = Documents.Open(fn)
   allText = target.Range.Text
   token = InputBox("APIに渡すトークンを入力して下さい。")
   If Len(token) <= 0 Then
      Call MsgBox("トークンが入力されませんでした。")
      Exit Sub
   End If 
   Set json = JsonConverter.ParseJson(GetApiResult(allText, token))

   Dim w As Object
   Dim sw As Object 
   Dim rs As New Collection
   Dim surface As String
   Dim furigana As String

   For Each w In json("result")("word")
      surface = w("surface")
      If w.Exists("furigana") Then
	 furigana = w("furigana")
	 If w.Exists("subword") Then
	    For Each sw In w("subword")
	       surface = sw("surface")
	       If IsKanji(surface) Then
		  furigana = sw("furigana")
		  rs.Add Array(surface, furigana)
	       End If
	    Next
	 Elseif IsKanji(surface) Then
	    rs.Add Array(surface, furigana)
	 End If
      End If
   Next

   If SetRuby(target, rs) Then
   Else 
      Exit Sub 
   End If 
   
   fn = GetNewFileName(fn, "ルビ付")
   If Len(fn) > 0 Then 
      Call target.SaveAs2(fn)
      Call MsgBox("ルビ付ファイルを作成しました。" & vbCrLf & "ファイル: (" & fn & ")")
   Else
      Call MsgBox("ルビ付のファイルが作成できませんでした。" & vbCrLf & "ルビ付ファイルと同名のファイルが既に存在します。")
   End If 
   target.Close

   Exit Sub
FuriganaGen_Error:
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "FuriganaGen: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
   Err.Clear

End Sub

'/////////////////////////////////////////////////////
'// GetApiResult(s, token)
'// 文字列と API のトークンを受け取り、API へ投げてその結果を返す関数
'// 引数:
'// s: String: 対象の文字列。4k まで。
'// token: String: API のトークン
'// 返り値:
'// String: API の返す結果
Private Function GetApiResult(ByVal s As String, ByVal token As String) As String

   On Error GoTo GetApiResult_Error

   Dim api As New Dictionary
   Dim apiData As New Dictionary

   api.Add "id", Format(Now(), "yyyyMMdd-HHmmss")
   api.Add "jsonrpc", "2.0"
   api.Add "method", "jlp.furiganaservice.furigana"
   apiData.Add "q", s
   apiData.Add "grade", G_API_GRADE
   api.Add "params", apiData

   Dim rq As Object: Set rq = CreateObject("MSXML2.XMLHTTP")
   With rq
      .Open "POST", G_API_URL, False
      .SetRequestHeader "Content-Type", "application/json"
      .SetRequestHeader "User-Agent", "Yahoo AppID: " & token
      .Send JsonConverter.ConvertToJson(api)
      GetApiResult = .ResponseText
   End With
   
   Exit Function
GetApiResult_Error:
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "GetApiResult: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
   Err.Clear
   
End Function

'/////////////////////////////////////////////////////
'// IsKanji(s)
'// 渡された文字列が漢字を含むかどうかをチェックする関数
'// 引数:
'// s: String: 検査する文字列
'// 返り値:
'// Boolean: 対象の文字列が漢字を含む場合は True / 含まなければ False
Private Function IsKanji(ByVal s As String) As Boolean

   On Error GoTo IsKanji_Error

   Dim code As Long
   code = AscW(s)
   If code < 0 Then
      code = code + 65536
   End If
   IsKanji = (19968 <= code And code <= 40959) ' Unicodeの範囲で漢字かどうかを判定(&H4E00-&H9FFF)
   
   Exit Function
IsKanji_Error:
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "IsKanji: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
   Err.Clear
   
End Function

'/////////////////////////////////////////////////////
'// SetRuby(rng, rubyList)
'// 渡されたルビのリストに従って、Word の選択範囲にルビを設定する。
'// ルビの設定中にターゲットの文字列が記載されるため、
'// 文字列の重複を避けるため、ルビのリストの最後から処理し、
'// 文書の末尾から1度だけ検索して、ヒットした文字列にルビを付けるようにする。
'// 引数:
'// target: Document: Word 文書
'// rubyList: Collection: ルビのリスト (文字列, ルビ)
'// 返り値:
'// Boolean: 処理の成否
Private Function SetRuby(ByRef target As Document, ByRef rubyList As Collection) As Boolean

   On Error GoTo SetRuby_Error

   SetRuby = True 

   Dim total As Long
   Dim ii As Long
   Dim ruby As Variant
   Dim rng As Range
   total = rubyList.Count
   For ii = total To 1 Step -1
      ruby = rubyList(ii)
      Set rng = target.Range
      With rng.Find
	 .Forward = False 
	 .Wrap = wdFindContinue
	 .Execute FindText:=ruby(0)
	 If .Found Then
	    rng.PhoneticGuide text:=ruby(1)  ', Alignment:=wdPhoneticGuideAlignmentCenter, Raise:=10, FontSize:=5 ' ルビ部分
	 Else
	    SetRuby = False
	    Call MsgBox("ルビを付けることができませんでした。" & vbCrLf & "対象の文字列:" & ruby(0) & "《" & ruby(1) & "》")
	    Exit Function 
	 End If 
      End With
      Set rng = Nothing 
   Next

   Exit Function
SetRuby_Error:
   SetRuby = False 
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "SetRuby: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
   Err.Clear
End Function 

'/////////////////////////////////////////////////////
'// FuriganaByAozora()
'// 指定された Word の文書と、その文書から作成されたルビ付文書(青空文庫形式)から
'// ルビ付の Word の文書を作成する
Public Sub FuriganaByAozora()
   
   On Error GoTo FuriganaByAozora_Error

   Dim fn As String
   fn = GetWordFile()
   If Len(fn) <= 0 Then Exit Sub
   
   Dim initialPath As String
   initialPath = Left(fn, Len(fn) - InstrRev(fn, "\") + 1)
   Dim rFn As String
   Dim fd As FileDialog
   Set fd = Application.FileDialog(msoFileDialogFilePicker)
   With fd
      .AllowMultiSelect = False
      .Title = "Wordの文書のルビ付のテキスト文書(青空文庫形式)選択してください。"
      .Filters.Clear
      .Filters.Add "テキストのファイル", "*.txt"
      .InitialFileName = initialPath 
      If .Show Then
	 rfn = .SelectedItems(1)
      Else
	 Call MsgBox("キャンセルされました。")
      End if
   End With
   If Len(rfn) <= 0 Then Exit Sub
   Dim s As String
   s = ReadFileToSJISText(rfn)
   If Len(s) <= 0 Then
      Call MsgBox("ルビ付のテキスト文書を読み込むことができませんでした。")
      Exit Sub
   End If

   Dim rgx As Object
   Set rgx = CreateObject("VBScript.RegExp")
   With rgx
      .Global = True
      .MultiLine = True
      .IgnoreCase = False
      .Pattern = "｜(.*?)《(.*?)》" ' 青空文庫形式のルビを探す正規表現
   End With
   If Not rgx.Test(s) Then
      Exit Sub
   End If 
   Dim ms As Object
   Dim ii As Long
   Dim m As Variant 
   Dim surface As String
   Dim furigana As String 
   Dim rs As New Collection
   Set ms = rgx.Execute(s)
   For ii = 0 To (ms.Cout - 1)
      Set m = ms(ii)
      surface = m.SubMatches(0)
      If IsKanji(surface) Then
	 furigana = m.SubMatches(1)
	 rs.Add Array(surface, furigana)
      End If 
   Next

   Dim target As Document
   Set target = Documents.Open(fn)
   If SetRuby(target, rs) Then
   Else 
      Exit Sub 
   End If 

   fn = GetNewFileName(fn, "ルビ付")
   If Len(fn) > 0 Then 
      Call target.SaveAs2(fn)
      Call MsgBox("ルビ付ファイルを作成しました。" & vbCrLf & "ファイル: (" & fn & ")")
   Else
      Call MsgBox("ルビ付のファイルが作成できませんでした。" & vbCrLf & "ルビ付ファイルと同名のファイルが既に存在します。")
   End If 

   target.Close
   
   Exit Sub
FuriganaByAozora_Error:
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "FuriganaByAozora: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
   Err.Clear
   
End Sub
'/////////////////////////////////////////////////////
'// GetWordFile()
'// ルビを付ける Word 文書を選択する関数
'// 引数:
'// 返り値:
'// String: 選択された文書名(絶対パス)。選択されなかった時は空文字列
Private Function GetWordFile() As String

   On Error GoTo GetWordFile_Error

   GetWordFile = ""

   Dim fd As FileDialog
   Set fd = Application.FileDialog(msoFileDialogFilePicker)
   With fd
      .AllowMultiSelect = False
      .Title = "ルビを付けるWordの文書を選択してください。"
      .Filters.Clear
      .Filters.Add "Wordのファイル", "*.docx"
      If .Show Then
	 GetWordFile = .SelectedItems(1)
      Else
	 Call MsgBox("キャンセルされました。")
      End if
   End With 
   
   Exit Function
GetWordFile_Error:
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "GetWordFile: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
   Err.Clear
   
End Function

'/////////////////////////////////////////////////////
'// GetNewFileName(fn, suffix)
'// 渡されたファイル名に付加文字をつけたファイル名を返す関数
'// 新しいファイル名のファイルが存在する時は、10 までの数値をつける
'// 従って、新しいファイル名作成のチャレンジは10回まで。
'// 引数:
'// fn: String: 新しくするファイル名(拡張子込み)
'// suffix: String: ファイル名に追加する文字列
'// 返り値:
'// String: 新しいファイル名。作成できなかった時は空文字列
Private Function GetNewFileName(ByVal fn As String, ByVal suffix As String) As String

   On Error GoTo GetNewFileName_Error

   GetNewFileName = ""
   If Len(fn) <= 0 Then Exit Function 
   Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
   Dim posOfPriod As Long
   posOfPriod = InstrRev(fn, ".")
   Dim fnOriginal As String
   Dim extOriginal As String
   fnOriginal = Left(fn, posOfPriod - 1)
   extOriginal = Right(fn, Len(fn) - posOfPriod + 1)
   GetNewFileName = fnOriginal & suffix & extOriginal
   If fn <> GetNewFileName And Not fso.FileExists(GetNewFileName) Then Exit Function
   Dim kaisu As Integer
   For kaisu = 1 To 9
      GetNewFileName = fnOriginal & suffix & CStr(kaisu) & extOriginal
      If Not fso.FileExists(GetNewFileName) Then
	 Exit Function
      Else
	 GetNewFileName = ""
      End If
   Next
   
   Exit Function

GetNewFileName_Error:
   GetNewFileName = ""
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "GetNewFileName: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
   Err.Clear
   
End Function

'/////////////////////////////////////////////////////
'// FuriganaGenByRuby()
'// 指定された Word の文書に、Word の機能でルビを付ける。
'// Word のバグのため、ルビが振られず、無限ループになってしまう漢字がある。
'// その場合は、その漢字に先にルビを振っておいてから実行する。
Public Sub FuriganaGenByRuby()

   On Error GoTo FuriganaGenByRuby_Error

   Dim fn As String 
   fn = GetWordFile()
   If Len(fn) <= 0 Then Exit Sub
   Dim target As Document
   Set target = Documents.Open(fn)
   Dim rng As Range
   Dim ii As Long 
   Dim startPos As Long
   Dim endPos As Long
   For Each rng In target.Range.Words
      'ルビが振られているか
      If rng.Fields.Count < 1 Then 
	 '漢字が含まれているか
	 If IsContainKanji(rng.Text, False) Then
	    For ii = 1 To rng.Characters.Count
	       If IsKanji(rng.Characters(ii).Text) Then
		  startPos = ii
		  endPos = ii
		  ii = ii + 1
		  Do While IsKanji(rng.Characters(ii).Text)
	             endPos = ii
	             ii = ii + 1
	             If ii > rng.Characters.Count Then
			Exit Do
		     End If 
		  Loop
		  With rng
		     .Start = startPos
		     .End = endPos
		     .Select
		     Application.Dialogs(wdDialogPhoneticGuide).Show 1
		  End With 
	       End If
	    Next 
	 End If
      End If 
   Next 
   
   fn = GetNewFileName(fn, "ルビ付")
   If Len(fn) > 0 Then 
      Call target.SaveAs2(fn)
      Call MsgBox("ルビ付ファイルを作成しました。" & vbCrLf & "ファイル: (" & fn & ")")
   Else
      Call MsgBox("ルビ付のファイルが作成できませんでした。" & vbCrLf & "ルビ付ファイルと同名のファイルが既に存在します。")
   End If 
   target.Close

   Exit Sub
FuriganaGenByRuby_Error:
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "FuriganaGenByRuby: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
   Err.Clear
   
End Sub

'/////////////////////////////////////////////////////
'// IsContainKanji(s, allKanji)
'// 渡された文字列に漢字が含まれているかどうかの検査
'// 引数:
'// s: String/: 検査対象の文字列
'// allKanji: Boolean: 文字列全部が漢字かどうかを判定するフラグ True = 全部を判定 / False = 一文字でも含まれているかを判定
'// 戻り値:
'// 検査結果 True = 含まれている / False = 含まれていない
Private Function IsContainKanji(ByVal s As String, ByVal allKanji As Boolean) As Boolean

   On Error GoTo IsContainKanji_Error

   IsContainKanji = False

   Dim ii As Long
   For ii = 1 To Len(s)
      IsContainKanji = IsKanji(Mid(s, ii, 1))
      If allKanji Then
	 If Not IsContainKanji Then Exit For
      Else
	 If IsContainKanji Then Exit For
      End If
   Next
   
   Exit Function
IsContainKanji_Error:
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "IsContainKanji: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
   Err.Clear
   
End Function
