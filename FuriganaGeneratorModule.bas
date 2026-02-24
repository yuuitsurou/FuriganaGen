Attribute VB_Name = "FuriganaGeneratorModule"
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

   'ルビを付ける文書を開く
   fn = GetWordFile()
   If Len(fn) <= 0 Then Exit Sub
   Set target = Documents.Open(fn)
   '対象の文書の全文を API に渡す
   allText = target.Range.Text
   token = InputBox("APIに渡すトークンを入力して下さい。")
   If Len(token) <= 0 Then
      Call MsgBox("トークンが入力されませんでした。")
      Exit Sub
   End If
   'API からのレスポンスをパース
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
         ElseIf IsKanji(surface) Then
            rs.Add Array(surface, furigana)
         End If
      End If
   Next
   'パースした API のレスポンスを元に対象の文書にルビを付ける
   If SetRuby(target, rs) Then
      '成功
   Else
      '失敗
      Exit Sub
   End If

   'ルビを付けた文書を新しい文書として保存
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
   GetApiResult = ""
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
   
   IsKanji = False
   Dim code As Long
   code = AscW(s)
   If code < 0 Then
      code = code + 65536
   End If
   IsKanji = (19968 <= code And code <= 40959) ' Unicodeの範囲で漢字かどうかを判定(&H4E00-&H9FFF)
   
   Exit Function
IsKanji_Error:
   IsKanji = False
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
               & "IsKanji: " & Err.Number & vbCrLf _
               & "( " & Err.Description & " )")
   
End Function

'/////////////////////////////////////////////////////
'// SetRuby(rng, rubyList)
'// 渡されたルビのリストに従って、Word の選択範囲にルビを設定する。
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
   Dim startPos As Long: startPos = target.Range.Start
   Dim endPos As Long: endPos = target.Range.End
   Set rng = target.Range
   'ルビのリストを最後から処理する
   For ii = 1 To rubyList.Count
      'ルビを付ける文字列 = 0 / ルビ = 1 の配列を取り出す
      ruby = rubyList(ii)
      '対象の文書を検索
      rng.Start = startPos
      rng.End = endPos
      With rng.Find
         .Forward = True
         .Wrap = wdFindContinue
         .Execute FindText:=ruby(0)
         If .Found Then
            'マッチするとレンジの範囲がマッチした部分になるので、
            'そのレンジにルビを付ける
            rng.PhoneticGuide Text:=ruby(1)  ', Alignment:=wdPhoneticGuideAlignmentCenter, Raise:=10, FontSize:=5 ' ルビ部分
            startPos = rng.End + 1
         Else
            SetRuby = False
            Call MsgBox("ルビを付けることができませんでした。" & vbCrLf & "対象の文字列:" & ruby(0) & "《" & ruby(1) & "》")
            Exit Function
         End If
      End With
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
   '対象の文書のパスを取得
   Dim fn As String
   fn = GetWordFile()
   If Len(fn) <= 0 Then Exit Sub
   'ルビ付文書を開いて読み込み
   Dim initialPath As String
   initialPath = Left(fn, InstrRev(fn, "\"))
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
         rFn = .SelectedItems(1)
      Else
         Call MsgBox("キャンセルされました。")
      End If
   End With
   If Len(rFn) <= 0 Then Exit Sub
   Dim s As String
   s = ReadFileToSJISText(rFn)
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
      .Pattern = "｜([^｜]+)《(.*?)》" ' 青空文庫形式のルビを探す正規表現
   End With
   If rgx.Test(s) Then
      '成功
   Else
      '失敗
      Call MsgBox("ルビ付のテキスト文書(青空文庫形式)として指定されたファイルからルビを取り出すことができません。" & vbCrLf & "ファイル: " & rFn)
      Exit Sub
   End If
   '正規表現による検索結果からルビのリストを作成
   Dim ms As Object
   Dim ii As Long
   Dim m As Variant
   Dim surface As String
   Dim furigana As String
   Dim rs As New Collection
   Set ms = rgx.Execute(s)
   For ii = 0 To (ms.Count - 1)
      Set m = ms(ii)
      surface = m.SubMatches(0)
      If IsKanji(surface) Then
         furigana = m.SubMatches(1)
         rs.Add Array(surface, furigana)
      End If
   Next
   'ルビのリストから対象の文書にルビを付ける
   Dim target As Document
   Set target = Documents.Open(fn)
   If SetRuby(target, rs) Then
      '成功
   Else
      '失敗
      Exit Sub
   End If

   'ルビを付けた文書を新しい文書として保存
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
         GetWordFile = ""
         Call MsgBox("キャンセルされました。")
      End If
   End With
   
   Exit Function
GetWordFile_Error:
   GetWordFile = ""
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
   
   '対象の文書を開く
   Dim fn As String
   fn = GetWordFile()
   If Len(fn) <= 0 Then Exit Sub
   Dim target As Document
   Set target = Documents.Open(fn)
   
   'ルビを振る
   Dim rng As Range
   Dim c As Range
   Dim ii As Long
   Dim kanji As Boolean
   For Each rng In target.Range.Words
      'ルビが振られているか
      If rng.Fields.Count < 1 Then
         '全て漢字か
         If IsContainKanji(rng.Text, True) Then
            rng.Select
            Application.Dialogs(wdDialogPhoneticGuide).Show 1
         Else
            If IsContainKanji(rng.Text, False) Then
               '漢字が含まれていたら、1文字ずつ処理
               For Each c In rng.Characters
                  ii = 1
                  kanji = False
                  Do While IsKanji(c.Text)
                     kanji = True
                     c.End = ii
                     ii = ii + 1
                     If ii > rng.Characters.Count Then Exit Do
                  Loop
                  If kanji Then
                     c.Select
                     Application.Dialogs(wdDialogPhoneticGuide).Show 1
                     Exit For
                  End If
               Next
            End If
         End If
      End If
   Next
   
   'ルビを付けた文書を新しい文書として保存
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
   IsContainKanji = False
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
               & "IsContainKanji: " & Err.Number & vbCrLf _
               & "( " & Err.Description & " )")
   Err.Clear
   
End Function
