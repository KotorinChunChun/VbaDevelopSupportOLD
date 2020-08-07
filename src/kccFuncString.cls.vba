VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kccFuncString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccFuncString
Rem
Rem  @description   文字列変換関数
Rem
Rem  @update        2020/08/07
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------

Rem  @description あらゆる初め括弧から閉じ括弧を返す関数
Rem
Rem  @param open_brackets       初め括弧（機種依存文字対応）
Rem
Rem  @return As String          閉じ括弧
Rem
Function OpenBracketsToClose(open_brackets) As String
    Dim stb As String: stb = open_brackets
    Dim etb As String: etb = ""
    Select Case stb
        Case "[", "{", "<", "［", "｛", "＜"
            etb = ChrW(AscW(stb) + 2)
        Case ChrW(171)
            etb = ChrW(AscW(stb) + 16)
        Case Else
            etb = ChrW(AscW(stb) + 1)
    End Select
    OpenBracketsToClose = etb
End Function

Rem 文字列に含まれる括弧をネストに応じて変化させる関数
Rem
Rem  @param base_str            入力文字列
Rem  @param open_Bracket        置換対象の初め括弧 (既定値:丸括弧)
Rem  @param replaced_brackets   置換後の初め括弧の配列 (既定値:[{(<の4段階)
Rem
Rem  @return As String          括弧を置換済みの文字列
Rem
Rem  @note
Rem      括弧のネストは文字列の先頭から順次変換するロジック
Rem      初め～閉じが不完全でも一切関知しないので注意すること
Rem
Rem  @example
Rem       IN : "Array(aaa, Array( hoge, fuga, piyo, Array(xxx), chun), bbb)"
Rem      OUT : "Array[aaa, Array{ hoge, fuga, piyo, Array(xxx), chun}, bbb]"
Rem
Function ReplaceBracketsNest( _
                ByVal base_str As String, _
                Optional open_bracket = "", _
                Optional replaced_brackets) As String
    If open_bracket = "" Then open_bracket = "("
    If IsMissing(replaced_brackets) Then replaced_brackets = VBA.Array("[", "{", "(", "<")
    Dim close_bracket
    close_bracket = OpenBracketsToClose(open_bracket)
    
    Dim nest As Long
    Dim i As Long
    nest = 0
    For i = 1 To Len(base_str)
        Select Case Mid(base_str, i, 1)
            Case open_bracket
                Mid(base_str, i, 1) = replaced_brackets(nest)
                nest = nest + 1
            Case close_bracket
                nest = nest - 1
                Mid(base_str, i, 1) = OpenBracketsToClose(replaced_brackets(nest))
        End Select
    Next
    ReplaceBracketsNest = base_str
End Function

Rem 区切り文字列のうちかっこに囲われた範囲だけの分割結果を返す
Rem
Rem  @param base_str        入力文字列
Rem  @param start_brackets  開始かっこの種類（終了カッコは自動判断）
Rem  @param remove_brackets カッコを...True:削除する(既定) False:残す
Rem
Rem  @return As Variant/Variant(0 To #)
Rem
Rem  @example
Rem          remove_brackets = False
Rem          Missing                              >> Variant(0 to -1) {}
Rem          String ""                            >> Variant(0 to -1) {}
Rem          String "abc,def,[ghi,jkl,mno],pqr"   >> String(0 to 2) {"ghi","jkl","mno"}
Rem          String "[abc,def],ghi[,jkl,mno],pqr" >> String(0 to 4) {"abc","def","","jkl","mno"}
Rem          String "abc,def,ghi,jkl,mno[,pqr]"   >> String(0 to 1) {"","pqr"}
Rem
Rem          remove_brackets = True
Rem          Missing                              >> Variant(0 to -1) {}
Rem          String ""                            >> Variant(0 to -1) {}
Rem          String "abc,def,[ghi,jkl,mno],pqr"   >> String(0 to 2) {"ghi","jkl","mno"}
Rem          String "[abc,def],ghi[,jkl,mno],pqr" >> String(0 to 4) {"abc","def","","jkl","mno"}
Rem          String "abc,def,ghi,jkl,mno[,pqr]"   >> String(0 to 1) {"","pqr"}
Rem
Rem  @note
Rem     入れ子には非対応
Rem
Public Function SplitWithInBrackets(ByVal base_str, _
                                        start_brackets, _
                                        Optional remove_brackets As Boolean = True _
                                        ) As Variant
    SplitWithInBrackets = VBA.Array()
    If IsMissing(base_str) Then Exit Function
    If base_str = "" Then Exit Function

    Dim reg     As Object: Set reg = CreateObject("VBScript.RegExp")
    Dim retVal     As String
    
    Const CashDelimiter = vbVerticalTab
    Dim openDelim As String, closeDelim As String
    Select Case start_brackets
        Case "(", "["
            openDelim = "\" & start_brackets
            closeDelim = "\" & OpenBracketsToClose(start_brackets)
        Case Else
            openDelim = start_brackets
            closeDelim = OpenBracketsToClose(start_brackets)
    End Select

    SplitWithInBrackets = Split(vbNullString)
    base_str = Replace(base_str, vbLf, "")

    ' 検索条件＝括弧内以外を抽出
    'reg.Pattern = "^(.*?)\(|\)(.*?)\(|\)(.*?).*$"
    reg.Pattern = "^(.*?)" & openDelim & "|" & closeDelim & "(.*?)" & openDelim & "|" & closeDelim & "(.*?).*$"
    'reg.Pattern = "\[[^\[\]]*(?=\])"
    ' 文字列の最後まで検索する
    reg.Global = True

    ' 検索一致文字をカンマに置き換える
    retVal = reg.Replace(base_str, CashDelimiter)

    If IsEmpty(retVal) Or retVal = "" Then Exit Function
    If reg.Execute(base_str).Count = 0 Then Exit Function

    ' 先頭と最後のカンマ文字を除去する
    retVal = Mid(retVal, 2, Len(retVal) - 2)

    ' 括弧内の文字列を括弧の数だけ配列として取得
    SplitWithInBrackets = Split(retVal, CashDelimiter)

End Function

Rem 通常トリムに加えて、文字列中の連続スペースをシングルスペースに変換する。
Rem Excel関数のTRIM互換
Rem
Rem  @param base_str       入力文字列
Rem
Rem  @return As String
Rem
Rem  @example
Rem
Public Function Trim2to1(ByVal base_str) As String
    Do
        Trim2to1 = Replace(Trim(base_str), "  ", " ")
        If Trim2to1 = base_str Then Exit Do
        base_str = Trim2to1
    Loop
End Function

Rem Right関数拡張  最後に出現する区切り文字列を切れ目として右側の文字を返す
Rem
Rem  @param base_str      取り出し元文字列
Rem  @param cut_str       切断文字列（末尾から検索して該当する文字列の手前までを取り出す）
Rem  @param cut_inc       切断文字列を含めて返すかどうか（通常は除外する）
Rem  @param shift_len     取り出し文字列を余分に取り出す文字数（プラス）、削り落とす文字数（マイナス）
Rem  @param should_fill   存在しない場合は入力文字列で埋めるか（既定True）
Rem
Rem  @return As String
Rem
Rem  @example
Rem
Public Function RightStrRev(base_str, cut_str, _
                                Optional cut_inc As Boolean = False, _
                                Optional shift_len As Long = 0, _
                                Optional should_fill = True) As String
    If InStrRev(base_str, cut_str, -1) > 0 And cut_str <> "" Then
        If cut_inc Then
            RightStrRev = Right(base_str, Len(base_str) - InStrRev(base_str, cut_str, -1) + shift_len + 1)
        Else
            RightStrRev = Right(base_str, Len(base_str) - InStrRev(base_str, cut_str, -1) + shift_len + 1 - Len(cut_str))
        End If
    ElseIf should_fill Then
        RightStrRev = base_str
    Else
        RightStrRev = ""
    End If
End Function

Rem フォルダの絶対パスとファイルの相対パスを合成して、目的のファイルの絶対パスを取得する関数
Rem
Rem  @name     AbsolutePathNameEx
Rem  @oldname  BuildPathEx
Rem
Rem  @param base_path      基準パス
Rem  @param ref_path       基準パスからの移動を示す相対パス（または上書きする絶対パス）
Rem
Rem  @return   As String   連結後の絶対パス
Rem
Rem  @note
Rem          fso.GetAbsolutePathName(fso.BuildPath(base_path, ref_path))の問題を解消した関数
Rem          * UNCに..\した時、PC直下には移動できない
Rem          * UNC解析が超低速
Rem          * フォルダ末尾に\が無い
Rem          *
Rem         ※UNCパス＝ネットワークコンピュータ上のファイルを参照するパスで\\から始まるアレ
Rem
Rem  @example
Rem     base_path = ""
Rem          Missing                            >> String ""
Rem          String ""                          >> String ""
Rem          String "C:\Book1.xlsx"             >> String "C:\Book1.xlsx"
Rem
Rem     base_path = "C:\hoge\fuga\"
Rem          Missing                            >> String ""
Rem          String ""                          >> String ""
Rem          String ".\"                        >> String "C:\hoge\fuga\"
Rem          String ".\Book1"                   >> String "C:\hoge\fuga\Book1"
Rem          String ".\Book1.xlsx"              >> String "C:\hoge\fuga\Book1.xlsx"
Rem          String "..\..\Book1.xlsx"          >> String "C:\Book1.xlsx"
Rem          String "..\..\Book1xlsx"           >> String "C:\Book1xlsx"
Rem          String "..\.\Book1.xlsx"           >> String "C:\hoge\Book1.xlsx"
Rem          String "..\Book1.xlsx"             >> String "C:\hoge\Book1.xlsx"
Rem          String "..\piyo\Book1.xlsx"        >> String "C:\hoge\piyo\Book1.xlsx"
Rem          String ".\fuga\piyo\..\Book1.xlsx" >> String "C:\hoge\fuga\fuga\Book1.xlsx"
Rem          String "\Book1.xlsx"               >> String "C:\hoge\fuga\Book1.xlsx"
Rem          String "C:\Book1.xlsx"             >> String "C:\Book1.xlsx"
Rem          String "\\hoge\fuga\"              >> String "\\hoge\fuga\"
Rem          String "\\127.0.0.1\hoge\fuga\"    >> String "\\127.0.0.1\hoge\fuga\"
Rem
Rem     base_path = "\\hoge\fuga\"
Rem          String ".\"                        >> String "\\hoge\fuga\"
Rem          String "\Book1.xlsx"               >> String "\\hoge\fuga\Book1.xlsx"
Rem
Rem     base_path = "\\127.0.0.1\hoge\fuga\"
Rem          String ".\Book1"                   >> String "\\127.0.0.1\hoge\fuga\Book1"
Rem          String ".\fuga\piyo\..\Book1.xlsx" >> String "\\127.0.0.1\hoge\fuga\fuga\Book1.xlsx"
Rem
Public Function AbsolutePathNameEx(ByVal base_path As String, ByVal ref_path As String) As String
    If IsMissing(ref_path) Then Exit Function
    If ref_path = "" Then Exit Function
    If ref_path Like "[A-Z]:\?*" Or ref_path Like "\\?*\?*" Then AbsolutePathNameEx = ref_path: Exit Function
    If IsMissing(base_path) Then Exit Function
    If base_path = "" Then Exit Function
    
    Dim i As Long
    
    base_path = Replace(base_path, "/", "\")
    base_path = Left(base_path, Len(base_path) - IIf(Right(base_path, 1) = "\", 1, 0))
    
    ref_path = Replace(ref_path, "/", "\")
    
    Dim retVal As String
    Dim rpArr() As String
    rpArr = Split(ref_path, "\")
    
    For i = LBound(rpArr) To UBound(rpArr)
        Select Case rpArr(i)
            Case "", "."
                If retVal = "" Then retVal = base_path
                rpArr(i) = ""
            Case ".."
                If retVal = "" Then retVal = base_path
                If InStrRev(retVal, "\") = 0 Then
                    'Err.Raise 8888, "AbsolutePathNameEx", "到達できないパスを指定しています。"
                    AbsolutePathNameEx = "到達不能"
                    Exit Function
                End If
                retVal = Left(retVal, InStrRev(retVal, "\") - 1)
                rpArr(i) = ""
            Case Else
                retVal = retVal & IIf(retVal = "", "", "\") & rpArr(i)
                rpArr(i) = ""
        End Select
        '相対パス部分が空欄、.\、..\で終わった時、末尾の\が不足するので補完が必要
        If i = UBound(rpArr) Then
            If ref_path <> "" Then
                If Right(ref_path, 1) = "\" Then
                    retVal = retVal & "\"
                End If
            End If
        End If
    Next
    '連続\の消去とネットワークパス対策
    retVal = Replace(retVal, "file:\\", "file://")
    retVal = Replace(retVal, "\\", "\")
    retVal = IIf(Left(retVal, 1) = "\", "\", "") & retVal
    AbsolutePathNameEx = retVal
End Function

Rem  パス名からファイル名を除いて､パスを取得します｡（最後に「\」はつきません。コロン「:」がなくかつ円記号「\」がない場合はファイルとします）
'Function GetPathName(PathName As String) As String
'  Dim l As Long ' 文字数
'  Dim yen As Long ' \ フォルダの区切り記号の位置
'  Dim colon As Long ' : ドライブの記号の位置
'
'  yen = InStrRev(PathName, Application.PathSeparator, compare:=vbBinaryCompare)
'  colon = InStrRev(PathName, ":", compare:=vbBinaryCompare)
'  l = Len(PathName)
'
'  GetPathName = PathName
'  If PathName = "." Then Exit Function
'  If PathName = ".." Then Exit Function
'
'  If yen > 0 Then
'    GetPathName = Left$(PathName, yen - 1)
'  ElseIf colon > 0 Then
'    GetPathName = PathName ' ドライブ
'  Else
'    GetPathName = vbNullString ' 円記号「\」がない場合はファイルとします
'  End If
'End Function

Rem ファイルパスを展開して、ディレクトリ、ファイル名、拡張子　をとりだす
Rem
Rem  @param FullPath        フルパスデータ
Rem  @param AddPath         戻り値にフォルダパスを含める
Rem  @param AddName         戻り値にベースファイル名を含める
Rem  @param AddExtension    戻り値に拡張子を含める
Rem  @param outPath         実引数にフォルダパスを返す(C:\hoge\)
Rem  @param outName         実引数にファイル名またはフォルダ名を返す("fuga")
Rem  @param outExtension    実引数に拡張子を返す(".ext")
Rem  @param outIsFolder     実引数にoutNameがフォルダの時Trueを返す
Rem
Rem  @return    As String   結合したパスデータ
Rem
Rem  @note
Rem     戻り値やoutNameには\が無いので注意すること
Rem
Rem  @example
Rem     | FullPath          | AddX3 | return            | outPath | outName | outExt | IsFolder |
Rem     | ----------------- | ----- | ----------------- | ------- | ------- | ------ | -------- |
Rem     | D:\vba\.txt       | TTT   | D:\vba\.txt       | D:\vba\ |         | .txt   | FALSE    |
Rem     | D:\vba\file       | TTT   | D:\vba\file       | D:\vba\ | file    |        | FALSE    |
Rem     | D:\vba\file.txt   | TTT   | D:\vba\file.txt   | D:\vba\ | file    | .txt   | FALSE    |
Rem     | D:\vba\file.2.txt | TTT   | D:\vba\file.2.txt | D:\vba\ | file.2  | .txt   | FALSE    |
Rem     | D:\vba\fol        | TTT   | D:\vba\fol        | D:\vba\ | fol     |        | TRUE     |
Rem     | D:\vba\fol\       | TTT   | D:\vba\fol        | D:\vba\ | fol     |        | TRUE     |
Rem     | D:\vba\fol.2      | TTT   | D:\vba\fol.2      | D:\vba\ | fol.2   |        | TRUE     |
Rem     | D:\vba\fol.2\     | TTT   | D:\vba\fol.2      | D:\vba\ | fol.2   |        | TRUE     |
Rem
Public Function GetPath( _
        ByVal FullPath, _
        ByVal AddPath As Boolean, _
        ByVal AddName As Boolean, _
        ByVal AddExtension As Boolean, _
        Optional ByRef outPath, _
        Optional ByRef outName, _
        Optional ByRef outExtension, _
        Optional ByRef outIsFolder) As String
    outPath = "": outName = "": outExtension = "": outIsFolder = False
'    outPath = "XXXX": outName = "XXXX": outExtension = "XXXX": outIsFolder = False
    
    If IsEmpty(FullPath) Then Exit Function
    If TypeName(FullPath) <> "String" Then Exit Function
    If Len(FullPath) = 0 Then Exit Function
    
'    FullPath = RenewalPath(FullPath)   'これするとファイルフォルダ判定がバグる
    If FullPath = "" Then Exit Function
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    '最後が\ならフォルダ扱い。
    '違ってもfsoで実物から判定する。
    '実在しないフォルダの場合、拡張子の有無で判定をする。
    'FullPathの末尾には\を付けない状態で後の処理に引き継ぐ
    outIsFolder = (FullPath Like "*\")
    If outIsFolder Then
        FullPath = Left$(FullPath, Len(FullPath) - 1)
    Else
        outIsFolder = fso.FolderExists(FullPath)
    End If
    
    'パス部とファイル部の抽出
    Dim NameAndExt As String
    outPath = Strings.Left(FullPath, Strings.InStrRev(FullPath, "\"))
    NameAndExt = Strings.Right(FullPath, Strings.Len(FullPath) - Strings.InStrRev(FullPath, "\"))
    If outIsFolder Then outName = NameAndExt: GoTo ExitProc
    
    'ファイル部と拡張子の抽出
    If InStr(NameAndExt, ".") = 0 Then outName = NameAndExt: GoTo ExitProc
    outName = Strings.Left(NameAndExt, Strings.InStrRev(NameAndExt, ".") - 1)
    outExtension = Strings.Right(NameAndExt, Strings.Len(NameAndExt) - Strings.InStrRev(NameAndExt, ".") + 1)
    
ExitProc:
    GetPath = ""
    If AddPath Then GetPath = GetPath & outPath
    If AddName Then GetPath = GetPath & outName
    If AddExtension Then GetPath = GetPath & outExtension
End Function

Rem パスを規定の書式に書き換える。（ネットワークドライブ対応）
'Public Function RenewalPath(ByVal Path As String, Optional AddYen As Boolean = False) As String
'    'ドットの有無でファイル or フォルダ判定　不完全。
'    If Strings.InStr(Path, ".") = 0 Then Path = Path & IIf(AddYen, "\", "")
'    RenewalPath = Strings.Left(Path, 2) & Strings.Replace(Strings.Replace(Path, "/", "\"), "\\", "\", 3)
'    RenewalPath = ToPathLastYen(RenewalPath, AddYen)
'End Function

Rem 親ディレクトリを返す。
Rem \マークは付与しない
Public Function ToPathParentFolder(ByVal Path As String, Optional AddYen As Boolean = False) As String
    ToPathParentFolder = ToPathLastYen(GetPath(Path, True, False, False), AddYen)
End Function

Rem パスの最後に\を付ける／消す
Public Function ToPathLastYen(Path, AddYen As Boolean) As String
    ToPathLastYen = Path
    If AddYen Then
        If Right(Path, 1) <> "\" Then
            ToPathLastYen = Path & "\"
        End If
    Else
        If Right(Path, 1) = "\" Then
            ToPathLastYen = Left(Path, Len(Path) - 1)
        End If
    End If
End Function

