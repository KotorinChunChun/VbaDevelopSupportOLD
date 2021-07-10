Attribute VB_Name = "Win32API_DeclareConverter"
Rem Win32APIのDeclare文を自動的に64bit対応コードに変換するプログラム
Rem
Rem ■公開先
Rem
Rem えくせるちゅんちゅん
Rem 2019/10/20
Rem VBAでWin32APIの64bit対応自動変換プログラムを作ってみた
Rem https://www.excel-chunchun.com/entry/vba-64bit-declare-convert
Rem
Rem ----------------------------------------------------------------------------------------------------
Rem
Rem ■参考資料
Rem
Rem 64 ビット Visual Basic for Applications の概要
Rem  https://docs.microsoft.com/ja-jp/office/vba/language/concepts/getting-started/64-bit-visual-basic-for-applications-overview
Rem
Rem Office の 32 ビット バージョンと 64 ビット バージョン間の互換性
Rem  https://docs.microsoft.com/ja-jp/office/client-developer/shared/compatibility-between-the-32-bit-and-64-bit-versions-of-office
Rem
Rem Declaring API functions in 64 bit Office
Rem  https://www.jkp-ads.com/articles/apideclarations.asp
Rem
Rem ----------------------------------------------------------------------------------------------------
Rem
Rem ■更新履歴
Rem
Rem  2019/10/20 : Declare文から各種環境に対応したDeclare文へ変換する関数
Rem  2019/10/21 : 関数名から各種環境に対応したDeclare文を生成する関数
Rem
Rem ----------------------------------------------------------------------------------------------------
Rem
Rem ■使い方
Rem
Rem VBAソースコードのDeclare文を32/64bit対応に変換する関数
Rem
Rem  @name ConvertVBACodeDeclare
Rem
Rem  @param vbaCodeText VBAソースコード文字列（vbCrLf）
Rem
Rem  @return As String  VBAソースコード文字列
Rem
Rem  @example
Rem    IN  : 適当なソースコード文字列
Rem    OUT : VBA6/7 Win32/64 対応ソースコード（動作は保証しない）
Rem
Rem  対応できない例
Rem  ・APIによってはパラメータが変わっている事もある。
Rem  ・パラメータがLongからLongPtr/LongLongに変化して呼び出し側も変更の必要がある。
Rem  ・構造体の仕様が変わっている/未定義の場合がある。
Rem  ・Win32API_PtrSafe.txtに掲載されていない関数には対応していない。
Rem  ・GetWindowLong等はGetWindowLongPtrに変更しないと使えない。
Rem
Rem  Function ConvertVBACodeDeclare(vbaCodeText) As String
Rem   IN  : 適当なソースコード文字列
Rem   OUT : VBA6/7 Win32/64 対応ソースコード（動作は保証しない）
Rem
Rem
Rem Win32API関数名を羅列したテキストをDeclareに変換する関数
Rem
Rem  @name GetDeclareCodeByText
Rem
Rem  @param base_str       関数名だけの行の含まれたテキスト
Rem  @param useVBA6        VBA6対応コードを生成するか
Rem  @param useVBA7        VBA7対応コードを生成するか(32bit/64bit)
Rem
Rem  @return As String     Declare宣言文
Rem
Rem
Rem Win32API関数名を渡したら全対応のDeclare文を返却する関数
Rem
Rem  name GetDeclareCodeByProcName
Rem
Rem  @param procName       検索対象の関数名
Rem  @param useVBA6        VBA6対応コードを生成するか
Rem  @param useVBA7        VBA7対応コードを生成するか(32bit/64bit)
Rem  @param indent_level   インデント幅（2~）
Rem
Rem  @return As String Declare文　未発見時はprocName
Rem
Rem  @note VBA6,7両方を使用する場合だけディレクティブによる分岐が生成される
Rem
Rem
Option Explicit

Private dicPtrSafe_ As Dictionary
Private dicPtrSafe32_ As Dictionary
Private dicPtrSafe64_ As Dictionary

Private Sub Sample()
    Const TEST_FILE = "Win32API変換テスト.bas"
    Const PARAM_INDENT_LEVEL = 10
    
    Dim fso As New FileSystemObject

    '変換前
    Dim vbaCodeText
    vbaCodeText = fso.OpenTextFile(ThisWorkbook.Path & "\" & TEST_FILE, ForReading, False).ReadAll()
    
    '変換後
    Dim replacedText
    replacedText = ConvertVBACodeDeclare(vbaCodeText, PARAM_INDENT_LEVEL)
    
    '先頭40行だけイミディエイトへ出力
    Dim idxs
    idxs = kccFuncString.InStrAll(replacedText, vbCrLf)
    Debug.Print Left(replacedText, idxs(40))
    
    'ファイル出力
    fso.OpenTextFile(ThisWorkbook.Path & "\" & TEST_FILE & "_conv.txt", ForWriting, True).Write replacedText
End Sub

Rem Microsoft公式の宣言文を解析して見本を辞書に保持する
Rem  @param bit =  0 : bitに依存しない
Rem               32 : 32bit専用
Rem               64 : 64bit専用
Rem  @return As Dictionary paramに対応した辞書 (※keyは小文字統一）
Rem
Rem https://docs.microsoft.com/ja-jp/office/client-developer/shared/compatibility-between-the-32-bit-and-64-bit-versions-of-office
Private Property Get DicDeclareCode(bit) As Dictionary
    Const PTRSAFEFILE = "Win32API_PtrSafe.txt"
    Dim PtrSafePath As String
    PtrSafePath = ThisWorkbook.Path & "\" & PTRSAFEFILE
    Dim fso As New FileSystemObject
    
    If dicPtrSafe_ Is Nothing Then
        Set dicPtrSafe_ = New Dictionary
        Set dicPtrSafe32_ = New Dictionary
        Set dicPtrSafe64_ = New Dictionary
        
        If Not fso.FileExists(PtrSafePath) Then
            MsgBox "Not Found : " & PtrSafePath
            Exit Property
        End If
        
        Dim vbaCodeText
        vbaCodeText = fso.OpenTextFile(PtrSafePath, ForReading, False).ReadAll()
        Debug.Print "Successfully loaded the " & PtrSafePath
        
        Dim v, ProcName
        Dim nowIndent As Long: nowIndent = 0
        Dim vbaMode As Long: vbaMode = 0    '0,6,7
        Dim vba7Indent As Long: vba7Indent = 0
        Dim winMode As Long: winMode = 0    '0,32,64
        Dim win64Indent As Long: win64Indent = 0
        
        Dim i As Long
        
        'これらの関数はTxtで二重定義されているので許容する。
        Dim oklist
        oklist = Array("GetUserName", "GetComputerName", _
                        "GetCurrentProcess", "OpenProcessToken", _
                        "GetTokenInformation", "LookupAccountSid", _
                        "UnhookWindowsHookEx")
        For i = LBound(oklist) To UBound(oklist): oklist(i) = LCase(oklist(i)): Next
        
        For Each v In Split(vbaCodeText, vbCrLf)
            i = i + 1
            If v Like "[#]If*" Then nowIndent = nowIndent + 1
            If v Like "[#]If *VBA7* Then" Then vbaMode = 7: vba7Indent = nowIndent
            If v Like "[#]If *Win64* Then" Then winMode = 64: win64Indent = nowIndent
            If v = "#Else" And vbaMode = 7 And nowIndent = vba7Indent Then vbaMode = 6
            If v = "#Else" And winMode = 64 And nowIndent = win64Indent Then winMode = 32
            If v = "#End If" And nowIndent = vba7Indent Then vbaMode = 0: vba7Indent = 0
            If v = "#End If" And nowIndent = win64Indent Then winMode = 0: win64Indent = 0
            If v = "#End If" Then nowIndent = nowIndent - 1
            
            ProcName = GetDeclareProcName(v)
            ProcName = LCase(ProcName)
            If ProcName <> "" Then
                If winMode = 32 Then
                    dicPtrSafe32_.Add ProcName, v
                ElseIf winMode = 64 Then
                    dicPtrSafe64_.Add ProcName, v
                Else
                    If UBound(Filter(oklist, ProcName)) >= 0 And dicPtrSafe_.Exists(ProcName) Then
'                         二重定義を許容
'                         独自に追加した関数で重複が見つかった場合に検知したいので敢えてこうした。
                    ElseIf dicPtrSafe_.Exists(ProcName) Then
                        Debug.Print ProcName & "は二重定義？"
                        Stop
                    Else
'                         Debug.Print procName
                        dicPtrSafe_.Add ProcName, v
                    End If
                End If
            End If
        Next
    End If
    
    If bit = 32 Then
        Set DicDeclareCode = dicPtrSafe32_
    ElseIf bit = 64 Then
        Set DicDeclareCode = dicPtrSafe64_
    Else
        Set DicDeclareCode = dicPtrSafe_
    End If
End Property

Rem VBAソースコードのDeclare文を32/64bit対応に変換
Rem
Rem  @name ConvertVBACodeDeclare
Rem
Rem  @param vbaCodeText VBAソースコード文字列（vbCrLf）
Rem
Rem  @return As String  VBAソースコード文字列
Rem
Rem  @example
Rem    IN  : 適当なソースコード文字列
Rem    OUT : VBA6/7 Win32/64 対応ソースコード（動作は保証しない）
Public Function ConvertVBACodeDeclare(vbaCodeText, indent_level As Long) As String
    If vbaCodeText = "" Then Exit Function

    Dim i As Long, j As Long
    Dim v
    Dim vbaLines
    vbaLines = Split(vbaCodeText, vbCrLf)
    
    Dim IsCommented() As Boolean
    ReDim IsCommented(LBound(vbaLines) To UBound(vbaLines))
    
    Dim SavedIndent1()
    ReDim SavedIndent1(LBound(vbaLines) To UBound(vbaLines))
    
    Dim SavedIndent2()
    ReDim SavedIndent2(LBound(vbaLines) To UBound(vbaLines))
    
    '宣言エリア最終行を特定
    Dim FinalRow As Long: FinalRow = 0
    For i = LBound(vbaLines) To UBound(vbaLines)
        v = vbaLines(i)
        If (v Like "*Sub*" Or v Like "*Function*" Or _
            v Like "*Property Get*" Or v Like "*Property Set*") And _
            (Not v Like "*Declare*") And (Not Trim(v) Like "'*") Then
            Exit For
        End If
    Next
    FinalRow = i - 1
    
    'コメントとインデント除去
    For i = LBound(vbaLines) To FinalRow
        v = vbaLines(i)
        SavedIndent1(i) = kccFuncString.InStrRept(v, " ")
        v = Trim(v)
        If v Like "'*" Then
            v = Mid(v, 2, Len(v))
            IsCommented(i) = True
            SavedIndent2(i) = kccFuncString.InStrRept(v, " ")
            v = Trim(v)
        End If
        vbaLines(i) = v
    Next
    
    'ステートメント改行を連結
    Dim vNow, vPrev
    For i = FinalRow To LBound(vbaLines) + 1 Step -1
        vNow = vbaLines(i)
        vPrev = vbaLines(i - 1)
        If vPrev Like "* _" Then
            vPrev = Left(vPrev, Len(vPrev) - 1) & Trim(vNow)
            vPrev = Replace(vPrev, "  ", " ")
            vNow = ""
        End If
        vbaLines(i) = vNow
        vbaLines(i - 1) = vPrev
    Next
    
    '---ここまで前処理
    
    Dim nowIndent As Long: nowIndent = 0
    Dim vbaMode As Long: vbaMode = 0    '0,6,7
    Dim vba7Indent As Long: vba7Indent = 0
    Dim winMode As Long: winMode = 0    '0,32,64
    Dim win64Indent As Long: win64Indent = 0
    
    Dim arr
    For i = LBound(vbaLines) To FinalRow
        v = Trim(vbaLines(i))
        
        If v Like "[#]If*" Then nowIndent = nowIndent + 1
        If v Like "[#]If *VBA7* Then" Then vbaMode = 7: vba7Indent = nowIndent
        If v Like "[#]If *Win64* Then" Then winMode = 64: win64Indent = nowIndent
        If v = "#Else" And vbaMode = 7 And nowIndent = vba7Indent Then vbaMode = 6
        If v = "#Else" And winMode = 64 And nowIndent = win64Indent Then winMode = 32
        If v = "#End If" And nowIndent = vba7Indent Then vbaMode = 0: vba7Indent = 0
        If v = "#End If" And nowIndent = win64Indent Then winMode = 0: win64Indent = 0
        If v = "#End If" Then nowIndent = nowIndent - 1
        
        '既存のDeclare文は正しいものと仮定してもう一方の文を追加する
        'VBA7ディレクティブ内に記述されている時は既に対処済みと判断し変換は行わない
        If v Like "*Declare *" Then
            If v Like "*Declare PtrSafe *" Then
                'VBA7宣言文
                If vbaMode = 0 Then
                    arr = Array("", _
                                "#If VBA7 Then", _
                                kccFuncString.InsertIndent(InsertDeclareIndent(v, indent_level)), _
                                "#Else", _
                                kccFuncString.InsertIndent(ReplaceDeclareTo6(v, indent_level)), _
                                "#End If")
                    v = Join(arr, vbCrLf)
                End If
            Else
                'VBA6(64bit非対応)宣言文
                If vbaMode = 0 Then
                    '非対応
                    arr = Array("", _
                                "#If VBA7 Then", _
                                kccFuncString.InsertIndent(ReplaceDeclareTo7(v, indent_level)), _
                                "#Else", _
                                kccFuncString.InsertIndent(InsertDeclareIndent(v, indent_level)), _
                                "#End If")
                    v = Join(arr, vbCrLf)
                ElseIf vbaMode = 7 Then
                    '対応漏れ(ディレクティブ内なのにPtrSafeしてない）
                    v = ReplaceDeclareTo7(v, indent_level)
                End If
            End If
        End If
        
        vbaLines(i) = v
    Next
    
    '---ここから後処理
    
    'コメントとインデントを復元
    For i = LBound(vbaLines) To FinalRow
        If vbaLines(i) <> "" Then
            v = vbaLines(i)
            v = kccFuncString.InsertString(v, String(SavedIndent2(i), " "))
            v = kccFuncString.InsertString(v, IIf(IsCommented(i), "'", ""))
            v = kccFuncString.InsertString(v, String(SavedIndent1(i), " "))
            vbaLines(i) = v
        End If
    Next
    
    Dim mergeVBA As String
    mergeVBA = Join(vbaLines, vbCrLf)
    For i = 1 To 10
        mergeVBA = Replace(mergeVBA, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
    Next
    
    ConvertVBACodeDeclare = mergeVBA
    
End Function

Private Sub Test_GetDeclareCodeByText()
'     Const TESTDATA = "GetWindowLong" & vbCrLf & "setwindowlong"
'     Const TESTDATA = "GetWindowLong"
    Const TESTDATA = "getwindow" & vbCrLf & "hoge"
    Debug.Print
    Debug.Print TESTDATA
'     Debug.Print
'     Debug.Print GetDeclareCodeByText(TESTDATA, False, False, 10)
    Debug.Print
    Debug.Print GetDeclareCodeByText(TESTDATA, True, True, 10)
    Debug.Print
    Debug.Print GetDeclareCodeByText(TESTDATA, False, True, 10)
    Debug.Print
    Debug.Print GetDeclareCodeByText(TESTDATA, True, False, 10)
End Sub

Rem Win32API関数名を羅列したテキストをDeclareに変換する関数
Rem
Rem  @name GetDeclareCodeByText
Rem
Rem  @param base_str       関数名だけの行の含まれたテキスト
Rem  @param useVBA6        VBA6対応コードを生成するか
Rem  @param useVBA7        VBA7対応コードを生成するか(32bit/64bit)
Rem
Rem  @return As String     Declare宣言文
Rem
Public Function GetDeclareCodeByText(base_str, useVBA6 As Boolean, useVBA7 As Boolean, indent_level As Long) As String
    Dim Rows, i
    Rows = Split(base_str, vbCrLf)
    For i = LBound(Rows) To UBound(Rows)
        Rows(i) = GetDeclareCodeByProcName(Rows(i), useVBA6, useVBA7, indent_level)
    Next
    GetDeclareCodeByText = Join(Rows, vbCrLf)
End Function

Rem 適当なDeclare文からVBA6対応コードに置換(不完全)
Rem  単純にVBA6非対応の文字を取り除くだけなので正しい書式になるとは限らない。
Private Function ReplaceDeclareTo6(ByVal base_str, Optional indent_level As Long = 0) As String
    base_str = Replace(base_str, "PtrSafe ", "")
    base_str = Replace(base_str, "LongPtr", "Long")
    ReplaceDeclareTo6 = InsertDeclareIndent(base_str, indent_level)
End Function

Rem 適当なDeclare文からVBA7対応(32/64bit両対応)コードに置換
Rem  「Win32API_PtrSafe.txt」を参照するため精度は高いがそのまま動かせるとは限らない。
Rem  元々の名前付き引数は保持されない
Private Function ReplaceDeclareTo7(ByVal base_str, Optional indent_level As Long = 0) As String
    Dim ProcName: ProcName = GetDeclareProcName(base_str)
    
    Dim lifeName: lifeName = ""
    If InStr(base_str, "Private") > 0 Then: lifeName = "Private "
    If InStr(base_str, "Public") > 0 Then: lifeName = "Public "
    If InStr(base_str, "Dim") > 0 Then: lifeName = "Dim "
    
    ReplaceDeclareTo7 = GetDeclareVBA7(ProcName, lifeName, indent_level)
End Function

Private Sub Test_ReplaceDeclareTo7()
    Const Teststr = "Declare Function WindowFromPoint Lib ""user32"" Alias ""WindowFromPoint"" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr"
    Debug.Print
    Debug.Print Teststr
    Debug.Print "  ↓"
    Debug.Print ReplaceDeclareTo7(Teststr, 10)
End Sub

Rem Win32API関数名を渡したら全対応のDeclare文を返却する関数
Rem
Rem  @name GetDeclareCodeByProcName
Rem
Rem  @param procName       検索対象の関数名
Rem  @param useVBA6        VBA6対応コードを生成するか
Rem  @param useVBA7        VBA7対応コードを生成するか(32bit/64bit)
Rem  @param indent_level   インデント幅（2~）
Rem
Rem  @return As String Declare文　未発見時はprocName
Rem
Rem  @note VBA6,7両方を使用する場合だけディレクティブによる分岐が生成される
Rem
Public Function GetDeclareCodeByProcName( _
        ProcName, useVBA6 As Boolean, useVBA7 As Boolean, _
        Optional indent_level As Long = 2) As String
    Dim arr
    If useVBA6 And useVBA7 Then
        arr = Array("", _
                "#If VBA7 Then", _
                kccFuncString.InsertIndent(GetDeclareVBA7(ProcName, "", indent_level)), _
                "#Else", _
                kccFuncString.InsertIndent(GetDeclareVBA6(ProcName, "", indent_level)), _
                "#End If")
    ElseIf useVBA7 Then
        arr = Array(GetDeclareVBA7(ProcName, "", indent_level))
    ElseIf useVBA6 Then
        arr = Array(GetDeclareVBA6(ProcName, "", indent_level))
    Else
        Err.Raise 9999, , "VBA6 VBA7両方非対応になっている"
    End If
    GetDeclareCodeByProcName = Join(arr, vbCrLf)
End Function

Rem Win32API関数名を渡したらVBA6(～Excel2007)対応のDeclare文を返す関数
Rem  32bit版の記法を改変することで生成
Public Function GetDeclareVBA6(ProcName, lifeName, Optional indent_level As Long = 0) As String
    Dim pn As String: pn = LCase(ProcName)
    If DicDeclareCode(0).Exists(pn) Then
        GetDeclareVBA6 = InsertDeclareIndent(lifeName & DicDeclareCode(0)(pn), indent_level)
    ElseIf DicDeclareCode(32).Exists(pn) Then
        GetDeclareVBA6 = InsertDeclareIndent(lifeName & DicDeclareCode(32)(pn), indent_level)
    Else
        GetDeclareVBA6 = ProcName
    End If
    GetDeclareVBA6 = Replace(GetDeclareVBA6, "PtrSafe ", "")
    GetDeclareVBA6 = Replace(GetDeclareVBA6, "LongPtr", "Long")
End Function

Private Sub Test_GetDeclareVBA7()
    Const TESTDATA = "getwindow"
    Debug.Print
    Debug.Print TESTDATA
    Debug.Print "  ↓"
    Debug.Print GetDeclareVBA7(TESTDATA, "", 2)
End Sub

Rem Win32API関数名を渡したらVBA7(Excel 2010～2016 32/64)対応のDeclare文を返す関数
Rem
Rem  @param procName       関数名
Rem  @param lifeName       公開範囲（空欄、Private 、Public ）
Rem  @param indent_level   インデント幅（0~）
Rem
Rem  @return As String     VBA7用のDeclare宣言文
Rem
Rem  @example
Rem    IN : GetWindow
Rem   OUT : Declare PtrSafe Function GetWindow Lib "user32" ( _
Rem                 ByVal hWnd As LongPtr, _
Rem                 ByVal wCmd As Long _
Rem                 ) As LongPtr
Rem
Public Function GetDeclareVBA7(ProcName, lifeName, Optional indent_level As Long = 0) As String
    Dim pn As String: pn = LCase(ProcName)
    Dim arr
    If DicDeclareCode(0).Exists(pn) Then
        GetDeclareVBA7 = InsertDeclareIndent(lifeName & DicDeclareCode(0)(pn), indent_level)
    ElseIf DicDeclareCode(64).Exists(pn) And DicDeclareCode(32).Exists(pn) Then
        arr = Array("#If Win64 Then", _
                    InsertDeclareIndent(kccFuncString.InsertIndent(lifeName & DicDeclareCode(64)(pn)), indent_level), _
                    "#Else", _
                    InsertDeclareIndent(kccFuncString.InsertIndent(lifeName & DicDeclareCode(32)(pn)), indent_level), _
                    "#End If")
        GetDeclareVBA7 = Join(arr, vbCrLf)
    ElseIf DicDeclareCode(64).Exists(pn) Then
        arr = Array("#If Win64 Then", _
                    InsertDeclareIndent(kccFuncString.InsertIndent(lifeName & DicDeclareCode(64)(pn)), indent_level), _
                    "#End If")
        GetDeclareVBA7 = Join(arr, vbCrLf)
    ElseIf DicDeclareCode(32).Exists(pn) Then
        'GetWindowLong等は64bit版の関数が無い。GetWindowLongPtrへの置き換えが必要。
        arr = Array("#If Win64 Then", _
                    "#Else", _
                    InsertDeclareIndent(kccFuncString.InsertIndent(lifeName & DicDeclareCode(32)(pn)), indent_level), _
                    "#End If")
        GetDeclareVBA7 = Join(arr, vbCrLf)
    Else
        GetDeclareVBA7 = ProcName
    End If
End Function

Rem 宣言文から関数名を取得
Rem
Rem  @param base_str   入力文字列（宣言文）
Rem  @return As String 関数名
Rem
Rem  @example
Rem    IN : Private Declare Function ReleaseDC Lib....
Rem   OUT : ReleaseDC
Rem
Private Function GetDeclareProcName(ByVal base_str) As String
    Dim sIdx As Long: sIdx = 0
    Dim eIdx As Long: eIdx = 0
    Dim fIdx As Long: fIdx = 0
    sIdx = InStr(base_str, "Sub "): If sIdx > 0 Then fIdx = sIdx + 4
    sIdx = InStr(base_str, "Function "): If sIdx > 0 Then fIdx = sIdx + 9
    If fIdx = 0 Then Exit Function
    eIdx = InStr(fIdx, base_str, " ")
    If eIdx = 0 Then eIdx = Len(base_str)
    GetDeclareProcName = Mid(base_str, fIdx, eIdx - fIdx)
End Function

Private Sub Test_GetDeclareProcName()
    Const s = "Private Declare PtrSafe Function ReleaseDC Lib ""user32"" ( ByVal hWnd As Long, ByVal hdc As Long ) As Long"
    Debug.Print GetDeclareProcName(s)
End Sub

Rem 宣言文のパラメータの自動改行とインデント
Rem
Rem  @param base_str       変換元文字列(Sub,Function,Property,Declare)
Rem  @param indent_level   先頭行以外インデントする幅(4*(2~#))
Rem                        -1の時、自動改行もインデントも行わない
Rem  @param delimiter      改行文字列（既定：CR+LF）
Rem
Rem  @return As String     整形後の文字列
Rem
Rem  @example
Rem    IN :
Rem         Function InsertDeclareIndent(ByVal base_str, Optional indent_level = 1, Optional delimiter = vbCrLf) As String
Rem   OUT :
Rem         Function InsertDeclareIndent( _
Rem                 ByVal base_str, _
Rem                 Optional indent_level = 1, _
Rem                 Optional delimiter = vbCrLf _
Rem                 ) As String
Rem
Private Function InsertDeclareIndent(ByVal base_str, Optional indent_level = 2, Optional Delimiter = vbCrLf) As String
    If InStr(base_str, "()") > 0 Then InsertDeclareIndent = base_str: Exit Function
    If indent_level < 0 Then InsertDeclareIndent = base_str: Exit Function
    base_str = Replace(base_str, "(", "( _" & Delimiter)
    base_str = Replace(base_str, ",", ", _" & Delimiter)
    base_str = Replace(base_str, ")", " _" & Delimiter & ")")
    base_str = Join(kccFuncString.TrimArray(Split(base_str, Delimiter)), Delimiter)
    InsertDeclareIndent = Replace(base_str, Delimiter, Delimiter & String(4 * indent_level, " "))
End Function

Private Sub Test_InsertDeclareIndent()
    Const TESTDATA = "Function InsertDeclareIndent(ByVal base_str, Optional indent_level = 1, Optional delimiter = vbCrLf) As String"
    Debug.Print InsertDeclareIndent(TESTDATA)
End Sub
