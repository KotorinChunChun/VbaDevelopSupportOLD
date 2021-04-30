Attribute VB_Name = "VbeDevelop"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        ExtDevelop
Rem
Rem  @description   開発環境VBE用のモジュール
Rem
Rem  @update        2020/08/06
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem  @references
Rem    Microsoft Visual Basic for Applications Extensibility 5.3
Rem
Rem --------------------------------------------------------------------------------
Rem  @refModules
Rem    kccFuncString
Rem    VbProcInfo
Rem      - VbProcParamInfo
Rem
Rem --------------------------------------------------------------------------------
Rem  @history
Rem    2020/08/01 再整備
Rem
Rem --------------------------------------------------------------------------------
Rem  @note
Rem     msdn
Rem       VBA で起動中のすべての Excel インスタンスを完全に取得したい
Rem       https://social.msdn.microsoft.com/Forums/ja-JP/7a46a3c9-f904-4fb0-a205-6112fba51fe6/vba-excel-?forum=vbajp
Rem
Rem     OKwave
Rem       別インスタンスのブック（個人用マクロブック以外）をすべて閉じる
Rem       MREXCEL.COM > Forum > Question Forums > Excel Questions > GetObject and HWND
Rem       https://okwave.jp/qa/q9196890.html
Rem
Rem     Qita
Rem       【ExcelVBA】VBAコードの情報や概要をシートに一覧出力する
Rem       https://qiita.com/Mikoshiba_Kyu/items/46b7243eb576848b3e55
Rem
Rem       excel Access VBA ２つか1つの設定でVBAの参照設定を完了するマクロ
Rem       https://qiita.com/Q11Q/items/67226e7c8b9def529668
Rem
Rem       VBAでExcelを使う
Rem       https://qiita.com/palglowr/items/04250eb1a8a873fbf9d2
Rem
Rem       GetRunningObjectTable
Rem       https://foren.activevb.de/forum/vb-classic/thread-409498/beitrag-409498/API-GetRunningObjectTable/
Rem
Rem       VBA 標準モジュールのマクロを読み取って起動時にVBEの
Rem       メニューに自動登録するアドインを自作する
Rem       https://thom.hateblo.jp/entry/2016/11/12/081256
Rem
Rem --------------------------------------------------------------------------------
Option Explicit
Option Private Module

Private Declare PtrSafe Function GetKeyboardState _
                        Lib "User32" (pbKeyState As Byte) As Long
Private Declare PtrSafe Function SetKeyboardState _
                        Lib "User32" (lppbKeyState As Byte) As Long
Private Declare PtrSafe Function PostMessage _
                        Lib "User32" Alias "PostMessageA" ( _
                        ByVal hWnd As LongPtr, ByVal wMsg As Long, _
                        ByVal wParam As Long, ByVal lParam As LongPtr _
                        ) As Long

Private Declare PtrSafe Function FindWindow Lib "User32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As LongPtr
    
Private Declare PtrSafe Function FindWindowEx Lib "User32" Alias "FindWindowExA" ( _
    ByVal hwndParent As LongPtr, _
    ByVal hwndChildAfter As LongPtr, _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As LongPtr
    
Private Declare PtrSafe Function GetWindow Lib "User32" ( _
    ByVal hWnd As LongPtr, _
    ByVal wCmd As Long) As LongPtr

                        
Private Const WM_KEYDOWN As Long = &H100
Private Const KEYSTATE_KEYDOWN As Long = &H80

Private Enum eRecord
    モジュール名 = 1
    モジュールタイプ
    プロシージャ名
    プロシージャタイプ
    行数
    引数
    戻り値
    概要
End Enum

Rem Sub Test_TextParse_VbProcedure()
Rem     Dim v
Rem '    v = TextParse_VbProcedure("Property Get RowKeys() As String()")
Rem '    v = TextParse_VbProcedure("Property Get RowKeys(p As Variant) As String()")
Rem     v = TextParse_VbProcedure("Property Get RowKeys(p As Variant, q As Variant()) As String()")
Rem     DpP "", v
Rem End Sub
Rem
Rem Sub Test_TextParse_VbProcedure()
Rem     Dim v
Rem '    v = TextParse_VbProcedure("Property Get RowKeys() As String()")
Rem '    v = TextParse_VbProcedure("Property Get RowKeys(p As Variant) As String()")
Rem     v = TextParse_VbProcedure("Property Get RowKeys(p As Variant, q As Variant()) As String()")
Rem     DpP "", v
Rem End Sub

Private fso As New FileSystemObject

Rem 本モジュール用終了時初期化処理
Public Sub Terminate()
    'CustomUI Import/Export用のZIP展開一時フォルダの初期化
    Call kccFuncZip.DeleteTempFolder
End Sub

Rem アクティブなプロジェクトの保存フォルダを開く
Public Sub OpenProjectFolder()
On Error Resume Next
    Dim fn: fn = Application.VBE.ActiveVBProject.FileName
    kccFuncWindowsProcess.ShellExplorer fn, True
End Sub

Rem WEBサイトを開く（関連付けプログラムで開く）
Public Sub OpenWebSite(URL)
    kccFuncWindowsProcess.OpenAssociationAPI URL
End Sub

Sub Test_VBP()
    Dim objFilePath As kccPath: Set objFilePath = kccPath.Init(Application.VBE.ActiveVBProject)
    Dim obj
    Set obj = objFilePath.VBProject
    Debug.Print obj.Name
    Stop
End Sub

Rem 現在アクティブなプロジェクトのワークブックを閉じる
Public Sub CloseProject()
    Dim objFilePath As kccPath: Set objFilePath = kccPath.Init(Application.VBE.ActiveVBProject)
    If objFilePath Is Nothing Then Exit Sub
    objFilePath.Workbook.Close
End Sub

Rem アクティブブックのソースコードのプロシージャ一覧を新規ブックへ出力
Public Sub VbeProcInfo_Output()
    Dim objFilePath As kccPath: Set objFilePath = kccPath.Init(Application.VBE.ActiveVBProject)
    
    'プロシージャ一覧を取得して二次元配列を取得する処理
    Dim data
    data = VbeProcInfo_GetTable(objFilePath.Workbook.VBProject)
    
    '二次元配列をブックに出力する処理
    
    'ここまでしてもブックのメモリが開放されない謎の減少発生中
    Dim outWb As Workbook:
'    Set outWb = ActiveWorkbook
    Set outWb = Workbooks.Add
    Dim outWs As Worksheet: Set outWs = outWb.Worksheets(1)
    Call VbeProcInfo_OutputWorksheet(data, outWs)
    DoEvents
    Set outWs = Nothing
    Set outWb = Nothing
End Sub

Rem ソースコードのプロシージャ一覧を指定シートへ出力
Private Function VbeProcInfo_GetTable(source_vbp As VBProject) As Variant
    Dim dicProcInfo As New Dictionary
    Dim i As Long
    Dim dKey
  
    'ブックの全モジュールを処理
    For i = 1 To source_vbp.VBComponents.Count
        Dim dic As Dictionary
        Set dic = GetProcInfoDictionary(source_vbp.VBComponents(i).CodeModule)
        For Each dKey In dic.Keys
            dicProcInfo.Add dKey, dic(dKey)
        Next
        Set dic = Nothing
    Next
    If dicProcInfo.Count = 0 Then MsgBox "VBAがありません。": Exit Function
  
    Dim data
    data = Array("モジュール", "行位置", "スコープ", "種別", "プロシージャー", "引数", "戻り値", "コメント", "宣言文")
    data = WorksheetFunction.Transpose(data)
    ReDim Preserve data(LBound(data) To UBound(data, 1), 1 To dicProcInfo.Count + 1)
    data = WorksheetFunction.Transpose(data)
    
    i = 2
    For Each dKey In dicProcInfo.Keys
        Dim v As VbProcInfo
        Set v = dicProcInfo(dKey)
        data(i, 1) = v.ModName
        data(i, 2) = v.LineNo
        data(i, 3) = v.Scope
        data(i, 4) = v.ProcKindName
        data(i, 5) = v.ProcName
        data(i, 6) = v.ParamsToString(vbLf)
        data(i, 7) = v.ReturnToString
        data(i, 8) = "'" & v.Comment
        data(i, 9) = v.Source
        Set v = Nothing
        i = i + 1
    Next

    Set dicProcInfo = Nothing
    VbeProcInfo_GetTable = data
End Function

Rem プロシージャ一覧二次元配列データをシートに出力する
Rem さらにJ,K列に数式を追加する
Private Sub VbeProcInfo_OutputWorksheet(data, output_ws As Worksheet)

    'Dictionaryよりシートに出力
    output_ws.Name = "プロシージャ一覧"
    output_ws.Parent.Activate
    output_ws.Parent.Windows(1).WindowState = xlMaximized
    
    With output_ws
        .Cells.Clear
        .Cells(1, 1).Resize(UBound(data, 1), UBound(data, 2)).Value = data
        
        '宣言文検証用の式
        .Range("J1").Value = "宣言文2"
        .Range("J2").FormulaR1C1 = "=RC[-7]&"" ""&RC[-6]&"" ""&RC[-5]&""(""&SUBSTITUTE(RC[-4],""" & Chr(10) & ""","", "")&"")""&IF(RC[-3]="""","""","" As ""&RC[-3])"
        .Range("K1").Value = "チェック"
        .Range("K2").FormulaR1C1 = "=RC[-2]=RC[-1]"
        '1行余分にフィルされる
        .Range("J2:K2").AutoFill Destination:=ResizeOffset(.UsedRange.Columns("J:K"), 1)
        
        .Range("A2").Select
        .Parent.Windows(1).FreezePanes = True
        .Cells.AutoFilter
        .Columns("A:K").EntireColumn.AutoFit
        .Columns("H:J").ColumnWidth = 16
        .Cells.WrapText = False
    End With
    
End Sub

Rem Rangeを指定座標だけOffsetしつつResizeする（先頭側を削るResize）
Rem
Rem  @param rng         対象Range
Rem  @param offsetRow   先頭からオフセット縮小する行数
Rem  @param offsetCol   先頭からオフセット縮小する列数
Rem
Rem  @return As Range   変形後のRange
Rem
Function ResizeOffset(rng As Range, Optional offsetRow As Long, _
                                    Optional offsetCol As Long) As Range
    Set ResizeOffset = Intersect(rng, rng.Offset(offsetRow, offsetCol))
End Function

'オートフィルタ方式
Sub Test_ResizeOffset_AutoFilter()
    Const TARGET_COL = "C:D"
    ResizeOffset(ActiveSheet.AutoFilter.Range.Columns(TARGET_COL), 1).Select
End Sub

Sub Test_ResizeOffset()
    Const HEAD_ROW = 3, TARGET_COL = "C:D"
    
    With ToWorksheet(ActiveSheet)
        Dim rng As Range
        
        'UsedRange方式
'        Set rng = ResizeOffset(.UsedRange.Columns(TARGET_COL), HEAD_ROW - .UsedRange.Row + 1)
        
        'オートフィルタ方式
'        Set rng = ResizeOffset(.AutoFilter.Range.Columns(TARGET_COL), HEAD_ROW - .AutoFilter.Range.Row + 1)

        'CurrentRegion方式
'        Set rng = Range("B3").CurrentRegion
'        Set rng = Intersect(rng, rng.Offset(1)).Columns("C:D")

        rng.Select
    End With
End Sub

        
        'CurrentRegion方式
'        Set rng = .Range(TARGET_COL).Cells(HEAD_ROW, 1).CurrentRegion
'        Set rng = ResizeOffset(rng.Columns(TARGET_COL), HEAD_ROW - rng.Row + 1)
        

'オートフィルタの有無
'UsedRange外の先頭行列の有無


'    Set OffsetResize = rng.Offset(offsetRow, offsetCol).Resize( _
'                            rng.Rows.CountLarge - offsetRow, _
'                            rng.Columns.CountLarge - offsetCol)

Rem アクティブシートのA列にプロシージャ情報が記載されているものとする
Private Sub プロシージャ一覧から引数を分解する()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim data: data = ws.Range(ws.Cells(1, 1), ws.UsedRange)
    
    Dim i As Long
    For i = 2 To UBound(data)
Rem         Debug.Print data(i, 1)
        
        Dim proc As VbProcInfo
        Set proc = VbProcInfo.Init("", "", "", 0, "", data(i, 1))
Rem         Debug.Print proc.ToString
        data(i, 2) = proc.ParamsToString(vbLf)
        data(i, 3) = proc.ReturnToString
        Set proc = Nothing
    Next
    
    ws.Cells(1, 1).Resize(UBound(data, 1), UBound(data, 2)).Value = data
End Sub

Rem Dictionaryにプロシージャー・プロパティ情報を格納
Public Function GetProcInfoDictionary(ByVal objCodeModule As CodeModule) As Dictionary
    Dim dic As Dictionary: Set dic = New Dictionary
    Dim sMod As String: sMod = objCodeModule.Name
    
    Dim codeLine As Long: codeLine = 1
    Do While codeLine <= objCodeModule.CountOfLines
    
        Dim sProcName As String
        Dim sProcKey As String
        Dim iProcKind As Long
        sProcName = objCodeModule.ProcOfLine(codeLine, iProcKind)
        sProcKey = sMod & "." & sProcName
        
        If sProcName <> "" Then
            If isProcLine(objCodeModule.Lines(codeLine, 1), sProcName) Then
                If Not dic.Exists(sProcKey) Then
                    Dim cProcInfo As VbProcInfo
                    Set cProcInfo = VbProcInfo.Init( _
                                        sMod, _
                                        sProcName, _
                                        iProcKind, _
                                        codeLine, _
                                        getProcComment(codeLine, objCodeModule), _
                                        getProcSource(codeLine, objCodeModule) _
                                        )
                    dic.Add sProcKey, cProcInfo
                End If
            End If
        End If
        codeLine = codeLine + 1
    Loop
    
    Set GetProcInfoDictionary = dic
End Function

Rem プロシージャー・プロパティ定義行かの判定
Private Function isProcLine(ByVal strLine As String, _
                            ByVal ProcName As String) As Boolean
    strLine = " " & Trim(strLine)
    Select Case True
        Case Left(strLine, 1) = " '"
            isProcLine = False
        Case Left(strLine, 1) = " Rem"
            isProcLine = False
        Case strLine Like "* Sub " & ProcName & "(*"
            isProcLine = True
        Case strLine Like "* Sub " & ProcName & " _"
            isProcLine = True
        Case strLine Like "* Function " & ProcName & "(*"
            isProcLine = True
        Case strLine Like "* Function " & ProcName & " _"
            isProcLine = True
        Case strLine Like "* Property * " & ProcName & "(*"
            isProcLine = True
        Case strLine Like "* Property * " & ProcName & " _"
            isProcLine = True
        Case Else
            isProcLine = False
    End Select
End Function

Rem Dictionaryにプロシージャー・プロパティ情報を格納
Public Function GetDecInfoDictionary(ByVal objCodeModule As CodeModule) As Dictionary
    Dim dic As Dictionary: Set dic = New Dictionary
    Dim sMod As String: sMod = objCodeModule.Name
    
    Dim codeLine As Long: codeLine = 1
    Do While codeLine <= objCodeModule.CountOfDeclarationLines()
        Dim strLine As String: strLine = objCodeModule.Lines(codeLine, 1)
        If isDecLine(strLine) Then
            dic.Add sMod & "." & codeLine & ":" & strLine, strLine
        End If
        codeLine = codeLine + 1
    Loop
    
    Set GetDecInfoDictionary = dic
End Function

Rem 宣言部で必要なデータか確認
Rem 予め  objCodeModule.CountOfDeclarationLines 判定を済ませた行であること。
Private Function isDecLine(ByVal strLine As String) As Boolean
    strLine = Trim(strLine)
    If Len(strLine) = 0 Then Exit Function
    
    Select Case Split(strLine, " ")(0)
        Case "Private", "Public", "Friend", "Dim", "Const", "Declare", "Type", "Enum", "'", "Rem"
            isDecLine = True
        Case "Option"
            isDecLine = False
        Case Else
            isDecLine = False
    End Select
End Function

Rem 文字列がコメント行か
Private Function isComment(ByVal strLine As String) As Boolean
    strLine = Trim(strLine)
    If Len(strLine) = 0 Then Exit Function
    
    Select Case True
        Case strLine Like "'*"
            isComment = True
        Case strLine = "Rem" Or strLine Like "Rem *"
            isComment = True
        Case Else
            isComment = False
    End Select
End Function

Rem 継続行( _)全てを連結した文字列で返す
Rem コロンやコメント以降は消す
Private Function getProcSource(ByRef codeLine As Long, _
                               ByVal aCodeModule As Object) As String
    getProcSource = ""
    Dim sTemp As String
    Do
        sTemp = Trim(aCodeModule.Lines(codeLine, 1))
        If Right(aCodeModule.Lines(codeLine, 1), 2) = " _" Then
            sTemp = Left(sTemp, Len(sTemp) - 1)
        End If
        getProcSource = getProcSource & sTemp
        If Right(aCodeModule.Lines(codeLine, 1), 2) <> " _" Then Exit Do
        codeLine = codeLine + 1
    Loop
    If InStr(getProcSource, ":") > 0 Then getProcSource = Left(getProcSource, InStr(getProcSource, ":") - 1)
    If InStr(getProcSource, "'") > 0 Then getProcSource = Left(getProcSource, InStr(getProcSource, "'") - 1)
    getProcSource = Trim(getProcSource)
End Function

Rem プロシージャーの直前のコメントを取得
Private Function getProcComment(ByVal codeLine As Long, _
                                ByVal aCodeModule As Object) As String
    getProcComment = ""
    codeLine = codeLine - 1
    If codeLine <= 0 Then Exit Function
    Do
        Dim strLine As String: strLine = Trim(aCodeModule.Lines(codeLine, 1))
        If Not strLine Like "'*" And Not strLine Like "Rem*" Then Exit Do
        If getProcComment <> "" Then getProcComment = vbLf & getProcComment
        getProcComment = aCodeModule.Lines(codeLine, 1) & getProcComment
        codeLine = codeLine - 1
    Loop
End Function

Private Sub ListUpProcs()
    Dim trgBook As Workbook: Set trgBook = ActiveWorkbook
    Dim trgSheet As Worksheet: Set trgSheet = trgBook.Worksheets.Add

    On Error GoTo hundler
    trgSheet.Name = "Procs"
    On Error GoTo 0

    'ヘッダーレコードをセットする
    Dim procRecords As Collection: Set procRecords = New Collection
    Dim procRecord(1 To 8) As String 'リストの列数
    procRecord(eRecord.モジュール名) = "モジュール名"
    procRecord(eRecord.モジュールタイプ) = "モジュールタイプ"
    procRecord(eRecord.プロシージャ名) = "プロシージャ名"
    procRecord(eRecord.プロシージャタイプ) = "プロシージャタイプ"
    procRecord(eRecord.行数) = "行数"
    procRecord(eRecord.引数) = "引数"
    procRecord(eRecord.戻り値) = "戻り値"
    procRecord(eRecord.概要) = "概要"
    procRecords.Add procRecord

    'Moduleを順次処理する
    Dim module As Object
    For Each module In trgBook.VBProject.VBComponents

        'モジュール名をセットする
        procRecord(eRecord.モジュール名) = module.Name

        'モジュールタイプをセットする
        procRecord(eRecord.モジュールタイプ) = FIX_MODULE_TYPE(module)

        'Module内のProcedure一覧をコレクションする
        Dim cModule As Object: Set cModule = module.CodeModule
        Dim procNames As Collection: Set procNames = COLLECT_PROCNAMES_IN_MODULE(cModule)

        'Procedureの内容を順次処理する
        Dim ProcName As Variant, procTop As String
        For Each ProcName In procNames

            'プロシージャ名をセットする
            procRecord(eRecord.プロシージャ名) = ProcName

            'プロシージャの1行目を取得する
            procTop = SET_PROC_TOP(CStr(ProcName), cModule)

            'プロシージャタイプをセットする
            procRecord(eRecord.プロシージャタイプ) = FIX_PROC_TYPE(CStr(ProcName), procTop)

            '行数をセットする
            procRecord(eRecord.行数) = cModule.ProcCountLines(ProcName, 0)

            '引数をセットする
Rem             procRecord(eRecord.引数) = FIX_PROC_ARGS(CStr(ProcName), procTop)

            '戻り値をセットする
Rem             procRecord(eRecord.戻り値) = FIX_PROC_RETURN(CStr(ProcName), procTop)

            '概要をセットする
            procRecord(eRecord.概要) = FIX_PROC_SUMMARY(CStr(ProcName), cModule)

            'レコードをコレクションする
            procRecords.Add procRecord

        Next
    Next

    'シートに書き出す
    Dim tmp As Variant, i As Long
    For Each tmp In procRecords
        i = i + 1
        With trgSheet
            .Cells(i, eRecord.モジュール名) = tmp(eRecord.モジュール名)
            .Cells(i, eRecord.モジュールタイプ) = tmp(eRecord.モジュールタイプ)
            .Cells(i, eRecord.プロシージャ名) = tmp(eRecord.プロシージャ名)
            .Cells(i, eRecord.プロシージャタイプ) = tmp(eRecord.プロシージャタイプ)
            .Cells(i, eRecord.行数) = tmp(eRecord.行数)
            .Cells(i, eRecord.引数) = tmp(eRecord.引数)
            .Cells(i, eRecord.戻り値) = tmp(eRecord.戻り値)
            .Cells(i, eRecord.概要) = tmp(eRecord.概要)
        End With
    Next


    '見た目を整える
    ActiveWindow.DisplayGridlines = False
    With trgSheet.Cells
        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop

         '折り返して表示Flse、Trueの順でAutoFitを2度行うとレイアウトをカッチリできる
        .WrapText = False
        .Columns.AutoFit
        .Rows.AutoFit
        .WrapText = True
        .Columns.AutoFit
        .Rows.AutoFit
    End With

    With trgSheet
        .ListObjects.Add(xlSrcRange, .Cells(1, 1).CurrentRegion, , xlYes).Name = "ProcList"
    End With

    Exit Sub
hundler:
    MsgBox "シート名「Procs」が存在しています。"

End Sub

Private Function COLLECT_PROCNAMES_IN_MODULE(cModule As Object) As Collection
Rem --------------------------------------------------------------------------------
Rem CodeModuleを受け取り、含まれるプロシージャ名の一覧をCollectionで返す。
Rem --------------------------------------------------------------------------------

    Dim procNames As Collection: Set procNames = New Collection
    Dim i As Long, buf As String
    For i = 1 To cModule.CountOfLines
        If buf <> cModule.ProcOfLine(i, 0) Then
            buf = cModule.ProcOfLine(i, 0)
            procNames.Add buf
        End If
    Next

    Set COLLECT_PROCNAMES_IN_MODULE = procNames

End Function

Private Function FIX_MODULE_TYPE(module As Object) As String
Rem --------------------------------------------------------------------------------
Rem Moduleを受け取りモジュールタイプを文字列で返す。
Rem --------------------------------------------------------------------------------

    Select Case module.Type
        Case 1
            FIX_MODULE_TYPE = "標準モジュール"
        Case 2
            FIX_MODULE_TYPE = "クラスモジュール"
        Case 3
            FIX_MODULE_TYPE = "ユーザーフォーム"
        Case 100
            FIX_MODULE_TYPE = "Excelオブジェクト"
        Case Else
            FIX_MODULE_TYPE = module.Type
    End Select
End Function

Private Function FIX_PROC_TYPE(ProcName As String, procTop As String) As String
Rem --------------------------------------------------------------------------------
Rem プロシージャの1行目を受け取り、プロシージャタイプを抽出してテキストで返す。
Rem --------------------------------------------------------------------------------

    Dim reg As Object: Set reg = CreateObject("VBScript.RegExp")
    With reg
        .Pattern = " " & ProcName & "\(.*"
        .IgnoreCase = False
        .Global = True
    End With

    FIX_PROC_TYPE = reg.Replace(procTop, "")

End Function

Rem Sub Test_FIX_PROC_ARGS()
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge()")
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge() As Variant")
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge() As Variant()")
Rem
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge(v)")
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge(v) As Variant")
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge(v) As Variant()")
Rem
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge(v As Long)")
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge(v As Long) As Variant")
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge(v As Long) As Variant()")
Rem
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge(v() As Long)")
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge(v() As Long) As Variant")
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge(v() As Long) As Variant()")
Rem
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge(ParamArray v())")
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge(ParamArray v()) As Variant")
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge(ParamArray v()) As Variant()")
Rem
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge(v() As Long, w() As Variant)")
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge(v() As Long, w() As Variant) As Variant")
Rem     Debug.Print FIX_PROC_ARGS("Hoge", "Private Sub Hoge(v() As Long, w() As Variant) As Variant()")
Rem End Sub
Rem
Rem Function FIX_PROC_ARGS(ProcName, ByVal procTop) As String
Rem Rem --------------------------------------------------------------------------------
Rem 'プロシージャの1行目を受け取り、引数を抽出してテキストで返す。
Rem '複数ある場合はセル内改行を付与する。
Rem Rem --------------------------------------------------------------------------------
Rem     If InStr(procTop, ":") > 0 Then procTop = Left(procTop, InStr(procTop, ":") - 1)
Rem     If InStr(procTop, "'") > 0 Then procTop = Left(procTop, InStr(procTop, "'") - 1)
Rem     procTop = Trim(procTop) '偶に先頭スペースがある
Rem     Dim reg As Object: Set reg = CreateObject("VBScript.RegExp")
Rem     With reg
Rem         .Pattern = "(.*" & ProcName & "\()" & "(.*)" & "(\).*)"
Rem '        .Pattern = "(.*" & procName & "\()" & "(.*)" & "((\)|\).*:).*)"
Rem         .IgnoreCase = False
Rem         .Global = True
Rem     End With
Rem
Rem     Dim tmp As String
Rem     tmp = Trim(reg.Replace(procTop, "$2"))
Rem
Rem     If tmp = "" Then
Rem         FIX_PROC_ARGS = "-"
Rem     Else
Rem         FIX_PROC_ARGS = Replace(tmp, ", ", vbLf)
Rem     End If
Rem
Rem End Function
Rem
Rem Sub Test_FIX_PROC_RETURN()
Rem     Debug.Print FIX_PROC_RETURN("RowKeys", "Property Get RowKeys() As String()")
Rem     Debug.Print FIX_PROC_RETURN("RowKeys", "Property Get RowKeys(p As Variant) As String()")
Rem     Debug.Print FIX_PROC_RETURN("RowKeys", "Property Get RowKeys(p As Variant, q As Variant()) As String()")
Rem End Sub
Rem
Rem Function FIX_PROC_RETURN(ProcName, ByVal procTop) As String
Rem Rem --------------------------------------------------------------------------------
Rem 'プロシージャの1行目を受け取り、戻り値の型を抽出してテキストで返す。
Rem Rem --------------------------------------------------------------------------------
Rem     If InStr(procTop, ":") > 0 Then procTop = Left(procTop, InStr(procTop, ":") - 1)
Rem     If InStr(procTop, "'") > 0 Then procTop = Left(procTop, InStr(procTop, "'") - 1)
Rem     procTop = Trim(procTop) '偶に先頭スペースがある
Rem
Rem     procTop = MidStrForRev(procTop, "(", ")", False, False)
Rem     Dim reg As Object: Set reg = CreateObject("VBScript.RegExp")
Rem     With reg
Rem         .Pattern = "(.*[\(]\) As )(.*)"
Rem         .IgnoreCase = False
Rem         .Global = True
Rem     End With
Rem     Dim Matches As Variant
Rem
Rem     Set Matches = reg.Execute(procTop)
Rem     If Matches.Count > 0 Then
Rem         FIX_PROC_RETURN = reg.Replace(procTop, "$2")
Rem     Else
Rem         FIX_PROC_RETURN = "-"
Rem     End If
Rem
Rem End Function

Private Function FIX_PROC_SUMMARY(ProcName As String, cModule As Object) As String
Rem --------------------------------------------------------------------------------
Rem ProcNameとCodeModuleを受け取り、そのプロシージャの概要を文字列で返す。
Rem --------------------------------------------------------------------------------

    Dim StartRow As Long: StartRow = cModule.ProcStartLine(ProcName, 0)
    Dim LastRow As Long: LastRow = StartRow + cModule.ProcCountLines(ProcName, 0) - 1

    Dim reg As Object: Set reg = CreateObject("VBScript.RegExp")
    With reg
        .Pattern = "'----------.*" 'ハイフン10個で判定
        .IgnoreCase = False
        .Global = True
    End With
    Dim Matches As Variant

    Dim i As Long, tmp As String, checker As Boolean
    For i = StartRow To LastRow
        If checker Then
            tmp = tmp & cModule.Lines(i, 1) & vbLf
            Set Matches = reg.Execute(cModule.Lines(i, 1))
            If Matches.Count > 0 Then
                Exit For
            End If
        Else
            Set Matches = reg.Execute(cModule.Lines(i, 1))
            If Matches.Count > 0 Then
                checker = True
            End If
        End If
    Next

    tmp = reg.Replace(tmp, "")

    If tmp = "" Then
        FIX_PROC_SUMMARY = "-"
    Else
        tmp = Replace(tmp, "'", "")
        FIX_PROC_SUMMARY = Left(tmp, Len(tmp) - 1)
    End If

End Function

Private Function SET_PROC_TOP(ProcName As String, cModule As Object) As String
Rem --------------------------------------------------------------------------------
Rem ProcNameとCodeModuleを受け取り、そのプロシージャの1行目の内容を文字列で返す。
Rem --------------------------------------------------------------------------------

    Dim StartRow As Long: StartRow = cModule.ProcStartLine(ProcName, 0)
    Dim LastRow As Long: LastRow = StartRow + cModule.ProcCountLines(ProcName, 0) - 1

    Dim reg As Object: Set reg = CreateObject("VBScript.RegExp")
    With reg
        .Pattern = " " & ProcName & "\(.*"
        .IgnoreCase = False
        .Global = True
    End With
    Dim Matches As Variant

    Dim tmp As String, i As Long
    For i = StartRow To LastRow
        tmp = cModule.Lines(i, 1)
        Set Matches = reg.Execute(tmp)
        If Matches.Count > 0 Then SET_PROC_TOP = tmp
    Next
End Function

Private Sub SetVBIDEAccess()
On Error Resume Next
Rem  Access 2010 Later
Rem  Microsoft Visual Basic for Applications Extensibility 5.3 をプログラムで参照設定するマクロ
Rem  Programatically Set VBIDE.
On Error Resume Next
Application.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
End Sub

Private Sub refsetAccee2010Later()
Rem  For Microsoft Access 2010 Later 64/32
Dim ref As Object, refs As Object
Dim i As Long
Set refs = Application.References
For i = refs.Count To 1 Step -1
With refs.Item(i)
Debug.Print .Name, , .FullPath ' この時はDescriptionは使えない
End With
Next
On Error Resume Next
For Each ref In refs
If ref.Name = "VBIDE" Then refs.Remove ref
Next
refs.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3

On Error Resume Next
For Each ref In Application.ActiveWorkbook.VBProject.References
Debug.Print ref.Name, ref.Description, ref.GUID, ref.Major, ref.Minor, ref.FullPath
Next
For Each ref In refs
If ref.BuiltIn = False Then
If ref.Name <> "VBIDE" Then
refs.Remove ref
End If
End If
Next
On Error Resume Next
Const MSO16_Pro64 = "C:\Program Files\Microsoft Office\Root\Office16\"
Const MSO16_Pro32 = "C:\Program Files(x86)\Microsoft Office\Root\Office16\"
Const MSO15_Pro64 = "C:\Program Files\Microsoft Office\Office15\"
Const MSO15_Pro32 = "C:\Program Files(x86)\Microsoft Office\Office15\"
Const cnsSys32 = "C:\WINDOWS\System32\"
Const cnsWow64 = "C:\WINDOWS\SysWOW64\"
Const MShared64 = "C:\Program Files\Common Files\Microsoft Shared\"
Const MShared32 = "C:\Program Files(x86)\Common Files\Microsoft Shared\"
Const Common64 = "C:\Program Files\Common Files\"
Const Common32 = "C:\Program Files(x86)\Common Files\"
Const GUID_DAO = "{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}": refs.AddFromGuid GUID_DAO, 12, 0 ' Microsoft Office 16.0 Access database engine Object Library  =12.0 Note: You need download and  Install
Const GUID_ADODB = "{B691E011-1797-432E-907A-4D8C69339129}": refs.AddFromGuid GUID_ADODB, 6, 1 'Microsoft ActiveX Data Objects 6.1 Library =6.1
Const GUID_ADOX = "{00000600-0000-0010-8000-00AA006D2EA4}": refs.AddFromGuid GUID_ADOX, 6, 0 'Microsoft ADO Ext. 6.0 for DDL and Security  =6.0
Const GUID_ADOR = "{00000300-0000-0010-8000-00AA006D2EA4}": refs.AddFromGuid GUID_ADOR, 6, 0 'Microsoft ActiveX Data Objects Recordset 6.0 Library
Const GUID_AccessApp = "{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}": refs.AddFromGuid GUID_ADOR, 9, 0  'Microsoft Access XX.0 Object Library
Const GUID_CDO = "{CD000000-8B95-11D1-82DB-00C04FB1625D}": refs.AddFromGuid GUID_CDO, 1, 0 'Microsoft CDO for Windows 2000 Library = 1.0
Const GUID_MSCoree24 = "{5477469E-83B1-11D2-8B49-00A0C9B7C9C4}": refs.AddFromGuid GUID_MSCoree24, 2, 4 'Common Language Runtime Execution Engine 2.4 Library  = 2.4
Const GUID_IMAPI2 = "{2735412F-7F64-5B0F-8F00-5D77AFBE261E}": refs.AddFromGuid GUID_IMAPI2, 1, 0 'Microsoft IMAPI2 Base Functionality = 1.0
Const GUID_IMAPI2FS = "{2C941FD0-975B-59BE-A960-9A2A262853A5}": refs.AddFromGuid GUID_IMAPI2FS, 1, 0 'Microsoft IMAPI2 File System Image Creator  = 1.0
Const GUID_JetES = "{2358C810-62BA-11D1-B3DB-00600832C573}": refs.AddFromGuid GUID_JetES, 4, 0  'JET Expression Service Type Library
Const GUID_JRO = "{AC3B8B4C-B6CA-11D1-9F31-00C04FC29D52}": refs.AddFromGuid GUID_JRO, 2, 6 ' Microsoft Jet and Replication Objects 2.6 Library =  2.6       C:\Program Files (x86)\Common Files\System\ado\msjro.dll
Const GUID_MsoEuro = "{76F6F3F5-9937-11D2-93BB-00105A994D2C}": refs.AddFromGuid GUID_JetES, 1, 0 'Microsoft Office Euro Converter Object Library = 1.0"
Const GUID_MSHTML = "{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}": refs.AddFromGuid GUID_MSHTML, 4, 0 'Microsoft HTML Object Library  = 4.0     C:\Windows\SysWOW64\msjtes40.dll 4.0"
Const GUID_MSXML2_V60 = "{F5078F18-C551-11D3-89B9-0000F81FE221}": refs.AddFromGuid GUID_MSXML2_V60, 6, 0  ' Microsoft XML, v6.0         C:\Windows\System32\msxml6.dll = 6.0
Const GUID_OfficeObject = "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}" 'Microsoft Office 16.0 Object Library      C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL = 2.8
refs.AddFromGuid "{F618C513-DFB8-11D1-A2CF-00805FC79235}", 1, 0
refs.AddFromGuid "{8E80422B-CAC4-472B-B272-9635F1DFEF3B}", 1, 0 'MMC20  Microsoft Management Console 2.0
Const GUID_VBRegExp55 = "{3F4DACA7-160D-11D2-A8E9-00104B365C9F}": refs.AddFromGuid GUID_VBRegExp55, 5, 5 'Microsoft VBScript Regular Expressions 5.5
Const GUID_Scripting = "{420B2830-E718-11CF-893D-00A0C9054228}": refs.AddFromGuid GUID_Scripting, 1, 0 ' Microsoft Scripting Runtime C:\Windows\System32\scrrun.dll =  1.0
Const GUID_Shell32 = "{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}": refs.AddFromGuid GUID_Shell32, 1, 0 ' Microsoft Shell Controls And Automation  = 1.0
Const GUID_SHDocVw = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}": refs.AddFromGuid GUID_SHDocVw, 1, 1   'Microsoft Internet Controls = 1.1
Const GUID_WIA = "{94A0E92D-43C0-494E-AC29-FD45948A5221}": refs.AddFromGuid GUID_WIA, 1, 0          ' Microsoft Windows Image Acquisition Library v2.0   = 1.0
Const GUID_WinHttp = "{662901FC-6951-4854-9EB2-D9A2570F2B2E}": refs.AddFromGuid GUID_WinHttp, 5, 1 'Microsoft WinHTTP Services, version 5.1   C:\WINDOWS\system32\winhttpcom.dll = 5.1
Const GUID_IWshRuntimeLibrary = "{F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}":  refs.AddFromGuid GUID_IWshRuntimeLibrary, 1, 0 ' Windows Script Host Object Model  = 1.0
Const GUID_WSHControllerLibrary = "{563DC060-B09A-11D2-A24D-00104BD35090}": refs.AddFromGuid GUID_WSHControllerLibrary, 1, 0    ' WSHControler Library = 1.0
Const GUID_RDS = "{BD96C556-65A3-11D0-983A-00C04FC29E30}": refs.AddFromGuid GUID_RDS, 1, 5       ' Microsoft Remote Data Services 6.0 Library  = 1.5
Const GUID_SpeechLib = "{C866CA3A-32F7-11D2-9602-00C04F8EE628}": refs.AddFromGuid GUID_SpeechLib, 5, 4 ' Microsoft Speech Object Library  =5.4
Const GUID_TTSEngineLib = "{EB2114C0-CB02-467A-AE4D-2ED171F05E6A}": refs.AddFromGuid GUID_SpeechLib, 10, 0 ' Microsoft TTS Engine 10.0 Type Library =10.0
Const GUID_System_Drawing = "{D37E2A3E-8545-3A39-9F4F-31827C9124AB}": refs.AddFromGuid GUID_System_Drawing, 2, 4           'System.Drawing.dll  2.4
Const GUID_System_EnterpriseServices = "{4FB2D46F-EFC8-4643-BCD0-6E5BFA6A174C}": refs.AddFromGuid GUID_System_EnterpriseServices, 2, 4  'System_EnterpriseServices = 2.4
Const GUID_System_Windows_Fomrs20 = "{215D64D2-031C-33C7-96E3-61794CD1EE61}": refs.AddFromGuid GUID_System_Windows_Fomrs20, 2, 0 'System Windows Forms 2.0 Object Library = 2.0
Const GUID_WbemScripting = "{565783C6-CB41-11D1-8B02-00600806D9B6}": refs.AddFromGuid GUID_WbemScripting, 1, 2 ' Microsoft WMI Scripting V1.2 Library  = 1.2
Const GUID_WMPLib = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}": refs.AddFromGuid GUID_WMPLib, 1, 0       ' Windows Media Player = 1.0
Const GUID_WinWord = "{00020905-0000-0000-C000-000000000046}" ' Microsoft Word 16.0 Object Library= 8.7
Const GUID_MSPub = "{0002123C-0000-0000-C000-000000000046}"      'Microsoft Publisher 16.0 Object Library   = 2.3
Const GUID_OUTLOOK = "{0006F062-0000-0000-C000-000000000046}" ' Microsoft Outlook 16.0 Object Library 9.6
Const GUID_OLXLib = "{0006F062-0000-0000-C000-000000000046}" ' Microsoft Outlook View Control = 1.2
Const GUID_POWERPOINT = "{91493440-5A91-11CF-8700-00AA0060263B}" ' Microsoft PowerPoint 16.0 Object Library = 2.12
Const GUID_Excel = "{00020813-0000-0000-C000-000000000046}"  ' Microsoft Excel 16.0 Object Library  = 1.9
Const GUID_GRAPH = "{00020802-0000-0000-C000-000000000046}" 'Microsoft Graph 16.0 Object Library  = 1.9tr
Const GUID_MSAccess16 = "{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}" ' Microsoft Access 16.0 Object Library = 9.0
Const GUID_BARCODELib = "{D9347025-9612-11D1-9D75-00C04FCC8CDC}": refs.AddFromGuid GUID_BARCODELib, 1, 0 ' Microsoft Access BarCode Control 14.0  = 1.0
Const GUID_eawfctrlLib16 = "{113D61B1-C7C0-4157-B694-43594E25DF45}" 'eawfctrl 1.2 Type Library = 1.2
#If VBA7 Then
refs.AddFromFile "C:\Windows\System32\tapi3.dll" 'Microsoft TAPI 3.0 Type Library
#Else
refs.AddFromFile "C:\Windows\SysWow64\tapi3.dll"
#End If
refs.AddFromGuid "{714DD4F6-7676-4BDE-925A-C2FEC2073F36}", 1, 0 ' AccessibilityCplAdminLib    AccessibilityCplAdmin 1.0 Type Library
refs.AddFromGuid "{44EC0535-400F-11D0-9DCD-00A0C90391D3}", 1, 0 ' ATLLib    ATL 2.0 Type Library
refs.AddFromGuid "{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}", 9, 0 ' Microsoft Access 14.0 -  16.0 Object Library
refs.AddFromGuid "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}", 1, 0 ' MSScriptControl  Microsoft Script Control 1.0
refs.AddFromGuid "{8D763331-F59C-46F5-99FF-F74CDC84AD0E}", 1, 0 ' Microsoft Project Task Launch Control
refs.AddFromGuid "{54AF9343-1923-11D3-9CA4-00C04F72C514}", 2, 50 'MACVer
refs.AddFromGuid "{0D452EE1-E08F-101A-852E-02608C4D0BB4}", 2, 0 ' Microsoft Forms 2.0 Object Library
refs.AddFromGuid "{7988B57C-EC89-11CF-9C00-00AA00A14F56}", 1, 0 ' Microsoft Disk Quota
refs.AddFromGuid "{06290C00-48AA-11D2-8432-006008C3FBFC}", 1, 0 ' Scriptlet
refs.AddFromGuid "{EB2114C0-CB02-467A-AE4D-2ED171F05E6A}", 10, 0 'Microsoft TTS Engine 10.0 Type Library
refs.AddFromGuid "{9B085638-018E-11D3-9D8E-00C04F72D980}", 1, 0 ' Microsoft Tuner 1.0
refs.AddFromGuid "{9B7C3E2E-25D5-4898-9D85-71CEA8B2B6DD}", 2, 0 ' FDATELib   FDate 2.0 Type Library      C:\Program Files\Common Files\Microsoft Shared\Smart Tag\FDATE.DLL
refs.AddFromGuid "{2206CEB0-19C1-11D1-89E0-00C04FD7A829}", 1, 0 ' MSDASC Microsoft OLE DB Service Component 1.0 Type Library
refs.AddFromGuid "{E0E270C2-C0BE-11D0-8FE4-00A0C90A6341}", 1, 5 ' MSDAOSP Microsoft OLE DB Simple Provider 1.5 Library
refs.AddFromGuid "{833E4000-AFF7-4AC3-AAC2-9F24C1457BCE}", 1, 0 ' HelpServiceTypeLib
refs.AddFromGuid "{2A005C00-A5DE-11CF-9E66-00AA00A3F464}", 1, 0 ' COMSVCSLib    COM+ Services Type Library
refs.AddFromGuid "{98315905-7BE5-11D2-ADC1-00A02463D6E7}", 1, 0 ' COMReplLib    ComPlus 1.0 Catalog Replication Type Library
refs.AddFromGuid "{6CAAAA3B-6502-40FE-97FC-72A290DC63CF}", 1, 0 ' CorrEngineLib CorrEngine 1.0 Type Library
refs.AddFromGuid "{87099223-C7AF-11D0-B225-00C04FB6C2F5}", 1, 0 ' FAXCOMLib   faxcom 1.0 Type Library
refs.AddFromGuid "{E4DE3030-0142-4ACA-BA48-8613B56A2555}", 1, 0 ' FAXCONTROLLib FaxControl 1.0 Type Library
refs.AddFromGuid "{2BF34C1A-8CAC-419F-8547-32FDF6505DB8}", 1, 0 ' Microsoft Fax Service Extended COM Type Library"
refs.AddFromGuid "{9CDCD9C9-BC40-41C6-89C5-230466DB0BD0}", 2, 0 ' Feed 2.0
refs.AddFromGuid "{0FFF9602-69CF-4728-9EA4-141514866CA2}", 1, 0 ' FIndPrinterslib
refs.AddFromGuid "{D8DC76AB-F007-49C6-B6FC-8392A3DF90C4}", 1, 0 ' LocalService 1.0 Type Library
refs.AddFromGuid "{BED7F4EA-1A96-11D2-8F08-00A0C9A6186D}", 2, 4 ' Microsoft Common Language Runtime Class Library System.Collection Arraylist
refs.AddFromGuid "{78530B68-61F9-11D2-8CAD-00A024580902}", 1, 0 ' DexterLib     Dexter 1.0 Type Library
refs.AddFromGuid "{9B085638-018E-11D3-9D8E-00C04F72D980}", 1, 0 ' ATLLib        ATL 2.0 Type Library
refs.AddFromGuid "{B30CDC65-4456-4FAA-93E3-F8A79E21891C}", 1, 0 ' ATLEntityPickerLib          ATLEntityPicker 1.0 Type Library
refs.AddFromGuid "{28854DE7-2CF8-4A60-A85A-C21184D76BB6}", 1, 0 ' InstallerMainShellLib       Installer Main Shell Lib
refs.AddFromGuid "{E34CB9F1-C7F7-424C-BE29-027DCC09363A}", 1, 0 ' TaskScheduler  1.0
refs.AddFromGuid "{28DCD85B-ACA4-11D0-A028-00AA00B605A4}", 1, 0 ' TERMMGRLib    TAPI3 Terminal Manager 1.0 Type Library
refs.AddFromGuid "{28DCD85B-ACA4-11D0-A028-00AA00B605A4}", 1, 1 ' TDCLib        Tabular Data Control 1.1 Type Library
refs.AddFromGuid "{8628F27C-64A2-4ED6-906B-E6155314C16A}", 1, 0 ' REMOTEPROXY6432Lib          RemoteProxy6432 1.0 Type Library
refs.AddFromGuid "{A87F050D-3FFD-4682-8E77-34E530624CB4}", 1, 0 ' SessionMsgLib
refs.AddFromGuid "{C3A407A9-3409-4028-ACCF-9225FD9688D7}", 1, 0 ' RdpCoreTSLib  Rdp Protocol Provider 1.0 Type Library
refs.AddFromGuid "{438EDB38-282C-435D-8BE3-4AB90B83CEF5}", 1, 0 ' PrintUIObjLib PrintUI Objects 1.0 Type Library
refs.AddFromGuid "{91CE54EE-C67C-4B46-A4FF-99416F27A8BF}", 1, 0 ' PrinterExtensionLib         Printer Extension 1.0 Type Library
refs.AddFromGuid "{C8B522D5-5CF3-11CE-ADE5-00AA0044773D}", 1, 0 ' OLEDBError      Microsoft OLE DB Error Library
refs.AddFromGuid "{FC5988CF-6D6A-4812-ADD9-2DDE4F47346F}", 1, 0 ' MSTSWebProxyLib Microsoft Terminal Services Web Proxy 1.0 Type Library
refs.AddFromGuid "{8C11EFA1-92C3-11D1-BC1E-00C04FA31489}", 1, 0 ' MSTSCLib        Microsoft Terminal Services Control Type Library
refs.AddFromGuid "{7E8BC440-AEFF-11D1-89C2-00C04FB6BFC4}", 1, 0 ' IEXTagLib       iextag 1.0 Type Library
refs.AddFromGuid "{06CA6721-CB57-449E-8097-E65B9F543A1A}", 1, 0 ' IETAGLib        ietag 1.0 Type Library
refs.AddFromGuid "{833E4000-AFF7-4AC3-AAC2-9F24C1457BCE}", 1, 0 ' HelpServiceTypeLib          Help Service 1.0 Type Library
refs.AddFromGuid "{BA35B84E-A623-471B-8B09-6D72DD072F25}", 1, 0 ' VisioViewer     Microsoft Visio Viewer 16.0 Type Library
refs.AddFromGuid "{B9164592-D558-4EE7-8B41-F1C9F66D683A}", 1, 0 ' OneNoteIEAddin  Microsoft OneNote IE Addin Object Library
refs.AddFromGuid "{1C82EAD8-508E-11D1-8DCF-00C04FB951F9}", 1, 0 ' MIMEEDIT        Microsoft MIMEEDIT Type Library 1.0
refs.AddFromGuid "{31411197-A502-11D2-BBCA-00C04F8EC294}", 1, 0 ' MSHelpServices  Microsoft Help Data Services 1.0 Type Library
refs.AddFromGuid "{F618C513-DFB8-11D1-A2CF-00805FC79235}", 1, 0 ' COMAdmin        COM + 1.0 Admin Type Library


#If Win32 Then
refs.AddFromGuid "{0109E0F4-91AE-4736-A2CE-9D63E89D0EF6}", 1, 0 'XPS_SHL_DLLLib XPS_SHL_DLL 1.0 Type Library 32 bit版のみ参照設定可能
#End If
With refs
If Application.Version >= 16 Then
.AddFromGuid GUID_OfficeObject, 2, 8
.AddFromGuid GUID_Excel, 1, 9
.AddFromGuid "{00062FFF-0000-0000-C000-000000000046}", 9, 6 'Microsoft Outlook 16.0 Object Library
.AddFromGuid GUID_POWERPOINT, 2, 12
.AddFromGuid GUID_MSPub, 2, 3
.AddFromGuid GUID_OLXLib, 1, 2

.AddFromGuid "{113D61B1-C7C0-4157-B694-43594E25DF45}", 1, 2 'eawfctrl 1.0 Type Library
.AddFromGuid "{E301A065-3DF5-4378-A829-57B1EA986631}", 1, 0 ' Microsoft Outlook SharePoint Social Provider
.AddFromGuid "{3120BA9F-4FC8-4A4F-AE1E-02114F421D0A}", 1, 0 'OneNoteEx     Microsoft OneNote 15.0 Extended Object Library
.AddFromGuid "{E301A065-3DF5-4378-A829-57B1EA986631}", 1, 1 'Microsoft Outlook SharePoint Social Provider 1.1
.AddFromGuid "{1F8E79BA-9268-4889-ADF3-6D2AABB3C32C}", 1, 1 'OutlookSocialProvider       Microsoft Outlook Social Provider Extensibility
.AddFromGuid "{9E175B61-F52A-11D8-B9A5-505054503030}", 1, 0 'Microsoft Search Interface Type Library(from 2016)   C:\WINDOWS\system32\mssitlb.dll
.AddFromGuid "{CBBC4772-C9A4-4FE8-B34B-5EFBD68F8E27}", 1, 0 'NoteLinkComLib 1.0 Type Library(from 2016)
.AddFromGuid "{3120BA9F-4FC8-4A4F-AE1E-02114F421D0A} ", 1, 0 'Microsoft OneNote 15.0 Extended Object Library 1.0
.AddFromGuid GUID_WinWord, 8, 7 'Microsoft Word 16.0 Object Library
.AddFromGuid "{73720012-33A0-11E4-9B9A-00155D152105}", 1, 0 ' Microsoft Office Screen Recorder 16.0)from 2016) Object Librar
.AddFromGuid "{6CC6A20E-96A4-4F94-A838-8E5EBE9E9925}", 1, 0 ' ScreenReaderHelper
.AddFromGuid "{22E0CB87-9325-4B0F-8ECC-21B271EC81AA}", 1, 0 ' DolbyDLLlib (from 2016 windows 10)
.AddFromGuid "{4486DF98-22A5-4F6B-BD5C-8CADCEC0A6DE}", 1, 0 'LocationApi 1.0 Type Library (from 2016 windows 10)
.AddFromGuid "{012F24C1-35B0-11D0-BF2D-0000E8D0D146}", 1, 0 ' ACTIVEXLib    Microsoft Office Template and Media Control 1.0 Type Library
.AddFromGuid "{00020802-0000-0000-C000-000000000046}", 1, 9 'Microsoft Graph 16.0 Object Library
ElseIf Application.Version = 15 Then
On Error Resume Next
.AddFromGuid GUID_OfficeObject, 2, 8
If Err.Number <> 0 Then
Err.Clear
.AddFromGuid GUID_OfficeObject, 2, 7
End If
If Err.Number <> 0 Then
Err.Clear
.AddFromGuid GUID_OfficeObject, 2, 6
End If
Rem OutlookSocialProvider
.AddFromGuid "{E301A065-3DF5-4378-A829-57B1EA986631}", 1, 1 'OutlookSocialProvider
If Err.Number <> 0 Then
Err.Clear
.AddFromGuid "{E301A065-3DF5-4378-A829-57B1EA986631}", 1, 0 'OutlookSocialProvider
End If

Rem  Office 2013 (version 15) でMajor,Minor番号が定まっていると考えられるもの
.AddFromGuid GUID_Excel, 1, 8
.AddFromGuid GUID_OUTLOOK, 9, 5
.AddFromGuid GUID_POWERPOINT, 2, 11
.AddFromGuid GUID_MSPub, 2, 2
.AddFromGuid "{3120BA9F-4FC8-4A4F-AE1E-02114F421D0A} ", 1, 0 'Microsoft OneNote 15.0 Extended Object Library 1.0
.AddFromGuid GUID_OLXLib, 1, 1
.AddFromGuid "{113D61B1-C7C0-4157-B694-43594E25DF45}", 1, 1 'eawfctrl 1.0 Type Library
.AddFromGuid GUID_WinWord, 8, 6 'Microsoft Word 15.0 Object Library

ElseIf Application.Version = 14 Then
.AddFromGuid "{E301A065-3DF5-4378-A829-57B1EA986631}", 1, 0 'OutlookSocialProvider 2013 以降はない
.AddFromGuid GUID_OfficeObject, 2, 5
.AddFromGuid GUID_Excel, 1, 7
.AddFromGuid GUID_OUTLOOK, 9, 4
.AddFromGuid GUID_POWERPOINT, 2, 10
.AddFromGuid GUID_MSPub, 2, 1
.AddFromGuid GUID_OLXLib, 1, 1
.AddFromGuid "{1F8E79BA-9268-4889-ADF3-6D2AABB3C32C}", 1, 0 'Microsoft Outlook Social Provider Extensibility
.AddFromGuid "{0EA692EE-BB50-4E3C-AEF0-356D91732725}", 1, 0 'Microsoft OneNote 14.0 Object Library
.AddFromGuid "{113D61B1-C7C0-4157-B694-43594E25DF45}", 1, 0 'eawfctrl 1.0 Type Library
.AddFromGuid GUID_WinWord, 8, 5 'Microsoft Word 14.0 Object Library
End If
End With

If Not refs Is Nothing Then Set refs = Nothing
Set refs = Application.VBE.ActiveVBProject.References 'Application.ReferencesではDescriptionが出ない。このため　Refs を Nothing にして、 左のように書き換える
For Each ref In refs
If ref.IsBroken = False Then
Debug.Print ref.Name, ref.GUID, ref.Major, ref.Minor, ref.Description, ref.FullPath
Else
refs.Remove ref
End If
Next
End Sub
    
Rem 14桁毎になるように右寄せにする
Rem ※14桁以上のデータは上位の桁が消える
Rem Function dpr(ParamArray vals() As Variant) As String
Rem     Dim v As Variant
Rem     For Each v In vals
Rem         dpr = dpr & Right(String(13, " ") & CStr(v), 14)
Rem     Next
Rem End Function

Private Function dpr(ParamArray vals() As Variant) As String
    Dim v As Variant, str14 As String * 14
    For Each v In vals
        RSet str14 = CStr(v)
        dpr = dpr & str14
    Next
End Function
Rem Debug.Print VBA.String(200, vbNewLine)

Rem http://beatdjam.hatenablog.com/entry/2014/10/08/023925
Rem /**
Rem  * OutputLog
Rem  * デバッグログをファイルに出力する
Rem  * @param varData              : 出力対象のデータ
Rem  * @param Optional strFileNm   :(出力ファイル名を指定する場合)ファイル名
Rem  * @param Optional lngDebugFLG :(0=デバッグ・ファイル出力,1=デバッグのみ出力,2=ファイルのみ出力)
Rem  */
Public Sub OutputLog(ByVal varData As Variant, _
                     Optional ByVal lngDebugFLG As Long = 1, _
                     Optional ByVal strFileNm As String = "")
    
    Dim lngFileNum As Long
    Dim strLogFile As String
      
    'ファイル出力対象の場合
    If lngDebugFLG = 0 Or lngDebugFLG = 2 Then
        ' ファイル名の指定がない場合、現在の年月日をファイル名とする
        ' 引数のファイル名に拡張子が存在しない場合、拡張子を付加する
        If strFileNm = "" Then
          strFileNm = Format(Now(), "yyyymmdd") & ".txt"
        ElseIf InStr(strFileNm, ".txt") = 0 Then
          strFileNm = strFileNm & ".txt"
        End If
        
        ' 出力先ファイル設定
        ' Accessで利用する場合はCurrentProjectオブジェクトを使う
        ' strLogFile = CurrentProject.Path & "\" & strFileNm
        strLogFile = ActiveWorkbook.Path & "\" & strFileNm
        lngFileNum = FreeFile()
        Open strLogFile For Append As #lngFileNum
        Print #lngFileNum, varData
        Close #lngFileNum
    End If
    
    'デバッグログ出力対象の場合
    If lngDebugFLG = 0 Or lngDebugFLG = 1 Then
        Debug.Print varData
    End If

End Sub

Rem msgをメッセージボックスに表示する
Private Sub proc(msg As String)
    MsgBox msg
End Sub

Rem nとmを足す関数
Private Function FuncSum(n As Long, M As Long) As Long
    FuncSum = n + M
End Function

Private Sub Test1()
    Dim i As Long
    For i = 1 To 10
        If ActiveSheet.Cells(i, 2) = "ことり" Then Stop
    Next
End Sub

Private Sub Test2()
    Dim arr As Variant
    ReDim arr(1 To 3)
    arr(1) = Array("1", "2", "3", "4", "5")
    arr(2) = Array("ひよこ", "ことり", "いぬ", "ひつじ", "ねこ")
    arr(3) = Array("ぴよぴよ", "ちゅんちゅん", "わんわん", "もふもふ", "にゃんにゃん")
End Sub

Rem Sub ForEachTest()
Rem     For Each R In Selection: Debug.Print """" & R & """,";: Next
Rem End Sub

Rem Sub Arr2Test()
Rem     Dim Arr
Rem     Arr = Selection
Rem     Stop
Rem End Sub
Rem https://www.moug.net/tech/exvba/0150101.html


Private Sub Format関数で数字の先頭に0を付ける()
  Debug.Print Format("123", "00000") ' 先頭に0を付ける（数字のみ）00123
  Debug.Print "[" & Format("ABC", "@@@@@") & "]"  ' 半角のスペースを埋めて右寄せ  [  ABC]
  Debug.Print "[" & Format("ABC", "!@@@@@") & "]" ' 半角のスペースを埋めて左寄せ  [ABC  ]
End Sub


Rem http://yumem.cocolog-nifty.com/excelvba/2011/05/post-82d3.html
Rem 空欄でログを流す
Rem カーソルが末尾にないと意味がない
Rem イミディエイトウィンドウは200行しか表示できないので199出力した時点で全滅する
Private Sub ImdFlush()
    Dim i As Long: For i = 1 To 199: Debug.Print: Next
End Sub

Rem イミディエイトに適当にデータを出力
Rem 　ただしカーソルの位置次第でダメ
Rem 　動作が重い
Private Sub ImdRandomData()
    Dim i As Long: For i = 1 To 10: Debug.Print Rnd: Next
End Sub

Rem イミディエイトウィンドウを全て削除する
Rem  非表示の時は動作しない。（必要ないから問題なし）
Rem  VBEオブジェクトアクセスの許可が必要
Rem  安定して動作しない
Private Sub ImdClear_G_Home_End_Del_F7()
    On Error GoTo ENDPOINT
    If Application.VBE.MainWindow.Visible And _
        Application.VBE.Windows("イミディエイト").Visible Then
            SendKeys "^{g}", True
Rem             DoEvents               'これを入れると、ポップアップ中はコードウィンドウが吹き飛ぶ
            SendKeys "^{Home}", True
            SendKeys "^+{End}", True
            SendKeys "{Del}", True
            SendKeys "{F7}", True
    End If
ENDPOINT:
End Sub

Rem この方法にはまだ問題があり、
Rem 続けてDebug.Print をすると、削除→出力の流れが、全て出力→VBA終了後に削除
Rem DoEventsを入れると、削除が一切働かない。
Rem
Rem
Private Sub ImdClear_G_A_Del_F7()
    On Error GoTo ENDPOINT
    If Application.VBE.MainWindow.Visible And _
        Application.VBE.Windows("イミディエイト").Visible Then
Rem             Application.VBE.Windows("イミディエイト").Visible = True
            SendKeys "^g", True
            SendKeys "^a", True
            SendKeys "{Del}", True
            SendKeys "{F7}", True
    End If
ENDPOINT:
End Sub

Rem イミディエイトウィンドウの先頭行を除いてすべて削除する
Rem  フォーカスをイミディエイトウィンドウに残す
Rem  安定して動作しない
Private Sub ImdClear_G_Home_Down_End_Del_F7()
    If Application.VBE.MainWindow.Visible And _
        Application.VBE.Windows("イミディエイト").Visible Then
            SendKeys "^{g}", True
Rem             DoEvents
            SendKeys "^{Home}", True
            SendKeys "{Down}", True
            SendKeys "^+{End}", True
            SendKeys "{Del}", True
            SendKeys "{F7}", True
    End If
End Sub

Rem イミディエイトウィンドウの末尾にフォーカスを移動する
Private Sub ImdCursolMoveToLast()
    If Application.VBE.MainWindow.Visible And _
        Application.VBE.Windows("イミディエイト").Visible Then
            SendKeys "^{g}", False
Rem             DoEvents
            SendKeys "^+{End}", False
            SendKeys "{F7}", False
            DoEvents
    End If
End Sub

Rem  イミディエイトウィンドウの内容を抹消
Rem 1. ウィンドウ切り替え
Rem 2. Ctrl+A
Rem 3. Delete
Rem 4. アクディブウィンドウを元へ
Public Sub ImdClear()
 
    Dim wd      As Object
    Dim wdwk    As Object
     
    Set wd = Application.VBE.Windows("イミディエイト")
    
    Application.VBE.Windows("イミディエイト").Visible = True
    
    Dim IsImdDocking As Boolean
    IsImdDocking = False
    
    'ドッキング中なら ※誤って実行するとコードが消える
    If IsImdDocking Then
        wd.SetFocus
        SendKeys "^a", False
        SendKeys "{Del}", False
        'Application.SendKeys "^g ^a {DEL}"
    Else
    'ポップアップ中なら
    
    End If
    
End Sub

Public Sub ImdClearGAX()
    SendKeys "^g", Wait:=True ' イミディエイト ウィンドウを表示します。
    SendKeys "^a", Wait:=True ' すべて選択
Rem     SendKeys "^x", Wait:=True ' 切り取り
    SendKeys "{Del}", Wait:=True ' 削除
End Sub

Private Sub Test_ImdCursolMoveToLast()
    Call ImdCursolMoveToLast
    Debug.Print "最後から出力"
End Sub

Private Sub VBEウィンドウを全て列挙()
    Dim Item
    For Each Item In Application.VBE.Windows
        Debug.Print Item.Caption
    Next
End Sub

Private Sub VBEウィンドウを指定した型だけ列挙()
    Dim Item
    For Each Item In GetVbeWindow(vbext_wt_Immediate)
        Debug.Print Item.Caption
    Next
End Sub

Private Sub VBEウィンドウのポップアップだけ列挙()
    Dim Item
    For Each Item In GetVbeWindow(vbext_wt_Immediate)
        Debug.Print Item.Caption
    Next
End Sub

Private Function GetVbeWindow(t As VBIDE.VBExt_WindowType) As Collection
    Dim retCol As Collection: Set retCol = New Collection
    Dim W As VBIDE.Window
    For Each W In Application.VBE.Windows
        If W.Type = t Then retCol.Add W
    Next
    Set GetVbeWindow = retCol
End Function

Rem 一瞬でVBEを開いて HomePersonal プロジェクトを選択し、イミディエイトをフォーカスする
Rem https://thom.hateblo.jp/entry/2015/08/16/025140
Rem １．VBIDEを使用する場合は「ツール」の「参照設定」メニューで
Rem 「Microsoft Visiual Basic for Applications Extensibility」を追加します。
Public Sub ShowImmediate()
    Application.VBE.MainWindow.Visible = True
    Dim W As VBIDE.Window
    Set Application.VBE.ActiveVBProject = Application.VBE.VBProjects("HomePersonal")
    For Each W In Application.VBE.Windows
        If W.Type = VBIDE.vbext_wt_Immediate Then
            W.SetFocus
        End If
    Next
End Sub

Private Sub DebugPrintClearProc(mode As String)
    'Adapted  by   keepITcool
    'Original from Jamie Collins fka "OneDayWhen"
    'http://www.dicks-blog.com/excel/2004/06/clear_the_immed.html

    Static savState(0 To 255) As Byte
    
    Select Case mode
        Case "Clear"
            Dim hPane As LongPtr
            Dim tmpState(0 To 255) As Byte
            
            hPane = GetImmHandle
            If hPane = 0 Then MsgBox "イミディエイトウィンドウが見つかりません。"
            If hPane < 1 Then Exit Sub
            
            'CtrlやShiftの状態を記憶
            GetKeyboardState savState(0)
            
            'Ctrl押し下げ
            tmpState(vbKeyControl) = KEYSTATE_KEYDOWN
            SetKeyboardState tmpState(0)
            'Ctrl+ENDを送信
            PostMessage hPane, WM_KEYDOWN, vbKeyEnd, 0&
            'SHIFT押し下げ
            tmpState(vbKeyShift) = KEYSTATE_KEYDOWN
            SetKeyboardState tmpState(0)
            'CTRL+SHIFT+Home
            PostMessage hPane, WM_KEYDOWN, vbKeyHome, 0&
            'CTRL+SHIFT+BackSpace
            PostMessage hPane, WM_KEYDOWN, vbKeyBack, 0&
            
            'CtrlやShiftの状態を復元
            Application.OnTime Now + TimeSerial(0, 0, 0), "DoCleanUp"
        Case "CleanUp"
            ' Restore keyboard state
            SetKeyboardState savState(0)
        Case Else
            Stop
    End Select
End Sub

Private Sub DebugPrintClear3()
    Call DebugPrintClearProc("Clear")
End Sub

Private Sub DebugPrintClear3_DoCleanUp()
    Call DebugPrintClearProc("CleanUp")
End Sub

Private Sub PopupGetImmHandle()
    MsgBox GetImmHandle
End Sub

Private Function GetImmHandle() As LongPtr
Rem This function finds the Immediate Pane and returns a handle.
Rem Docked or MDI, Desked or Floating, Visible or Hidden


    Dim oWnd As Object, bDock As Boolean, bShow As Boolean
    Dim sMain$, sDock$, sPane$
    Dim lMain As LongPtr
    Dim lDock As LongPtr
    Dim lPane As LongPtr
    
    On Error Resume Next
    sMain = Application.VBE.MainWindow.Caption
    If Err <> 0 Then
        MsgBox "VBAプロジェクトにアクセスできません。"
        GetImmHandle = -1
        Exit Function
        ' Excel2003: Registry Editor (Regedit.exe)
        '    HKLM\SOFTWARE\Microsoft\Office\11.0\Excel\Security
        '    Change or add a DWORD called 'AccessVBOM', set to 1
        ' Excel2002: Tools/Macro/Security
        '    Tab 'Trusted Sources', Check 'Trust access..'
    End If
    
    
    For Each oWnd In Application.VBE.Windows
        If oWnd.Type = 5 Then
            bShow = oWnd.Visible
            sPane = oWnd.Caption
            If Not oWnd.LinkedWindowFrame Is Nothing Then
                bDock = True
                sDock = oWnd.LinkedWindowFrame.Caption
            End If
            Exit For
        End If
    Next
    
    lMain = FindWindow("wndclass_desked_gsk", sMain)
    If bDock Then
        'Docked within the VBE
        lPane = FindWindowEx(lMain, 0&, "VbaWindow", sPane)
        If lPane = 0 Then
            'Floating Pane.. which MAY have it's own frame
            lDock = FindWindow("VbFloatingPalette", vbNullString)
            lPane = FindWindowEx(lDock, 0&, "VbaWindow", sPane)
            While lDock > 0 And lPane = 0
                lDock = GetWindow(lDock, 2) 'GW_HWNDNEXT = 2
                lPane = FindWindowEx(lDock, 0&, "VbaWindow", sPane)
            Wend
        End If
    ElseIf bShow Then
        lDock = FindWindowEx(lMain, 0&, "MDIClient", _
        vbNullString)
        lDock = FindWindowEx(lDock, 0&, "DockingView", _
        vbNullString)
        lPane = FindWindowEx(lDock, 0&, "VbaWindow", sPane)
    Else
        lPane = FindWindowEx(lMain, 0&, "VbaWindow", sPane)
    End If
    
    
    GetImmHandle = lPane


End Function

Private Sub CheckImdVisible()
    'VBEが表示されているか
    MsgBox Application.VBE.MainWindow.Visible
    
    'イミディエイトが表示されているか
    MsgBox Application.VBE.Windows("イミディエイト").Visible
    
    'イミディエイトがドッキングされているか
    'イミディエイトはポップアップ表示か
    
    '現在フォーカスのあるウィンドウはどれか
    
    
End Sub

Rem イミディエイトウィンドウを非表示
Private Sub ImdClose()
    Application.VBE.Windows("イミディエイト").Visible = False
    Debug.Print Application.VBE.ActiveWindow.Caption
End Sub

Rem イミディエイトウィンドウを表示
Private Sub ImdShow()
    'イミディエイトウィンドウを表示
    'Trueにするとフォーカスがイミディエイトに移る
    Application.VBE.Windows("イミディエイト").Visible = True
    Debug.Print Application.VBE.ActiveWindow.Caption
    'イミディエイト　と出力
    'ただし、VBA終了後のフォーカスが
    '　ドッキング中はイミディエイト
    '　ポップアップ中はコードウィンドウ
    'に戻る
End Sub

Rem イミディエイトウィンドウを表示してフォーカスをVBEに戻す
Private Sub ImdShow_UnFocus()
    Dim win As Object
    Set win = Application.VBE.ActiveWindow
    Application.VBE.Windows("イミディエイト").Visible = True
    Debug.Print Application.VBE.ActiveWindow.Caption
    win.SetFocus
    Debug.Print Application.VBE.ActiveWindow.Caption
End Sub


Rem ----------

Rem http://suyamasoft.blue.coocan.jp/ExcelVBA/Sample/VBProject/index.html

Rem  VBEのバージョンを表示します。
Private Sub Display_VBE_Version_Sample()
  MsgBox Prompt:="VBE.Version = " & Application.VBE.Version, Buttons:=vbInformation, Title:="VBE.Version"
End Sub

Rem  VBEのプロパティをイミディエイト ウィンドウに表示します。
Private Sub VBE_Sample()
  With Application.VBE
    Debug.Print "ActiveCodePane.TopLine: " & .ActiveCodePane.TopLine
    Debug.Print "ActiveVBProject.Name: " & .ActiveVBProject.Name
    Debug.Print "ActiveWindow.Caption: " & .ActiveWindow.Caption
    Debug.Print "Addins.Count: " & .AddIns.Count
    Debug.Print "CodePanes.Count: " & .CodePanes.Count
    Debug.Print "CommandBars.Count: " & .CommandBars.Count
    Debug.Print "MainWindow.Caption: " & .MainWindow.Caption
    Debug.Print "SelectedVBComponent.Name: " & .SelectedVBComponent.Name
    Debug.Print "VBProjects.Count: " & .VBProjects.Count
    Debug.Print "Version: " & .Version
    Debug.Print "Windows.Count: " & .Windows.Count
  End With
End Sub

Rem  VBEのコマンドバーの一覧のブックを作成します。
Private Sub Crate_CommandBars_List()
  Dim i As Long
  Dim wb As Workbook

  Set wb = Workbooks.Add
  wb.Worksheets(1).Cells(1, 1) = "Type"
  wb.Worksheets(1).Cells(1, 2) = "Name"
  wb.Worksheets(1).Cells(1, 3) = "NameLocal"
  With Application.VBE.CommandBars
    For i = 1 To .Count
      wb.Worksheets(1).Cells(i + 1, 1) = .Item(i).Type
      wb.Worksheets(1).Cells(i + 1, 2) = .Item(i).Name
      wb.Worksheets(1).Cells(i + 1, 3) = .Item(i).NameLocal
    Next i
  End With

  wb.Worksheets(1).Range("B:C").Columns.AutoFit
  wb.Worksheets(1).Range("A1").Select
End Sub

Rem  コマンド バーのリセットします。
Private Sub ResetCommandBars()
  Dim cb As CommandBar

  If MsgBox(Prompt:="やり直しできませんが、すべてのVBEのコマンドバーをリセットしますか？", Buttons:=vbYesNo + vbQuestion, Title:="確認") <> vbYes Then Exit Sub
  On Error Resume Next
  Application.Cursor = xlWait ' 砂時計型カーソルポインタ
  Application.StatusBar = "すべてのVBEのコマンドバーをリセットしてます。しばらくお待ちください..."
  For Each cb In Application.VBE.CommandBars
    If cb.BuiltIn Then
      cb.Reset ' 標準のコマンド バーはリセット
    Else
      cb.Delete ' ユーザーのコマンド バーは削除
    End If
  Next
  Application.StatusBar = ""
  Application.Cursor = xlDefault ' 標準のカーソルポインタ
  On Error GoTo 0
End Sub

Rem  VBEのアドイン一覧をイミディエイト ウィンドウに表示します。
Private Sub Addin_Sample()
  Dim i As Long

  With Application.VBE.AddIns
    If .Count < 1 Then
      MsgBox Prompt:="VBEのアドインはインストールしてません！", Buttons:=vbInformation, Title:="VBE.AddIns.Count"
      Exit Sub
    End If
    For i = 1 To .Count
      Debug.Print "progID:" & .Item(i).progID
      Debug.Print "Connect:" & .Item(i).Connect
      Debug.Print "Description:" & .Item(i).Description
      Debug.Print "GUID:" & .Item(i).GUID
      Debug.Print ""
    Next i
  End With
End Sub

Rem  Windowの一覧をイミディエイト ウィンドウに表示します。
Rem  vbext_WindowType（0=vbext_wt_CodeWindow, 5=vbext_wt_Immediate, 6=vbext_wt_ProjectWindowなど）
Private Sub Windows_Sample()
  Dim i As Long

  With Application.VBE.Windows
    For i = 1 To .Count
      Debug.Print "Caption:" & .Item(i).Caption
      Debug.Print "Top:" & .Item(i).Top
      Debug.Print "Left:" & .Item(i).Left
      Debug.Print "Width:" & .Item(i).Width
      Debug.Print "Height:" & .Item(i).Height
      Debug.Print "Visible:" & .Item(i).Visible
      Debug.Print "Type:" & .Item(i).Type
      Debug.Print "WindowState:" & .Item(i).WindowState
      Debug.Print ""
    Next i
  End With
End Sub

Rem  イミディエイト ウィンドウを表示しアクティブにします。
Private Sub Immediate_Window_SetFocus_Sample()
  Dim i As Long

  With Application.VBE.Windows
    For i = 1 To .Count
      If .Item(i).Type = vbext_wt_Immediate Then
        .Item(i).Visible = True
        .Item(i).SetFocus
        Exit For
      End If
    Next i
  End With
End Sub

Rem  すべてのプロジェクトのファイル名をイミディエイト ウィンドウに表示します。
Private Sub VBProjects_Sample()
  Dim i As Long

  For i = 1 To Application.VBE.VBProjects.Count
    Debug.Print Application.VBE.VBProjects(i).FileName
  Next i
End Sub

Rem  アクティブ プロジェクトのプロパティをイミディエイト ウィンドウに表示します。
Private Sub ActiveVBProject_Sample()
  With Application.VBE.ActiveVBProject
    Debug.Print "BuildFileName:" & .BuildFileName
    Debug.Print "Description:" & .Description
    Debug.Print "FileName:" & .FileName
    Debug.Print "Name:" & .Name
    Debug.Print "References.Count:" & .References.Count
    Debug.Print "Saved:" & .Saved
    Debug.Print "Type:" & .Type ' vbext_pt_HostProject = 100  or  vbext_pt_StandAlone = 101
    Debug.Print "VBComponents.Count:" & .VBComponents.Count
  End With
End Sub

Rem  アクティブ プロジェクトのモードをイミディエイト ウィンドウに表示します。
Private Sub VBAMode_Sample()
  Select Case Application.VBE.ActiveVBProject.mode
    Case vbext_vm_Run
      MsgBox Prompt:="vbext_vm_Run", Buttons:=vbInformation, Title:="ActiveVBProject.Mode"
    Case vbext_vm_Break
      MsgBox Prompt:="vbext_vm_Break", Buttons:=vbInformation, Title:="ActiveVBProject.Mode"
    Case vbext_vm_Design
      MsgBox Prompt:="vbext_vm_Design", Buttons:=vbInformation, Title:="ActiveVBProject.Mode"
  End Select
End Sub

Rem  プロジェクトを保存したか表示します。
Private Sub Display_VBE_ActiveVBProject_Saved()
  MsgBox Prompt:="VBE.ActiveVBProject.Saved = " & Application.VBE.ActiveVBProject.Saved, Buttons:=vbInformation, Title:="VBE.ActiveVBProject.Saved"
End Sub

Rem   アクティブ プロジェクト「参照設定」の一覧をイミディエイト ウィンドウに表示します。
Rem  エクセルのシートに貼り付けた後で複数の列に分けるには、「データ」タブの「区切り位置」を実行しカンマを選択します。
Private Sub Debug_Print_References()
  Dim i As Long

  With Application.VBE.ActiveVBProject
    Debug.Print "BuiltIn, Name, Description, FullPath, GUID"
    For i = 1 To .References.Count
      Debug.Print .References(i).BuiltIn & ", " & .References(i).Name & ", """ & .References(i).Description _
                  & """, """ & .References(i).FullPath & """, " & .References(i).GUID
    Next i
  End With
End Sub

Rem  コンポーネントのタイプを取得します。
Rem  vbext_ComponentType
Rem      1 vbext_ct_StdModule = 標準モジュール
Rem      2 vbext_ct_ClassModule = クラスモジュール
Rem      3 vbext_ct_MSForm = フォーム
Rem     11 vbext_ct_ActiveXDesigner = ActiveXDesigner
Rem    100 vbext_ct_Document = ドキュメント（Workbook,Worksheetなど）
Private Sub VBComponents_Type_Sample()
  Dim i As Long

  If Excel.ActiveWorkbook Is Nothing Then Exit Sub
  With Excel.ActiveWorkbook.VBProject
    For i = 1 To .VBComponents.Count
      Debug.Print .VBComponents(i).Type & ", " & .VBComponents(i).Name
    Next i
  End With
End Sub
Rem  DeleteModuleモジュールを削除します。
Private Sub VBComponents_Remove_Sample()
  Const DeleteName = "DeleteModule"
  Dim ret As VbMsgBoxResult
  Dim vbc As VBIDE.VBComponent

  On Error Resume Next
  Set vbc = ThisWorkbook.VBProject.VBComponents(DeleteName)
  On Error GoTo 0
  If vbc Is Nothing Then Exit Sub ' モジュールは存在しない！
  ret = MsgBox(Prompt:=DeleteName & " モジュールを削除しますか？", Buttons:=vbYesNo + vbQuestion, Title:="確認")
  If ret <> vbYes Then Exit Sub
  With ThisWorkbook.VBProject.VBComponents
    .Remove .Item(DeleteName)
  End With
End Sub
Rem  選択してるモジュール名を表示します。
Private Sub SelectedVBComponent_Name_Sample()
  MsgBox Prompt:="選択してるモジュール名：" & Application.VBE.SelectedVBComponent.Name, Buttons:=vbYesNo + vbQuestion, Title:="SelectedVBComponent.Name"
End Sub

Sub Test_kccPath_ParentFolderPath()
    Dim p As kccPath
    
    '明示的にis_file:=Falseとすればフォルダ認識
    Set p = kccPath.Init("C:\vba\hoge", False)
    Debug.Print p.CurrentFolderPath, p.ParentFolderPath
    
    'パスの末尾が￥ならフォルダ認識
    Set p = kccPath.Init("C:\vba\hoge\")
    Debug.Print p.CurrentFolderPath, p.ParentFolderPath
    
    '未指定は原則ファイル認識
    Set p = kccPath.Init("C:\vba\hoge\a.xlsm")
    Debug.Print p.CurrentFolderPath, p.ParentFolderPath
End Sub

Sub Test_AbsolutePathNameEx()
    Dim s As String
    Debug.Print kccFuncString.AbsolutePathNameEx("C:\vba\hoge", ".\hoge.xls")
    Debug.Print kccFuncString.AbsolutePathNameEx("C:\vba\hoge\", ".\hoge.xls")
    Debug.Print kccFuncString.AbsolutePathNameEx("C:\vba\hoge", "hoge.xls")
    Debug.Print kccFuncString.AbsolutePathNameEx("C:\vba\hoge\", "hoge.xls")
End Sub

Sub Test_Load_kccsettings()
    Const SETTINGS_FILE_NAME = "kccsettings.json"
    
    Rem JsonテキストをUTF8で読み込み、規約違反のコメント行を削除
    Dim jsonText As String
    jsonText = kccPath.ReadUTF8Text(ThisWorkbook.Path & "\" & SETTINGS_FILE_NAME)
    jsonText = kccWsFuncRegExp.RegexReplace(jsonText, "[ ]*//.*\r\n", "")
    Debug.Print jsonText
    Stop
    
    Rem Jsonをパース
    Dim dic As Dictionary
    Set dic = JsonConverter.ParseJson(jsonText)
    Dim dKey, dItem
    For Each dKey In dic
        Debug.Print dKey, dic(dKey)
    Next
    Stop
End Sub

Sub Test_Load_kccsettings_class()
    Dim st As kccSettings
    Set st = kccSettings.Init(ThisWorkbook.Path)
    Debug.Print st.ExportBinFolder
    Debug.Print st.ExportSrcFolder
    Debug.Print st.BackupBinFile
    Debug.Print st.BackupSrcFile
    Stop
End Sub

Sub Test_Load_kccsettings_default()
    Dim st As kccSettings
    Set st = kccSettings.Init(ThisWorkbook.Path)
    st.CreateDefaultSetting
    Debug.Print st.ExportBinFolder
    Debug.Print st.ExportSrcFolder
    Debug.Print st.BackupBinFile
    Debug.Print st.BackupSrcFile
    Stop
End Sub

Rem アクティブなプロジェクトへソースをSRCからインポート
Rem
Rem  /src/CodeName.bas.vba
Rem
Public Sub VBComponents_Import_SRC()
'    Call VBComponents_BackupAndInport_Sub( _
'            Application.VBE.ActiveVBProject, _
'            ".\..\src", _
'            "", "")
    MsgBox "未実装", vbOKOnly, "VBComponents_Import_SRC"
End Sub

Rem アクティブなプロジェクトのソースコードを配下にエクスポート
Rem
Rem  /AddinName.xlam
Rem  /YYYYMMDD_HHMMSS/CodeName.bas.vba
Rem
Public Sub VBComponents_Export_YYYYMMDD()
    Call VBComponents_BackupAndExport_Sub( _
            Application.VBE.ActiveVBProject, _
            "", _
            ".\src\[YYYYMMDD]_[HHMMSS]\", _
            "", "")
End Sub

Rem アクティブなプロジェクトをgit用にエクスポート
Rem
Rem  /bin/AddinName.xlam
Rem  /src/CodeName.bas.vba
Rem
Public Sub VBComponents_Export_SRC()
    Dim obj As Object: Set obj = Application.VBE.ActiveVBProject
    Dim fn As String: fn = kccPath.Init(obj).CurrentFolder.FullPath
    Dim st As kccSettings: Set st = kccSettings.Init(fn)
    With st
        Call VBComponents_BackupAndExport_Sub( _
                obj, _
                .ExportBinFolder, _
                .ExportSrcFolder, _
                "", "")
    End With
End Sub

Rem アクティブなプロジェクトをGIT用バックアップ＆エクスポート
Rem
Rem  /bin/AddinName.xlam
Rem  /src/CodeName.bas.vba
Rem  /backup/bin/YYYYMMDD_HHMMSS_AddinName.xlam
Rem  /backup/src/CodeName.bas.vba
Rem
Public Sub VBComponents_BackupAndExport()
    Dim obj As Object: Set obj = Application.VBE.ActiveVBProject
    Dim st As kccSettings: Set st = kccSettings.Init(kccPath.Init(obj).FullPath)
    With st
        Call VBComponents_BackupAndExport_Sub( _
                obj, _
                .ExportBinFolder, _
                .ExportSrcFolder, _
                .BackupBinFile, _
                .BackupSrcFile)
    End With
End Sub

Public Sub VBComponents_BackupAndExportForAccess(): Call VBComponents_BackupAndExportForApps("Access.Application"): End Sub
Public Sub VBComponents_BackupAndExportForPowerPoint(): Call VBComponents_BackupAndExportForApps("PowerPoint.Application"): End Sub
Public Sub VBComponents_BackupAndExportForWord(): Call VBComponents_BackupAndExportForApps("Word.Application"): End Sub

Private Sub VBComponents_BackupAndExportForApps(AppClass As String)
    Dim objApplication As Object
    On Error Resume Next
    Set objApplication = GetObject(, AppClass)
    On Error GoTo 0
    If objApplication Is Nothing Then
        MsgBox "実行中の" & AppClass & "が見つかりませんでした。", vbCritical + vbOKOnly, "BackupAndExport"
        Exit Sub
    End If
    
    Call VBComponents_BackupAndExport_Sub( _
            objApplication.VBE.ActiveVBProject, _
            ".\.\bin", _
            ".\.\src", _
            ".\.\backup\bin\[YYYYMMDD]_[HHMMSS]_[FILENAME]", _
            ".\.\backup\src\[YYYYMMDD]_[HHMMSS]\[FILENAME]")
End Sub

Rem  プロジェクトのソースコードをエクスポートしたりバックアップする処理
Rem
Rem  @param ExportObject    出力プロジェクト（Workbook,VBProject)
Rem  @param ExportBinFolder エクスポートbinフォルダ
Rem  @param ExportSrcFolder エクスポートsrcフォルダ
Rem  @param BackupBinFile   バックアップbinファイル命名規則
Rem  @param BackupSrcFile   エクスポートsrcファイル命名規則
Rem
Public Sub VBComponents_BackupAndExport_Sub( _
            ExportObject As Object, _
            ExportBinFolder As String, _
            ExportSrcFolder As String, _
            BackupBinFile As String, _
            BackupSrcFile As String)
    Const PROC_NAME = "VBComponents_Export"
    
    Dim NowDateTime As Date: NowDateTime = Now()
    Dim objFilePath As kccPath: Set objFilePath = kccPath.Init(ExportObject)
    
    If Not objFilePath.Workbook Is Nothing Then
        If objFilePath.Workbook.ReadOnly Then
            MsgBox "[" & objFilePath.FileName & "] は読み取り専用です。処理を中止します。", vbOKOnly + vbCritical, PROC_NAME
            Exit Sub
        End If
        
        'プロジェクトの上書き保存
        Dim res As VbMsgBoxResult
        res = MsgBox(Join(Array( _
            objFilePath.FileName, _
            "エクスポートを実行します。", _
            "実行前にブックを保存しますか？"), vbLf), vbYesNoCancel, PROC_NAME)
        Select Case res
            Case vbYes
                Call UserNameStackPush(" ")
                objFilePath.Workbook.Save
                Call UserNameStackPush
            Case vbNo
                '何もしない
            Case vbCancel
                Exit Sub
        End Select
    End If
    
    'プロジェクトをリリースフォルダへ複製
    If ExportBinFolder <> "" Then
        Dim binPath As kccPath
        Set binPath = objFilePath.SelectPathToFolderPath(ExportBinFolder).ReplacePathAuto(DateTime:=NowDateTime)
        If binPath.FullPath <> objFilePath.CurrentFolder.FullPath Then
            binPath.DeleteItems
            binPath.CreateFolder
            If objFilePath.CurrentFolder.CopyTo(binPath, UseIgnoreFile:=True).IsAbort Then Exit Sub
        End If
    End If
    
    '既存ソースの削除とエクスポート
    '既存ソースを一旦別のフォルダに移動して、出力後に比較して、完全一致なら巻き戻す。
    If ExportSrcFolder <> "" Then
        Dim srcPath As kccPath
        Set srcPath = objFilePath.SelectPathToFolderPath(ExportSrcFolder)
        Set srcPath = srcPath.ReplacePathAuto(DateTime:=NowDateTime, FileName:=objFilePath.Name)
        srcPath.CreateFolder
        
        'src_backフォルダを作成してsrcの中身をsrc_backへ
        Dim backPath As kccPath
        Set backPath = srcPath.SelectPathToFolderPath("..\" & srcPath.Name & "_back\")
        backPath.CreateFolder
        backPath.DeleteFiles
        backPath.DeleteFolders
        srcPath.MoveTo backPath
        
        'srcフォルダを作成して中にexport
        Call VBComponents_Export(ExportObject, srcPath)
        Call CustomUI_Export(objFilePath, srcPath)
        
        'backから変更がないfrxを復元
        Dim f1 As File, f2 As File
        For Each f1 In srcPath.Folder.Files
            If f1.Name Like "*.frx" Then
                Dim isRestore As Boolean: isRestore = False
                For Each f2 In backPath.Folder.Files
                    If f1.Name = f2.Name Then
                        If f1.Size = f2.Size Then
                            '一致
                            Debug.Print "restore : " & f1.Name
                            f2.Copy f1.Path, True
                            isRestore = True
                        End If
                    End If
                Next
#If DEBUG_MODE Then
                'frxが何故か全部更新されてしまうときの確認用
                If Not isRestore Then Stop
#End If
            End If
        Next
        
        'タイムスタンプの復元
        
        
        'backフォルダの削除
        backPath.DeleteFolder
    End If
    
    'binとsrcのバックアップ
    If BackupBinFile <> "" Then
        binPath.CopyTo objFilePath.SelectPathToFilePath(BackupBinFile).ReplacePathAuto(DateTime:=NowDateTime), withoutFilterString:="*~$*"
    End If
    If BackupSrcFile <> "" Then
        srcPath.CopyTo objFilePath.SelectPathToFilePath(BackupSrcFile).ReplacePathAuto(DateTime:=NowDateTime)
    End If
    
    Debug.Print "VBA Exported : " & objFilePath.FileName
End Sub

Rem Application.UserNameを一時的に上書きする
Rem
Rem @param OverrideUserName 指定時:一時的に上書きする名前
Rem                         省略時:元の名前に復元
Rem
Sub UserNameStackPush(Optional OverrideUserName)
    Static lastUserName
    If IsMissing(OverrideUserName) Then
        Application.UserName = lastUserName
    Else
        If OverrideUserName = "" Then _
            Err.Raise "ユーザー名を空欄にするのはログイン名に置き換えられるため禁止です"
        lastUserName = Application.UserName
        Application.UserName = OverrideUserName
    End If
End Sub

Rem プロジェクトのCustomUIを指定フォルダにエクスポート
Rem
Rem  @param prj_path        エクスポートしたいブックのパス
Rem  @param output_path     エクスポート先のフォルダ
Rem
Private Sub CustomUI_Export(prj_path As kccPath, output_path As kccPath)
    
    Dim inFilePath As String
'    inFilePath = Path
    
    Dim tempPath As String
    With kccFuncZip.DecompZip(prj_path.FullPath)
        tempPath = .DecompFolder
        
        Dim xml1 As kccPath: Set xml1 = kccPath.Init(tempPath & "\" & "customUI\customUI.xml")
        Dim xml2 As kccPath: Set xml2 = kccPath.Init(tempPath & "\" & "customUI\customUI14.xml")
        
        xml1.CopyTo output_path
        xml2.CopyTo output_path
    End With
    
End Sub

Rem プロジェクトのCustomUIを指定フォルダにエクスポート
Rem
Rem  @param prj_path        エクスポートしたいブックのパス
Rem
Private Sub CustomUI_ExportAndOpen(prj_path As kccPath)
    Const PROC_NAME = "CustomUI_ExportAndOpen"
    
    Dim tempPath As String
    With kccFuncZip.DecompZip(prj_path.FullPath, AutoDelete:=False)
        tempPath = .DecompFolder
    
        Dim xml1 As kccPath: Set xml1 = kccPath.Init(tempPath & "\" & "customUI\customUI.xml")
        Dim xml2 As kccPath: Set xml2 = kccPath.Init(tempPath & "\" & "customUI\customUI14.xml")
        
        If xml1.Exists Or xml2.Exists Then
            Select Case MsgBox(Replace("はい：ファイルを開く\nいいえ：フォルダを開く\n", "\n", vbLf), vbYesNo)
                Case VbMsgBoxResult.vbYes
                    If xml1.Exists Then kccFuncWindowsProcess.OpenAssociationAPI xml1.FullPath
                    If xml2.Exists Then kccFuncWindowsProcess.OpenAssociationAPI xml2.FullPath
                Case VbMsgBoxResult.vbNo
                    Shell "explorer " & tempPath, vbNormalFocus
            End Select
        Else
            MsgBox "CustomUIは含まれていないようです。", vbOKOnly, PROC_NAME
        End If
    End With
    
End Sub

Public Sub CurrentProject_CustomUI_Import()
    MsgBox "未実装", vbOKOnly, "CurrentProject_CustomUI_Import"
End Sub

Public Sub CurrentProject_CustomUI_Export()
    Call CustomUI_ExportAndOpen(kccPath.Init(Application.VBE.ActiveVBProject))
End Sub

Private Sub Test_CustomUIをtempに展開して開いてみるだけ()
    Const Path = "C:\vba\test_CustomUI_Export.xlam"
    
    Dim inFilePath As String
    inFilePath = Path
    
    Dim tempPath As String
    With kccFuncZip.DecompZip(inFilePath, "\")
        tempPath = .DecompFolder
        Shell "explorer " & tempPath, vbNormalFocus
        Shell "notepad " & tempPath & "\customUI\customUI14.xml", vbNormalFocus
    End With
End Sub

Private Sub Test_Zip_一時フォルダの自動削除検証()
    Const Path = "C:\vba\test_CustomUI_Export.xlam"
    Dim tempPath As String
    
    Debug.Print "-----tempへの展開-----"
    
    With kccFuncZip.DecompZip(Path)
        tempPath = .DecompFolder
    End With
    Debug.Print "自動削除未指定(ON)", fso.FolderExists(tempPath) = False, tempPath
    Application.Wait [Now() + "00:00:01"]
    
    With kccFuncZip.DecompZip(Path, , AutoDelete:=False)
        tempPath = .DecompFolder
    End With
    Debug.Print "自動削除OFF", fso.FolderExists(tempPath) = True, tempPath
    Application.Wait [Now() + "00:00:01"]
    
    With kccFuncZip.DecompZip(Path, , AutoDelete:=True)
        tempPath = .DecompFolder
    End With
    Debug.Print "自動削除ON", fso.FolderExists(tempPath) = False, tempPath
    Application.Wait [Now() + "00:00:01"]
    
    
    Debug.Print "-----xlamと同じフォルダへの展開-----"
    
    With kccFuncZip.DecompZip(Path, "\")
        tempPath = .DecompFolder
    End With
    Debug.Print "自動削除未指定(OFF)", fso.FolderExists(tempPath) = True, tempPath
    Application.Wait [Now() + "00:00:01"]
    
    With kccFuncZip.DecompZip(Path, "\", AutoDelete:=False)
        tempPath = .DecompFolder
    End With
    Debug.Print "自動削除OFF", fso.FolderExists(tempPath) = True, tempPath
    Application.Wait [Now() + "00:00:01"]
    
    With kccFuncZip.DecompZip(Path, "\", AutoDelete:=True)
        tempPath = .DecompFolder
    End With
    Debug.Print "自動削除ON", fso.FolderExists(tempPath) = False, tempPath
    Application.Wait [Now() + "00:00:01"]
    
End Sub

Private Function isVBProjectProtected(prj As VBProject) As Boolean
    On Error Resume Next
    Dim dummy
    Set dummy = prj.VBComponents
    On Error GoTo 0
    isVBProjectProtected = IsEmpty(dummy)
End Function

Rem プロジェクトのソースコードを指定フォルダにエクスポート
Rem
Rem
Private Sub VBComponents_Export(prj As VBProject, output_path As kccPath)
    If prj Is Nothing Then MsgBox "VBAプロジェクト無し", vbOKOnly, "Export Error": Exit Sub
    If isVBProjectProtected(prj) Then MsgBox "VBAプロジェクトのロックを解除してください", vbOKOnly, "Export Error": Exit Sub
    output_path.CreateFolder
    
    Dim i As Long
    Dim cmp As VBComponent
    With prj
        For i = 1 To .VBComponents.Count
            Set cmp = .VBComponents(i)
            Dim declDic: Set declDic = GetDecInfoDictionary(cmp.CodeModule)
            Dim procDic: Set procDic = GetProcInfoDictionary(cmp.CodeModule)
            If declDic.Count = 0 And procDic.Count = 0 Then
                Debug.Print "Skip", cmp.Name
                GoTo ForContinue
            End If
            
            Debug.Print "Export", cmp.Name, , "宣言部", declDic.Count, , "関数部", procDic.Count
            
            Dim oFullPath As String: oFullPath = ""
            Select Case cmp.Type
                Case Is = vbext_ct_StdModule
                    oFullPath = output_path.FullPath & "\" & cmp.Name & ".bas" & ".vba"
                  
                Case Is = vbext_ct_ClassModule
                    oFullPath = output_path.FullPath & "\" & cmp.Name & ".cls" & ".vba"
                  
                Case Is = vbext_ct_MSForm
                    oFullPath = output_path.FullPath & "\" & cmp.Name & ".frm" & ".vba"
                  
                ' Workbook, Worksheetなど
                Case Is = vbext_ct_Document
                    oFullPath = output_path.FullPath & "\" & cmp.Name & ".cls" & ".vba"
                  
                ' ActiveX デザイナ
                Case Is = vbext_ct_ActiveXDesigner
                    oFullPath = output_path.FullPath & "\" & cmp.Name & ".cls" & ".vba"
            End Select
            
            If oFullPath <> "" Then
                cmp.Export oFullPath
                
                '環境によってfrmの座標に .001 が付与される現象の解消
                Call RepairFrm(oFullPath)
                
                'コードの末尾に不要な改行が入りがちな問題の解消
                Call CleanSource(oFullPath)
                
                'UTF-8への変換
'                kccPath.Init(oFullPath, True).ConvertCharCode_SJIS_to_utf8
            End If
ForContinue:
        Next
    End With
End Sub

'環境によってfrmの座標に .001 が付与される現象の解消
Private Function RepairFrm(frmFullPath)
    If Not frmFullPath Like "*.frm.vba" Then Exit Function
    Dim FileLines: FileLines = Split(fso.OpenTextFile(frmFullPath, ForReading).ReadAll, vbNewLine)
    
    Dim IsRepaired As Boolean
    Dim i As Long
    For i = 1 To 10
        If FileLines(i) Like "*.001" Then
            IsRepaired = True
            Debug.Print kccFuncString.GetPath(frmFullPath, False, True, True), FileLines(i)
'            Stop
        End If
        FileLines(i) = Replace(FileLines(i), ".001", "")
    Next
    If IsRepaired Then
        Dim ts As TextStream
        Set ts = fso.OpenTextFile(frmFullPath, ForWriting, True)
        ts.Write Join(FileLines, vbNewLine)
        ts.Close
    End If
End Function

Rem テキストファイル末尾の不要な改行を取り除く
Private Function CleanSource(oFullPath)
    Dim code As String
    code = fso.OpenTextFile(oFullPath, ForReading).ReadAll
    
    Dim IsChanged As Boolean
    Do
        If code Like "*" & vbCrLf & vbCrLf Then
            code = Left(code, Len(code) - 2)
            IsChanged = True
        Else
            Exit Do
        End If
    Loop
    
    If IsChanged Then
        fso.OpenTextFile(oFullPath, ForWriting).Write code
    End If
End Function

Rem  指定した名前のVBComponentが存在しているか調べます。
Private Function ExistsVBComponent(VBComponentName As String, Optional pVBProject As Variant)
  Dim VBPro As VBIDE.VBProject
  Dim VBCom As VBIDE.VBComponent

  ExistsVBComponent = False
  On Error Resume Next
  If IsMissing(pVBProject) Then
    If Application.VBE.ActiveVBProject Is Nothing Then Exit Function
    Set VBPro = Application.VBE.ActiveVBProject
  Else
    Set VBPro = pVBProject
  End If
  Set VBCom = VBPro.VBComponents(VBComponentName)
  ExistsVBComponent = Not (VBCom Is Nothing)
  On Error GoTo 0
  Set VBCom = Nothing
  Set VBPro = Nothing
End Function

Rem  アクティブ モジュールの宣言セクション部分の行数を返します。
Private Sub CountOfDeclarationLines_Sample()
  Dim Line As Long

  Line = Application.VBE.ActiveCodePane.CodeModule.CountOfDeclarationLines
  MsgBox Prompt:="宣言セクション部分の行数：" & Line, Buttons:=vbInformation, Title:="CodeModule.CountOfDeclarationLines"
End Sub

Rem  アクティブ モジュールの行数を返します。
Private Sub CountOfLines_Sample()
  Dim Line As Long

  Line = Application.VBE.ActiveCodePane.CodeModule.CountOfLines
  MsgBox Prompt:="モジュールの行数：" & Line, Buttons:=vbInformation, Title:="CodeModule.CountOfLines"
End Sub

Rem  プロシージャーの行数を返します。
Rem  【注意】プロシージャーの前の行にコメントがある場合は、コメントの行を含めます。
Private Sub ProcCountLines_Sample()
  Dim StartLine As Long

  StartLine = ThisWorkbook.VBProject.VBComponents("Code_Module").CodeModule.ProcCountLines(ProcName:="ProcCountLines_Sample", ProcKind:=vbext_pk_Proc)
  MsgBox Prompt:="コメントの行を含むプロシージャーの行数：" & StartLine, Buttons:=vbInformation, Title:="CodeModule.ProcCountLines"
End Sub

Rem  プロシージャーの開始行を返します。（プロシージャーの前の行にあるコメント行を含みます。）
Rem  【注意】前のプロシージャーの次の行を返します。
Rem   vbext_ProcKind
Rem     vbext_pk_Get
Rem     vbext_pk_Let
Rem     vbext_pk_Proc
Rem     vbext_pk_Set
Private Sub ProcStartLine_Sample()
  Dim StartLine As Long

  StartLine = ThisWorkbook.VBProject.VBComponents("Code_Module").CodeModule.ProcStartLine(ProcName:="ProcStartLine_Sample", ProcKind:=vbext_pk_Proc)
  MsgBox Prompt:="コメント行を含むプロシージャーの開始行：" & StartLine, Buttons:=vbInformation, Title:="CodeModule.ProcStartLine"
End Sub

Rem  プロシージャーの開始行を返します。
Private Sub ProcBodyLine_Sample()
  Dim StartLine As Long

  StartLine = ThisWorkbook.VBProject.VBComponents("Code_Module").CodeModule.ProcBodyLine(ProcName:="ProcBodyLine_Sample", ProcKind:=vbext_pk_Proc)
  MsgBox Prompt:="プロシージャーの開始行：" & StartLine, Buttons:=vbInformation, Title:="CodeModule.ProcBodyLine"
End Sub

Rem  コードモジュールの指定行から指定した行数のテキストを取得します。
Rem  【注意】CodeModule.LinesはUnicodeなので、半角でも2バイトです。
Private Sub Lines_Sample()
  Dim StartLine As Long, Count As Long

  StartLine = 3
  Count = 8
  MsgBox Prompt:=ThisWorkbook.VBProject.VBComponents("Code_Module").CodeModule.Lines(StartLine, Count), Buttons:=vbInformation, Title:="CodeModule.Lines"
End Sub

Rem  指定した行が含まれるプロシージャー名を取得します。
Private Sub ProcOfLine_Sample()
  Dim num As Variant
  Dim ProcName As String

  num = Application.InputBox(Prompt:="行数：", Title:="プロシージャー名の行数の入力", Default:=57, Type:=1)
  If TypeName(num) <> "Double" Then Exit Sub ' [キャンセル]ボタン
  ProcName = ThisWorkbook.VBProject.VBComponents("Code_Module").CodeModule.ProcOfLine(Line:=num, ProcKind:=vbext_pk_Proc)
  MsgBox Prompt:="プロシージャー名：" & ProcName, Buttons:=vbInformation, Title:="CodeModule.ProcOfLine"
End Sub

Rem  選択したテキストファイルをTempModuleモジュールの最初のプロシージャーの前に挿入します。
Private Sub AddFromFile_Sample()
  Dim FileName As Variant

  FileName = Application.GetOpenFileName(FileFilter:="テキストファイル（*.txt）, *.txt,すべてのファイル（*.*）,*.*", FilterIndex:=1, Title:="ファイルのインポート", ButtonText:="インポート", MultiSelect:=False)
  If TypeName(FileName) = "Boolean" Then Exit Sub ' [キャンセル]ボタン
  ThisWorkbook.VBProject.VBComponents("TempModule").CodeModule.AddFromFile FileName
End Sub

Rem  テキストをTempModuleモジュールの最初のプロシージャーの前に挿入します。
Private Sub AddFromString_Sample()
  Dim Str As String

  Str = "'" & String(50, "=") & vbCrLf
  Str = Str & "'AddFromStringで挿入しました。 " & Format(Now, "yyyy/mm/dd hh:mm:ss") & vbCrLf & Str
  ThisWorkbook.VBProject.VBComponents("TempModule").CodeModule.AddFromString Str
End Sub

Rem  テキストをTempModuleモジュールの5行目に挿入します。
Private Sub InsertLines_Sample()
  Dim Str As String

  Str = "' 5行目にInsertLinesで挿入しました。" & vbCrLf & "' vbCrLfを使用すると複数の行を挿入できます。"
  ThisWorkbook.VBProject.VBComponents("TempModule").CodeModule.InsertLines 5, Str
End Sub

Rem  現在のカーソルの開始行に日付と時間を挿入します。
Private Sub Insert_Text()
  Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
  Dim Text As String

  Text = Format(Now, "' ggge年mm月dd日 hh時mm分ss秒")
  With Application.VBE.ActiveCodePane
    .getSelection StartLine, StartColumn, EndLine, EndColumn
    .CodeModule.InsertLines StartLine, Text
  End With
End Sub

Rem  TempModuleモジュールの5行目と6行目の2行を削除します。
Private Sub DeleteLines_Sample()
  Dim StartLine As Long, CountLine As Long

  StartLine = 5
  CountLine = 2
  ThisWorkbook.VBProject.VBComponents("TempModule").CodeModule.DeleteLines StartLine, CountLine
End Sub

Rem  検索した文字列があるかを表示します。
Rem  Find(Target As String, StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long, [WholeWord As Boolean = False], [MatchCase As Boolean = False], [PatternSearch As Boolean = False]) As Boolean
Rem  【注意】StartColumnとEndColumnの桁は半角は1、全角は2で計算します。
Private Sub Find_Sample()
  Dim ret As Boolean
  Dim FindText As Variant

  FindText = Application.InputBox(Prompt:="文字列：", Title:="文字列の検索", Type:=2)
  If TypeName(FindText) = "Boolean" Then Exit Sub
  If Len(FindText) < 1 Then Exit Sub
  With ThisWorkbook.VBProject.VBComponents("Code_Module").CodeModule
    ret = .Find(FindText, 1, 1, .CountOfLines, LenB(.Lines(.CountOfLines, 1)), False, False, False) ' 【注意】LenではなくLenBを使います。
  End With
  MsgBox Prompt:=FindText & "の検索結果 = " & ret, Buttons:=vbInformation, Title:="文字列の検索"
End Sub

Rem  TempModuleモジュールの5行目を文字列で置き換えます。
Private Sub ReplaceLine_Sample()
  Dim Str As String

  Str = "' 5行目をReplaceLineで置き換えしました。 " & Format(Now, "yyyy/mm/dd hh:mm:ss")
  ThisWorkbook.VBProject.VBComponents("TempModule").CodeModule.ReplaceLine 5, Str
End Sub

Rem  アクティブ コンポーネントのすべてのプロシージャ名を表示します。（クラスのGet, Set, Letは除きます）
Private Sub Display_ProcName_Sample()
  Dim msg As String
  Dim ProcName As String
  Dim i As Long

  ProcName = vbNullString
  With Application.VBE.ActiveCodePane.CodeModule
    For i = 1 To .CountOfLines
      If ProcName <> .ProcOfLine(i, ProcKind:=vbext_pk_Proc) Then ' プロシージャ名が変わった場合は
        ProcName = .ProcOfLine(i, ProcKind:=vbext_pk_Proc)
Rem         Debug.Print buf
        msg = msg & ProcName & vbCrLf
      End If
    Next i
  End With

  MsgBox Prompt:=msg, Buttons:=vbInformation, Title:="プロシージャ名の一覧"
End Sub

Rem  選択範囲のカーソル位置を取得します。
Rem  【注意】StartColumnとEndColumnの桁は半角は1、全角は2で計算します。
Private Sub GetSelection_Sample()
  Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
  Dim msg As String

  Application.VBE.ActiveCodePane.getSelection StartLine, StartColumn, EndLine, EndColumn
  msg = "開始：" & StartLine & "行 " & StartColumn & "桁" & vbCrLf & vbCrLf
  msg = msg & "終了：" & EndLine & "行" & EndColumn & "桁"
  MsgBox Prompt:=msg, Buttons:=vbInformation, Title:="カーソル位置"
End Sub

Rem  選択範囲を設定します。
Rem  【注意】StartColumnとEndColumnの桁は半角は1、全角は2で計算します。
Rem  【注意】左端の桁は0ではなく1です。
Private Sub SetSelection_Sample()
  Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long

  StartLine = 14: StartColumn = 7: EndLine = 22: EndColumn = 8 ' ◆ StartColumnは全角は2で計算してます。
  Application.VBE.ActiveCodePane.SetSelection StartLine, StartColumn, EndLine, EndColumn
End Sub

Rem  アクティブ コード ペインの画面に表示できる行数を表示します。
Private Sub CountOfVisibleLines_Sample()
  MsgBox Prompt:="画面に表示できる行数：" & Application.VBE.ActiveCodePane.CountOfVisibleLines, Buttons:=vbInformation, Title:="ActiveCodePane.CountOfVisibleLines"
End Sub

Rem  アクティブ コード ペインの画面の最上行を表示します。
Private Sub TopLine_Sample()
  MsgBox Prompt:="画面の最上行：" & Application.VBE.ActiveCodePane.TopLine, Buttons:=vbInformation, Title:="ActiveCodePane.TopLine"
End Sub

Rem 全てのコードウインドウを閉じる
Public Sub CloseCodePanes()
    Dim C As CodePane
    For Each C In Application.VBE.CodePanes
        C.Window.Close
    Next
End Sub

Rem 現在のカーソルにある関数のテストを実行する
Public Sub TestExecute()
    Run GetCursolFunctionName()
    MsgBox "未完成"
End Sub

Rem 現在のカーソルにある関数のテストへジャンプする
Public Sub TestJump()
    ProcJump GetCursolFunctionName()
    MsgBox "未完成"
End Sub

Rem 現在のカーソルにある関数名を返す
Private Function GetCursolFunctionName()
    
End Function

Rem 指定した関数名の場所にジャンプする
Private Sub ProcJump(func As String)
    
    Rem モジュールを開く
    Rem カーソル位置を変える
End Sub

Rem ファイル化されていないブック全てを保存せずに閉じる
Public Sub CloseNofileWorkbook()
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Path = "" Then
            wb.Close False
        End If
    Next
End Sub

Rem VBAプロジェクトのパスワードを1234へ変更する
Public Sub BreakPassword1234Project()
    Dim beforePath As kccPath: Set beforePath = kccPath.Init(Application.VBE.ActiveVBProject)
    Dim afterPath As kccPath: Set afterPath = beforePath.SelectPathToFilePath("|t_1234|e")
    Select Case MsgBox(beforePath.FileName & "を" & afterPath.FileName & "へ出力します。", vbOKCancel)
        Case vbOK
            Dim res: res = BrokenVbaPassword(beforePath.FullPath, afterPath.FullPath)
            afterPath.OpenExplorer
            MsgBox "完了！！！" & res, vbOKOnly
        Case vbCancel
    End Select
End Sub

Public Sub OpenFormDeclareSourceGenerate()
    FormDeclareSourceGenerate.Show
End Sub

Public Sub OpenFormDeclareSourceTo64bit()
    FormDeclareSourceTo64bit.Show
End Sub

Rem 同じフォルダ、又は上位フォルダの大文字小文字ファイルを開く
Public Sub OpenTextFileBy大文字小文字()
    Dim targetPath As kccPath
    Set targetPath = kccPath.Init(ThisWorkbook.Path, False).SelectPathToFilePath(DEF_大文字小文字ファイル)
    If Not targetPath.Exists Then
        Set targetPath = targetPath.SelectParentFolder("..\")
    End If
    targetPath.OpenAssociation
End Sub

'.vbaをつけていなかったファイルに付け足す
Sub Test_AddVBA()
    Const TARGET_PATH = "C:\Users\hogehoge\src\20190416\"
    Dim p As kccPath: Set p = kccPath.Init(TARGET_PATH)
    Dim fl As File
    For Each fl In p.Folder.Files
        Select Case VBA.Right(fl.Name, 3)
            Case "bas", "cls", "frm"
                fl.Name = fl.Name & ".vba"
            Case "frx"
                fl.Name = Replace(fl.Name, ".frx", ".frm.frx")
            Case "vba"
                'nochange
            Case Else
                Stop
        End Select
    Next
End Sub
