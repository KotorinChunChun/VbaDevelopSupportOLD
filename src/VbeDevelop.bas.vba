Attribute VB_Name = "VbeDevelop"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        ExtDevelop
Rem
Rem  @description   �J����VBE�p�̃��W���[��
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
Rem    2020/08/01 �Đ���
Rem
Rem --------------------------------------------------------------------------------
Rem  @note
Rem     msdn
Rem       VBA �ŋN�����̂��ׂĂ� Excel �C���X�^���X�����S�Ɏ擾������
Rem       https://social.msdn.microsoft.com/Forums/ja-JP/7a46a3c9-f904-4fb0-a205-6112fba51fe6/vba-excel-?forum=vbajp
Rem
Rem     OKwave
Rem       �ʃC���X�^���X�̃u�b�N�i�l�p�}�N���u�b�N�ȊO�j�����ׂĕ���
Rem       MREXCEL.COM > Forum > Question Forums > Excel Questions > GetObject and HWND
Rem       https://okwave.jp/qa/q9196890.html
Rem
Rem     Qita
Rem       �yExcelVBA�zVBA�R�[�h�̏���T�v���V�[�g�Ɉꗗ�o�͂���
Rem       https://qiita.com/Mikoshiba_Kyu/items/46b7243eb576848b3e55
Rem
Rem       excel Access VBA �Q��1�̐ݒ��VBA�̎Q�Ɛݒ����������}�N��
Rem       https://qiita.com/Q11Q/items/67226e7c8b9def529668
Rem
Rem       VBA��Excel���g��
Rem       https://qiita.com/palglowr/items/04250eb1a8a873fbf9d2
Rem
Rem       GetRunningObjectTable
Rem       https://foren.activevb.de/forum/vb-classic/thread-409498/beitrag-409498/API-GetRunningObjectTable/
Rem
Rem       VBA �W�����W���[���̃}�N����ǂݎ���ċN������VBE��
Rem       ���j���[�Ɏ����o�^����A�h�C�������삷��
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
    ���W���[���� = 1
    ���W���[���^�C�v
    �v���V�[�W����
    �v���V�[�W���^�C�v
    �s��
    ����
    �߂�l
    �T�v
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

Rem �{���W���[���p�I��������������
Public Sub Terminate()
    'CustomUI Import/Export�p��ZIP�W�J�ꎞ�t�H���_�̏�����
    Call kccFuncZip.DeleteTempFolder
End Sub

Rem �A�N�e�B�u�ȃv���W�F�N�g�̕ۑ��t�H���_���J��
Public Sub OpenProjectFolder()
On Error Resume Next
    Dim fn: fn = Application.VBE.ActiveVBProject.FileName
    kccFuncWindowsProcess.ShellExplorer fn, True
End Sub

Rem WEB�T�C�g���J���i�֘A�t���v���O�����ŊJ���j
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

Rem ���݃A�N�e�B�u�ȃv���W�F�N�g�̃��[�N�u�b�N�����
Public Sub CloseProject()
    Dim objFilePath As kccPath: Set objFilePath = kccPath.Init(Application.VBE.ActiveVBProject)
    If objFilePath Is Nothing Then Exit Sub
    objFilePath.Workbook.Close
End Sub

Rem �A�N�e�B�u�u�b�N�̃\�[�X�R�[�h�̃v���V�[�W���ꗗ��V�K�u�b�N�֏o��
Public Sub VbeProcInfo_Output()
    Dim objFilePath As kccPath: Set objFilePath = kccPath.Init(Application.VBE.ActiveVBProject)
    
    '�v���V�[�W���ꗗ���擾���ē񎟌��z����擾���鏈��
    Dim data
    data = VbeProcInfo_GetTable(objFilePath.Workbook.VBProject)
    
    '�񎟌��z����u�b�N�ɏo�͂��鏈��
    
    '�����܂ł��Ă��u�b�N�̃��������J������Ȃ���̌���������
    Dim outWb As Workbook:
'    Set outWb = ActiveWorkbook
    Set outWb = Workbooks.Add
    Dim outWs As Worksheet: Set outWs = outWb.Worksheets(1)
    Call VbeProcInfo_OutputWorksheet(data, outWs)
    DoEvents
    Set outWs = Nothing
    Set outWb = Nothing
End Sub

Rem �\�[�X�R�[�h�̃v���V�[�W���ꗗ���w��V�[�g�֏o��
Private Function VbeProcInfo_GetTable(source_vbp As VBProject) As Variant
    Dim dicProcInfo As New Dictionary
    Dim i As Long
    Dim dKey
  
    '�u�b�N�̑S���W���[��������
    For i = 1 To source_vbp.VBComponents.Count
        Dim dic As Dictionary
        Set dic = GetProcInfoDictionary(source_vbp.VBComponents(i).CodeModule)
        For Each dKey In dic.Keys
            dicProcInfo.Add dKey, dic(dKey)
        Next
        Set dic = Nothing
    Next
    If dicProcInfo.Count = 0 Then MsgBox "VBA������܂���B": Exit Function
  
    Dim data
    data = Array("���W���[��", "�s�ʒu", "�X�R�[�v", "���", "�v���V�[�W���[", "����", "�߂�l", "�R�����g", "�錾��")
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

Rem �v���V�[�W���ꗗ�񎟌��z��f�[�^���V�[�g�ɏo�͂���
Rem �����J,K��ɐ�����ǉ�����
Private Sub VbeProcInfo_OutputWorksheet(data, output_ws As Worksheet)

    'Dictionary���V�[�g�ɏo��
    output_ws.Name = "�v���V�[�W���ꗗ"
    output_ws.Parent.Activate
    output_ws.Parent.Windows(1).WindowState = xlMaximized
    
    With output_ws
        .Cells.Clear
        .Cells(1, 1).Resize(UBound(data, 1), UBound(data, 2)).Value = data
        
        '�錾�����ؗp�̎�
        .Range("J1").Value = "�錾��2"
        .Range("J2").FormulaR1C1 = "=RC[-7]&"" ""&RC[-6]&"" ""&RC[-5]&""(""&SUBSTITUTE(RC[-4],""" & Chr(10) & ""","", "")&"")""&IF(RC[-3]="""","""","" As ""&RC[-3])"
        .Range("K1").Value = "�`�F�b�N"
        .Range("K2").FormulaR1C1 = "=RC[-2]=RC[-1]"
        '1�s�]���Ƀt�B�������
        .Range("J2:K2").AutoFill Destination:=ResizeOffset(.UsedRange.Columns("J:K"), 1)
        
        .Range("A2").Select
        .Parent.Windows(1).FreezePanes = True
        .Cells.AutoFilter
        .Columns("A:K").EntireColumn.AutoFit
        .Columns("H:J").ColumnWidth = 16
        .Cells.WrapText = False
    End With
    
End Sub

Rem Range���w����W����Offset����Resize����i�擪�������Resize�j
Rem
Rem  @param rng         �Ώ�Range
Rem  @param offsetRow   �擪����I�t�Z�b�g�k������s��
Rem  @param offsetCol   �擪����I�t�Z�b�g�k�������
Rem
Rem  @return As Range   �ό`���Range
Rem
Function ResizeOffset(rng As Range, Optional offsetRow As Long, _
                                    Optional offsetCol As Long) As Range
    Set ResizeOffset = Intersect(rng, rng.Offset(offsetRow, offsetCol))
End Function

'�I�[�g�t�B���^����
Sub Test_ResizeOffset_AutoFilter()
    Const TARGET_COL = "C:D"
    ResizeOffset(ActiveSheet.AutoFilter.Range.Columns(TARGET_COL), 1).Select
End Sub

Sub Test_ResizeOffset()
    Const HEAD_ROW = 3, TARGET_COL = "C:D"
    
    With ToWorksheet(ActiveSheet)
        Dim rng As Range
        
        'UsedRange����
'        Set rng = ResizeOffset(.UsedRange.Columns(TARGET_COL), HEAD_ROW - .UsedRange.Row + 1)
        
        '�I�[�g�t�B���^����
'        Set rng = ResizeOffset(.AutoFilter.Range.Columns(TARGET_COL), HEAD_ROW - .AutoFilter.Range.Row + 1)

        'CurrentRegion����
'        Set rng = Range("B3").CurrentRegion
'        Set rng = Intersect(rng, rng.Offset(1)).Columns("C:D")

        rng.Select
    End With
End Sub

        
        'CurrentRegion����
'        Set rng = .Range(TARGET_COL).Cells(HEAD_ROW, 1).CurrentRegion
'        Set rng = ResizeOffset(rng.Columns(TARGET_COL), HEAD_ROW - rng.Row + 1)
        

'�I�[�g�t�B���^�̗L��
'UsedRange�O�̐擪�s��̗L��


'    Set OffsetResize = rng.Offset(offsetRow, offsetCol).Resize( _
'                            rng.Rows.CountLarge - offsetRow, _
'                            rng.Columns.CountLarge - offsetCol)

Rem �A�N�e�B�u�V�[�g��A��Ƀv���V�[�W����񂪋L�ڂ���Ă�����̂Ƃ���
Private Sub �v���V�[�W���ꗗ��������𕪉�����()
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

Rem Dictionary�Ƀv���V�[�W���[�E�v���p�e�B�����i�[
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

Rem �v���V�[�W���[�E�v���p�e�B��`�s���̔���
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

Rem Dictionary�Ƀv���V�[�W���[�E�v���p�e�B�����i�[
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

Rem �錾���ŕK�v�ȃf�[�^���m�F
Rem �\��  objCodeModule.CountOfDeclarationLines ������ς܂����s�ł��邱�ƁB
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

Rem �����񂪃R�����g�s��
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

Rem �p���s( _)�S�Ă�A������������ŕԂ�
Rem �R������R�����g�ȍ~�͏���
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

Rem �v���V�[�W���[�̒��O�̃R�����g���擾
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

    '�w�b�_�[���R�[�h���Z�b�g����
    Dim procRecords As Collection: Set procRecords = New Collection
    Dim procRecord(1 To 8) As String '���X�g�̗�
    procRecord(eRecord.���W���[����) = "���W���[����"
    procRecord(eRecord.���W���[���^�C�v) = "���W���[���^�C�v"
    procRecord(eRecord.�v���V�[�W����) = "�v���V�[�W����"
    procRecord(eRecord.�v���V�[�W���^�C�v) = "�v���V�[�W���^�C�v"
    procRecord(eRecord.�s��) = "�s��"
    procRecord(eRecord.����) = "����"
    procRecord(eRecord.�߂�l) = "�߂�l"
    procRecord(eRecord.�T�v) = "�T�v"
    procRecords.Add procRecord

    'Module��������������
    Dim module As Object
    For Each module In trgBook.VBProject.VBComponents

        '���W���[�������Z�b�g����
        procRecord(eRecord.���W���[����) = module.Name

        '���W���[���^�C�v���Z�b�g����
        procRecord(eRecord.���W���[���^�C�v) = FIX_MODULE_TYPE(module)

        'Module����Procedure�ꗗ���R���N�V��������
        Dim cModule As Object: Set cModule = module.CodeModule
        Dim procNames As Collection: Set procNames = COLLECT_PROCNAMES_IN_MODULE(cModule)

        'Procedure�̓��e��������������
        Dim ProcName As Variant, procTop As String
        For Each ProcName In procNames

            '�v���V�[�W�������Z�b�g����
            procRecord(eRecord.�v���V�[�W����) = ProcName

            '�v���V�[�W����1�s�ڂ��擾����
            procTop = SET_PROC_TOP(CStr(ProcName), cModule)

            '�v���V�[�W���^�C�v���Z�b�g����
            procRecord(eRecord.�v���V�[�W���^�C�v) = FIX_PROC_TYPE(CStr(ProcName), procTop)

            '�s�����Z�b�g����
            procRecord(eRecord.�s��) = cModule.ProcCountLines(ProcName, 0)

            '�������Z�b�g����
Rem             procRecord(eRecord.����) = FIX_PROC_ARGS(CStr(ProcName), procTop)

            '�߂�l���Z�b�g����
Rem             procRecord(eRecord.�߂�l) = FIX_PROC_RETURN(CStr(ProcName), procTop)

            '�T�v���Z�b�g����
            procRecord(eRecord.�T�v) = FIX_PROC_SUMMARY(CStr(ProcName), cModule)

            '���R�[�h���R���N�V��������
            procRecords.Add procRecord

        Next
    Next

    '�V�[�g�ɏ����o��
    Dim tmp As Variant, i As Long
    For Each tmp In procRecords
        i = i + 1
        With trgSheet
            .Cells(i, eRecord.���W���[����) = tmp(eRecord.���W���[����)
            .Cells(i, eRecord.���W���[���^�C�v) = tmp(eRecord.���W���[���^�C�v)
            .Cells(i, eRecord.�v���V�[�W����) = tmp(eRecord.�v���V�[�W����)
            .Cells(i, eRecord.�v���V�[�W���^�C�v) = tmp(eRecord.�v���V�[�W���^�C�v)
            .Cells(i, eRecord.�s��) = tmp(eRecord.�s��)
            .Cells(i, eRecord.����) = tmp(eRecord.����)
            .Cells(i, eRecord.�߂�l) = tmp(eRecord.�߂�l)
            .Cells(i, eRecord.�T�v) = tmp(eRecord.�T�v)
        End With
    Next


    '�����ڂ𐮂���
    ActiveWindow.DisplayGridlines = False
    With trgSheet.Cells
        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop

         '�܂�Ԃ��ĕ\��Flse�ATrue�̏���AutoFit��2�x�s���ƃ��C�A�E�g���J�b�`���ł���
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
    MsgBox "�V�[�g���uProcs�v�����݂��Ă��܂��B"

End Sub

Private Function COLLECT_PROCNAMES_IN_MODULE(cModule As Object) As Collection
Rem --------------------------------------------------------------------------------
Rem CodeModule���󂯎��A�܂܂��v���V�[�W�����̈ꗗ��Collection�ŕԂ��B
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
Rem Module���󂯎�胂�W���[���^�C�v�𕶎���ŕԂ��B
Rem --------------------------------------------------------------------------------

    Select Case module.Type
        Case 1
            FIX_MODULE_TYPE = "�W�����W���[��"
        Case 2
            FIX_MODULE_TYPE = "�N���X���W���[��"
        Case 3
            FIX_MODULE_TYPE = "���[�U�[�t�H�[��"
        Case 100
            FIX_MODULE_TYPE = "Excel�I�u�W�F�N�g"
        Case Else
            FIX_MODULE_TYPE = module.Type
    End Select
End Function

Private Function FIX_PROC_TYPE(ProcName As String, procTop As String) As String
Rem --------------------------------------------------------------------------------
Rem �v���V�[�W����1�s�ڂ��󂯎��A�v���V�[�W���^�C�v�𒊏o���ăe�L�X�g�ŕԂ��B
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
Rem '�v���V�[�W����1�s�ڂ��󂯎��A�����𒊏o���ăe�L�X�g�ŕԂ��B
Rem '��������ꍇ�̓Z�������s��t�^����B
Rem Rem --------------------------------------------------------------------------------
Rem     If InStr(procTop, ":") > 0 Then procTop = Left(procTop, InStr(procTop, ":") - 1)
Rem     If InStr(procTop, "'") > 0 Then procTop = Left(procTop, InStr(procTop, "'") - 1)
Rem     procTop = Trim(procTop) '��ɐ擪�X�y�[�X������
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
Rem '�v���V�[�W����1�s�ڂ��󂯎��A�߂�l�̌^�𒊏o���ăe�L�X�g�ŕԂ��B
Rem Rem --------------------------------------------------------------------------------
Rem     If InStr(procTop, ":") > 0 Then procTop = Left(procTop, InStr(procTop, ":") - 1)
Rem     If InStr(procTop, "'") > 0 Then procTop = Left(procTop, InStr(procTop, "'") - 1)
Rem     procTop = Trim(procTop) '��ɐ擪�X�y�[�X������
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
Rem ProcName��CodeModule���󂯎��A���̃v���V�[�W���̊T�v�𕶎���ŕԂ��B
Rem --------------------------------------------------------------------------------

    Dim StartRow As Long: StartRow = cModule.ProcStartLine(ProcName, 0)
    Dim LastRow As Long: LastRow = StartRow + cModule.ProcCountLines(ProcName, 0) - 1

    Dim reg As Object: Set reg = CreateObject("VBScript.RegExp")
    With reg
        .Pattern = "'----------.*" '�n�C�t��10�Ŕ���
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
Rem ProcName��CodeModule���󂯎��A���̃v���V�[�W����1�s�ڂ̓��e�𕶎���ŕԂ��B
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
Rem  Microsoft Visual Basic for Applications Extensibility 5.3 ���v���O�����ŎQ�Ɛݒ肷��}�N��
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
Debug.Print .Name, , .FullPath ' ���̎���Description�͎g���Ȃ�
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
refs.AddFromGuid "{0109E0F4-91AE-4736-A2CE-9D63E89D0EF6}", 1, 0 'XPS_SHL_DLLLib XPS_SHL_DLL 1.0 Type Library 32 bit�ł̂ݎQ�Ɛݒ�\
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

Rem  Office 2013 (version 15) ��Major,Minor�ԍ�����܂��Ă���ƍl���������
.AddFromGuid GUID_Excel, 1, 8
.AddFromGuid GUID_OUTLOOK, 9, 5
.AddFromGuid GUID_POWERPOINT, 2, 11
.AddFromGuid GUID_MSPub, 2, 2
.AddFromGuid "{3120BA9F-4FC8-4A4F-AE1E-02114F421D0A} ", 1, 0 'Microsoft OneNote 15.0 Extended Object Library 1.0
.AddFromGuid GUID_OLXLib, 1, 1
.AddFromGuid "{113D61B1-C7C0-4157-B694-43594E25DF45}", 1, 1 'eawfctrl 1.0 Type Library
.AddFromGuid GUID_WinWord, 8, 6 'Microsoft Word 15.0 Object Library

ElseIf Application.Version = 14 Then
.AddFromGuid "{E301A065-3DF5-4378-A829-57B1EA986631}", 1, 0 'OutlookSocialProvider 2013 �ȍ~�͂Ȃ�
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
Set refs = Application.VBE.ActiveVBProject.References 'Application.References�ł�Description���o�Ȃ��B���̂��߁@Refs �� Nothing �ɂ��āA ���̂悤�ɏ���������
For Each ref In refs
If ref.IsBroken = False Then
Debug.Print ref.Name, ref.GUID, ref.Major, ref.Minor, ref.Description, ref.FullPath
Else
refs.Remove ref
End If
Next
End Sub
    
Rem 14�����ɂȂ�悤�ɉE�񂹂ɂ���
Rem ��14���ȏ�̃f�[�^�͏�ʂ̌���������
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
Rem  * �f�o�b�O���O���t�@�C���ɏo�͂���
Rem  * @param varData              : �o�͑Ώۂ̃f�[�^
Rem  * @param Optional strFileNm   :(�o�̓t�@�C�������w�肷��ꍇ)�t�@�C����
Rem  * @param Optional lngDebugFLG :(0=�f�o�b�O�E�t�@�C���o��,1=�f�o�b�O�̂ݏo��,2=�t�@�C���̂ݏo��)
Rem  */
Public Sub OutputLog(ByVal varData As Variant, _
                     Optional ByVal lngDebugFLG As Long = 1, _
                     Optional ByVal strFileNm As String = "")
    
    Dim lngFileNum As Long
    Dim strLogFile As String
      
    '�t�@�C���o�͑Ώۂ̏ꍇ
    If lngDebugFLG = 0 Or lngDebugFLG = 2 Then
        ' �t�@�C�����̎w�肪�Ȃ��ꍇ�A���݂̔N�������t�@�C�����Ƃ���
        ' �����̃t�@�C�����Ɋg���q�����݂��Ȃ��ꍇ�A�g���q��t������
        If strFileNm = "" Then
          strFileNm = Format(Now(), "yyyymmdd") & ".txt"
        ElseIf InStr(strFileNm, ".txt") = 0 Then
          strFileNm = strFileNm & ".txt"
        End If
        
        ' �o�͐�t�@�C���ݒ�
        ' Access�ŗ��p����ꍇ��CurrentProject�I�u�W�F�N�g���g��
        ' strLogFile = CurrentProject.Path & "\" & strFileNm
        strLogFile = ActiveWorkbook.Path & "\" & strFileNm
        lngFileNum = FreeFile()
        Open strLogFile For Append As #lngFileNum
        Print #lngFileNum, varData
        Close #lngFileNum
    End If
    
    '�f�o�b�O���O�o�͑Ώۂ̏ꍇ
    If lngDebugFLG = 0 Or lngDebugFLG = 1 Then
        Debug.Print varData
    End If

End Sub

Rem msg�����b�Z�[�W�{�b�N�X�ɕ\������
Private Sub proc(msg As String)
    MsgBox msg
End Sub

Rem n��m�𑫂��֐�
Private Function FuncSum(n As Long, M As Long) As Long
    FuncSum = n + M
End Function

Private Sub Test1()
    Dim i As Long
    For i = 1 To 10
        If ActiveSheet.Cells(i, 2) = "���Ƃ�" Then Stop
    Next
End Sub

Private Sub Test2()
    Dim arr As Variant
    ReDim arr(1 To 3)
    arr(1) = Array("1", "2", "3", "4", "5")
    arr(2) = Array("�Ђ悱", "���Ƃ�", "����", "�Ђ�", "�˂�")
    arr(3) = Array("�҂�҂�", "����񂿂��", "�����", "���ӂ���", "�ɂ��ɂ��")
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


Private Sub Format�֐��Ő����̐擪��0��t����()
  Debug.Print Format("123", "00000") ' �擪��0��t����i�����̂݁j00123
  Debug.Print "[" & Format("ABC", "@@@@@") & "]"  ' ���p�̃X�y�[�X�𖄂߂ĉE��  [  ABC]
  Debug.Print "[" & Format("ABC", "!@@@@@") & "]" ' ���p�̃X�y�[�X�𖄂߂č���  [ABC  ]
End Sub


Rem http://yumem.cocolog-nifty.com/excelvba/2011/05/post-82d3.html
Rem �󗓂Ń��O�𗬂�
Rem �J�[�\���������ɂȂ��ƈӖ����Ȃ�
Rem �C�~�f�B�G�C�g�E�B���h�E��200�s�����\���ł��Ȃ��̂�199�o�͂������_�őS�ł���
Private Sub ImdFlush()
    Dim i As Long: For i = 1 To 199: Debug.Print: Next
End Sub

Rem �C�~�f�B�G�C�g�ɓK���Ƀf�[�^���o��
Rem �@�������J�[�\���̈ʒu����Ń_��
Rem �@���삪�d��
Private Sub ImdRandomData()
    Dim i As Long: For i = 1 To 10: Debug.Print Rnd: Next
End Sub

Rem �C�~�f�B�G�C�g�E�B���h�E��S�č폜����
Rem  ��\���̎��͓��삵�Ȃ��B�i�K�v�Ȃ�������Ȃ��j
Rem  VBE�I�u�W�F�N�g�A�N�Z�X�̋����K�v
Rem  ���肵�ē��삵�Ȃ�
Private Sub ImdClear_G_Home_End_Del_F7()
    On Error GoTo ENDPOINT
    If Application.VBE.MainWindow.Visible And _
        Application.VBE.Windows("�C�~�f�B�G�C�g").Visible Then
            SendKeys "^{g}", True
Rem             DoEvents               '���������ƁA�|�b�v�A�b�v���̓R�[�h�E�B���h�E���������
            SendKeys "^{Home}", True
            SendKeys "^+{End}", True
            SendKeys "{Del}", True
            SendKeys "{F7}", True
    End If
ENDPOINT:
End Sub

Rem ���̕��@�ɂ͂܂���肪����A
Rem ������Debug.Print ������ƁA�폜���o�̗͂��ꂪ�A�S�ďo�́�VBA�I����ɍ폜
Rem DoEvents������ƁA�폜����ؓ����Ȃ��B
Rem
Rem
Private Sub ImdClear_G_A_Del_F7()
    On Error GoTo ENDPOINT
    If Application.VBE.MainWindow.Visible And _
        Application.VBE.Windows("�C�~�f�B�G�C�g").Visible Then
Rem             Application.VBE.Windows("�C�~�f�B�G�C�g").Visible = True
            SendKeys "^g", True
            SendKeys "^a", True
            SendKeys "{Del}", True
            SendKeys "{F7}", True
    End If
ENDPOINT:
End Sub

Rem �C�~�f�B�G�C�g�E�B���h�E�̐擪�s�������Ă��ׂč폜����
Rem  �t�H�[�J�X���C�~�f�B�G�C�g�E�B���h�E�Ɏc��
Rem  ���肵�ē��삵�Ȃ�
Private Sub ImdClear_G_Home_Down_End_Del_F7()
    If Application.VBE.MainWindow.Visible And _
        Application.VBE.Windows("�C�~�f�B�G�C�g").Visible Then
            SendKeys "^{g}", True
Rem             DoEvents
            SendKeys "^{Home}", True
            SendKeys "{Down}", True
            SendKeys "^+{End}", True
            SendKeys "{Del}", True
            SendKeys "{F7}", True
    End If
End Sub

Rem �C�~�f�B�G�C�g�E�B���h�E�̖����Ƀt�H�[�J�X���ړ�����
Private Sub ImdCursolMoveToLast()
    If Application.VBE.MainWindow.Visible And _
        Application.VBE.Windows("�C�~�f�B�G�C�g").Visible Then
            SendKeys "^{g}", False
Rem             DoEvents
            SendKeys "^+{End}", False
            SendKeys "{F7}", False
            DoEvents
    End If
End Sub

Rem  �C�~�f�B�G�C�g�E�B���h�E�̓��e�𖕏�
Rem 1. �E�B���h�E�؂�ւ�
Rem 2. Ctrl+A
Rem 3. Delete
Rem 4. �A�N�f�B�u�E�B���h�E������
Public Sub ImdClear()
 
    Dim wd      As Object
    Dim wdwk    As Object
     
    Set wd = Application.VBE.Windows("�C�~�f�B�G�C�g")
    
    Application.VBE.Windows("�C�~�f�B�G�C�g").Visible = True
    
    Dim IsImdDocking As Boolean
    IsImdDocking = False
    
    '�h�b�L���O���Ȃ� ������Ď��s����ƃR�[�h��������
    If IsImdDocking Then
        wd.SetFocus
        SendKeys "^a", False
        SendKeys "{Del}", False
        'Application.SendKeys "^g ^a {DEL}"
    Else
    '�|�b�v�A�b�v���Ȃ�
    
    End If
    
End Sub

Public Sub ImdClearGAX()
    SendKeys "^g", Wait:=True ' �C�~�f�B�G�C�g �E�B���h�E��\�����܂��B
    SendKeys "^a", Wait:=True ' ���ׂđI��
Rem     SendKeys "^x", Wait:=True ' �؂���
    SendKeys "{Del}", Wait:=True ' �폜
End Sub

Private Sub Test_ImdCursolMoveToLast()
    Call ImdCursolMoveToLast
    Debug.Print "�Ōォ��o��"
End Sub

Private Sub VBE�E�B���h�E��S�ė�()
    Dim Item
    For Each Item In Application.VBE.Windows
        Debug.Print Item.Caption
    Next
End Sub

Private Sub VBE�E�B���h�E���w�肵���^������()
    Dim Item
    For Each Item In GetVbeWindow(vbext_wt_Immediate)
        Debug.Print Item.Caption
    Next
End Sub

Private Sub VBE�E�B���h�E�̃|�b�v�A�b�v������()
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

Rem ��u��VBE���J���� HomePersonal �v���W�F�N�g��I�����A�C�~�f�B�G�C�g���t�H�[�J�X����
Rem https://thom.hateblo.jp/entry/2015/08/16/025140
Rem �P�DVBIDE���g�p����ꍇ�́u�c�[���v�́u�Q�Ɛݒ�v���j���[��
Rem �uMicrosoft Visiual Basic for Applications Extensibility�v��ǉ����܂��B
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
            If hPane = 0 Then MsgBox "�C�~�f�B�G�C�g�E�B���h�E��������܂���B"
            If hPane < 1 Then Exit Sub
            
            'Ctrl��Shift�̏�Ԃ��L��
            GetKeyboardState savState(0)
            
            'Ctrl��������
            tmpState(vbKeyControl) = KEYSTATE_KEYDOWN
            SetKeyboardState tmpState(0)
            'Ctrl+END�𑗐M
            PostMessage hPane, WM_KEYDOWN, vbKeyEnd, 0&
            'SHIFT��������
            tmpState(vbKeyShift) = KEYSTATE_KEYDOWN
            SetKeyboardState tmpState(0)
            'CTRL+SHIFT+Home
            PostMessage hPane, WM_KEYDOWN, vbKeyHome, 0&
            'CTRL+SHIFT+BackSpace
            PostMessage hPane, WM_KEYDOWN, vbKeyBack, 0&
            
            'Ctrl��Shift�̏�Ԃ𕜌�
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
        MsgBox "VBA�v���W�F�N�g�ɃA�N�Z�X�ł��܂���B"
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
    'VBE���\������Ă��邩
    MsgBox Application.VBE.MainWindow.Visible
    
    '�C�~�f�B�G�C�g���\������Ă��邩
    MsgBox Application.VBE.Windows("�C�~�f�B�G�C�g").Visible
    
    '�C�~�f�B�G�C�g���h�b�L���O����Ă��邩
    '�C�~�f�B�G�C�g�̓|�b�v�A�b�v�\����
    
    '���݃t�H�[�J�X�̂���E�B���h�E�͂ǂꂩ
    
    
End Sub

Rem �C�~�f�B�G�C�g�E�B���h�E���\��
Private Sub ImdClose()
    Application.VBE.Windows("�C�~�f�B�G�C�g").Visible = False
    Debug.Print Application.VBE.ActiveWindow.Caption
End Sub

Rem �C�~�f�B�G�C�g�E�B���h�E��\��
Private Sub ImdShow()
    '�C�~�f�B�G�C�g�E�B���h�E��\��
    'True�ɂ���ƃt�H�[�J�X���C�~�f�B�G�C�g�Ɉڂ�
    Application.VBE.Windows("�C�~�f�B�G�C�g").Visible = True
    Debug.Print Application.VBE.ActiveWindow.Caption
    '�C�~�f�B�G�C�g�@�Əo��
    '�������AVBA�I����̃t�H�[�J�X��
    '�@�h�b�L���O���̓C�~�f�B�G�C�g
    '�@�|�b�v�A�b�v���̓R�[�h�E�B���h�E
    '�ɖ߂�
End Sub

Rem �C�~�f�B�G�C�g�E�B���h�E��\�����ăt�H�[�J�X��VBE�ɖ߂�
Private Sub ImdShow_UnFocus()
    Dim win As Object
    Set win = Application.VBE.ActiveWindow
    Application.VBE.Windows("�C�~�f�B�G�C�g").Visible = True
    Debug.Print Application.VBE.ActiveWindow.Caption
    win.SetFocus
    Debug.Print Application.VBE.ActiveWindow.Caption
End Sub


Rem ----------

Rem http://suyamasoft.blue.coocan.jp/ExcelVBA/Sample/VBProject/index.html

Rem  VBE�̃o�[�W������\�����܂��B
Private Sub Display_VBE_Version_Sample()
  MsgBox Prompt:="VBE.Version = " & Application.VBE.Version, Buttons:=vbInformation, Title:="VBE.Version"
End Sub

Rem  VBE�̃v���p�e�B���C�~�f�B�G�C�g �E�B���h�E�ɕ\�����܂��B
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

Rem  VBE�̃R�}���h�o�[�̈ꗗ�̃u�b�N���쐬���܂��B
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

Rem  �R�}���h �o�[�̃��Z�b�g���܂��B
Private Sub ResetCommandBars()
  Dim cb As CommandBar

  If MsgBox(Prompt:="��蒼���ł��܂��񂪁A���ׂĂ�VBE�̃R�}���h�o�[�����Z�b�g���܂����H", Buttons:=vbYesNo + vbQuestion, Title:="�m�F") <> vbYes Then Exit Sub
  On Error Resume Next
  Application.Cursor = xlWait ' �����v�^�J�[�\���|�C���^
  Application.StatusBar = "���ׂĂ�VBE�̃R�}���h�o�[�����Z�b�g���Ă܂��B���΂炭���҂���������..."
  For Each cb In Application.VBE.CommandBars
    If cb.BuiltIn Then
      cb.Reset ' �W���̃R�}���h �o�[�̓��Z�b�g
    Else
      cb.Delete ' ���[�U�[�̃R�}���h �o�[�͍폜
    End If
  Next
  Application.StatusBar = ""
  Application.Cursor = xlDefault ' �W���̃J�[�\���|�C���^
  On Error GoTo 0
End Sub

Rem  VBE�̃A�h�C���ꗗ���C�~�f�B�G�C�g �E�B���h�E�ɕ\�����܂��B
Private Sub Addin_Sample()
  Dim i As Long

  With Application.VBE.AddIns
    If .Count < 1 Then
      MsgBox Prompt:="VBE�̃A�h�C���̓C���X�g�[�����Ă܂���I", Buttons:=vbInformation, Title:="VBE.AddIns.Count"
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

Rem  Window�̈ꗗ���C�~�f�B�G�C�g �E�B���h�E�ɕ\�����܂��B
Rem  vbext_WindowType�i0=vbext_wt_CodeWindow, 5=vbext_wt_Immediate, 6=vbext_wt_ProjectWindow�Ȃǁj
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

Rem  �C�~�f�B�G�C�g �E�B���h�E��\�����A�N�e�B�u�ɂ��܂��B
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

Rem  ���ׂẴv���W�F�N�g�̃t�@�C�������C�~�f�B�G�C�g �E�B���h�E�ɕ\�����܂��B
Private Sub VBProjects_Sample()
  Dim i As Long

  For i = 1 To Application.VBE.VBProjects.Count
    Debug.Print Application.VBE.VBProjects(i).FileName
  Next i
End Sub

Rem  �A�N�e�B�u �v���W�F�N�g�̃v���p�e�B���C�~�f�B�G�C�g �E�B���h�E�ɕ\�����܂��B
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

Rem  �A�N�e�B�u �v���W�F�N�g�̃��[�h���C�~�f�B�G�C�g �E�B���h�E�ɕ\�����܂��B
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

Rem  �v���W�F�N�g��ۑ��������\�����܂��B
Private Sub Display_VBE_ActiveVBProject_Saved()
  MsgBox Prompt:="VBE.ActiveVBProject.Saved = " & Application.VBE.ActiveVBProject.Saved, Buttons:=vbInformation, Title:="VBE.ActiveVBProject.Saved"
End Sub

Rem   �A�N�e�B�u �v���W�F�N�g�u�Q�Ɛݒ�v�̈ꗗ���C�~�f�B�G�C�g �E�B���h�E�ɕ\�����܂��B
Rem  �G�N�Z���̃V�[�g�ɓ\��t������ŕ����̗�ɕ�����ɂ́A�u�f�[�^�v�^�u�́u��؂�ʒu�v�����s���J���}��I�����܂��B
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

Rem  �R���|�[�l���g�̃^�C�v���擾���܂��B
Rem  vbext_ComponentType
Rem      1 vbext_ct_StdModule = �W�����W���[��
Rem      2 vbext_ct_ClassModule = �N���X���W���[��
Rem      3 vbext_ct_MSForm = �t�H�[��
Rem     11 vbext_ct_ActiveXDesigner = ActiveXDesigner
Rem    100 vbext_ct_Document = �h�L�������g�iWorkbook,Worksheet�Ȃǁj
Private Sub VBComponents_Type_Sample()
  Dim i As Long

  If Excel.ActiveWorkbook Is Nothing Then Exit Sub
  With Excel.ActiveWorkbook.VBProject
    For i = 1 To .VBComponents.Count
      Debug.Print .VBComponents(i).Type & ", " & .VBComponents(i).Name
    Next i
  End With
End Sub
Rem  DeleteModule���W���[�����폜���܂��B
Private Sub VBComponents_Remove_Sample()
  Const DeleteName = "DeleteModule"
  Dim ret As VbMsgBoxResult
  Dim vbc As VBIDE.VBComponent

  On Error Resume Next
  Set vbc = ThisWorkbook.VBProject.VBComponents(DeleteName)
  On Error GoTo 0
  If vbc Is Nothing Then Exit Sub ' ���W���[���͑��݂��Ȃ��I
  ret = MsgBox(Prompt:=DeleteName & " ���W���[�����폜���܂����H", Buttons:=vbYesNo + vbQuestion, Title:="�m�F")
  If ret <> vbYes Then Exit Sub
  With ThisWorkbook.VBProject.VBComponents
    .Remove .Item(DeleteName)
  End With
End Sub
Rem  �I�����Ă郂�W���[������\�����܂��B
Private Sub SelectedVBComponent_Name_Sample()
  MsgBox Prompt:="�I�����Ă郂�W���[�����F" & Application.VBE.SelectedVBComponent.Name, Buttons:=vbYesNo + vbQuestion, Title:="SelectedVBComponent.Name"
End Sub

Sub Test_kccPath_ParentFolderPath()
    Dim p As kccPath
    
    '�����I��is_file:=False�Ƃ���΃t�H���_�F��
    Set p = kccPath.Init("C:\vba\hoge", False)
    Debug.Print p.CurrentFolderPath, p.ParentFolderPath
    
    '�p�X�̖��������Ȃ�t�H���_�F��
    Set p = kccPath.Init("C:\vba\hoge\")
    Debug.Print p.CurrentFolderPath, p.ParentFolderPath
    
    '���w��͌����t�@�C���F��
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
    
    Rem Json�e�L�X�g��UTF8�œǂݍ��݁A�K��ᔽ�̃R�����g�s���폜
    Dim jsonText As String
    jsonText = kccPath.ReadUTF8Text(ThisWorkbook.Path & "\" & SETTINGS_FILE_NAME)
    jsonText = kccWsFuncRegExp.RegexReplace(jsonText, "[ ]*//.*\r\n", "")
    Debug.Print jsonText
    Stop
    
    Rem Json���p�[�X
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

Rem �A�N�e�B�u�ȃv���W�F�N�g�փ\�[�X��SRC����C���|�[�g
Rem
Rem  /src/CodeName.bas.vba
Rem
Public Sub VBComponents_Import_SRC()
'    Call VBComponents_BackupAndInport_Sub( _
'            Application.VBE.ActiveVBProject, _
'            ".\..\src", _
'            "", "")
    MsgBox "������", vbOKOnly, "VBComponents_Import_SRC"
End Sub

Rem �A�N�e�B�u�ȃv���W�F�N�g�̃\�[�X�R�[�h��z���ɃG�N�X�|�[�g
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

Rem �A�N�e�B�u�ȃv���W�F�N�g��git�p�ɃG�N�X�|�[�g
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

Rem �A�N�e�B�u�ȃv���W�F�N�g��GIT�p�o�b�N�A�b�v���G�N�X�|�[�g
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
        MsgBox "���s����" & AppClass & "��������܂���ł����B", vbCritical + vbOKOnly, "BackupAndExport"
        Exit Sub
    End If
    
    Call VBComponents_BackupAndExport_Sub( _
            objApplication.VBE.ActiveVBProject, _
            ".\.\bin", _
            ".\.\src", _
            ".\.\backup\bin\[YYYYMMDD]_[HHMMSS]_[FILENAME]", _
            ".\.\backup\src\[YYYYMMDD]_[HHMMSS]\[FILENAME]")
End Sub

Rem  �v���W�F�N�g�̃\�[�X�R�[�h���G�N�X�|�[�g������o�b�N�A�b�v���鏈��
Rem
Rem  @param ExportObject    �o�̓v���W�F�N�g�iWorkbook,VBProject)
Rem  @param ExportBinFolder �G�N�X�|�[�gbin�t�H���_
Rem  @param ExportSrcFolder �G�N�X�|�[�gsrc�t�H���_
Rem  @param BackupBinFile   �o�b�N�A�b�vbin�t�@�C�������K��
Rem  @param BackupSrcFile   �G�N�X�|�[�gsrc�t�@�C�������K��
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
            MsgBox "[" & objFilePath.FileName & "] �͓ǂݎ���p�ł��B�����𒆎~���܂��B", vbOKOnly + vbCritical, PROC_NAME
            Exit Sub
        End If
        
        '�v���W�F�N�g�̏㏑���ۑ�
        Dim res As VbMsgBoxResult
        res = MsgBox(Join(Array( _
            objFilePath.FileName, _
            "�G�N�X�|�[�g�����s���܂��B", _
            "���s�O�Ƀu�b�N��ۑ����܂����H"), vbLf), vbYesNoCancel, PROC_NAME)
        Select Case res
            Case vbYes
                Call UserNameStackPush(" ")
                objFilePath.Workbook.Save
                Call UserNameStackPush
            Case vbNo
                '�������Ȃ�
            Case vbCancel
                Exit Sub
        End Select
    End If
    
    '�v���W�F�N�g�������[�X�t�H���_�֕���
    If ExportBinFolder <> "" Then
        Dim binPath As kccPath
        Set binPath = objFilePath.SelectPathToFolderPath(ExportBinFolder).ReplacePathAuto(DateTime:=NowDateTime)
        If binPath.FullPath <> objFilePath.CurrentFolder.FullPath Then
            binPath.DeleteItems
            binPath.CreateFolder
            If objFilePath.CurrentFolder.CopyTo(binPath, UseIgnoreFile:=True).IsAbort Then Exit Sub
        End If
    End If
    
    '�����\�[�X�̍폜�ƃG�N�X�|�[�g
    '�����\�[�X����U�ʂ̃t�H���_�Ɉړ����āA�o�͌�ɔ�r���āA���S��v�Ȃ犪���߂��B
    If ExportSrcFolder <> "" Then
        Dim srcPath As kccPath
        Set srcPath = objFilePath.SelectPathToFolderPath(ExportSrcFolder)
        Set srcPath = srcPath.ReplacePathAuto(DateTime:=NowDateTime, FileName:=objFilePath.Name)
        srcPath.CreateFolder
        
        'src_back�t�H���_���쐬����src�̒��g��src_back��
        Dim backPath As kccPath
        Set backPath = srcPath.SelectPathToFolderPath("..\" & srcPath.Name & "_back\")
        backPath.CreateFolder
        backPath.DeleteFiles
        backPath.DeleteFolders
        srcPath.MoveTo backPath
        
        'src�t�H���_���쐬���Ē���export
        Call VBComponents_Export(ExportObject, srcPath)
        Call CustomUI_Export(objFilePath, srcPath)
        
        'back����ύX���Ȃ�frx�𕜌�
        Dim f1 As File, f2 As File
        For Each f1 In srcPath.Folder.Files
            If f1.Name Like "*.frx" Then
                Dim isRestore As Boolean: isRestore = False
                For Each f2 In backPath.Folder.Files
                    If f1.Name = f2.Name Then
                        If f1.Size = f2.Size Then
                            '��v
                            Debug.Print "restore : " & f1.Name
                            f2.Copy f1.Path, True
                            isRestore = True
                        End If
                    End If
                Next
#If DEBUG_MODE Then
                'frx�����̂��S���X�V����Ă��܂��Ƃ��̊m�F�p
                If Not isRestore Then Stop
#End If
            End If
        Next
        
        '�^�C���X�^���v�̕���
        
        
        'back�t�H���_�̍폜
        backPath.DeleteFolder
    End If
    
    'bin��src�̃o�b�N�A�b�v
    If BackupBinFile <> "" Then
        binPath.CopyTo objFilePath.SelectPathToFilePath(BackupBinFile).ReplacePathAuto(DateTime:=NowDateTime), withoutFilterString:="*~$*"
    End If
    If BackupSrcFile <> "" Then
        srcPath.CopyTo objFilePath.SelectPathToFilePath(BackupSrcFile).ReplacePathAuto(DateTime:=NowDateTime)
    End If
    
    Debug.Print "VBA Exported : " & objFilePath.FileName
End Sub

Rem Application.UserName���ꎞ�I�ɏ㏑������
Rem
Rem @param OverrideUserName �w�莞:�ꎞ�I�ɏ㏑�����閼�O
Rem                         �ȗ���:���̖��O�ɕ���
Rem
Sub UserNameStackPush(Optional OverrideUserName)
    Static lastUserName
    If IsMissing(OverrideUserName) Then
        Application.UserName = lastUserName
    Else
        If OverrideUserName = "" Then _
            Err.Raise "���[�U�[�����󗓂ɂ���̂̓��O�C�����ɒu���������邽�ߋ֎~�ł�"
        lastUserName = Application.UserName
        Application.UserName = OverrideUserName
    End If
End Sub

Rem �v���W�F�N�g��CustomUI���w��t�H���_�ɃG�N�X�|�[�g
Rem
Rem  @param prj_path        �G�N�X�|�[�g�������u�b�N�̃p�X
Rem  @param output_path     �G�N�X�|�[�g��̃t�H���_
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

Rem �v���W�F�N�g��CustomUI���w��t�H���_�ɃG�N�X�|�[�g
Rem
Rem  @param prj_path        �G�N�X�|�[�g�������u�b�N�̃p�X
Rem
Private Sub CustomUI_ExportAndOpen(prj_path As kccPath)
    Const PROC_NAME = "CustomUI_ExportAndOpen"
    
    Dim tempPath As String
    With kccFuncZip.DecompZip(prj_path.FullPath, AutoDelete:=False)
        tempPath = .DecompFolder
    
        Dim xml1 As kccPath: Set xml1 = kccPath.Init(tempPath & "\" & "customUI\customUI.xml")
        Dim xml2 As kccPath: Set xml2 = kccPath.Init(tempPath & "\" & "customUI\customUI14.xml")
        
        If xml1.Exists Or xml2.Exists Then
            Select Case MsgBox(Replace("�͂��F�t�@�C�����J��\n�������F�t�H���_���J��\n", "\n", vbLf), vbYesNo)
                Case VbMsgBoxResult.vbYes
                    If xml1.Exists Then kccFuncWindowsProcess.OpenAssociationAPI xml1.FullPath
                    If xml2.Exists Then kccFuncWindowsProcess.OpenAssociationAPI xml2.FullPath
                Case VbMsgBoxResult.vbNo
                    Shell "explorer " & tempPath, vbNormalFocus
            End Select
        Else
            MsgBox "CustomUI�͊܂܂�Ă��Ȃ��悤�ł��B", vbOKOnly, PROC_NAME
        End If
    End With
    
End Sub

Public Sub CurrentProject_CustomUI_Import()
    MsgBox "������", vbOKOnly, "CurrentProject_CustomUI_Import"
End Sub

Public Sub CurrentProject_CustomUI_Export()
    Call CustomUI_ExportAndOpen(kccPath.Init(Application.VBE.ActiveVBProject))
End Sub

Private Sub Test_CustomUI��temp�ɓW�J���ĊJ���Ă݂邾��()
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

Private Sub Test_Zip_�ꎞ�t�H���_�̎����폜����()
    Const Path = "C:\vba\test_CustomUI_Export.xlam"
    Dim tempPath As String
    
    Debug.Print "-----temp�ւ̓W�J-----"
    
    With kccFuncZip.DecompZip(Path)
        tempPath = .DecompFolder
    End With
    Debug.Print "�����폜���w��(ON)", fso.FolderExists(tempPath) = False, tempPath
    Application.Wait [Now() + "00:00:01"]
    
    With kccFuncZip.DecompZip(Path, , AutoDelete:=False)
        tempPath = .DecompFolder
    End With
    Debug.Print "�����폜OFF", fso.FolderExists(tempPath) = True, tempPath
    Application.Wait [Now() + "00:00:01"]
    
    With kccFuncZip.DecompZip(Path, , AutoDelete:=True)
        tempPath = .DecompFolder
    End With
    Debug.Print "�����폜ON", fso.FolderExists(tempPath) = False, tempPath
    Application.Wait [Now() + "00:00:01"]
    
    
    Debug.Print "-----xlam�Ɠ����t�H���_�ւ̓W�J-----"
    
    With kccFuncZip.DecompZip(Path, "\")
        tempPath = .DecompFolder
    End With
    Debug.Print "�����폜���w��(OFF)", fso.FolderExists(tempPath) = True, tempPath
    Application.Wait [Now() + "00:00:01"]
    
    With kccFuncZip.DecompZip(Path, "\", AutoDelete:=False)
        tempPath = .DecompFolder
    End With
    Debug.Print "�����폜OFF", fso.FolderExists(tempPath) = True, tempPath
    Application.Wait [Now() + "00:00:01"]
    
    With kccFuncZip.DecompZip(Path, "\", AutoDelete:=True)
        tempPath = .DecompFolder
    End With
    Debug.Print "�����폜ON", fso.FolderExists(tempPath) = False, tempPath
    Application.Wait [Now() + "00:00:01"]
    
End Sub

Private Function isVBProjectProtected(prj As VBProject) As Boolean
    On Error Resume Next
    Dim dummy
    Set dummy = prj.VBComponents
    On Error GoTo 0
    isVBProjectProtected = IsEmpty(dummy)
End Function

Rem �v���W�F�N�g�̃\�[�X�R�[�h���w��t�H���_�ɃG�N�X�|�[�g
Rem
Rem
Private Sub VBComponents_Export(prj As VBProject, output_path As kccPath)
    If prj Is Nothing Then MsgBox "VBA�v���W�F�N�g����", vbOKOnly, "Export Error": Exit Sub
    If isVBProjectProtected(prj) Then MsgBox "VBA�v���W�F�N�g�̃��b�N���������Ă�������", vbOKOnly, "Export Error": Exit Sub
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
            
            Debug.Print "Export", cmp.Name, , "�錾��", declDic.Count, , "�֐���", procDic.Count
            
            Dim oFullPath As String: oFullPath = ""
            Select Case cmp.Type
                Case Is = vbext_ct_StdModule
                    oFullPath = output_path.FullPath & "\" & cmp.Name & ".bas" & ".vba"
                  
                Case Is = vbext_ct_ClassModule
                    oFullPath = output_path.FullPath & "\" & cmp.Name & ".cls" & ".vba"
                  
                Case Is = vbext_ct_MSForm
                    oFullPath = output_path.FullPath & "\" & cmp.Name & ".frm" & ".vba"
                  
                ' Workbook, Worksheet�Ȃ�
                Case Is = vbext_ct_Document
                    oFullPath = output_path.FullPath & "\" & cmp.Name & ".cls" & ".vba"
                  
                ' ActiveX �f�U�C�i
                Case Is = vbext_ct_ActiveXDesigner
                    oFullPath = output_path.FullPath & "\" & cmp.Name & ".cls" & ".vba"
            End Select
            
            If oFullPath <> "" Then
                cmp.Export oFullPath
                
                '���ɂ����frm�̍��W�� .001 ���t�^����錻�ۂ̉���
                Call RepairFrm(oFullPath)
                
                '�R�[�h�̖����ɕs�v�ȉ��s�����肪���Ȗ��̉���
                Call CleanSource(oFullPath)
                
                'UTF-8�ւ̕ϊ�
'                kccPath.Init(oFullPath, True).ConvertCharCode_SJIS_to_utf8
            End If
ForContinue:
        Next
    End With
End Sub

'���ɂ����frm�̍��W�� .001 ���t�^����錻�ۂ̉���
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

Rem �e�L�X�g�t�@�C�������̕s�v�ȉ��s����菜��
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

Rem  �w�肵�����O��VBComponent�����݂��Ă��邩���ׂ܂��B
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

Rem  �A�N�e�B�u ���W���[���̐錾�Z�N�V���������̍s����Ԃ��܂��B
Private Sub CountOfDeclarationLines_Sample()
  Dim Line As Long

  Line = Application.VBE.ActiveCodePane.CodeModule.CountOfDeclarationLines
  MsgBox Prompt:="�錾�Z�N�V���������̍s���F" & Line, Buttons:=vbInformation, Title:="CodeModule.CountOfDeclarationLines"
End Sub

Rem  �A�N�e�B�u ���W���[���̍s����Ԃ��܂��B
Private Sub CountOfLines_Sample()
  Dim Line As Long

  Line = Application.VBE.ActiveCodePane.CodeModule.CountOfLines
  MsgBox Prompt:="���W���[���̍s���F" & Line, Buttons:=vbInformation, Title:="CodeModule.CountOfLines"
End Sub

Rem  �v���V�[�W���[�̍s����Ԃ��܂��B
Rem  �y���Ӂz�v���V�[�W���[�̑O�̍s�ɃR�����g������ꍇ�́A�R�����g�̍s���܂߂܂��B
Private Sub ProcCountLines_Sample()
  Dim StartLine As Long

  StartLine = ThisWorkbook.VBProject.VBComponents("Code_Module").CodeModule.ProcCountLines(ProcName:="ProcCountLines_Sample", ProcKind:=vbext_pk_Proc)
  MsgBox Prompt:="�R�����g�̍s���܂ރv���V�[�W���[�̍s���F" & StartLine, Buttons:=vbInformation, Title:="CodeModule.ProcCountLines"
End Sub

Rem  �v���V�[�W���[�̊J�n�s��Ԃ��܂��B�i�v���V�[�W���[�̑O�̍s�ɂ���R�����g�s���܂݂܂��B�j
Rem  �y���Ӂz�O�̃v���V�[�W���[�̎��̍s��Ԃ��܂��B
Rem   vbext_ProcKind
Rem     vbext_pk_Get
Rem     vbext_pk_Let
Rem     vbext_pk_Proc
Rem     vbext_pk_Set
Private Sub ProcStartLine_Sample()
  Dim StartLine As Long

  StartLine = ThisWorkbook.VBProject.VBComponents("Code_Module").CodeModule.ProcStartLine(ProcName:="ProcStartLine_Sample", ProcKind:=vbext_pk_Proc)
  MsgBox Prompt:="�R�����g�s���܂ރv���V�[�W���[�̊J�n�s�F" & StartLine, Buttons:=vbInformation, Title:="CodeModule.ProcStartLine"
End Sub

Rem  �v���V�[�W���[�̊J�n�s��Ԃ��܂��B
Private Sub ProcBodyLine_Sample()
  Dim StartLine As Long

  StartLine = ThisWorkbook.VBProject.VBComponents("Code_Module").CodeModule.ProcBodyLine(ProcName:="ProcBodyLine_Sample", ProcKind:=vbext_pk_Proc)
  MsgBox Prompt:="�v���V�[�W���[�̊J�n�s�F" & StartLine, Buttons:=vbInformation, Title:="CodeModule.ProcBodyLine"
End Sub

Rem  �R�[�h���W���[���̎w��s����w�肵���s���̃e�L�X�g���擾���܂��B
Rem  �y���ӁzCodeModule.Lines��Unicode�Ȃ̂ŁA���p�ł�2�o�C�g�ł��B
Private Sub Lines_Sample()
  Dim StartLine As Long, Count As Long

  StartLine = 3
  Count = 8
  MsgBox Prompt:=ThisWorkbook.VBProject.VBComponents("Code_Module").CodeModule.Lines(StartLine, Count), Buttons:=vbInformation, Title:="CodeModule.Lines"
End Sub

Rem  �w�肵���s���܂܂��v���V�[�W���[�����擾���܂��B
Private Sub ProcOfLine_Sample()
  Dim num As Variant
  Dim ProcName As String

  num = Application.InputBox(Prompt:="�s���F", Title:="�v���V�[�W���[���̍s���̓���", Default:=57, Type:=1)
  If TypeName(num) <> "Double" Then Exit Sub ' [�L�����Z��]�{�^��
  ProcName = ThisWorkbook.VBProject.VBComponents("Code_Module").CodeModule.ProcOfLine(Line:=num, ProcKind:=vbext_pk_Proc)
  MsgBox Prompt:="�v���V�[�W���[���F" & ProcName, Buttons:=vbInformation, Title:="CodeModule.ProcOfLine"
End Sub

Rem  �I�������e�L�X�g�t�@�C����TempModule���W���[���̍ŏ��̃v���V�[�W���[�̑O�ɑ}�����܂��B
Private Sub AddFromFile_Sample()
  Dim FileName As Variant

  FileName = Application.GetOpenFileName(FileFilter:="�e�L�X�g�t�@�C���i*.txt�j, *.txt,���ׂẴt�@�C���i*.*�j,*.*", FilterIndex:=1, Title:="�t�@�C���̃C���|�[�g", ButtonText:="�C���|�[�g", MultiSelect:=False)
  If TypeName(FileName) = "Boolean" Then Exit Sub ' [�L�����Z��]�{�^��
  ThisWorkbook.VBProject.VBComponents("TempModule").CodeModule.AddFromFile FileName
End Sub

Rem  �e�L�X�g��TempModule���W���[���̍ŏ��̃v���V�[�W���[�̑O�ɑ}�����܂��B
Private Sub AddFromString_Sample()
  Dim Str As String

  Str = "'" & String(50, "=") & vbCrLf
  Str = Str & "'AddFromString�ő}�����܂����B " & Format(Now, "yyyy/mm/dd hh:mm:ss") & vbCrLf & Str
  ThisWorkbook.VBProject.VBComponents("TempModule").CodeModule.AddFromString Str
End Sub

Rem  �e�L�X�g��TempModule���W���[����5�s�ڂɑ}�����܂��B
Private Sub InsertLines_Sample()
  Dim Str As String

  Str = "' 5�s�ڂ�InsertLines�ő}�����܂����B" & vbCrLf & "' vbCrLf���g�p����ƕ����̍s��}���ł��܂��B"
  ThisWorkbook.VBProject.VBComponents("TempModule").CodeModule.InsertLines 5, Str
End Sub

Rem  ���݂̃J�[�\���̊J�n�s�ɓ��t�Ǝ��Ԃ�}�����܂��B
Private Sub Insert_Text()
  Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
  Dim Text As String

  Text = Format(Now, "' ggge�Nmm��dd�� hh��mm��ss�b")
  With Application.VBE.ActiveCodePane
    .getSelection StartLine, StartColumn, EndLine, EndColumn
    .CodeModule.InsertLines StartLine, Text
  End With
End Sub

Rem  TempModule���W���[����5�s�ڂ�6�s�ڂ�2�s���폜���܂��B
Private Sub DeleteLines_Sample()
  Dim StartLine As Long, CountLine As Long

  StartLine = 5
  CountLine = 2
  ThisWorkbook.VBProject.VBComponents("TempModule").CodeModule.DeleteLines StartLine, CountLine
End Sub

Rem  �������������񂪂��邩��\�����܂��B
Rem  Find(Target As String, StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long, [WholeWord As Boolean = False], [MatchCase As Boolean = False], [PatternSearch As Boolean = False]) As Boolean
Rem  �y���ӁzStartColumn��EndColumn�̌��͔��p��1�A�S�p��2�Ōv�Z���܂��B
Private Sub Find_Sample()
  Dim ret As Boolean
  Dim FindText As Variant

  FindText = Application.InputBox(Prompt:="������F", Title:="������̌���", Type:=2)
  If TypeName(FindText) = "Boolean" Then Exit Sub
  If Len(FindText) < 1 Then Exit Sub
  With ThisWorkbook.VBProject.VBComponents("Code_Module").CodeModule
    ret = .Find(FindText, 1, 1, .CountOfLines, LenB(.Lines(.CountOfLines, 1)), False, False, False) ' �y���ӁzLen�ł͂Ȃ�LenB���g���܂��B
  End With
  MsgBox Prompt:=FindText & "�̌������� = " & ret, Buttons:=vbInformation, Title:="������̌���"
End Sub

Rem  TempModule���W���[����5�s�ڂ𕶎���Œu�������܂��B
Private Sub ReplaceLine_Sample()
  Dim Str As String

  Str = "' 5�s�ڂ�ReplaceLine�Œu���������܂����B " & Format(Now, "yyyy/mm/dd hh:mm:ss")
  ThisWorkbook.VBProject.VBComponents("TempModule").CodeModule.ReplaceLine 5, Str
End Sub

Rem  �A�N�e�B�u �R���|�[�l���g�̂��ׂẴv���V�[�W������\�����܂��B�i�N���X��Get, Set, Let�͏����܂��j
Private Sub Display_ProcName_Sample()
  Dim msg As String
  Dim ProcName As String
  Dim i As Long

  ProcName = vbNullString
  With Application.VBE.ActiveCodePane.CodeModule
    For i = 1 To .CountOfLines
      If ProcName <> .ProcOfLine(i, ProcKind:=vbext_pk_Proc) Then ' �v���V�[�W�������ς�����ꍇ��
        ProcName = .ProcOfLine(i, ProcKind:=vbext_pk_Proc)
Rem         Debug.Print buf
        msg = msg & ProcName & vbCrLf
      End If
    Next i
  End With

  MsgBox Prompt:=msg, Buttons:=vbInformation, Title:="�v���V�[�W�����̈ꗗ"
End Sub

Rem  �I��͈͂̃J�[�\���ʒu���擾���܂��B
Rem  �y���ӁzStartColumn��EndColumn�̌��͔��p��1�A�S�p��2�Ōv�Z���܂��B
Private Sub GetSelection_Sample()
  Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
  Dim msg As String

  Application.VBE.ActiveCodePane.getSelection StartLine, StartColumn, EndLine, EndColumn
  msg = "�J�n�F" & StartLine & "�s " & StartColumn & "��" & vbCrLf & vbCrLf
  msg = msg & "�I���F" & EndLine & "�s" & EndColumn & "��"
  MsgBox Prompt:=msg, Buttons:=vbInformation, Title:="�J�[�\���ʒu"
End Sub

Rem  �I��͈͂�ݒ肵�܂��B
Rem  �y���ӁzStartColumn��EndColumn�̌��͔��p��1�A�S�p��2�Ōv�Z���܂��B
Rem  �y���Ӂz���[�̌���0�ł͂Ȃ�1�ł��B
Private Sub SetSelection_Sample()
  Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long

  StartLine = 14: StartColumn = 7: EndLine = 22: EndColumn = 8 ' �� StartColumn�͑S�p��2�Ōv�Z���Ă܂��B
  Application.VBE.ActiveCodePane.SetSelection StartLine, StartColumn, EndLine, EndColumn
End Sub

Rem  �A�N�e�B�u �R�[�h �y�C���̉�ʂɕ\���ł���s����\�����܂��B
Private Sub CountOfVisibleLines_Sample()
  MsgBox Prompt:="��ʂɕ\���ł���s���F" & Application.VBE.ActiveCodePane.CountOfVisibleLines, Buttons:=vbInformation, Title:="ActiveCodePane.CountOfVisibleLines"
End Sub

Rem  �A�N�e�B�u �R�[�h �y�C���̉�ʂ̍ŏ�s��\�����܂��B
Private Sub TopLine_Sample()
  MsgBox Prompt:="��ʂ̍ŏ�s�F" & Application.VBE.ActiveCodePane.TopLine, Buttons:=vbInformation, Title:="ActiveCodePane.TopLine"
End Sub

Rem �S�ẴR�[�h�E�C���h�E�����
Public Sub CloseCodePanes()
    Dim C As CodePane
    For Each C In Application.VBE.CodePanes
        C.Window.Close
    Next
End Sub

Rem ���݂̃J�[�\���ɂ���֐��̃e�X�g�����s����
Public Sub TestExecute()
    Run GetCursolFunctionName()
    MsgBox "������"
End Sub

Rem ���݂̃J�[�\���ɂ���֐��̃e�X�g�փW�����v����
Public Sub TestJump()
    ProcJump GetCursolFunctionName()
    MsgBox "������"
End Sub

Rem ���݂̃J�[�\���ɂ���֐�����Ԃ�
Private Function GetCursolFunctionName()
    
End Function

Rem �w�肵���֐����̏ꏊ�ɃW�����v����
Private Sub ProcJump(func As String)
    
    Rem ���W���[�����J��
    Rem �J�[�\���ʒu��ς���
End Sub

Rem �t�@�C��������Ă��Ȃ��u�b�N�S�Ă�ۑ������ɕ���
Public Sub CloseNofileWorkbook()
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Path = "" Then
            wb.Close False
        End If
    Next
End Sub

Rem VBA�v���W�F�N�g�̃p�X���[�h��1234�֕ύX����
Public Sub BreakPassword1234Project()
    Dim beforePath As kccPath: Set beforePath = kccPath.Init(Application.VBE.ActiveVBProject)
    Dim afterPath As kccPath: Set afterPath = beforePath.SelectPathToFilePath("|t_1234|e")
    Select Case MsgBox(beforePath.FileName & "��" & afterPath.FileName & "�֏o�͂��܂��B", vbOKCancel)
        Case vbOK
            Dim res: res = BrokenVbaPassword(beforePath.FullPath, afterPath.FullPath)
            afterPath.OpenExplorer
            MsgBox "�����I�I�I" & res, vbOKOnly
        Case vbCancel
    End Select
End Sub

Public Sub OpenFormDeclareSourceGenerate()
    FormDeclareSourceGenerate.Show
End Sub

Public Sub OpenFormDeclareSourceTo64bit()
    FormDeclareSourceTo64bit.Show
End Sub

Rem �����t�H���_�A���͏�ʃt�H���_�̑啶���������t�@�C�����J��
Public Sub OpenTextFileBy�啶��������()
    Dim targetPath As kccPath
    Set targetPath = kccPath.Init(ThisWorkbook.Path, False).SelectPathToFilePath(DEF_�啶���������t�@�C��)
    If Not targetPath.Exists Then
        Set targetPath = targetPath.SelectParentFolder("..\")
    End If
    targetPath.OpenAssociation
End Sub

'.vba�����Ă��Ȃ������t�@�C���ɕt������
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
