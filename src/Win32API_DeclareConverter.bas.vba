Attribute VB_Name = "Win32API_DeclareConverter"
Rem Win32API��Declare���������I��64bit�Ή��R�[�h�ɕϊ�����v���O����
Rem
Rem �����J��
Rem
Rem �������邿��񂿂��
Rem 2019/10/20
Rem VBA��Win32API��64bit�Ή������ϊ��v���O����������Ă݂�
Rem https://www.excel-chunchun.com/entry/vba-64bit-declare-convert
Rem
Rem ----------------------------------------------------------------------------------------------------
Rem
Rem ���Q�l����
Rem
Rem 64 �r�b�g Visual Basic for Applications �̊T�v
Rem  https://docs.microsoft.com/ja-jp/office/vba/language/concepts/getting-started/64-bit-visual-basic-for-applications-overview
Rem
Rem Office �� 32 �r�b�g �o�[�W������ 64 �r�b�g �o�[�W�����Ԃ̌݊���
Rem  https://docs.microsoft.com/ja-jp/office/client-developer/shared/compatibility-between-the-32-bit-and-64-bit-versions-of-office
Rem
Rem Declaring API functions in 64 bit Office
Rem  https://www.jkp-ads.com/articles/apideclarations.asp
Rem
Rem ----------------------------------------------------------------------------------------------------
Rem
Rem ���X�V����
Rem
Rem  2019/10/20 : Declare������e����ɑΉ�����Declare���֕ϊ�����֐�
Rem  2019/10/21 : �֐�������e����ɑΉ�����Declare���𐶐�����֐�
Rem
Rem ----------------------------------------------------------------------------------------------------
Rem
Rem ���g����
Rem
Rem VBA�\�[�X�R�[�h��Declare����32/64bit�Ή��ɕϊ�����֐�
Rem
Rem  @name ConvertVBACodeDeclare
Rem
Rem  @param vbaCodeText VBA�\�[�X�R�[�h������ivbCrLf�j
Rem
Rem  @return As String  VBA�\�[�X�R�[�h������
Rem
Rem  @example
Rem    IN  : �K���ȃ\�[�X�R�[�h������
Rem    OUT : VBA6/7 Win32/64 �Ή��\�[�X�R�[�h�i����͕ۏ؂��Ȃ��j
Rem
Rem  �Ή��ł��Ȃ���
Rem  �EAPI�ɂ���Ă̓p�����[�^���ς���Ă��鎖������B
Rem  �E�p�����[�^��Long����LongPtr/LongLong�ɕω����ČĂяo�������ύX�̕K�v������B
Rem  �E�\���̂̎d�l���ς���Ă���/����`�̏ꍇ������B
Rem  �EWin32API_PtrSafe.txt�Ɍf�ڂ���Ă��Ȃ��֐��ɂ͑Ή����Ă��Ȃ��B
Rem  �EGetWindowLong����GetWindowLongPtr�ɕύX���Ȃ��Ǝg���Ȃ��B
Rem
Rem  Function ConvertVBACodeDeclare(vbaCodeText) As String
Rem   IN  : �K���ȃ\�[�X�R�[�h������
Rem   OUT : VBA6/7 Win32/64 �Ή��\�[�X�R�[�h�i����͕ۏ؂��Ȃ��j
Rem
Rem
Rem Win32API�֐����𗅗񂵂��e�L�X�g��Declare�ɕϊ�����֐�
Rem
Rem  @name GetDeclareCodeByText
Rem
Rem  @param base_str       �֐��������̍s�̊܂܂ꂽ�e�L�X�g
Rem  @param useVBA6        VBA6�Ή��R�[�h�𐶐����邩
Rem  @param useVBA7        VBA7�Ή��R�[�h�𐶐����邩(32bit/64bit)
Rem
Rem  @return As String     Declare�錾��
Rem
Rem
Rem Win32API�֐�����n������S�Ή���Declare����ԋp����֐�
Rem
Rem  name GetDeclareCodeByProcName
Rem
Rem  @param procName       �����Ώۂ̊֐���
Rem  @param useVBA6        VBA6�Ή��R�[�h�𐶐����邩
Rem  @param useVBA7        VBA7�Ή��R�[�h�𐶐����邩(32bit/64bit)
Rem  @param indent_level   �C���f���g���i2~�j
Rem
Rem  @return As String Declare���@����������procName
Rem
Rem  @note VBA6,7�������g�p����ꍇ�����f�B���N�e�B�u�ɂ�镪�򂪐��������
Rem
Rem
Option Explicit

Private dicPtrSafe_ As Dictionary
Private dicPtrSafe32_ As Dictionary
Private dicPtrSafe64_ As Dictionary

Private Sub Sample()
    Const TEST_FILE = "Win32API�ϊ��e�X�g.bas"
    Const PARAM_INDENT_LEVEL = 10
    
    Dim fso As New FileSystemObject

    '�ϊ��O
    Dim vbaCodeText
    vbaCodeText = fso.OpenTextFile(ThisWorkbook.Path & "\" & TEST_FILE, ForReading, False).ReadAll()
    
    '�ϊ���
    Dim replacedText
    replacedText = ConvertVBACodeDeclare(vbaCodeText, PARAM_INDENT_LEVEL)
    
    '�擪40�s�����C�~�f�B�G�C�g�֏o��
    Dim idxs
    idxs = kccFuncString.InStrAll(replacedText, vbCrLf)
    Debug.Print Left(replacedText, idxs(40))
    
    '�t�@�C���o��
    fso.OpenTextFile(ThisWorkbook.Path & "\" & TEST_FILE & "_conv.txt", ForWriting, True).Write replacedText
End Sub

Rem Microsoft�����̐錾������͂��Č��{�������ɕێ�����
Rem  @param bit =  0 : bit�Ɉˑ����Ȃ�
Rem               32 : 32bit��p
Rem               64 : 64bit��p
Rem  @return As Dictionary param�ɑΉ��������� (��key�͏���������j
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
        
        '�����̊֐���Txt�œ�d��`����Ă���̂ŋ��e����B
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
'                         ��d��`�����e
'                         �Ǝ��ɒǉ������֐��ŏd�������������ꍇ�Ɍ��m�������̂Ŋ����Ă��������B
                    ElseIf dicPtrSafe_.Exists(ProcName) Then
                        Debug.Print ProcName & "�͓�d��`�H"
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

Rem VBA�\�[�X�R�[�h��Declare����32/64bit�Ή��ɕϊ�
Rem
Rem  @name ConvertVBACodeDeclare
Rem
Rem  @param vbaCodeText VBA�\�[�X�R�[�h������ivbCrLf�j
Rem
Rem  @return As String  VBA�\�[�X�R�[�h������
Rem
Rem  @example
Rem    IN  : �K���ȃ\�[�X�R�[�h������
Rem    OUT : VBA6/7 Win32/64 �Ή��\�[�X�R�[�h�i����͕ۏ؂��Ȃ��j
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
    
    '�錾�G���A�ŏI�s�����
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
    
    '�R�����g�ƃC���f���g����
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
    
    '�X�e�[�g�����g���s��A��
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
    
    '---�����܂őO����
    
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
        
        '������Declare���͐��������̂Ɖ��肵�Ă�������̕���ǉ�����
        'VBA7�f�B���N�e�B�u���ɋL�q����Ă��鎞�͊��ɑΏ��ς݂Ɣ��f���ϊ��͍s��Ȃ�
        If v Like "*Declare *" Then
            If v Like "*Declare PtrSafe *" Then
                'VBA7�錾��
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
                'VBA6(64bit��Ή�)�錾��
                If vbaMode = 0 Then
                    '��Ή�
                    arr = Array("", _
                                "#If VBA7 Then", _
                                kccFuncString.InsertIndent(ReplaceDeclareTo7(v, indent_level)), _
                                "#Else", _
                                kccFuncString.InsertIndent(InsertDeclareIndent(v, indent_level)), _
                                "#End If")
                    v = Join(arr, vbCrLf)
                ElseIf vbaMode = 7 Then
                    '�Ή��R��(�f�B���N�e�B�u���Ȃ̂�PtrSafe���ĂȂ��j
                    v = ReplaceDeclareTo7(v, indent_level)
                End If
            End If
        End If
        
        vbaLines(i) = v
    Next
    
    '---��������㏈��
    
    '�R�����g�ƃC���f���g�𕜌�
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

Rem Win32API�֐����𗅗񂵂��e�L�X�g��Declare�ɕϊ�����֐�
Rem
Rem  @name GetDeclareCodeByText
Rem
Rem  @param base_str       �֐��������̍s�̊܂܂ꂽ�e�L�X�g
Rem  @param useVBA6        VBA6�Ή��R�[�h�𐶐����邩
Rem  @param useVBA7        VBA7�Ή��R�[�h�𐶐����邩(32bit/64bit)
Rem
Rem  @return As String     Declare�錾��
Rem
Public Function GetDeclareCodeByText(base_str, useVBA6 As Boolean, useVBA7 As Boolean, indent_level As Long) As String
    Dim Rows, i
    Rows = Split(base_str, vbCrLf)
    For i = LBound(Rows) To UBound(Rows)
        Rows(i) = GetDeclareCodeByProcName(Rows(i), useVBA6, useVBA7, indent_level)
    Next
    GetDeclareCodeByText = Join(Rows, vbCrLf)
End Function

Rem �K����Declare������VBA6�Ή��R�[�h�ɒu��(�s���S)
Rem  �P����VBA6��Ή��̕�������菜�������Ȃ̂Ő����������ɂȂ�Ƃ͌���Ȃ��B
Private Function ReplaceDeclareTo6(ByVal base_str, Optional indent_level As Long = 0) As String
    base_str = Replace(base_str, "PtrSafe ", "")
    base_str = Replace(base_str, "LongPtr", "Long")
    ReplaceDeclareTo6 = InsertDeclareIndent(base_str, indent_level)
End Function

Rem �K����Declare������VBA7�Ή�(32/64bit���Ή�)�R�[�h�ɒu��
Rem  �uWin32API_PtrSafe.txt�v���Q�Ƃ��邽�ߐ��x�͍��������̂܂ܓ�������Ƃ͌���Ȃ��B
Rem  ���X�̖��O�t�������͕ێ�����Ȃ�
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
    Debug.Print "  ��"
    Debug.Print ReplaceDeclareTo7(Teststr, 10)
End Sub

Rem Win32API�֐�����n������S�Ή���Declare����ԋp����֐�
Rem
Rem  @name GetDeclareCodeByProcName
Rem
Rem  @param procName       �����Ώۂ̊֐���
Rem  @param useVBA6        VBA6�Ή��R�[�h�𐶐����邩
Rem  @param useVBA7        VBA7�Ή��R�[�h�𐶐����邩(32bit/64bit)
Rem  @param indent_level   �C���f���g���i2~�j
Rem
Rem  @return As String Declare���@����������procName
Rem
Rem  @note VBA6,7�������g�p����ꍇ�����f�B���N�e�B�u�ɂ�镪�򂪐��������
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
        Err.Raise 9999, , "VBA6 VBA7������Ή��ɂȂ��Ă���"
    End If
    GetDeclareCodeByProcName = Join(arr, vbCrLf)
End Function

Rem Win32API�֐�����n������VBA6(�`Excel2007)�Ή���Declare����Ԃ��֐�
Rem  32bit�ł̋L�@�����ς��邱�ƂŐ���
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
    Debug.Print "  ��"
    Debug.Print GetDeclareVBA7(TESTDATA, "", 2)
End Sub

Rem Win32API�֐�����n������VBA7(Excel 2010�`2016 32/64)�Ή���Declare����Ԃ��֐�
Rem
Rem  @param procName       �֐���
Rem  @param lifeName       ���J�͈́i�󗓁APrivate �APublic �j
Rem  @param indent_level   �C���f���g���i0~�j
Rem
Rem  @return As String     VBA7�p��Declare�錾��
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
        'GetWindowLong����64bit�ł̊֐��������BGetWindowLongPtr�ւ̒u���������K�v�B
        arr = Array("#If Win64 Then", _
                    "#Else", _
                    InsertDeclareIndent(kccFuncString.InsertIndent(lifeName & DicDeclareCode(32)(pn)), indent_level), _
                    "#End If")
        GetDeclareVBA7 = Join(arr, vbCrLf)
    Else
        GetDeclareVBA7 = ProcName
    End If
End Function

Rem �錾������֐������擾
Rem
Rem  @param base_str   ���͕�����i�錾���j
Rem  @return As String �֐���
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

Rem �錾���̃p�����[�^�̎������s�ƃC���f���g
Rem
Rem  @param base_str       �ϊ���������(Sub,Function,Property,Declare)
Rem  @param indent_level   �擪�s�ȊO�C���f���g���镝(4*(2~#))
Rem                        -1�̎��A�������s���C���f���g���s��Ȃ�
Rem  @param delimiter      ���s������i����FCR+LF�j
Rem
Rem  @return As String     ���`��̕�����
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
