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
Rem  @description   ������ϊ��֐�
Rem
Rem  @update        2020/08/07
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------

Rem  @description �����鏉�ߊ��ʂ�������ʂ�Ԃ��֐�
Rem
Rem  @param open_brackets       ���ߊ��ʁi�@��ˑ������Ή��j
Rem
Rem  @return As String          ������
Rem
Function OpenBracketsToClose(open_brackets) As String
    Dim stb As String: stb = open_brackets
    Dim etb As String: etb = ""
    Select Case stb
        Case "[", "{", "<", "�m", "�o", "��"
            etb = ChrW(AscW(stb) + 2)
        Case ChrW(171)
            etb = ChrW(AscW(stb) + 16)
        Case Else
            etb = ChrW(AscW(stb) + 1)
    End Select
    OpenBracketsToClose = etb
End Function

Rem ������Ɋ܂܂�銇�ʂ��l�X�g�ɉ����ĕω�������֐�
Rem
Rem  @param base_str            ���͕�����
Rem  @param open_Bracket        �u���Ώۂ̏��ߊ��� (����l:�ۊ���)
Rem  @param replaced_brackets   �u����̏��ߊ��ʂ̔z�� (����l:[{(<��4�i�K)
Rem
Rem  @return As String          ���ʂ�u���ς݂̕�����
Rem
Rem  @note
Rem      ���ʂ̃l�X�g�͕�����̐擪���珇���ϊ����郍�W�b�N
Rem      ���߁`�����s���S�ł���؊֒m���Ȃ��̂Œ��ӂ��邱��
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

Rem ��؂蕶����̂����������Ɉ͂�ꂽ�͈͂����̕������ʂ�Ԃ�
Rem
Rem  @param base_str        ���͕�����
Rem  @param start_brackets  �J�n�������̎�ށi�I���J�b�R�͎������f�j
Rem  @param remove_brackets �J�b�R��...True:�폜����(����) False:�c��
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
Rem     ����q�ɂ͔�Ή�
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

    ' �������������ʓ��ȊO�𒊏o
    'reg.Pattern = "^(.*?)\(|\)(.*?)\(|\)(.*?).*$"
    reg.Pattern = "^(.*?)" & openDelim & "|" & closeDelim & "(.*?)" & openDelim & "|" & closeDelim & "(.*?).*$"
    'reg.Pattern = "\[[^\[\]]*(?=\])"
    ' ������̍Ō�܂Ō�������
    reg.Global = True

    ' ������v�������J���}�ɒu��������
    retVal = reg.Replace(base_str, CashDelimiter)

    If IsEmpty(retVal) Or retVal = "" Then Exit Function
    If reg.Execute(base_str).Count = 0 Then Exit Function

    ' �擪�ƍŌ�̃J���}��������������
    retVal = Mid(retVal, 2, Len(retVal) - 2)

    ' ���ʓ��̕���������ʂ̐������z��Ƃ��Ď擾
    SplitWithInBrackets = Split(retVal, CashDelimiter)

End Function

Rem ������Ɋ܂܂�镶����̏o���ʒu�S�Ă�Ԃ��֐�
Rem
Rem  @param base_str ���͕�����
Rem  @param find_str ����������
Rem
Rem  @return As Variant/Long(1 To #) ����������̐擪�C���f�b�N�X�̔z��
Rem
Rem  @example
Rem          find_str = "a"
Rem          Missing              >> Variant(0 to -1) {}
Rem          String ""            >> Variant(0 to -1) {}
Rem          String "a"           >> Long(1 to 1) {1}
Rem          String "abacda"      >> Long(1 to 3) {1,3,6}
Rem
Rem          find_str = "bc"
Rem          String "abacda"      >> Variant(0 to -1) {}
Rem          String "dsbcffdgrbc" >> Long(1 to 2) {3,10}
Rem
Rem  @note
Rem     �ő�65535���܂łȂ��Ƃɒ���
Rem
Public Function InStrAll(base_str, find_str) As Variant
    If IsMissing(base_str) Then Exit Function
    If base_str = "" Or find_str = "" Then Exit Function
    Dim n As Long: n = 0
    Dim retVal As Long: retVal = 0
    Dim retIndexs() As Long
    ReDim retIndexs(1 To 65535)
    Do
        n = InStr(n + 1, base_str, find_str)
        If n = 0 Then
            Exit Do
        Else
            retVal = retVal + 1
            If UBound(retIndexs) > retVal Then
                retIndexs(retVal) = n
            End If
        End If
    Loop
    If retVal = 0 Then
        InStrAll = VBA.Array()
    Else
        ReDim Preserve retIndexs(1 To retVal)
        InStrAll = retIndexs
    End If
End Function

Rem �����������J��Ԃ��ꂽ��������Ԃ�
Rem
Rem  @param base_str       ���͕�����
Rem  @param find_str       ����������
Rem  @param start_index    �����J�n�ʒu(1~)
Rem
Rem  @retuen As Long ����������������������(����������*��)
Rem                   �S�Ă�find_str�Ȃ�len(base_str)
Rem
Rem  @example
Rem          find_str = "a"
Rem          Missing         >> Long 0
Rem          String ""       >> Long 0
Rem
Rem          start_index = 1
Rem          String "a"      >> Long 1
Rem          String "abaa"   >> Long 1
Rem          String "xyzaaa" >> Long 0
Rem
Rem          start_index = 3
Rem          String "a"      >> Long 0
Rem          String "abaa"   >> Long 2
Rem          String "xyzaaa" >> Long 0
Rem
Rem          start_index = 4
Rem          String "a"      >> Long 0
Rem          String "abaa"   >> Long 1
Rem          String "xyzaaa" >> Long 3
Rem
Public Function InStrRept(base_str, find_str, Optional start_index = 1) As Long
    If IsMissing(base_str) Then Exit Function
    If base_str = "" Then Exit Function
    If start_index < 0 Then Err.Raise 9999, , "start_index�͕�����̊J�n�ʒu(1~)���w�肵�ĉ�����"
    Dim i As Long
    For i = start_index To Len(base_str) Step Len(find_str)
        If Mid(base_str, i, Len(find_str)) <> find_str Then Exit For
    Next
    InStrRept = i - start_index
End Function

Rem ���[�̃X�y�[�X����������Trim��z��S�̂ɓK�p����
Rem
Rem  @param As Variant/String() arr_base_str ���͕�����z��
Rem
Rem  @return As Variant/String()             �o�͕�����z��
Rem
Public Function TrimArray(ByRef arr_base_str) As Variant
    Dim i As Long
    For i = LBound(arr_base_str) + 1 To UBound(arr_base_str)
        arr_base_str(i) = Trim(arr_base_str(i))
    Next
    TrimArray = arr_base_str
End Function

Rem �ʏ�g�����ɉ����āA�����񒆂̘A���X�y�[�X���V���O���X�y�[�X�ɕϊ�����B
Rem Excel�֐���TRIM�݊�
Rem
Rem  @param base_str       ���͕�����
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

Rem ��؂蕶�����Ƃɐ擪�ɏ���̕�����ǋL����
Rem
Rem  @param base_str       �ϊ���������(Declare��)
Rem  @param delimiter      ���s������i����FCR+LF�j
Rem
Rem  @return As String     ���`��̕�����
Rem
Public Function InsertString(base_str, add_str, Optional Delimiter = vbCrLf) As String
    InsertString = add_str & Replace(base_str, Delimiter, Delimiter & add_str)
End Function
Rem   �R�����g�u'�v��}��
Public Function InsertComment(ByVal base_str, Optional Delimiter = vbCrLf) As String
    InsertComment = InsertString(base_str, "'")
End Function
Rem   �C���f���g�u    �v��}��
Rem  @param indent_level   �C���f���g���镝(4*(1~#))
Public Function InsertIndent(ByVal base_str, Optional indent_level = 1, Optional Delimiter = vbCrLf) As String
    InsertIndent = InsertString(base_str, String(4 * indent_level, " "))
End Function

Rem Right�֐��g��  �Ō�ɏo�������؂蕶�����؂�ڂƂ��ĉE���̕�����Ԃ�
Rem
Rem  @param base_str      ���o����������
Rem  @param cut_str       �ؒf������i�������猟�����ĊY�����镶����̎�O�܂ł����o���j
Rem  @param cut_inc       �ؒf��������܂߂ĕԂ����ǂ����i�ʏ�͏��O����j
Rem  @param shift_len     ���o���������]���Ɏ��o���������i�v���X�j�A��藎�Ƃ��������i�}�C�i�X�j
Rem  @param should_fill   ���݂��Ȃ��ꍇ�͓��͕�����Ŗ��߂邩�i����True�j
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



Rem Left�֐��g��  �Ō�ɏo�������؂蕶�����؂�ڂƂ��č����̕�����Ԃ�
Rem
Rem  @param base_str      ���o����������
Rem  @param cut_str       �ؒf������i�Y�����镶����̎�O�܂ł����o���j  ���݂��Ȃ��ꍇ�͓��͕�����Ŗ��߂邩
Rem  @param cut_inc       �ؒf��������܂߂ĕԂ����ǂ����i�ʏ�͏��O����j
Rem  @param shift_len     ���o���������]���Ɏ��o���������i�v���X�j�A��藎�Ƃ��������i�}�C�i�X�j
Rem  @param should_fill   ���݂��Ȃ��ꍇ�͓��͕�����Ŗ��߂邩�i����True�j
Rem
Rem  @return As String
Rem
Rem  @example
Rem
Public Function LeftStrRev(base_str, cut_str, Optional cut_inc As Boolean = False, _
                                Optional shift_len As Long = 0, Optional should_fill = True) As String
    If InStrRev(base_str, cut_str, -1) > 0 And cut_str <> "" Then
        If cut_inc Then
            LeftStrRev = Left(base_str, InStrRev(base_str, cut_str, -1) - 1 + shift_len + Len(cut_str))
        Else
            LeftStrRev = Left(base_str, InStrRev(base_str, cut_str, -1) - 1 + shift_len)
        End If
    ElseIf should_fill Then
        LeftStrRev = base_str
    Else
        LeftStrRev = ""
    End If
End Function
Rem �t�H���_�̐�΃p�X�ƃt�@�C���̑��΃p�X���������āA�ړI�̃t�@�C���̐�΃p�X���擾����֐�
Rem
Rem  @name     AbsolutePathNameEx
Rem  @oldname  BuildPathEx
Rem
Rem  @param base_path      ��p�X
Rem  @param ref_path       ��p�X����̈ړ����������΃p�X�i�܂��͏㏑�������΃p�X�j
Rem
Rem  @return   As String   �A����̐�΃p�X
Rem
Rem  @note
Rem          fso.GetAbsolutePathName(fso.BuildPath(base_path, ref_path))�̖������������֐�
Rem          * UNC��..\�������APC�����ɂ͈ړ��ł��Ȃ�
Rem          * UNC��͂����ᑬ
Rem          * �t�H���_������\������
Rem          *
Rem         ��UNC�p�X���l�b�g���[�N�R���s���[�^��̃t�@�C�����Q�Ƃ���p�X��\\����n�܂�A��
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
                    'Err.Raise 8888, "AbsolutePathNameEx", "���B�ł��Ȃ��p�X���w�肵�Ă��܂��B"
                    AbsolutePathNameEx = "���B�s�\"
                    Exit Function
                End If
                retVal = Left(retVal, InStrRev(retVal, "\") - 1)
                rpArr(i) = ""
            Case Else
                retVal = retVal & IIf(retVal = "", "", "\") & rpArr(i)
                rpArr(i) = ""
        End Select
        '���΃p�X�������󗓁A.\�A..\�ŏI��������A������\���s������̂ŕ⊮���K�v
        If i = UBound(rpArr) Then
            If ref_path <> "" Then
                If Right(ref_path, 1) = "\" Then
                    retVal = retVal & "\"
                End If
            End If
        End If
    Next
    '�A��\�̏����ƃl�b�g���[�N�p�X�΍�
    retVal = Replace(retVal, "file:\\", "file://")
    retVal = Replace(retVal, "\\", "\")
    retVal = IIf(Left(retVal, 1) = "\", "\", "") & retVal
    AbsolutePathNameEx = retVal
End Function

Rem �p�X�����񂪃��[�g�i�h���C�u or UNC�j����n�܂��Ă��邩
Function IsRootStart(ByVal p)
    p = Replace(UCase(p), "/", "\")
    IsRootStart = ((p Like "[A-Z]:") Or (p Like "[A-Z]:\*") Or (p Like "\\?*"))
End Function

Rem  �p�X������t�@�C�����������Ĥ�p�X���擾���܂���i�Ō�Ɂu\�v�͂��܂���B�R�����u:�v���Ȃ����~�L���u\�v���Ȃ��ꍇ�̓t�@�C���Ƃ��܂��j
'Function GetPathName(PathName As String) As String
'  Dim l As Long ' ������
'  Dim yen As Long ' \ �t�H���_�̋�؂�L���̈ʒu
'  Dim colon As Long ' : �h���C�u�̋L���̈ʒu
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
'    GetPathName = PathName ' �h���C�u
'  Else
'    GetPathName = vbNullString ' �~�L���u\�v���Ȃ��ꍇ�̓t�@�C���Ƃ��܂�
'  End If
'End Function

Rem �t�@�C���p�X��W�J���āA�f�B���N�g���A�t�@�C�����A�g���q�@���Ƃ肾��
Rem
Rem  @param FullPath        �t���p�X�f�[�^
Rem  @param AddPath         �߂�l�Ƀt�H���_�p�X���܂߂�
Rem  @param AddName         �߂�l�Ƀx�[�X�t�@�C�������܂߂�
Rem  @param AddExtension    �߂�l�Ɋg���q���܂߂�
Rem  @param outPath         �������Ƀt�H���_�p�X��Ԃ�(C:\hoge\)
Rem  @param outName         �������Ƀt�@�C�����܂��̓t�H���_����Ԃ�("fuga")
Rem  @param outExtension    �������Ɋg���q��Ԃ�(".ext")
Rem  @param outIsFolder     ��������outName���t�H���_�̎�True��Ԃ�
Rem
Rem  @return    As String   ���������p�X�f�[�^
Rem
Rem  @note
Rem     �߂�l��outName�ɂ�\�������̂Œ��ӂ��邱��
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
    
'    FullPath = RenewalPath(FullPath)   '���ꂷ��ƃt�@�C���t�H���_���肪�o�O��
    If FullPath = "" Then Exit Function
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    '�Ōオ\�Ȃ�t�H���_�����B
    '����Ă�fso�Ŏ������画�肷��B
    '���݂��Ȃ��t�H���_�̏ꍇ�A�g���q�̗L���Ŕ��������B
    'FullPath�̖����ɂ�\��t���Ȃ���ԂŌ�̏����Ɉ����p��
    outIsFolder = (FullPath Like "*\")
    If outIsFolder Then
        FullPath = Left$(FullPath, Len(FullPath) - 1)
    Else
        outIsFolder = fso.FolderExists(FullPath)
    End If
    
    '�p�X���ƃt�@�C�����̒��o
    Dim NameAndExt As String
    outPath = Strings.Left(FullPath, Strings.InStrRev(FullPath, "\"))
    NameAndExt = Strings.Right(FullPath, Strings.Len(FullPath) - Strings.InStrRev(FullPath, "\"))
    If outIsFolder Then outName = NameAndExt: GoTo ExitProc
    
    '�t�@�C�����Ɗg���q�̒��o
    If InStr(NameAndExt, ".") = 0 Then outName = NameAndExt: GoTo ExitProc
    outName = Strings.Left(NameAndExt, Strings.InStrRev(NameAndExt, ".") - 1)
    outExtension = Strings.Right(NameAndExt, Strings.Len(NameAndExt) - Strings.InStrRev(NameAndExt, ".") + 1)
    
ExitProc:
    GetPath = ""
    If AddPath Then GetPath = GetPath & outPath
    If AddName Then GetPath = GetPath & outName
    If AddExtension Then GetPath = GetPath & outExtension
End Function

Rem �p�X���K��̏����ɏ���������B�i�l�b�g���[�N�h���C�u�Ή��j
'Public Function RenewalPath(ByVal Path As String, Optional AddYen As Boolean = False) As String
'    '�h�b�g�̗L���Ńt�@�C�� or �t�H���_����@�s���S�B
'    If Strings.InStr(Path, ".") = 0 Then Path = Path & IIf(AddYen, "\", "")
'    RenewalPath = Strings.Left(Path, 2) & Strings.Replace(Strings.Replace(Path, "/", "\"), "\\", "\", 3)
'    RenewalPath = ToPathLastYen(RenewalPath, AddYen)
'End Function

Rem �e�f�B���N�g����Ԃ��B
Rem \�}�[�N�͕t�^���Ȃ�
Public Function ToPathParentFolder(ByVal Path As String, Optional AddYen As Boolean = False) As String
    ToPathParentFolder = ToPathLastYen(GetPath(Path, True, False, False), AddYen)
End Function

Rem �p�X�̍Ō��\��t����^����
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
