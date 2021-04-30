Attribute VB_Name = "kccWsFuncRegExp"
Option Explicit

Rem �}�b�`���邩
Rem
Rem  @param strSource       �����Ώە�����
Rem  @param strPattern      �����p�^�[��
Rem
Rem  @return As Boolean     True:�}�b�`�����BFalse:�}�b�`���Ȃ�����
Rem
Function RegexIsMatch(strSource As String, strPattern As String) As Boolean
    Dim re As RegExp
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = strPattern       ''�����p�^�[����ݒ�
        .IgnoreCase = True          ''�啶���Ə���������ʂ��Ȃ�
        .Global = True              ''������S�̂�����
        RegexIsMatch = re.Test(strSource)
    End With
End Function

Sub Test_RegexIsMatch()
    Const src = "abc def ghi jkl abc ghi"
    Debug.Print RegexIsMatch(src, "abc")
    Debug.Print RegexIsMatch(src, "dgh")
End Sub

Rem �}�b�`�����������u��
Rem
Rem  @param strSource       �����Ώە�����
Rem  @param strPattern      �����p�^�[��
Rem  @param strReplace      �u��������
Rem
Rem  @return As String      �u����̕�����
Rem
Function RegexReplace(strSource As String, strPattern As String, strReplace As String) As String
    Dim re As RegExp
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = strPattern       ''�����p�^�[����ݒ�
        .IgnoreCase = True          ''�啶���Ə���������ʂ��Ȃ�
        .Global = True              ''������S�̂�����
        RegexReplace = re.Replace(strSource, strReplace)
    End With
End Function

Sub Test_RegexReplace()
    Const src = "abc def ghi jkl abc ghi"
    Debug.Print RegexReplace(src, "abc", "XXX")
    Debug.Print RegexReplace(src, "xyz", "XXX")
End Sub

Rem �}�b�`�����ӏ���z��ŕԂ�
Rem
Rem  @param strSource       �����Ώە�����
Rem  @param strPattern      �����p�^�[��
Rem  @param strProperty     �擾�������v���p�e�B
Rem
Rem  @return As VBScript_RegExp_55.MatchCollection
Rem                         �v���p�e�B���w��ł�mc�R���N�V���������̂܂ܕԂ�
Rem
Function RegexMatches(strSource As String, strPattern As String, strProperty As String) As Variant
    Dim re As RegExp
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = strPattern       ''�����p�^�[����ݒ�
        .IgnoreCase = True          ''�啶���Ə���������ʂ��Ȃ�
        .Global = True              ''������S�̂�����
        
        Dim mc As VBScript_RegExp_55.MatchCollection
        Set mc = re.Execute(strSource)
        If strProperty = "" Then Set RegexMatches = mc: Exit Function
        If strProperty = "Count" Then RegexMatches = mc.Count: Exit Function
        If mc.Count = 0 Then: RegexMatches = Array(): Exit Function
        
        Dim arr()
        ReDim arr(0 To mc.Count - 1)
        Dim i As Long
        For i = 0 To mc.Count - 1
            If strProperty = "SubMatches" Then
                Dim sm As VBScript_RegExp_55.SubMatches
                Set sm = mc.Item(i).SubMatches
                Dim subarr()
                ReDim subarr(0 To sm.Count - 1)
                Dim j As Long
                For j = 0 To sm.Count - 1
                    subarr(j) = sm.Item(j)
                Next
                arr(i) = subarr
            Else
                arr(i) = CallByName(mc.Item(i), strProperty, VbGet)
            End If
        Next
        RegexMatches = arr
    End With
End Function

Rem �}�b�`�����ӏ��̌�
Function RegexMatchCount(strSource As String, strPattern As String)
    RegexMatchCount = RegexMatches(strSource, strPattern, "Count")
End Function

Rem �}�b�`�����ӏ��̊J�n�C���f�b�N�X�z��
Function RegexMatchIndexs(strSource As String, strPattern As String)
    RegexMatchIndexs = RegexMatches(strSource, strPattern, "FirstIndex")
End Function

Rem �}�b�`�����ӏ��̕����񒷔z��
Function RegexMatchLengths(strSource As String, strPattern As String)
    RegexMatchLengths = RegexMatches(strSource, strPattern, "Length")
End Function

Rem �}�b�`�����ӏ��̒l�z��
Function RegexMatchValues(strSource As String, strPattern As String)
    RegexMatchValues = RegexMatches(strSource, strPattern, "Value")
End Function

Sub Test_RegexMatches()
    Const src = "aabbcc axxyyzzc ghi jkl abbaac ghi"
    Const ptn = "a.+?c" '�ua�v�Ŏn�܂�uc�v�ŏI��镶����i�ŒZ�j�Ɉ�v
    Debug.Print RegexMatchCount(src, ptn)
    Debug.Print Join(RegexMatchIndexs(src, ptn), ",")
    Debug.Print Join(RegexMatchLengths(src, ptn), ",")
    Debug.Print Join(RegexMatchValues(src, ptn), ",")
End Sub

Rem �}�b�`�����ӏ��̔z��̃T�u�}�b�`�z��
Function RegexSubMatches(strSource As String, strPattern As String)
    RegexSubMatches = RegexMatches(strSource, strPattern, "SubMatches")
End Function

Sub Test_RegexSubMatches()
    Const src = "AAAAA BB001 AA202 jk345 abcde i030k X12345"
    Const ptn = "([A-Z]+)([0-9]+)" '�u�A���t�@�x�b�g�啶���̃O���[�v�v�u���l�̃O���[�v�v�Ɉ�v
    
    Dim jagArr
    jagArr = RegexSubMatches(src, ptn)
    Stop
End Sub
