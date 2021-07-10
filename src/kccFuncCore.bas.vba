Attribute VB_Name = "kccFuncCore"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccFuncCore
Rem
Rem  @description   �K�{�֐��������W�߂����W���[��
Rem
Rem  @update        2020/09/09
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Rem �o�C�i���z�񂩂�w�肵���f�[�^�̊J�n����ʒu��Ԃ�
Rem
Rem  @param sourceBytes     �������o�C�i���f�[�^�z��
Rem  @param findData        �����Ώۃf�[�^
Rem  @param startIndex      �����݊J�n�ʒu�̃C���f�b�N�X 0~
Rem
Rem  @return As Long        �z��̂�����v�����ӏ��̐擪�v�f�ԍ�
Rem
Function IndexOfBinary(sourceBytes() As Byte, findData, Optional startIndex) As Long
    Dim findBytes() As Byte
    Select Case TypeName(findData)
        Case "String"
            'UNICODE >> JIS
            findBytes = StrConv(findData, vbFromUnicode)
        Case "Byte"
            ReDim findBytes(0 To 0)
            findBytes(0) = findData
        Case "Byte()"
            findBytes = findData
        Case Else
            Stop
    End Select
    
    Dim ret As Boolean
    Dim i As Long, j As Long
    If VBA.IsMissing(startIndex) Then startIndex = LBound(sourceBytes)
    For i = startIndex To UBound(sourceBytes)
        For j = LBound(findBytes) To UBound(findBytes)
            ret = False
            If sourceBytes(i + j) <> findBytes(j) Then Exit For
            ret = True
        Next
        If ret Then IndexOfBinary = i: Exit Function
    Next
    IndexOfBinary = -1
End Function

Rem �o�C�i���z��̎w�肵���C���f�b�N�X�ȍ~�Ƀo�C�i���z����R�s�[����
Rem
Rem  @param arrBytes()          �����ݐ�o�C�g�z�� 0~
Rem  @param writeBytes()        �����݂����~�J�Ɣz��
Rem  @param startIndex          �����݊J�n�ʒu�̃C���f�b�N�X 0~
Rem
Sub WriteBinary(ByRef arrBytes() As Byte, writeBytes() As Byte, startIndex)
    Dim i As Long
    For i = LBound(writeBytes) To UBound(writeBytes)
        arrBytes(startIndex + i) = writeBytes(i)
    Next
End Sub

Rem �o�C�g�z��f�[�^���f�o�b�O�p������ɕϊ�
Rem
Rem  @param bData               ���炩�̃f�[�^
Rem
Rem  @return As String          �ϊ���̕�����
Rem
Function ToStringByte(bData) As String
    Select Case TypeName(bData)
        Case "Byte"
            ToStringByte = Right("   " & bData, 4) & " - 0x" & Right("00" & Hex(bData), 2) & " - Char[" & Chr(bData) & "]"
        Case "Byte()"
            Dim i As Long
            Dim arr()
            ReDim arr(LBound(bData) To UBound(bData))
            For i = LBound(bData) To UBound(bData)
                arr(i) = ToStringByte(bData(i))
            Next
            ToStringByte = Join(arr, vbLf)
        Case Else
            Stop
    End Select
End Function

Rem 1�o�C�g�Ǎ� - �f�o�b�O�p������o�͔�
Function ReadByteToString(bfr As kccBinaryFileIO, Optional FileIndex) As String
    ReadByteToString = ToStringByte(bfr.ReadByte(FileIndex))
End Function

Rem �w��T�C�Y���o�C�g�z��ɓǍ� - �f�o�b�O�p������o�͔�
Function ReadBytesToString(bfr As kccBinaryFileIO, Optional FileIndex, Optional ReadSize = 1) As Byte()
    ReadBytesToString = ToStringByte(bfr.ReadBytes(FileIndex, ReadSize))
End Function

Rem ��x�ɂ��ׂēǂݍ���ŏo��
Rem
Rem  @param arrBytes()          �o�͂������o�C�g�z��
Rem  @param BreakCount          �C�~�f�B�G�C�g�E�B���h�E�ɏo�͂���f�[�^����
Rem
Sub DebugPrintByteArray(arrBytes() As Byte, Optional BreakCount)
    
    Debug.Print "----------DebugPrintByteArray----------"
    Debug.Print "No.      - 10�i - 16�i - ������"
    Dim i As Long
    If VBA.IsMissing(BreakCount) Then BreakCount = UBound(arrBytes)
    For i = 0 To BreakCount
        '// �����[�v�̔z��l���擾
        Dim bData
        bData = arrBytes(i)
        
        '// ���s�R�[�h�̏ꍇ
        If bData = 10 Or bData = 13 Then
            Debug.Print "���s�ł�"
        End If
        
        '// �o��
        Debug.Print "No." & Left(i & String(5, " "), 5) & " - " & ToStringByte(bData)
        DoEvents
    Next
    Debug.Print
End Sub

Rem MsgBox�̃��b�p�[
Function MsgBox(ByVal Prompt, _
        Optional ByVal Buttons As VbMsgBoxStyle, _
        Optional ByVal Title, _
        Optional ByVal HelpFile, _
        Optional ByVal Context)
    Buttons = Buttons Or vbMsgBoxSetForeground
    If IsMissing(Title) Then Title = APP_NAME
    VBA.MsgBox Prompt, Buttons, Title, HelpFile, Context
End Function
