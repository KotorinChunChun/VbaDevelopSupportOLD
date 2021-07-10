VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kccSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SETTINGS_FILE_NAME = "kccsettings.json"

Private fso As New FileSystemObject

Rem
Private TargetFolderPath As String

Public IsBackupProject As Boolean
Public ExportSrcFolder As String
Public ExportBinFolder As String
Public ExportBackupSrcFolders As Collection
Public ExportBackupBinFolders As Collection

Public IgnoreEmptyModule As Boolean
Public HasExtension As Boolean
Public StrExtension As String
Public IsExportCustomUI As Boolean

Public Function Init(sProjectPath) As kccSettings
    Set Init = Me
    TargetFolderPath = fso.GetParentFolderName(sProjectPath) & "\"
End Function

Rem �ݒ�t�@�C���̑��݂����ԋ߂���ʂ̊K�w�����߂�
Public Property Get Path()
    Dim fd As String: fd = TargetFolderPath
    Do
        If fso.FileExists(fd & SETTINGS_FILE_NAME) Then
            Path = fd & SETTINGS_FILE_NAME
            Exit Do
        End If
        fd = fso.GetParentFolderName(fd) & "\"
        If fd = "" Or fd = "\" Then Path = "": Exit Do
    Loop
End Property

Rem �v���W�F�N�g�t�H���_�p�X
Public Property Get ProjectFolder() As String
    If Path <> "" Then
        Call LoadFile
        ProjectFolder = fso.GetFile(Path).ParentFolder.Path
    Else
        ProjectFolder = TargetFolderPath
    End If
End Property

Rem ������
Public Sub ClearAllField()
    IsBackupProject = False
    ExportBinFolder = ""
    ExportSrcFolder = ""
    ExportSrcFolder = ""
    Set ExportBackupSrcFolders = New Collection
    IgnoreEmptyModule = False
    HasExtension = False
    StrExtension = ""
    IsExportCustomUI = False
End Sub

Rem �ݒ�t�@�C�������݂��Ȃ��ꍇ�̊���l
Public Function CreateDefaultSetting(Optional AddFileName = False) As kccSettings
    Dim dic As Dictionary
    Set dic = New Dictionary
    
    Call SetDefaultSetting(dic)
    
    Set CreateDefaultSetting = Me
End Function

Rem Dictionary������`�̂Ƃ������l��ݒ�
Private Sub SetDefaultSetting(dic As Dictionary, Optional AddFileName = False)
    If Not dic.Exists("ExportBinFolder") Then dic("ExportBinFolder") = ".\..\bin"
    If Not dic.Exists("ExportSrcFolder") Then dic("ExportSrcFolder") = IIf(AddFileName, ".\..\src\[FILENAME]", ".\..\src")

    If Not dic.Exists("ExportBackupBinFolders") Then Set dic("ExportBackupBinFolders") = New Collection
    If Not dic.Exists("ExportBackupSrcFolders") Then Set dic("ExportBackupSrcFolders") = New Collection
    
    If Not dic.Exists("IgnoreEmptyModule") Then dic("IgnoreEmptyModule") = True
    If Not dic.Exists("HasExtension") Then dic("HasExtension") = True
    If Not dic.Exists("StrExtension") Then dic("StrExtension") = ".vba"
    If Not dic.Exists("IsExportCustomUI") Then dic("IsExportCustomUI") = False
End Sub

Rem Json�ݒ�l��Dictionary�Ŏ擾
Public Function GetDictionaryBySettingFile() As Dictionary
    Dim txt As String
    txt = kccPath.ReadUTF8Text(Path)
    txt = kccWsFuncRegExp.RegexReplace(txt, "[ ]*//.*\r\n", "")
    If txt = "" Then
        Set GetDictionaryBySettingFile = New Dictionary
    Else
        Set GetDictionaryBySettingFile = JsonConverter.ParseJson(txt)
    End If
    Call SetDefaultSetting(GetDictionaryBySettingFile)
End Function

Rem �ݒ�l��ǂݍ���
Rem  @return    �Ǎ�������������(True:����or�Ǎ��� / False:���s)
Public Function LoadFile() As Boolean
    Static dt As Date
    If Not fso.FileExists(Path) Then dt = 0: Call ClearAllField: Exit Function
    
    LoadFile = True
    If dt = fso.GetFile(Path).DateLastModified Then Exit Function
    dt = fso.GetFile(Path).DateLastModified
    
    Dim dic As Dictionary
    Set dic = GetDictionaryBySettingFile()
    IsBackupProject = dic("IsBackupProject")
    ExportBinFolder = dic("ExportBinFolder")
    ExportSrcFolder = dic("ExportSrcFolder")
    Set ExportBackupBinFolders = dic("ExportBackupBinFolders")
    Set ExportBackupSrcFolders = dic("ExportBackupSrcFolders")
    IgnoreEmptyModule = dic("IgnoreEmptyModule")
    HasExtension = dic("HasExtension")
    StrExtension = dic("StrExtension")
    IsExportCustomUI = dic("IsExportCustomUI")
End Function

Rem ���݂̐ݒ�l���t�@�C���ɏ����o��
Public Sub SaveFile()
    If TargetFolderPath = "" Then Err.Raise 9999
    
    Dim dic As Dictionary
    Set dic = GetDictionaryBySettingFile()
    
    dic("ExportBinFolder") = ExportBinFolder
    dic("ExportSrcFolder") = ExportSrcFolder
    Set dic("ExportBackupBinFolders") = ExportBackupBinFolders
    Set dic("ExportBackupSrcFolders") = ExportBackupSrcFolders
    
    dic("IgnoreEmptyModule") = IgnoreEmptyModule
    dic("HasExtension") = HasExtension
    dic("StrExtension") = StrExtension
    dic("IsExportCustomUI") = IsExportCustomUI
    
    Dim txt As String
    txt = JsonConverter.ConvertToJson(dic, " ")
    Call kccPath.Init(Path, True).WriteUTF8Text(txt)
End Sub
