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

Private TargetFilePath As String
Private dic As Dictionary
Private dt As Date

Public Function Init(sProjectPath) As kccSettings
    Set Init = Me
    TargetFilePath = sProjectPath
End Function

Public Property Get fn()
    Rem 再帰的に所在を求める
    Dim fd As String: fd = TargetFilePath
    Do
        If fso.FileExists(fd & SETTINGS_FILE_NAME) Then
            fn = TargetFilePath & SETTINGS_FILE_NAME
            Exit Do
        End If
        fd = fso.GetParentFolderName(fd)
        If fd = "" Then fn = "": Exit Do
    Loop
End Property

Private Function CheckInit() As Boolean
    If TargetFilePath = "" Then Err.Raise 9999
End Function

Public Sub LoadFile()
    Call CheckInit
    
    If Not fso.FileExists(fn) Then Exit Sub
    If dt = fso.GetFile(fn).DateLastModified Then Exit Sub
    dt = fso.GetFile(fn).DateLastModified
    
    Dim txt As String
    txt = kccPath.ReadUTF8Text(fn)
    txt = kccWsFuncRegExp.RegexReplace(txt, "[ ]*//.*\r\n", "")
    Set dic = JsonConverter.ParseJson(txt)
End Sub

Public Sub SaveFile()
    Call CheckInit
    Dim txt As String
    txt = JsonConverter.ConvertToJson(dic)
'    Call kccPath.WriteUTF8Txt(fn, txt)
End Sub

Rem 設定ファイルが存在しない場合の既定値
Public Function CreateDefaultSetting(Optional AddFileName = False) As kccSettings
    Set dic = New Dictionary
    dt = 0
    
    dic("ExportBinFolder") = ".\..\bin"
    dic("ExportSrcFolder") = IIf(AddFileName, _
                                ".\..\src\[FILENAME]", _
                                ".\..\src")
    dic("BackupBinFile") = ".\..\backup\bin\[YYYYMMDD]_[HHMMSS]_[FILENAME]"
    dic("BackupSrcFile") = IIf(AddFileName, _
                                ".\..\backup\src\[FILENAME]\[YYYYMMDD]_[HHMMSS]_[FILENAME]", _
                                ".\..\backup\src\[YYYYMMDD]_[HHMMSS]_[FILENAME]")
    Set CreateDefaultSetting = Me
End Function

Public Property Get ProjectFolder() As String
    If fn <> "" Then
        Call LoadFile
        ProjectFolder = fso.GetFile(fn).ParentFolder.Path
    Else
        ProjectFolder = TargetFilePath
    End If
End Property

Public Property Get ExportBinFolder() As String
    If fn <> "" Then Call LoadFile
    ExportBinFolder = dic("ExportBinFolder")
End Property

Public Property Get ExportSrcFolder() As String
    If fn <> "" Then Call LoadFile
    ExportSrcFolder = dic("ExportSrcFolder")
End Property

Public Property Get BackupBinFile() As String
    If fn <> "" Then Call LoadFile
    BackupBinFile = dic("BackupBinFile")
End Property

Public Property Get BackupSrcFile() As String
    If fn <> "" Then Call LoadFile
    BackupSrcFile = dic("BackupSrcFile")
End Property
