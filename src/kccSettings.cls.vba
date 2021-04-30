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

Private ProjectPath As String
Private dic As Dictionary
Private dt As Date

Public Function Init(sProjectPath) As kccSettings
    Set Init = Me
    ProjectPath = sProjectPath
End Function

Public Property Get fn()
    fn = ProjectPath & "\" & SETTINGS_FILE_NAME
End Property

Private Function CheckInit() As Boolean
    If ProjectPath = "" Then Err.Raise 9999
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
Public Function CreateDefaultSetting() As kccSettings
    Set dic = New Dictionary
    dt = 0
    
    dic("ExportBinFolder") = ".\..\bin"
    dic("ExportSrcFolder") = ".\..\src"
'    dic("ExportSrcFolder") = ".\..\src\[FILENAME]"
    dic("BackupBinFile") = ".\..\backup\bin\[YYYYMMDD]_[HHMMSS]_[FILENAME]"
    dic("BackupSrcFile") = ".\..\backup\src\[YYYYMMDD]_[HHMMSS]_[FILENAME]"
'    dic("BackupSrcFile") = ".\..\backup\src\[FILENAME]\[YYYYMMDD]_[HHMMSS]_[FILENAME]"
End Function

Public Property Get ExportBinFolder()
    Call LoadFile
    ExportBinFolder = dic("ExportBinFolder")
End Property

Public Property Get ExportSrcFolder()
    Call LoadFile
    ExportSrcFolder = dic("ExportSrcFolder")
End Property

Public Property Get BackupBinFile()
    Call LoadFile
    BackupBinFile = dic("BackupBinFile")
End Property

Public Property Get BackupSrcFile()
    Call LoadFile
    BackupSrcFile = dic("BackupSrcFile")
End Property
