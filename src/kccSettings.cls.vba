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
Private ExportBackupSrcFolders__ As Collection
Private ExportBackupBinFolders__ As Collection

Public IgnoreEmptyModule As Boolean
Public HasExtension As Boolean
Public StrExtension As String
Public IsExportCustomUI As Boolean

Public Property Let ExportBackupSrcFolders(obj)
    Set ExportBackupSrcFolders__ = ToCollection(obj)
End Property

Public Property Get ExportBackupSrcFolders() As Collection
    If ExportBackupSrcFolders__ Is Nothing Then Set ExportBackupSrcFolders__ = New Collection
    Set ExportBackupSrcFolders = ExportBackupSrcFolders__
End Property

Public Property Let ExportBackupBinFolders(obj)
    Set ExportBackupBinFolders__ = ToCollection(obj)
End Property

Public Property Get ExportBackupBinFolders() As Collection
    If ExportBackupBinFolders__ Is Nothing Then Set ExportBackupBinFolders__ = New Collection
    Set ExportBackupBinFolders = ExportBackupBinFolders__
End Property

Private Function ToCollection(obj) As Collection
    Set ToCollection = New Collection
    If TypeName(obj) = "Collection" Then
        Set ToCollection = obj
    ElseIf IsArray(obj) Then
        Dim v
        For Each v In obj
            ToCollection.Add v
        Next
    Else
        Err.Raise 9999
    End If
End Function

Public Function Init(sProjectPath) As kccSettings
    Set Init = Me
    TargetFolderPath = fso.GetParentFolderName(sProjectPath) & "\"
End Function

Rem 設定ファイルの存在する一番近い上位の階層を求める
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

Rem プロジェクトフォルダパス
Public Property Get ProjectFolder() As String
    If Path <> "" Then
        Call LoadFile
        ProjectFolder = fso.GetFile(Path).ParentFolder.Path
    Else
        ProjectFolder = TargetFolderPath
    End If
End Property

Rem 初期化
Public Sub ClearAllField()
    IsBackupProject = False
    ExportBinFolder = ""
    ExportSrcFolder = ""
    ExportSrcFolder = ""
    ExportBackupBinFolders = New Collection
    ExportBackupSrcFolders = New Collection
    IgnoreEmptyModule = False
    HasExtension = False
    StrExtension = ""
    IsExportCustomUI = False
End Sub

Rem 設定ファイルが存在しない場合の既定値
Public Function CreateDefaultSetting(Optional AddFileName = False) As kccSettings
    Dim dic As Dictionary
    Set dic = New Dictionary
    
    Call SetDefaultSetting(dic)
    
    Set CreateDefaultSetting = Me
End Function

Rem Dictionaryが未定義のとき初期値を設定
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

Rem Json設定値をDictionaryで取得
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

Rem 設定値を読み込み
Rem  @return    読込が成功したか(True:成功or読込済 / False:失敗)
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
    ExportBackupBinFolders = dic("ExportBackupBinFolders")
    ExportBackupSrcFolders = dic("ExportBackupSrcFolders")
    IgnoreEmptyModule = dic("IgnoreEmptyModule")
    HasExtension = dic("HasExtension")
    StrExtension = dic("StrExtension")
    IsExportCustomUI = dic("IsExportCustomUI")
    IsBackupProject = dic("IsBackupProject")
End Function

Rem 現在の設定値をファイルに書き出し
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
    dic("IsBackupProject") = IsBackupProject
    
    Dim txt As String
    txt = JsonConverter.ConvertToJson(dic, " ")
    Call kccPath.Init(Path, True).WriteUTF8Text(txt)
End Sub
