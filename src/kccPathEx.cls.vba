VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kccPathEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccPathEx
Rem
Rem  @description   パス情報管理クラス
Rem
Rem  @update        2020/08/06
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem  @references
Rem    Microsoft Scripting Runtime
Rem
Rem --------------------------------------------------------------------------------
Rem  @refModules
Rem    kccFuncString
Rem    kccFuncFileFolderPath
Rem
Rem --------------------------------------------------------------------------------
Rem  @history
Rem
Rem --------------------------------------------------------------------------------
Rem  @functions
Rem
Rem --------------------------------------------------------------------------------
Rem  @note
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Public FullPath__ As String
Public IsFile     As Boolean

Public Property Get fso() As FileSystemObject
    Static xxFso As Object  'FileSystemObject
    If xxFso Is Nothing Then Set xxFso = CreateObject("Scripting.FileSystemObject")
    Set fso = xxFso
End Property

Rem オブジェクトの作成
Public Function Init(obj, Optional is_file As Boolean = True) As kccPathEx
    Set Init = Me
    Select Case TypeName(obj)
        Case "String":    IsFile = is_file: FullPath = obj
        Case "File":      IsFile = True:    FullPath = ToFile(obj).Path
        Case "Folder":    IsFile = False:   FullPath = ToFolder(obj).Path
        Case "Range":     IsFile = True:    FullPath = ToRange(obj).Worksheet.Parent.FullName
        Case "Worksheet": IsFile = True:    FullPath = ToWorksheet(obj).Parent.FullName
        Case "Workbook":  IsFile = True:    FullPath = ToWorkbook(obj).FullName
        Case "Window":    IsFile = True:    FullPath = ToWindow(obj).Parent.FullName
        Case "VBProject": IsFile = True:    FullPath = VBEProjectFileName(ToVBProject(obj))
        Case Else
            On Error Resume Next
            FullPath = obj.CreateClass.FullPath: If FullPath <> "" Then IsFile = True: Exit Function
            FullPath = obj.Path: If FullPath <> "" Then IsFile = False: Exit Function
            On Error GoTo 0
            Debug.Print TypeName(obj)
            Stop
    End Select
End Function

Property Get Self() As kccPathEx: Set Self = Me: End Property

Public Function Clone() As kccPathEx
    Set Clone = VBA.CVar(New kccPathEx).Init(Me.FullPath, Me.IsFile)
End Function

Rem VBProjectから名前を取得する関数
Rem
Rem  未保存のブックではVBProject.FileNameがエラーになる。
Rem  VBProjectから直接名前を取得する手段は他に存在しない。
Rem  未保存のブックでWorkbook.FullPathなどは[Book1]と言った単純な名前しか返さない。
Rem
Private Property Get VBEProjectFileName(prj As VBProject) As String
On Error Resume Next
    VBEProjectFileName = prj.FileName
On Error GoTo 0
    If VBEProjectFileName <> "" Then Exit Property
    
    Dim wb As Excel.Workbook
    For Each wb In Workbooks
        If prj Is wb.VBProject Then
            VBEProjectFileName = wb.FullName
            Exit For
        End If
    Next
End Property

Rem ブック名からWorkbookを返す。
Rem
Rem  もしかしたらこの方法では取得できない事例があるかもしれない。
Rem
Public Function GetWorkbook(book_str_name) As Excel.Workbook
    Set GetWorkbook = Workbooks(book_str_name)
'    Dim wb As Workbook
'    For Each wb In Workbooks
'        If wb.Name = book_str_name Then
'            Set GetWorkbook = wb
'            Exit Function
'        End If
'    Next
End Function

Rem フルパス名
Property Get FullPath() As String: FullPath = FullPath__: End Function
Property Let FullPath(Path As String)
    If Path Like "*\" Then Me.IsFile = False
    'フルパス、UNC、相対、カレントを自動認識してフルパス化
    FullPath__ = kccFuncString.ToPathLastYen(Path, False)
End Property

Rem ファイルまたはフォルダ名
Property Get Name() As String
    Name = kccFuncString.GetPath(FullPath, False, True, True)
End Property

Rem ファイル名
Rem  フォルダのとき空欄
Property Get FileName() As String
    Dim IsFolder As Boolean
    FileName = kccFuncString.GetPath(FullPath, False, True, True, outIsFolder:=IsFolder)
    If IsFolder Then FileName = ""
End Property

Rem 拡張子を除く名前
Property Get BaseName() As String
    BaseName = kccFuncString.GetPath(FullPath, False, True, False)
End Property

Rem 拡張子の名前（.ext）
Property Get Extension() As String
    Extension = kccFuncString.GetPath(FullPath, False, False, True)
End Property

Rem フォルダ名
Rem  ファイルのとき空欄
Property Get FolderName() As String
    Dim IsFolder As Boolean
    FolderName = kccFuncString.GetPath(FullPath, False, True, True, outIsFolder:=IsFolder)
    If IsFolder Then Else FolderName = ""
End Property

Rem 現フォルダフルパス
Property Get CurrentFolderPath(Optional AddYen As Boolean = False) As String
    If Me.IsFile Then
        CurrentFolderPath = kccFuncString.GetPath(Me.FullPath, True, False, False)
    Else
        CurrentFolderPath = Me.FullPath
    End If
    CurrentFolderPath = kccFuncString.ToPathLastYen(CurrentFolderPath, AddYen)
End Property

Rem 親フォルダ名
Property Get ParentFolderPath(Optional AddYen As Boolean = False) As String
    If Me.IsFile Then
        ParentFolderPath = kccFuncString.GetPath(Me.CurrentFolderPath(AddYen:=False), True, False, False)
    Else
        ParentFolderPath = kccFuncString.GetPath(Me.FullPath, True, False, False)
    End If
    ParentFolderPath = kccFuncString.ToPathLastYen(ParentFolderPath, AddYen)
End Property

Rem 親フォルダオブジェクト
Property Get CurrentFolder() As kccPathEx
    Set CurrentFolder = VBA.CVar(New kccPathEx).Init(Me.CurrentFolderPath, False)
End Property

Rem 親フォルダオブジェクト
Property Get ParentFolder() As kccPathEx
    Set ParentFolder = VBA.CVar(New kccPathEx).Init(Me.ParentFolderPath, False)
End Property

Rem 存在しないとエラーになるかも
Rem FSOファイルオブジェクト
Public Function File() As Scripting.File: Set File = fso.GetFile(FullPath): End Function
Rem FSOフォルダオブジェクト
Public Function Folder() As Scripting.Folder: Set Folder = fso.GetFolder(Me.CurrentFolderPath): End Function

Rem VBプロジェクト
Public Function VBProject() As VBIDE.VBProject
On Error Resume Next
'    Dim VBP As VBProject
'    For Each VBP In Application.VBE.VBProjects
'        Dim prjName As String
'        prjName = VBP.FileName
'        If Err.Number = 0 Then
'            If prjName = Me.FullPath Then
'                Set VBProject = VBP
'            End If
'        End If
'    Next
    Dim wb As Workbook
    Set wb = Me.Workbook
    If wb Is Nothing Then Stop
    Set VBProject = wb.VBProject
End Function

Rem Excelワークブック
Public Function Workbook() As Excel.Workbook
    '[Workbooks("Book1.xlsx")]
    '[Workbooks("Book1")]
    Set Workbook = GetWorkbook(Me.FileName)
End Function

Rem 相対パスにより移動したフォルダのパス
Public Function MoveFolderPath(relative_path) As String
    MoveFolderPath = kccFuncString.AbsolutePathNameEx(Me.FullPath, relative_path)
End Function

Rem 相対パスにより移動したフォルダのインスタンスを新規生成
Public Function MoveFolder(relative_path) As kccPathEx
    Dim bas As String: bas = Me.CurrentFolderPath
    Dim ref As String: ref = relative_path
    Dim ppp As String: ppp = kccFuncString.AbsolutePathNameEx(bas, ref)
    Set MoveFolder = VBA.CVar(New kccPathEx).Init(ppp, False)
End Function

Rem 相対パスにより移動したファイルのインスタンスを新規生成
Rem   既存がフォルダのとき：「現パス\ファイル名」
Rem   既存がファイルのとき：「カレントフォルダ\ファイル名」
Public Function MoveFile(FileName) As kccPathEx
    Dim bas As String: bas = Me.CurrentFolderPath
    Dim ref As String: ref = IIf(FileName Like "*\*", "", ".\") & FileName
    Dim ppp As String: ppp = kccFuncString.AbsolutePathNameEx(bas, ref)
    Set MoveFile = VBA.CVar(New kccPathEx).Init(ppp, True)
End Function

Rem フォルダを一気に作成
Rem  成功した場合
Rem  成功:既に存在した場合
Rem  失敗:ファイルが既に存在した場合
Rem  失敗:それ以外の理由
Public Function CreateFolder() As kccPathEx
    Set CreateFolder = Me
    If Not kccFuncFileFolderPath.CreateDirectoryEx(Me.CurrentFolderPath) Then
        Debug.Print "CreateFolder 失敗 : " & Me.CurrentFolderPath
    End If
End Function

Public Function DeleteFolder()
'    On Error Resume Next
    If fso.FolderExists(Me.CurrentFolderPath) Then
        fso.DeleteFolder Me.CurrentFolderPath
    End If
End Function

Rem フォルダのファイルをまとめてコピーする
Rem 速度は無視。
Public Function CopyFiles(dest As kccPathEx, _
        Optional withFilterString As String = "*", _
        Optional withoutFilterString As String = "")
    Dim f As File
    For Each f In Me.Folder.Files
        If f.Name Like withFilterString And _
            Not f.Name Like withoutFilterString Then
            If dest.IsFile Then
                f.Copy dest.ReplacePathAuto(FileName:=f.Name).CreateFolder.FullPath
            Else
                f.Copy dest.ReplacePathAuto(FileName:=f.Name).MoveFile(f.Name).CreateFolder.FullPath
            End If
        End If
    Next
End Function

Rem フォルダが存在するか否か
Public Function FolderExists() As Boolean
    FolderExists = fso.FolderExists(Me.FullPath)
End Function

Rem パス文字列を単純に置換
Public Function ReplacePath(src, dest) As kccPathEx
    Set ReplacePath = Me.Clone
    ReplacePath.FullPath = Replace(ReplacePath.FullPath, src, dest)
End Function

Rem パス文字列をマジックナンバーにより置換
Public Function ReplacePathAuto(Optional DateTime, Optional FileName) As kccPathEx
    Dim obj As kccPathEx: Set obj = Me.Clone
    If VBA.IsMissing(DateTime) Then
    Else
        Set obj = obj.ReplacePath("[YYYYMMDD]_[HHMMSS]", Format$(DateTime, "yyyymmdd_hhmmss"))
        Set obj = obj.ReplacePath("[YYYYMMDD]", Format$(DateTime, "yyyymmdd"))
        Set obj = obj.ReplacePath("[HHMMSS]", Format$(DateTime, "hhmmss"))
    End If
    If VBA.IsMissing(FileName) Then Else Set obj = obj.ReplacePath("[FILENAME]", FileName)
    Set ReplacePathAuto = obj
End Function
