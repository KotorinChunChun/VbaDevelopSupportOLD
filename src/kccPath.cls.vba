VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kccPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccPath
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
Rem    kccFuncPath
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
Public Function Init(obj, Optional is_file As Boolean = True) As kccPath
    If Me Is kccPath Then
        With New kccPath
            Set Init = .Init(obj, is_file)
        End With
        Exit Function
    End If
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

Property Get Self() As kccPath: Set Self = Me: End Property

Public Function Clone() As kccPath
    Set Clone = kccPath.Init(Me.FullPath, Me.IsFile)
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
Property Get FullPath() As String: FullPath = Me.FullPath__ & IIf(Me.IsFile, "", "\"): End Function
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

Rem 現在のフォルダ名の変更
Property Let CurrentFolderName(FolderName As String)
    Dim cur As Scripting.Folder
    Set cur = Me.CurrentFolder.Folder
    cur.Name = FolderName
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
Property Get CurrentFolder() As kccPath
    Set CurrentFolder = kccPath.Init(Me.CurrentFolderPath, False)
End Property

Rem 親フォルダオブジェクト
Property Get ParentFolder() As kccPath
    Set ParentFolder = kccPath.Init(Me.ParentFolderPath, False)
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
Public Function MovePathByFolder(relative_path, Optional KeepFileName As Boolean = False) As kccPath
    Dim bas As String: bas = Me.CurrentFolderPath
    Dim ref As String: ref = relative_path
    Dim ppp As String: ppp = kccFuncString.AbsolutePathNameEx(bas, ref)
    Set MovePathByFolder = kccPath.Init(ppp, False)
End Function

Rem 相対パスにより移動したファイルのインスタンスを新規生成
Rem   既存がフォルダのとき：「現パス\ファイル名」
Rem   既存がファイルのとき：「カレントフォルダ\ファイル名」
Public Function MovePathByFile(relative_path) As kccPath
    Dim bas As String: bas = Me.CurrentFolderPath
    Dim ref As String: ref = IIf(relative_path Like "*\*", "", ".\") & relative_path
    Dim ppp As String: ppp = kccFuncString.AbsolutePathNameEx(bas, ref)
    Set MovePathByFile = kccPath.Init(ppp, True)
End Function

Rem フォルダを一気に作成
Rem  成功した場合
Rem  成功:既に存在した場合
Rem  失敗:ファイルが既に存在した場合
Rem  失敗:それ以外の理由
Public Function CreateFolder() As kccPath
    Set CreateFolder = Me
    If Not kccFuncPath.CreateDirectoryEx(Me.CurrentFolderPath) Then
        Debug.Print "CreateFolder 失敗 : " & Me.CurrentFolderPath
    End If
End Function

Rem フォルダを削除
Rem
Rem  @return As Boolean 削除結果
Rem                         True  : 成功(削除に成功 or 元々フォルダが無い)
Rem                         False : 失敗(フォルダが残っている)
Rem
Public Function DeleteFolder() As Boolean
    Dim cPath As String: cPath = Me.CurrentFolderPath
    If fso.FolderExists(cPath) Then
        '1秒空けて3回リトライ
        Dim n As Long: n = 3
        Do
            On Error Resume Next
            fso.DeleteFolder cPath
            If Err.number = 0 Then Exit Do
            On Error GoTo 0
            Application.Wait [Now() + "00:00:01"]
            n = n - 1
            If n = 0 Then Exit Do 'Err.Raise 9999, "DeleteFolder", "削除できません"
        Loop
        DoEvents
    End If
    DeleteFolder = Not fso.FolderExists(cPath)
End Function

Rem フォルダのファイルをまとめてコピーする
Rem 速度は無視。
Public Function CopyFiles(dest As kccPath, _
        Optional withFilterString As String = "*", _
        Optional withoutFilterString As String = "")
    Dim f As File
    If Me.CurrentFolderPath = "" Then Stop
    For Each f In Me.Folder.Files
        If f.Name Like withFilterString And _
            Not f.Name Like withoutFilterString Then
            'このCopyでは失敗してもエラーが起こらないらしい
            If dest.IsFile Then
                f.Copy dest.ReplacePathAuto(FileName:=f.Name).CreateFolder.FullPath
            Else
                f.Copy dest.ReplacePathAuto(FileName:=f.Name).MovePathByFile(f.Name).CreateFolder.FullPath, True
            End If
        End If
    Next
End Function

Rem フォルダが存在するか否か
Public Function FolderExists() As Boolean
    FolderExists = fso.FolderExists(Me.FullPath)
End Function

Rem パス文字列を単純に置換
Public Function ReplacePath(src, dest) As kccPath
    Set ReplacePath = Me.Clone
    ReplacePath.FullPath = Replace(ReplacePath.FullPath, src, dest)
End Function

Rem パス文字列をマジックナンバーにより置換
Public Function ReplacePathAuto(Optional DateTime, Optional FileName) As kccPath
    Dim obj As kccPath: Set obj = Me.Clone
    If VBA.IsMissing(DateTime) Then
    Else
        Set obj = obj.ReplacePath("[YYYYMMDD]_[HHMMSS]", Format$(DateTime, "yyyymmdd_hhmmss"))
        Set obj = obj.ReplacePath("[YYYYMMDD]", Format$(DateTime, "yyyymmdd"))
        Set obj = obj.ReplacePath("[HHMMSS]", Format$(DateTime, "hhmmss"))
    End If
    If VBA.IsMissing(FileName) Then Else Set obj = obj.ReplacePath("[FILENAME]", FileName)
    Set ReplacePathAuto = obj
End Function

' ファイルの文字コードをSJISからUTF8(BOM無し)に変換する
Public Sub ConvertCharCode_SJIS_to_utf8()
    If Me.IsFile Then Else Exit Sub
    If fso.FileExists(Me.FullPath) Then Else Exit Sub
    
    Dim fn As String: fn = Me.FullPath
    Dim destWithBOM As Object: Set destWithBOM = CreateObject("ADODB.Stream")
    With destWithBOM
        .Type = 2
        .Charset = "utf-8"
        .Open
        
        ' ファイルをSJIS で開いて、dest へ 出力
        With CreateObject("ADODB.Stream")
            .Type = 2
            .Charset = "shift-jis"
            .Open
            .LoadFromFile fn
            .Position = 0
            .copyTo destWithBOM
            .Close
        End With
        
        ' BOM消去
        ' 3バイト無視してからバイナリとして出力
        .Position = 0
        .Type = 1 ' adTypeBinary
        .Position = 3
        
        Dim dest: Set dest = CreateObject("ADODB.Stream")
        With dest
            .Type = 1 ' adTypeBinary
            .Open
            destWithBOM.copyTo dest
            .savetofile fn, 2
            .Close
        End With
        
        .Close
    End With
End Sub

