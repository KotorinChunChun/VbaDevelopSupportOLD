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
        Case "kccPath":   IsFile = obj.IsFile: FullPath = obj.FullPath
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
Rem  この関数を使うには[VBA プロジェクト オブジェクトモデルへのアクセス]の許可が必要
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
    On Error Resume Next
    Set GetWorkbook = Workbooks(book_str_name)
    On Error GoTo 0
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

Rem FSOファイルオブジェクト
Public Function File() As Scripting.File
    On Error Resume Next
    Set File = fso.GetFile(FullPath)
End Function

Rem FSOフォルダオブジェクト
Public Function Folder() As Scripting.Folder
    On Error Resume Next
    Set Folder = fso.GetFolder(Me.CurrentFolderPath)
End Function

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
    Select Case Me.Extension
        Case ".xls", ".xlsm", ".xla", ".xlam", ".xlsb"
            Dim wb As Workbook
            Set wb = Me.Workbook
            If wb Is Nothing Then Stop
            Set VBProject = wb.VBProject
        Case ".mdb", ".accdb"
            Stop
        Case ".doc", "docm", ".dotm"
            Stop
        Case ".ppt", ".pptm", ".ppa", ".ppam"
            Stop
        Case Else
            Stop
    End Select
End Function

Rem Excelワークブック
Public Function Workbook() As Excel.Workbook
    If Me.FileName = "" Then Exit Function
    '[Workbooks("Book1.xlsx")]
    '[Workbooks("Book1")]
    Set Workbook = GetWorkbook(Me.FileName)
End Function

Rem 相対パスにより移動したフォルダのパス
Public Function MoveFolderPath(relative_path) As String
    If Me.IsFile Then
        MoveFolderPath = kccFuncString.AbsolutePathNameEx(Me.CurrentFolder.FullPath, relative_path)
    Else
        MoveFolderPath = kccFuncString.AbsolutePathNameEx(Me.FullPath, relative_path)
    End If
End Function

Rem 相対パスにより移動したフォルダのインスタンスを新規生成
Public Function MovePathToFolderPath(relative_path) As kccPath
    Dim basePath As String: basePath = Me.CurrentFolderPath
    Dim refePath As String: refePath = relative_path
    Dim absoPath As String: absoPath = kccFuncString.AbsolutePathNameEx(basePath, refePath)
    Set MovePathToFolderPath = kccPath.Init(absoPath, False)
End Function

Rem 相対パスにより移動したファイルのインスタンスを新規生成
Rem   既存がフォルダのとき：「現パス\relative_path」
Rem   既存がファイルのとき：「カレントフォルダ\relative_path」
Rem
Rem  特別に使用できる文字 : |t |e
Rem    エスケープ文字は、ファイル名に使用できないパイプ | とする。
Rem    hoge.ext
Rem      元のファイル名       : hoge : |t : titleの略
Rem      元のファイルの拡張子 : .ext : |e : extensionの略
Public Function MovePathToFilePath(ByVal relative_path) As kccPath
    If VarType(relative_path) <> vbString Then Err.Raise 9999, , "型が違います"
    If relative_path = "" Then Set MovePathToFilePath = kccPath.Init(Me)
    relative_path = Replace(relative_path, "|t", Me.BaseName)
    relative_path = Replace(relative_path, "|e", Me.Extension)
    '自身がフォルダで移動先ファイルがファイル名のみしか指定されなかった場合、カレントを示す\を追記
    If Not Me.IsFile And Not relative_path Like "\*" Then relative_path = "\" & relative_path
    Dim basePath As String: basePath = Me.CurrentFolderPath
    Dim refePath As String: refePath = IIf(relative_path Like "*\*", "", ".\") & relative_path
    Dim absoPath As String: absoPath = kccFuncString.AbsolutePathNameEx(basePath, refePath)
    Set MovePathToFilePath = kccPath.Init(absoPath, True)
End Function

Rem 相対パスによりファイル名を維持したまま親フォルダを移動する
Public Function MoveParentFolder(ByVal relative_path) As kccPath
    Set MoveParentFolder = Me.MovePathToFilePath(relative_path & "|t|e")
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
            If Err.Number = 0 Then Exit Do
            On Error GoTo 0
            Application.Wait [Now() + "00:00:01"]
            n = n - 1
            If n = 0 Then Exit Do 'Err.Raise 9999, "DeleteFolder", "削除できません"
        Loop
        DoEvents
    End If
    DeleteFolder = Not fso.FolderExists(cPath)
End Function

Rem ファイル・フォルダが存在するか
Public Function Exists() As Boolean
    If Me.IsFile Then
        Exists = Not (Me.File Is Nothing)
    Else
        Exists = Not (Me.Folder Is Nothing)
    End If
End Function

Rem ファイルをコピーする
Rem
Rem  @param dest            コピー先ファイル名またはフォルダ
Rem  @param OverWriteFiles  コピー先ファイルが存在する時上書きするか(既定:True)
Rem
Rem  @note
Rem    dest:=ファイル指定・・・対象ファイル名で書き込み
Rem    dest:=フォルダ指定・・・対象フォルダに元と同じファイル名で書き込み
Rem
Public Function CopyFile(dest As kccPath, _
                         Optional OverWriteFiles As Boolean = True) As kccResult
    Const PROC_NAME = "CopyFile"
    
    'コピー元ファイル不在：無視して終了
    If Not Me.Exists Then
        Set CopyFile = kccResult.Init(False, "コピー元ファイルがありません")
        Exit Function
    End If
    
    'フォルダの場合、フォルダそのものをコピーする？未実装
    If Me.IsFile Then Else Stop
    
    Dim fl As File:   Set fl = Me.File
    Dim destFile As kccPath: Set destFile = dest
    If dest.IsFile Then Else Set destFile = destFile.MovePathToFilePath(".\" & Me.File.Name)
    
    Set CopyFile = kccResult.Init(True)
    
    If dest.Exists Then
        If OverWriteFiles Then
            CopyFile.Add True, dest.FullPath & " 上書きします"
        Else
            Set CopyFile = kccResult.Init(False, dest.FullPath & " 既に存在するため失敗しました")
            Exit Function
        End If
    End If
    
    On Error GoTo CopyFileError
    fl.Copy dest.FullPath, OverWriteFiles:=OverWriteFiles
    On Error GoTo 0
    
    CopyFile.Add True, PROC_NAME & " 完了しました。"
    
CopyFileEnd:
    Exit Function
    
CopyFileError:
    CopyFile.IsSuccess = False
    Select Case MsgBox( _
            "[" & dest.FullPath & "]" & "へファイルをコピーできません。" & vbLf & _
            "ファイルまたがロックされていないか確認してください。", _
            vbAbortRetryIgnore, PROC_NAME)
        Case VbMsgBoxResult.vbAbort: CopyFile.Add False, dest.FullPath & " 失敗し中止されました", True: Resume CopyFileEnd
        Case VbMsgBoxResult.vbRetry: CopyFile.Add False, dest.FullPath & " 失敗し再試行しました": Resume
        Case VbMsgBoxResult.vbIgnore: CopyFile.Add False, dest.FullPath & " 失敗し省略されました": Resume Next
    End Select
End Function

Rem フォルダ内のファイル・フォルダをまとめてコピーする
Rem
Rem  @param dest                コピー先ファイル名またはフォルダ
Rem  @param withFilterString    含めるファイルを表すLike比較用文字列(既定:全て)
Rem  @param withoutFilterString 除外するファイルを表すLike比較用文字列(既定:無し)
Rem  @param OverWriteFiles      コピー先ファイルが存在する時上書きするか(既定:True)
Rem
Rem  @note
Rem    dest:=ファイル指定・・・対象ファイル名のフォルダを作成
Rem    dest:=フォルダ指定・・・対象フォルダ名のフォルダを作成
Rem    速度は低下するが無視している
Rem
Public Function CopyFiles(dest As kccPath, _
                          Optional withFilterString As String = "*", _
                          Optional withoutFilterString As String = "", _
                          Optional OverWriteFiles = True) As kccResult
    Const PROC_NAME = "CopyFiles"
    
    If Me.IsFile Then: Set CopyFiles = Me.CopyFile(dest): Exit Function
    
    On Error Resume Next
    Dim strIgnore: strIgnore = Me.Folder.Files(".kccignore").OpenAsTextStream.ReadAll()
    Dim arrIgnore: arrIgnore = ToArrayByIgnoreFileText(strIgnore)
    On Error GoTo 0
    
    Set CopyFiles = kccResult.Init(True)
    
    If Me.CurrentFolderPath = "" Then Stop
    Dim fl As File
    Dim fd As String
    For Each fl In Me.Folder.Files
        If fl.Name Like withFilterString And _
            Not fl.Name Like withoutFilterString Then
            
            Dim fn As String: fn = fl.Path
            If MatchLike(arrIgnore, fn) = 0 Then
                'このCopyでは失敗してもエラーが起こらないらしい？
                'ロックされてるとエラーが出る。
                If dest.IsFile Then
                    fd = dest.ReplacePathAuto(FileName:=fl.Name).CreateFolder.FullPath
                Else
                    fd = dest.ReplacePathAuto(FileName:=fl.Name).MovePathToFilePath(fl.Name).CreateFolder.FullPath
                End If
                
                On Error GoTo CopyFilesError
                    fl.Copy fd, True
                On Error GoTo 0
            End If
        End If
    Next
    
    CopyFiles.Add True, PROC_NAME & " 完了しました。"
    
CopyFilesEnd:
    Exit Function
    
CopyFilesError:
    Dim res As VbMsgBoxResult
    Select Case MsgBox( _
            "[" & fd & "]" & "へファイルをコピーできません。" & vbLf & _
            "ファイルまたは上位のフォルダがロックされていないか確認してください。", _
            vbAbortRetryIgnore, PROC_NAME)
        Case VbMsgBoxResult.vbAbort: CopyFiles.Add False, dest.FullPath & " 失敗し中止されました", True: Resume CopyFilesEnd
        Case VbMsgBoxResult.vbRetry: CopyFiles.Add False, dest.FullPath & " 失敗し再試行しました": Resume
        Case VbMsgBoxResult.vbIgnore: CopyFiles.Add False, dest.FullPath & " 失敗し省略されました": Resume Next
    End Select
End Function

Public Function MatchLike(arr, v) As Long
    MatchLike = 0
    Dim xx
    For Each xx In arr
        MatchLike = MatchLike + 1
        If v Like "*" & xx Then Exit Function
    Next
    MatchLike = 0
End Function

Rem ignoreファイルのテキストを配列に変換する
Public Function ToArrayByIgnoreFile(ignoreFile) As Variant
    Dim s As String
    s = fso.OpenTextFile(ignoreFile).ReadAll()
    ToArrayByIgnoreFile = ToArrayByIgnoreFileText(s)
End Function

Public Function ToArrayByIgnoreFileText(strIgnore) As Variant
    Dim sss() As String: sss = Split(Replace(strIgnore, vbCrLf, vbLf), vbLf)
    Dim arr() As Variant: ReDim arr(1 To UBound(sss))
    
    Dim i As Long, n As Long
    For i = LBound(sss) To UBound(sss)
        If sss(i) Like "[#]*" Or sss(i) = "" Then
            'コメント
        Else
            n = n + 1
            arr(n) = sss(i)
        End If
    Next
    ReDim Preserve arr(1 To n)
    ToArrayByIgnoreFileText = arr
End Function

Public Sub Test_kccPath()
    Dim p1 As kccPath
    Dim p2 As kccPath
    Set p1 = kccPath.Init(ThisWorkbook).CurrentFolder
    Dim sp2 As String
    sp2 = p1.MoveFolderPath("..\bin\")
    Set p2 = kccPath.Init(sp2)
    Dim p3 As kccPath
    
    Dim res As kccResult
    Set res = p1.CopyFiles(p2)
End Sub

Public Sub Test_IgnoreFile()
    Dim ignoreFile As String
    ignoreFile = ThisWorkbook.Path & "\.kccignore"
    
    Dim s As String
    s = Join(kccPath.ToArrayByIgnoreFile(ignoreFile), vbLf)
    
    Debug.Print s
End Sub

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

Rem エクスプローラで開く
Rem  @param IsSelected  対象を選択状態にするか
Rem                       既定は対象によって変化
Rem                         ファイル：True
Rem                         フォルダ：False
Rem
Public Sub OpenExplorer(Optional IsSelected)
    If VBA.IsMissing(IsSelected) Then
        IsSelected = Me.IsFile
    End If
    kccFuncWindowsProcess.ShellExplorer Me.FullPath, (IsSelected = True)
End Sub

Rem 関連付けられたプログラムで開く
Rem   ファイル：関連付けられたプログラム
Rem   フォルダ：エクスプローラ
Public Sub OpenAssociation()
    Debug.Print kccFuncWindowsProcess.OpenAssociationShell32(Me.FullPath)
End Sub

Rem ファイルの文字コードをSJISからUTF8(BOM無し)に変換する
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
            .SaveToFile fn, 2
            .Close
        End With
        
        .Close
    End With
End Sub

Rem UTF-8で作成されたファイルを読む
Public Function ReadUTF8Text(argPath As String) As String

    Dim buf  As String

    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Type = 2           'adTypeText
        .LineSeparator = -1 'adCrLf
        .Open
        .LoadFromFile argPath
        buf = .ReadText(-1) 'adReadAll
        .Close
    End With

    ReadUTF8Text = buf

End Function
