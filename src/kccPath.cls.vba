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
Rem    2021/04/12 名称がMoveから始まるパスを移動する関数の名前をSelectに変更（実体の移動と混乱する）
Rem
Rem --------------------------------------------------------------------------------
Rem  @functions
Rem
Rem --------------------------------------------------------------------------------
Rem  @note
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Private Const IGNORE_FILE = ".kccignore"

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
Public Function SelectFolderPath(relative_path) As String
    If Me.IsFile Then
        SelectFolderPath = kccFuncString.AbsolutePathNameEx(Me.CurrentFolder.FullPath, relative_path)
    Else
        SelectFolderPath = kccFuncString.AbsolutePathNameEx(Me.FullPath, relative_path)
    End If
End Function

Rem 相対パスにより移動したフォルダのインスタンスを新規生成
Public Function SelectPathToFolderPath(relative_path) As kccPath
    Dim basePath As String: basePath = Me.CurrentFolderPath
    Dim refePath As String: refePath = relative_path
    Dim absoPath As String: absoPath = kccFuncString.AbsolutePathNameEx(basePath, refePath)
    Set SelectPathToFolderPath = kccPath.Init(absoPath, False)
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
Public Function SelectPathToFilePath(ByVal relative_path) As kccPath
    If VarType(relative_path) <> vbString Then Err.Raise 9999, , "型が違います"
    If relative_path = "" Then Set SelectPathToFilePath = kccPath.Init(Me)
    relative_path = Replace(relative_path, "|t", Me.BaseName)
    relative_path = Replace(relative_path, "|e", Me.Extension)
    '自身がフォルダで移動先ファイルがファイル名のみしか指定されなかった場合、カレントを示す\を追記
    If Not Me.IsFile And Not relative_path Like "\*" Then relative_path = "\" & relative_path
    Dim basePath As String: basePath = Me.CurrentFolderPath
    Dim refePath As String: refePath = IIf(relative_path Like "*\*", "", ".\") & relative_path
    Dim absoPath As String: absoPath = kccFuncString.AbsolutePathNameEx(basePath, refePath)
    Set SelectPathToFilePath = kccPath.Init(absoPath, True)
End Function

Rem 相対パスによりファイル名を維持したまま親フォルダを移動する
Public Function SelectParentFolder(ByVal relative_path) As kccPath
    Set SelectParentFolder = Me.SelectPathToFilePath(relative_path & "|t|e")
End Function

Rem フォルダを一気に作成
Rem  成功した場合
Rem  成功:既に存在した場合
Rem  失敗:ファイルが既に存在した場合
Rem  失敗:それ以外の理由
Public Function CreateFolder() As kccPath
    Set CreateFolder = Me
    Dim errValue
    If Not kccFuncPath.CreateDirectoryEx(Me.CurrentFolderPath, errValue) Then
        Debug.Print "CreateFolder 失敗 : " & errValue & ":" & Me.CurrentFolderPath
        Err.Raise errValue, "CreateFolder", "CreateFolder 失敗 : " & errValue & ":" & Me.CurrentFolderPath
    End If
End Function

Rem フォルダを削除
Rem
Rem  @return As Boolean 削除結果
Rem                         True  : 成功(削除に成功 or 元々フォルダが無い)
Rem                         False : 失敗(フォルダが残っている)
Rem
Public Function DeleteFolder() As Boolean
    DeleteFolder = kccFuncPath.DeleteFolderReplay(Me.CurrentFolderPath)
End Function

Rem ファイル・フォルダが存在するか
Public Function Exists() As Boolean
    If Me.IsFile Then
        Exists = Not (Me.File Is Nothing)
    Else
        Exists = Not (Me.Folder Is Nothing)
    End If
End Function

Rem ファイルをコピーする（fso仕様準拠）
Rem
Rem  @param dest            コピー先ファイル名またはフォルダパス
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
    If dest.IsFile Then Else Set destFile = destFile.SelectPathToFilePath(".\" & Me.File.Name)
    
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
'    fl.Copy dest.FullPath, OverWriteFiles:=OverWriteFiles
    kccFuncPath.CopyFile fl.Path, dest.FullPath, OverWriteFiles:=OverWriteFiles
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

Rem ファイルをすべて削除する
Rem エラー処理は保留
Public Function DeleteFiles()
    On Error Resume Next
    If Me.IsFile Then
        fso.DeleteFile Me.FullPath
    Else
        fso.DeleteFile Me.FullPath & "\*"
    End If
End Function

Public Function DeleteFolders()
    On Error Resume Next
    fso.DeleteFolder Me.FullPath & "\*"
End Function

Public Function DeleteItems()
    Call DeleteFiles
    Call DeleteFolders
End Function

Public Function MoveTo(dest As kccPath, _
                          Optional withFilterString As String = "*", _
                          Optional withoutFilterString As String = "", _
                          Optional UseIgnoreFile As Boolean = False, _
                          Optional OverWriteFiles = True) As kccResult
    Set MoveTo = Me.MoveCopyTo( _
                            dest, _
                            IsCopy:=False, _
                            withFilterString:=withFilterString, _
                            withoutFilterString:=withoutFilterString, _
                            UseIgnoreFile:=UseIgnoreFile, _
                            OverWriteFiles:=OverWriteFiles)
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
Public Function CopyTo(dest As kccPath, _
                          Optional withFilterString As String = "*", _
                          Optional withoutFilterString As String = "", _
                          Optional UseIgnoreFile As Boolean = False, _
                          Optional OverWriteFiles = True) As kccResult
    Set CopyTo = Me.MoveCopyTo( _
                            dest, _
                            IsCopy:=True, _
                            withFilterString:=withFilterString, _
                            withoutFilterString:=withoutFilterString, _
                            UseIgnoreFile:=UseIgnoreFile, _
                            OverWriteFiles:=OverWriteFiles)
End Function

Public Function GetIgnoreFile() As kccPath
    Set GetIgnoreFile = Me.SelectPathToFilePath(".\" & IGNORE_FILE)
End Function

Rem フォルダ内のファイル・フォルダをまとめてコピーする
Rem
Rem  @param dest                コピー先ファイル名またはフォルダ
Rem  @param IsCopy              コピーするのか移動するのか(既定:False:移動)
Rem  @param withFilterString    含めるファイルを表すLike比較用文字列(既定:全て)
Rem  @param withoutFilterString 除外するファイルを表すLike比較用文字列(既定:無し)
Rem  @param OverWriteFiles      コピー先ファイルが存在する時上書きするか(既定:True)
Rem
Rem  @note
Rem    dest:=ファイル指定・・・対象ファイル名のフォルダを作成
Rem    dest:=フォルダ指定・・・対象フォルダ名のフォルダを作成
Rem    速度は低下するが無視している
Rem
Public Function MoveCopyTo(dest As kccPath, _
                          Optional IsCopy As Boolean = False, _
                          Optional withFilterString As String = "*", _
                          Optional withoutFilterString As String = "", _
                          Optional UseIgnoreFile As Boolean = False, _
                          Optional OverWriteFiles = True) As kccResult
    Const PROC_NAME = "MoveCopyTo"
    
    If Me.IsFile Then: Set MoveCopyTo = Me.CopyFile(dest): Exit Function
    
    Dim strIgnore, arrIgnore
    If UseIgnoreFile Then
        On Error Resume Next
        strIgnore = Me.Folder.Files(IGNORE_FILE).OpenAsTextStream.ReadAll()
        arrIgnore = ToArrayByIgnoreFileText(strIgnore)
        On Error GoTo 0
    End If
    
    Set MoveCopyTo = kccResult.Init(True)
    
    If Me.CurrentFolderPath = "" Then Stop
    Dim MePath As String: MePath = Me.Folder.Path & "\"
    Dim fl As File
    Dim fd As String
    Dim vf
    Dim cll As Collection
    Set cll = kccFuncPath.GetFileFolderList(MePath, add_files:=True, add_folders:=False, search_min_layer:=-1, search_max_layer:=-1)
'    Set cll = kccFuncArray.Concat(cll, MePath)
    Dim newCll As Collection: Set newCll = New Collection
    For Each vf In cll
        If MatchLike(arrIgnore, vf) = 0 Then
'            If vf Like "\" Then
'                'folder
'                newCll.Add
'            Else
                Set fl = fso.GetFile(MePath & vf)
                If fl.Name Like withFilterString And Not fl.Name Like withoutFilterString Then
                    newCll.Add vf
                End If
'            End If
        End If
    Next
    
    For Each vf In newCll
        'このCopyでは失敗してもエラーが起こらないらしい？
        'ロックされてるとエラーが出る。
        
        '事前にフォルダがロックされていて、フォルダ作成の工程でエラー5が発生。その後コピーでエラーが出るらしい。
        '再現が非常に難しいため、解消できておらず。おそらくDropbox等の同期関係
        If dest.IsFile Then
            fd = dest.ReplacePathAuto(FileName:=vf).CreateFolder.FullPath
        Else
            fd = dest.ReplacePathAuto(FileName:=vf).SelectPathToFilePath(vf).CreateFolder.FullPath
        End If
        
        On Error GoTo CopyFilesError
            Set fl = fso.GetFile(MePath & vf)
            If IsCopy Then
                kccFuncPath.CopyFile fl.Path, fd, True
            Else
                fl.Move fd
            End If
            Debug.Print "MoveCopyTo : " & fl.Path & " to " & fd
            If Err Then Stop
        On Error GoTo 0
    Next
    
    MoveCopyTo.Add True, PROC_NAME & " 完了しました。"
    
CopyFilesEnd:
    Exit Function
    
CopyFilesError:
    Dim res As VbMsgBoxResult
    Select Case MsgBox( _
            "[" & fd & "]" & "へファイルをコピーできません。" & vbLf & _
            "ファイルまたは上位のフォルダがロックされていないか確認してください。", _
            vbAbortRetryIgnore, PROC_NAME)
        Case VbMsgBoxResult.vbAbort: MoveCopyTo.Add False, dest.FullPath & " 失敗し中止されました", True: Resume CopyFilesEnd
        Case VbMsgBoxResult.vbRetry: MoveCopyTo.Add False, dest.FullPath & " 失敗し再試行しました": Resume
        Case VbMsgBoxResult.vbIgnore: MoveCopyTo.Add False, dest.FullPath & " 失敗し省略されました": Resume Next
    End Select
End Function

Rem 指定フォルダにファイル一式を同期
Rem
Rem  @param src 同期先
Rem  @param filter_body_type ファイル内容チェックに基づく同期 セミコロン区切り
Rem  @param filter_size_type ファイルサイズチェックに基づく同期 セミコロン区切り
Rem  @param pair_type ヒットした時セットでコピー処理を行う形式 名前置換前後セミコロン区切り、LF区切り（未実装）
Rem
Rem 同一の名前のファイルは変化していたら上書き
Rem 消滅したファイルは同期先からも削除
Rem
Public Function SyncTo( _
            srcPath As kccPath, _
            filter_body_type As String, _
            filter_size_type As String, _
            pair_type As String)
    Dim bufPath As kccPath
    Set bufPath = Me
    srcPath.CreateFolder
    
    'ファイルリストを取得して静的文字列コレクションに変換
    Dim bufFL As Collection: Set bufFL = kccFuncPath.GetFileFolderList(bufPath.FullPath, add_files:=True)
    Dim srcFL As Collection: Set srcFL = kccFuncPath.GetFileFolderList(srcPath.FullPath, add_files:=True)
    
    'buf有 src有 … 内容を比較して変更があれば上書き
    Dim ext
    Dim bufFN, srcFN
    Dim bufKccPath As kccPath
    Dim srcKccPath As kccPath
    Dim IsExists As Boolean
    For Each bufFN In bufFL
        IsExists = False
        For Each ext In Split(filter_body_type, ";")
            If bufFN Like ext Then IsExists = True: Exit For
        Next
        If IsExists Then
            For Each srcFN In srcFL
                If bufFN = srcFN Then
                    Set bufKccPath = bufPath.SelectPathToFilePath(".\" & bufFN)
                    Set srcKccPath = srcPath.SelectPathToFilePath(".\" & srcFN)
                    
                    Dim bufTxt As String: bufTxt = fso.OpenTextFile(bufKccPath.FullPath, ForReading).ReadAll()
                    Dim srcTxt As String: srcTxt = fso.OpenTextFile(srcKccPath.FullPath, ForReading).ReadAll()
                    
                    If bufTxt <> srcTxt Then
                        bufKccPath.CopyTo srcKccPath, OverWriteFiles:=True
                        
                        'セットで移動
                        Dim filetype_pair: filetype_pair = Split(pair_type, ";")
                        If bufFN Like "*" & filetype_pair(0) Then
                            Set bufKccPath = bufPath.SelectPathToFilePath(".\" & Replace(bufFN, filetype_pair(0), filetype_pair(1)))
                            Set srcKccPath = srcPath.SelectPathToFilePath(".\" & Replace(srcFN, filetype_pair(0), filetype_pair(1)))
                            bufKccPath.CopyTo srcKccPath, OverWriteFiles:=True
                        End If
                    End If
                End If
            Next
        End If
    Next
    
    'frxのファイルサイズが変化…frxのみ差し替え
    For Each bufFN In bufFL
        IsExists = False
        For Each ext In Split(filter_size_type, ";")
            If bufFN Like ext Then IsExists = True: Exit For
        Next
        If IsExists Then
            For Each srcFN In srcFL
                If bufFN = srcFN Then
                    Set bufKccPath = bufPath.SelectPathToFilePath(".\" & bufFN)
                    Set srcKccPath = srcPath.SelectPathToFilePath(".\" & srcFN)
                    If bufKccPath.File.Size <> srcKccPath.File.Size Then
                        bufKccPath.CopyTo srcKccPath, OverWriteFiles:=True
                    End If
                End If
            Next
        End If
    Next
    
    'buf有 src無 … bufからコピー
    For Each bufFN In bufFL
        IsExists = False
        For Each srcFN In srcFL
            If bufFN = srcFN Then IsExists = True: Exit For
        Next
        If Not IsExists Then
            Set bufKccPath = bufPath.SelectPathToFilePath(".\" & bufFN)
            Set srcKccPath = srcPath.SelectPathToFilePath(".\" & bufFN)
            bufKccPath.CopyTo srcKccPath, OverWriteFiles:=True
        End If
    Next
    
    'buf無 src有 … srcから消去
    For Each srcFN In srcFL
        IsExists = False
        For Each bufFN In bufFL
            If bufFN = srcFN Then IsExists = True: Exit For
        Next
        If Not IsExists Then
            Set srcKccPath = srcPath.SelectPathToFilePath(".\" & srcFN)
            srcKccPath.DeleteFiles
        End If
    Next
End Function

Rem ignoreでファイルリストを絞り込み
Public Function IgnoreFilter(cll As Collection, arr) As Collection

End Function

Rem gitignore準拠判定
Rem パスの区切は\限定
Rem パスの文字列は小文字限定
Public Function MatchLike(arr, ByVal test_value) As Long
    MatchLike = 0
    If IsEmpty(arr) Then Exit Function
    Dim tv: tv = LCase(Replace(test_value, "/", "\"))
    
    Dim i As Long
    For i = 1 To UBound(arr)
        Dim ptn: ptn = arr(i)
        If ptn Like "!*" Then
            Rem 除外しない
            Dim inPtn
            inPtn = Mid(ptn, 2, Len(ptn))
            If tv Like inPtn Then MatchLike = 0
        Else
            Rem 除外する
            If Right(ptn, 1) = "\" Then
                If tv Like ptn & "*" Then MatchLike = i
            Else
                If LCase(tv) Like LCase(ptn) Then MatchLike = i
            End If
        End If
    Next
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
            arr(n) = LCase(Replace(sss(i), "/", "\"))
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
    sp2 = p1.SelectFolderPath("..\bin\")
    Set p2 = kccPath.Init(sp2)
    Dim p3 As kccPath
    
    Dim res As kccResult
    Set res = p1.CopyTo(p2)
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

Rem SJISで作成されたファイルの文字コードをUTF8(BOM無し)に変換する
Public Function ConvertCharCode_SJIS_to_utf8() As Boolean
    If Not Me.IsFile Then Exit Function
    If fso.FileExists(Me.FullPath) Then Else Exit Function
    ConvertCharCode_SJIS_to_utf8 = kccFuncPath.ConvertCharCode_SJIS_to_utf8(Me.FullPath)
End Function

Rem UTF-8で作成されたファイルを読み込む
Public Function ReadUTF8Text() As String
    If Not Me.IsFile Then Exit Function
    ReadUTF8Text = kccFuncPath.ReadUTF8Text(Me.FullPath)
End Function

Rem UTF-8でファイルへ書き込む
Public Function WriteUTF8Text(strText As String) As Boolean
    If Not Me.IsFile Then Exit Function
    WriteUTF8Text = kccFuncPath.WriteUTF8Text(Me.FullPath, strText)
End Function
