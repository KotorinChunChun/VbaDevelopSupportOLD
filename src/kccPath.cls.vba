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
Rem  @description   �p�X���Ǘ��N���X
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
Rem    2021/04/12 ���̂�Move����n�܂�p�X���ړ�����֐��̖��O��Select�ɕύX�i���̂̈ړ��ƍ�������j
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

Rem �I�u�W�F�N�g�̍쐬
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

Rem VBProject���疼�O���擾����֐�
Rem
Rem  ���ۑ��̃u�b�N�ł�VBProject.FileName���G���[�ɂȂ�B
Rem  VBProject���璼�ږ��O���擾�����i�͑��ɑ��݂��Ȃ��B
Rem  ���ۑ��̃u�b�N��Workbook.FullPath�Ȃǂ�[Book1]�ƌ������P���Ȗ��O�����Ԃ��Ȃ��B
Rem
Rem  ���̊֐����g���ɂ�[VBA �v���W�F�N�g �I�u�W�F�N�g���f���ւ̃A�N�Z�X]�̋����K�v
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

Rem �u�b�N������Workbook��Ԃ��B
Rem
Rem  �����������炱�̕��@�ł͎擾�ł��Ȃ����Ⴊ���邩������Ȃ��B
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

Rem �t���p�X��
Property Get FullPath() As String: FullPath = Me.FullPath__ & IIf(Me.IsFile, "", "\"): End Function
Property Let FullPath(Path As String)
    If Path Like "*\" Then Me.IsFile = False
    '�t���p�X�AUNC�A���΁A�J�����g�������F�����ăt���p�X��
    FullPath__ = kccFuncString.ToPathLastYen(Path, False)
End Property

Rem �t�@�C���܂��̓t�H���_��
Property Get Name() As String
    Name = kccFuncString.GetPath(FullPath, False, True, True)
End Property

Rem �t�@�C����
Rem  �t�H���_�̂Ƃ���
Property Get FileName() As String
    Dim IsFolder As Boolean
    FileName = kccFuncString.GetPath(FullPath, False, True, True, outIsFolder:=IsFolder)
    If IsFolder Then FileName = ""
End Property

Rem �g���q���������O
Property Get BaseName() As String
    BaseName = kccFuncString.GetPath(FullPath, False, True, False)
End Property

Rem �g���q�̖��O�i.ext�j
Property Get Extension() As String
    Extension = kccFuncString.GetPath(FullPath, False, False, True)
End Property

Rem �t�H���_��
Rem  �t�@�C���̂Ƃ���
Property Get FolderName() As String
    Dim IsFolder As Boolean
    FolderName = kccFuncString.GetPath(FullPath, False, True, True, outIsFolder:=IsFolder)
    If IsFolder Then Else FolderName = ""
End Property

Rem ���t�H���_�t���p�X
Property Get CurrentFolderPath(Optional AddYen As Boolean = False) As String
    If Me.IsFile Then
        CurrentFolderPath = kccFuncString.GetPath(Me.FullPath, True, False, False)
    Else
        CurrentFolderPath = Me.FullPath
    End If
    CurrentFolderPath = kccFuncString.ToPathLastYen(CurrentFolderPath, AddYen)
End Property

Rem ���݂̃t�H���_���̕ύX
Property Let CurrentFolderName(FolderName As String)
    Dim cur As Scripting.Folder
    Set cur = Me.CurrentFolder.Folder
    cur.Name = FolderName
End Property

Rem �e�t�H���_��
Property Get ParentFolderPath(Optional AddYen As Boolean = False) As String
    If Me.IsFile Then
        ParentFolderPath = kccFuncString.GetPath(Me.CurrentFolderPath(AddYen:=False), True, False, False)
    Else
        ParentFolderPath = kccFuncString.GetPath(Me.FullPath, True, False, False)
    End If
    ParentFolderPath = kccFuncString.ToPathLastYen(ParentFolderPath, AddYen)
End Property

Rem �e�t�H���_�I�u�W�F�N�g
Property Get CurrentFolder() As kccPath
    Set CurrentFolder = kccPath.Init(Me.CurrentFolderPath, False)
End Property

Rem �e�t�H���_�I�u�W�F�N�g
Property Get ParentFolder() As kccPath
    Set ParentFolder = kccPath.Init(Me.ParentFolderPath, False)
End Property

Rem FSO�t�@�C���I�u�W�F�N�g
Public Function File() As Scripting.File
    On Error Resume Next
    Set File = fso.GetFile(FullPath)
End Function

Rem FSO�t�H���_�I�u�W�F�N�g
Public Function Folder() As Scripting.Folder
    On Error Resume Next
    Set Folder = fso.GetFolder(Me.CurrentFolderPath)
End Function

Rem VB�v���W�F�N�g
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

Rem Excel���[�N�u�b�N
Public Function Workbook() As Excel.Workbook
    If Me.FileName = "" Then Exit Function
    '[Workbooks("Book1.xlsx")]
    '[Workbooks("Book1")]
    Set Workbook = GetWorkbook(Me.FileName)
End Function

Rem ���΃p�X�ɂ��ړ������t�H���_�̃p�X
Public Function SelectFolderPath(relative_path) As String
    If Me.IsFile Then
        SelectFolderPath = kccFuncString.AbsolutePathNameEx(Me.CurrentFolder.FullPath, relative_path)
    Else
        SelectFolderPath = kccFuncString.AbsolutePathNameEx(Me.FullPath, relative_path)
    End If
End Function

Rem ���΃p�X�ɂ��ړ������t�H���_�̃C���X�^���X��V�K����
Public Function SelectPathToFolderPath(relative_path) As kccPath
    Dim basePath As String: basePath = Me.CurrentFolderPath
    Dim refePath As String: refePath = relative_path
    Dim absoPath As String: absoPath = kccFuncString.AbsolutePathNameEx(basePath, refePath)
    Set SelectPathToFolderPath = kccPath.Init(absoPath, False)
End Function

Rem ���΃p�X�ɂ��ړ������t�@�C���̃C���X�^���X��V�K����
Rem   �������t�H���_�̂Ƃ��F�u���p�X\relative_path�v
Rem   �������t�@�C���̂Ƃ��F�u�J�����g�t�H���_\relative_path�v
Rem
Rem  ���ʂɎg�p�ł��镶�� : |t |e
Rem    �G�X�P�[�v�����́A�t�@�C�����Ɏg�p�ł��Ȃ��p�C�v | �Ƃ���B
Rem    hoge.ext
Rem      ���̃t�@�C����       : hoge : |t : title�̗�
Rem      ���̃t�@�C���̊g���q : .ext : |e : extension�̗�
Public Function SelectPathToFilePath(ByVal relative_path) As kccPath
    If VarType(relative_path) <> vbString Then Err.Raise 9999, , "�^���Ⴂ�܂�"
    If relative_path = "" Then Set SelectPathToFilePath = kccPath.Init(Me)
    relative_path = Replace(relative_path, "|t", Me.BaseName)
    relative_path = Replace(relative_path, "|e", Me.Extension)
    '���g���t�H���_�ňړ���t�@�C�����t�@�C�����݂̂����w�肳��Ȃ������ꍇ�A�J�����g������\��ǋL
    If Not Me.IsFile And Not relative_path Like "\*" Then relative_path = "\" & relative_path
    Dim basePath As String: basePath = Me.CurrentFolderPath
    Dim refePath As String: refePath = IIf(relative_path Like "*\*", "", ".\") & relative_path
    Dim absoPath As String: absoPath = kccFuncString.AbsolutePathNameEx(basePath, refePath)
    Set SelectPathToFilePath = kccPath.Init(absoPath, True)
End Function

Rem ���΃p�X�ɂ��t�@�C�������ێ������܂ܐe�t�H���_���ړ�����
Public Function SelectParentFolder(ByVal relative_path) As kccPath
    Set SelectParentFolder = Me.SelectPathToFilePath(relative_path & "|t|e")
End Function

Rem �t�H���_����C�ɍ쐬
Rem  ���������ꍇ
Rem  ����:���ɑ��݂����ꍇ
Rem  ���s:�t�@�C�������ɑ��݂����ꍇ
Rem  ���s:����ȊO�̗��R
Public Function CreateFolder() As kccPath
    Set CreateFolder = Me
    Dim errValue
    If Not kccFuncPath.CreateDirectoryEx(Me.CurrentFolderPath, errValue) Then
        Debug.Print "CreateFolder ���s : " & errValue & ":" & Me.CurrentFolderPath
        Err.Raise errValue, "CreateFolder", "CreateFolder ���s : " & errValue & ":" & Me.CurrentFolderPath
    End If
End Function

Rem �t�H���_���폜
Rem
Rem  @return As Boolean �폜����
Rem                         True  : ����(�폜�ɐ��� or ���X�t�H���_������)
Rem                         False : ���s(�t�H���_���c���Ă���)
Rem
Public Function DeleteFolder() As Boolean
    Dim cPath As String: cPath = Me.CurrentFolderPath
    If fso.FolderExists(cPath) Then
        '1�b�󂯂�3�񃊃g���C
        Dim n As Long: n = 3
        Do
            On Error Resume Next
            fso.DeleteFolder cPath
            If Err.Number = 0 Then Exit Do
            On Error GoTo 0
            Application.Wait [Now() + "00:00:01"]
            n = n - 1
            If n = 0 Then Exit Do 'Err.Raise 9999, "DeleteFolder", "�폜�ł��܂���"
        Loop
        DoEvents
    End If
    DeleteFolder = Not fso.FolderExists(cPath)
End Function

Rem �t�@�C���E�t�H���_�����݂��邩
Public Function Exists() As Boolean
    If Me.IsFile Then
        Exists = Not (Me.File Is Nothing)
    Else
        Exists = Not (Me.Folder Is Nothing)
    End If
End Function

Rem �t�@�C�����R�s�[����
Rem
Rem  @param dest            �R�s�[��t�@�C�����܂��̓t�H���_
Rem  @param OverWriteFiles  �R�s�[��t�@�C�������݂��鎞�㏑�����邩(����:True)
Rem
Rem  @note
Rem    dest:=�t�@�C���w��E�E�E�Ώۃt�@�C�����ŏ�������
Rem    dest:=�t�H���_�w��E�E�E�Ώۃt�H���_�Ɍ��Ɠ����t�@�C�����ŏ�������
Rem
Public Function CopyFile(dest As kccPath, _
                         Optional OverWriteFiles As Boolean = True) As kccResult
    Const PROC_NAME = "CopyFile"
    
    '�R�s�[���t�@�C���s�݁F�������ďI��
    If Not Me.Exists Then
        Set CopyFile = kccResult.Init(False, "�R�s�[���t�@�C��������܂���")
        Exit Function
    End If
    
    '�t�H���_�̏ꍇ�A�t�H���_���̂��̂��R�s�[����H������
    If Me.IsFile Then Else Stop
    
    Dim fl As File:   Set fl = Me.File
    Dim destFile As kccPath: Set destFile = dest
    If dest.IsFile Then Else Set destFile = destFile.SelectPathToFilePath(".\" & Me.File.Name)
    
    Set CopyFile = kccResult.Init(True)
    
    If dest.Exists Then
        If OverWriteFiles Then
            CopyFile.Add True, dest.FullPath & " �㏑�����܂�"
        Else
            Set CopyFile = kccResult.Init(False, dest.FullPath & " ���ɑ��݂��邽�ߎ��s���܂���")
            Exit Function
        End If
    End If
    
    On Error GoTo CopyFileError
    fl.Copy dest.FullPath, OverWriteFiles:=OverWriteFiles
    On Error GoTo 0
    
    CopyFile.Add True, PROC_NAME & " �������܂����B"
    
CopyFileEnd:
    Exit Function
    
CopyFileError:
    CopyFile.IsSuccess = False
    Select Case MsgBox( _
            "[" & dest.FullPath & "]" & "�փt�@�C�����R�s�[�ł��܂���B" & vbLf & _
            "�t�@�C���܂������b�N����Ă��Ȃ����m�F���Ă��������B", _
            vbAbortRetryIgnore, PROC_NAME)
        Case VbMsgBoxResult.vbAbort: CopyFile.Add False, dest.FullPath & " ���s�����~����܂���", True: Resume CopyFileEnd
        Case VbMsgBoxResult.vbRetry: CopyFile.Add False, dest.FullPath & " ���s���Ď��s���܂���": Resume
        Case VbMsgBoxResult.vbIgnore: CopyFile.Add False, dest.FullPath & " ���s���ȗ�����܂���": Resume Next
    End Select
End Function

Rem �t�@�C�������ׂč폜����
Rem �G���[�����͕ۗ�
Public Function DeleteFiles()
    fso.DeleteFile Me.FullPath & "\*"
End Function

Public Function DeleteFolders()
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

Rem �t�H���_���̃t�@�C���E�t�H���_���܂Ƃ߂ăR�s�[����
Rem
Rem  @param dest                �R�s�[��t�@�C�����܂��̓t�H���_
Rem  @param withFilterString    �܂߂�t�@�C����\��Like��r�p������(����:�S��)
Rem  @param withoutFilterString ���O����t�@�C����\��Like��r�p������(����:����)
Rem  @param OverWriteFiles      �R�s�[��t�@�C�������݂��鎞�㏑�����邩(����:True)
Rem
Rem  @note
Rem    dest:=�t�@�C���w��E�E�E�Ώۃt�@�C�����̃t�H���_���쐬
Rem    dest:=�t�H���_�w��E�E�E�Ώۃt�H���_���̃t�H���_���쐬
Rem    ���x�͒ቺ���邪�������Ă���
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

Rem �t�H���_���̃t�@�C���E�t�H���_���܂Ƃ߂ăR�s�[����
Rem
Rem  @param dest                �R�s�[��t�@�C�����܂��̓t�H���_
Rem  @param IsCopy              �R�s�[����̂��ړ�����̂�(����:False:�ړ�)
Rem  @param withFilterString    �܂߂�t�@�C����\��Like��r�p������(����:�S��)
Rem  @param withoutFilterString ���O����t�@�C����\��Like��r�p������(����:����)
Rem  @param OverWriteFiles      �R�s�[��t�@�C�������݂��鎞�㏑�����邩(����:True)
Rem
Rem  @note
Rem    dest:=�t�@�C���w��E�E�E�Ώۃt�@�C�����̃t�H���_���쐬
Rem    dest:=�t�H���_�w��E�E�E�Ώۃt�H���_���̃t�H���_���쐬
Rem    ���x�͒ቺ���邪�������Ă���
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
        '����Copy�ł͎��s���Ă��G���[���N����Ȃ��炵���H
        '���b�N����Ă�ƃG���[���o��B
        
        '���O�Ƀt�H���_�����b�N����Ă��āA�t�H���_�쐬�̍H���ŃG���[5�������B���̌�R�s�[�ŃG���[���o��炵���B
        '�Č������ɓ�����߁A�����ł��Ă��炸�B�����炭Dropbox���̓����֌W
        If dest.IsFile Then
            fd = dest.ReplacePathAuto(FileName:=vf).CreateFolder.FullPath
        Else
            fd = dest.ReplacePathAuto(FileName:=vf).SelectPathToFilePath(vf).CreateFolder.FullPath
        End If
        
        On Error GoTo CopyFilesError
            Set fl = fso.GetFile(MePath & vf)
            If IsCopy Then
                fl.Copy fd, True
            Else
                fl.Move fd
            End If
            Debug.Print "MoveCopyTo : " & fl.Path & " to " & fd
            If Err Then Stop
        On Error GoTo 0
    Next
    
    MoveCopyTo.Add True, PROC_NAME & " �������܂����B"
    
CopyFilesEnd:
    Exit Function
    
CopyFilesError:
    Dim res As VbMsgBoxResult
    Select Case MsgBox( _
            "[" & fd & "]" & "�փt�@�C�����R�s�[�ł��܂���B" & vbLf & _
            "�t�@�C���܂��͏�ʂ̃t�H���_�����b�N����Ă��Ȃ����m�F���Ă��������B", _
            vbAbortRetryIgnore, PROC_NAME)
        Case VbMsgBoxResult.vbAbort: MoveCopyTo.Add False, dest.FullPath & " ���s�����~����܂���", True: Resume CopyFilesEnd
        Case VbMsgBoxResult.vbRetry: MoveCopyTo.Add False, dest.FullPath & " ���s���Ď��s���܂���": Resume
        Case VbMsgBoxResult.vbIgnore: MoveCopyTo.Add False, dest.FullPath & " ���s���ȗ�����܂���": Resume Next
    End Select
End Function

Rem ignore�Ńt�@�C�����X�g���i�荞��
Public Function IgnoreFilter(cll As Collection, arr) As Collection

End Function

Rem gitignore��������
Rem �p�X�̋�؂�\����
Rem �p�X�̕�����͏���������
Public Function MatchLike(arr, ByVal v) As Long
    MatchLike = 0
    If IsEmpty(arr) Then Exit Function
'    Dim fln As String
'    fln = Mid(v, Len(v) - InStr(v, "\"))
    v = LCase(Replace(v, "/", "\"))
    Dim xx
    For Each xx In arr
        MatchLike = MatchLike + 1
        If Right(xx, 1) = "\" Then
            If v Like xx & "*" Then Exit Function
        Else
            If LCase(v) Like LCase(xx) Then Exit Function
        End If
    Next
    MatchLike = 0
End Function

Rem ignore�t�@�C���̃e�L�X�g��z��ɕϊ�����
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
            '�R�����g
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

Rem �p�X�������P���ɒu��
Public Function ReplacePath(src, dest) As kccPath
    Set ReplacePath = Me.Clone
    ReplacePath.FullPath = Replace(ReplacePath.FullPath, src, dest)
End Function

Rem �p�X��������}�W�b�N�i���o�[�ɂ��u��
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

Rem �G�N�X�v���[���ŊJ��
Rem  @param IsSelected  �Ώۂ�I����Ԃɂ��邩
Rem                       ����͑Ώۂɂ���ĕω�
Rem                         �t�@�C���FTrue
Rem                         �t�H���_�FFalse
Rem
Public Sub OpenExplorer(Optional IsSelected)
    If VBA.IsMissing(IsSelected) Then
        IsSelected = Me.IsFile
    End If
    kccFuncWindowsProcess.ShellExplorer Me.FullPath, (IsSelected = True)
End Sub

Rem �֘A�t����ꂽ�v���O�����ŊJ��
Rem   �t�@�C���F�֘A�t����ꂽ�v���O����
Rem   �t�H���_�F�G�N�X�v���[��
Public Sub OpenAssociation()
    Debug.Print kccFuncWindowsProcess.OpenAssociationShell32(Me.FullPath)
End Sub

Rem �t�@�C���̕����R�[�h��SJIS����UTF8(BOM����)�ɕϊ�����
Public Sub ConvertCharCode_SJIS_to_utf8()
    If Me.IsFile Then Else Exit Sub
    If fso.FileExists(Me.FullPath) Then Else Exit Sub
    
    Dim fn As String: fn = Me.FullPath
    Dim destWithBOM As Object: Set destWithBOM = CreateObject("ADODB.Stream")
    With destWithBOM
        .Type = 2
        .Charset = "utf-8"
        .Open
        
        ' �t�@�C����SJIS �ŊJ���āAdest �� �o��
        With CreateObject("ADODB.Stream")
            .Type = 2
            .Charset = "shift-jis"
            .Open
            .LoadFromFile fn
            .Position = 0
            .CopyTo destWithBOM
            .Close
        End With
        
        ' BOM����
        ' 3�o�C�g�������Ă���o�C�i���Ƃ��ďo��
        .Position = 0
        .Type = 1 ' adTypeBinary
        .Position = 3
        
        Dim dest: Set dest = CreateObject("ADODB.Stream")
        With dest
            .Type = 1 ' adTypeBinary
            .Open
            destWithBOM.CopyTo dest
            .SaveToFile fn, 2
            .Close
        End With
        
        .Close
    End With
End Sub

Rem UTF-8�ō쐬���ꂽ�t�@�C����ǂ�
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
