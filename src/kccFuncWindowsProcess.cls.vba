VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kccFuncWindowsProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetActiveWindow Lib "User32" () As LongPtr
#Else
    Private Declare Function GetActiveWindow Lib "User32" () As Long
#End If

Rem --------------------------------------------------------------------------------
Rem ShellExecute�֐�
Rem https://docs.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shellexecutea
Rem �@�\   �w��t�@�C�����w�肵������Ŏ��s���܂��
Rem ����
Rem  hWnd           ShellExecute���Ăяo���E�B���h�E�̃n���h��
Rem  lpOperation    �������䕶���� open edit explore find print runas NULL
Rem                 NULL�̋���:�f�t�H���g����(��`����Ă�����)��OPEN(��`����Ă�����)�����W�X�g���ŗ񋓂���Ă���ŏ��̓���
Rem                 �ݒ�l��SHELLEXECUTEINFO�\���̂�lpVerb�����o���Q�Ƃ��Ă��������B�A���A"properties"�͐ݒ�ł��܂���B
Rem  lpFile         �N������t�@�C���̖��O
Rem  lpParameters   �N��������s�t�@�C���ւ̃p�����[�^�ilpFile�����s�\�t�@�C���̂Ƃ��j�BlpFile���h�L�������g�t�@�C���̂Ƃ��͐ݒ肵�Ȃ��ŉ������B
Rem  lpDirectory    ��Ɨp�f�B���N�g����ݒ肵�Ȃ��Ƃ��̓J�����g�f�B���N�g���ɂȂ�܂��
Rem  nShowCmd       �N��������s�\�t�@�C���̃E�B���h�E�̏�ԡ�ݒ�l��SHELLEXECUTEINFO�\���̂�nShow�����o SW_****
Private Const SW_HIDE = 0            '�E�B���h�E���\���ɂ��āA���̃E�B���h�E���A�N�e�B�u�ɂ��܂��B
Private Const SW_SHOWNORMAL = 1      '�E�B���h�E���A�N�e�B�u�ɂ��ĕ\�����܂��B�E�B���h�E���ŏ����܂��͍ő剻����Ă���ꍇ�́A�E�B���h�E�̈ʒu�ƃT�C�Y�����ɖ߂��܂��B�A�v���P�[�V�����́A�ŏ��ɃE�B���h�E��\��������Ƃ��ɂ��̃t���O���w�肷��ׂ��ł��B
Private Const SW_SHOWMINIMIZED = 2   '�E�B���h�E���A�N�e�B�u�ɂ��āA�ŏ������ꂽ�E�B���h�E�Ƃ��ĕ\�����܂��B
Private Const SW_SHOWMAXIMIZED = 3   '�E�B���h�E���A�N�e�B�u�ɂ��āA�ő剻���ꂽ�E�B���h�E�Ƃ��ĕ\�����܂��B
Private Const SW_MAXIMIZE = 3        '�E�B���h�E���A�N�e�B�u�ɂ��āA�ő剻���ꂽ�E�B���h�E�Ƃ��ĕ\�����܂��B
Private Const SW_SHOWNOACTIVATE = 4  '�E�B���h�E���A�N�e�B�u�ɂ͂����ɕ\�����܂��B
Private Const SW_SHOW = 5            '�E�B���h�E���A�N�e�B�u�ɂ��āA���݂̈ʒu�ƃT�C�Y�ŕ\�����܂��B
Private Const SW_MINIMIZE = 6        '�w�肳�ꂽ�E�B���h�E���ŏ������āA���� Z �I�[�_�[�ɂ���g�b�v���x���E�B���h�E���A�N�e�B�u�ɂ��܂��B
Private Const SW_SHOWMINNOACTIVE = 7 '�E�B���h�E���ŏ������ꂽ�E�B���h�E�Ƃ��ĕ\�����܂��B�E�B���h�E�̓A�N�e�B�u������܂���B
Private Const SW_SHOWNA = 8          '�E�B���h�E�����݂̈ʒu�ƃT�C�Y�ŕ\�����܂��B�E�B���h�E�̓A�N�e�B�u������܂���B
Private Const SW_RESTORE = 9         '�E�B���h�E���A�N�e�B�u�ɂ��ĕ\�����܂��B�E�B���h�E���ŏ����܂��͍ő剻����Ă���ꍇ�́A�E�B���h�E�̈ʒu�ƃT�C�Y�����ɖ߂��܂��B�A�v���P�[�V�����́A�ŏ������ꂽ�E�B���h�E�̈ʒu�ƃT�C�Y�����ɖ߂��Ƃ��ɂ��̃t���O���w�肷��ׂ��ł��B
Private Const SW_SHOWDEFAULT = 10    '�A�v���P�[�V�������N�������v���O������CreateProcess�֐��Ƀp�����[�^�Ƃ��ēn����STARTUPINFO�\���̂Ŏw�肳��Ă��� SW_ �l�Ɋ�Â��ĕ\����Ԃ��ݒ肳��܂��B
Private Const SW_FORCEMINIMIZE = 11  'Windows 2000/XP�F�E�B���h�E�����L���Ă���X���b�h���n���O���Ă����Ԃł����Ă��A�E�B���h�E���ŏ������܂��B���̃X���b�h����E�B���h�E���ŏ���������ꍇ�ɂ̂݁A���̃t���O���g�p����ׂ��ł��B
Rem �߂�l
Rem  33�ȏ� : �J�����t�@�C���̃C���X�^���X�n���h��
Rem  32�ȉ� : �G���[�R�[�h
Rem                                = 0    '�������܂��̓��\�[�X���s�����Ă��܂��B
Private Const ERROR_FILE_NOT_FOUND = 2    '�w�肳�ꂽ�t�@�C����������܂���ł����B
Private Const ERROR_PATH_NOT_FOUND = 3    '�w�肳�ꂽ�p�X��������܂���ł����B
Private Const ERROR_BAD_FORMAT = 11       '.exe �t�@�C���������ł��BWin32 �� .exe �ł͂Ȃ����A.exe �C���[�W���ɃG���[������܂��B
Private Const SE_ERR_ACCESSDENIED = 5     '�I�y���[�e�B���O�V�X�e�����A�w�肳�ꂽ�t�@�C���ւ̃A�N�Z�X�����ۂ��܂����B
Private Const SE_ERR_ASSOCINCOMPLETE = 27 '�t�@�C�����̊֘A�t�����s���S�܂��͖����ł��B
Private Const SE_ERR_DDEBUSY = 30         '�ق��� DDE �g�����U�N�V���������ݏ������Ȃ̂ŁADDE �g�����U�N�V�����������ł��܂���ł����B
Private Const SE_ERR_DDEFAIL = 29         'DDE �g�����U�N�V���������s���܂����B
Private Const SE_ERR_DDETIMEOUT = 28      '�v�����^�C���A�E�g�����̂ŁADDE �g�����U�N�V�����������ł��܂���ł����B
Private Const SE_ERR_DLLNOTFOUND = 32     '�w�肳�ꂽ�_�C�i�~�b�N�����N���C�u�����iDLL�j��������܂���ł����B
Private Const SE_ERR_FNF = 2              '�w�肳�ꂽ�t�@�C����������܂���ł����B
Private Const SE_ERR_NOASSOC = 31         '�w�肳�ꂽ�t�@�C���g���q�Ɋ֘A�t����ꂽ�A�v���P�[�V����������܂���B
                                         '����\�ł͂Ȃ��t�@�C����������悤�Ƃ����ꍇ���A���̃G���[���Ԃ�܂��B
Private Const SE_ERR_OOM = 8              '�������������̂ɏ\���ȃ�����������܂���B
Private Const SE_ERR_PNF = 3              '�w�肳�ꂽ�p�X���A������܂���ł����B
Private Const SE_ERR_SHARE = 26           '���L�ᔽ���������܂����B

#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "Shell32" Alias "ShellExecuteA" ( _
        ByVal hWnd As LongPtr, _
        ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As LongPtr) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "Shell32" Alias "ShellExecuteA" ( _
        ByVal hWnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd_ As Long) As Long
#End If
Rem --------------------------------------------------------------------------------

Rem https://excel-ubara.com/excelvba4/EXCEL295.html
Rem

Rem Shell
Rem ���s�t�@�C���̎w�肪�K�{�ŁA�Ώۃt�@�C���̓p�����[�^�Ƃ��ĕt�^���Ȃ��Ƃ����Ȃ�
Public Function OpenAssociationShell(ByVal ExeFileName As String)
    VBA.Shell ExeFileName, vbNormalFocus
'5
'�v���V�[�W���̌Ăяo����܂��͈������s���ł��
End Function

Rem VBA����̓o�b�`�𐶐����N���A���̃o�b�`�̒��Ńt�@�C�����J�����@
Rem Win10�֘A�t�����s����
Public Function OpenAssociationCmdExe(ByVal FileName As String)
    Dim batFile As String
    FileName = """" & FileName & """"
    batFile = ThisWorkbook.Path & "\vba_temp.bat"
    Open batFile For Output As #1
    Print #1, FileName
    Close #1
    VBA.Shell batFile, vbMinimizedNoFocus
End Function

Rem Shell32
Rem Win10�֘A�t�����s����
Public Function OpenAssociationShell32(ByVal FileName As String)
    Dim Sh As Object 'Shell32.Shell '�Q�Ɛݒ�uMicrosoft Shell Controls And Automation�v
    Set Sh = CreateObject("Shell.Application")
    Sh.ShellExecute FileName
    Set Sh = Nothing
End Function

Rem �w��t�@�C�����֘A�t����ꂽ�A�v���P�[�V�����ŊJ��(API)
Rem
Rem  ��"open"���삪�ݒ肳��Ă��Ȃ��ƃG���[�ɂȂ�
Rem   ret = 31 : �w�肳�ꂽ�t�@�C���g���q�Ɋ֘A�t����ꂽ�A�v���P�[�V����������܂���B
Public Function OpenAssociationAPI(ByVal FileName As String)

    'Application.hwnd���g����̂�Excel2002�ȍ~
#If VBA7 Then
    Dim ret As LongPtr
#Else
    Dim ret As Long
#End If
    ret = ShellExecute(GetActiveWindow(), vbNullString, FileName, vbNullString, vbNullString, SW_SHOW)
'    ret = ShellExecute(GetActiveWindow(), "Open", FileName, vbNullString, vbNullString, SW_SHOW)
    OpenAssociationAPI = CLng(ret) 'SR_ERR
End Function

Rem �w��t�@�C�����֘A�t����ꂽ�A�v���P�[�V�����ŊJ��(WSH����)
Rem
Rem  ��"open"���삪�ݒ肳��Ă��Ȃ��ƃG���[�ɂȂ�
Rem   -2147023741
Rem   Run' ���\�b�h�͎��s���܂���: 'IWshShell3' �I�u�W�F�N�g
Public Function OpenAssociationWSH(ByVal strFileName, Optional strParam = "", Optional nMode = SW_SHOWMAXIMIZED)
    Dim strP As String: strP = strFileName
    If strParam <> "" Then strP = strP & " " & strParam
    OpenAssociationWSH = CreateObject("Wscript.Shell").Run(strP, nMode) 'SW_HIDE SW_SHOWMAXIMIZED
End Function

Rem �w��t�@�C�����֘A�t����ꂽ�A�v���P�[�V�����ŊJ��(Excel�̃n�C�p�[�����N�@�\)
Rem
Rem �K���m�F���b�Z�[�W���o��
Rem
Rem  ��"open"���삪�ݒ肳��Ă��Ȃ��ƃG���[�ɂȂ�
Rem   -2147221018
Rem   ���̃t�@�C�����J�����߂̃v���O�������o�^����Ă��܂���
Public Function OpenAssociationExcelHyperlink(ByVal FileName)
    ThisWorkbook.FollowHyperlink FileName
End Function

Rem �t�@�C���E�t�H���_���G�N�X�v���[���ŊJ��
Rem
Rem  @param full_path   �Ώۃt�@�C���E�t�H���_�̃t���p�X
Rem  @param IsSelected  �I����ԂŊJ����
Rem
Rem  @note
Rem    �I����Ԃɂ���Ƌ����������ɕω�����̂Œ���
Rem
Public Sub ShellExplorer(full_path, Optional IsSelected As Boolean = False)
    If IsSelected Then
        VBA.Shell "explorer " & full_path & ",/select", vbNormalFocus
    Else
        VBA.Shell "explorer " & full_path & "", vbNormalFocus
    End If
End Sub
