VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kccFuncPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccFuncPath
Rem
Rem  @description   �t�@�C���E�t�H���_�E�p�X��͊֐�
Rem
Rem  @update        2020/09/22
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem  @references
Rem    Microsoft Scripting Runtime
Rem    Microsoft VBScript Regular Expressions 5.5
Rem
Rem --------------------------------------------------------------------------------
Rem  @refModules
Rem    �s�v
Rem
Rem --------------------------------------------------------------------------------
Rem  @history
Rem    2019/06/24 ���W���[����������
Rem    2019/09/28 FuncFileList��FuncPath�𓝍���FuncFileFolderPath�Ƃ��čĒ�`
Rem    2019/11/12 SpecialFolders�ǉ�
Rem    2019/12/05 CreateAllFolder���X�V
Rem    2020/02/22 �����֐��ɔėp�t�B���^������ǉ�
Rem    2020/05/10 ���� ModIOStream�Aoutlook_path_hyperlink_unc
Rem
Rem --------------------------------------------------------------------------------
Rem  @functions
Rem
Rem --------------------------------------------------------------------------------
Rem  @note
Rem
Rem --------------------------------------------------------------------------------

Rem --------------------------------------------------------------------------------
Rem
Rem Unicode�Ή��Ńt�@�C�����X�g�쐬�֐�
Rem
Rem --------------------------------------------------------------------------------
Rem
Rem  @history
Rem
Rem 2019/04/24 : ���񃊃��[�X
Rem 2019/04/26 : 4/25�u���O�R�����g�̎w�E�����ɏC��
Rem 2019/04/27 : 64bit�Ή��BAPI��Ex�ɕύX�B�G�N�X�v���[�����\�[�g�ɑΉ��͂ł��Ă��Ȃ��B
Rem
Rem --------------------------------------------------------------------------------
Rem
Rem ���g����
Rem
Rem �ᑬ��FileSystemObject���g�킸�ɁAWindowsAPI�݂̂��g�p���ăt�@�C�����X�g��
Rem �쐬���邽�߂̊֐��ł��B
Rem �����K�v�Ƃ���Ɠ��ȋ@�\�𓋍ڂ��Ă��܂��B
Rem
Rem �ȉ��A���ӎ����ł��B
Rem
Rem 1.parentFolder�͖�����\��t�����t�H���_�p�X���w�肵�Ă��������B
Rem     parentFolder :     -   :  �K�{  : �����Ώۃt�H���_�̖�����\�ŏI���p�X
Rem
Rem     ������\�������Ǝ��s���G���[�𔭐������܂��B
Rem
Rem 2.AddFile��AddFolder���ȗ�����ƁA�����擾����܂���B
Rem     AddFile      :  False  : �ȗ��� : �t�@�C����ΏۂɊ܂߂邩
Rem     AddFolder    :  False  : �ȗ��� : �t�H���_��ΏۂɊ܂߂邩
Rem
Rem     ���Ȃ��Ƃ������������ǂ��炩��True�ɂ��Ă��������B
Rem
Rem 3.SubMin��SubMax���ȗ�����ƁA�����̃��m�����擾���܂���B
Rem     SubMin       :      0  : �ȗ��� : ���K�w�ȍ~��T�����邩�i0�`n�A-1�̎��͖������j
Rem     SubMax       :      0  : �ȗ��� : ���K�w�ȑO��T�����邩�i0�`n�A-1�̎��͖������j
Rem
Rem     parentFolder�Ŏw�肵���p�X�������0�K�w�Ƃ��ăJ�E���g���܂��B
Rem     ����āASubMin�̏ȗ��A0�A-1�͑S�ē��`�ł��B
Rem
Rem   �z���̑S�Ẵt�@�C�����擾�������ꍇ�́A-1,-1�ɂȂ�܂��B
Rem     ���邢�́A0,9999�Ƃ��Ă������I�ɓ������ʂ������܂��B
Rem
Rem   ������l��z���S�Ẵt�@�C���Ƃ���ƁA����Ȏ��Ԃ������鋰�ꂪ���邽�߂ł��B
Rem
Rem 4.�߂�l��parentFolder���猩���y���΃p�X�z�ɂȂ�܂��B
Rem     ��΃p�X��Ԃ��悤�ɂ���ƁA�S�Ẵ��m�ɓ���̕����񂪕t�^����邽�߁A
Rem     �[���K�w�Ō������J�n�������Ƀ������𖳑ʂɏ����̂�h�����߂ł��B
Rem     ���������āA���o�����A�C�e����parentFolder�ƘA�����Ă���g�p���܂��B
Rem
Rem     �܂��A�t�H���_�̖����ɂ͕K��\��t�^������ԂŕԂ��܂��B
Rem     ���p�X�����񂩂�t�@�C���ƃt�H���_�����ʂł���悤�ɂ��邽�߂ł��B
Rem
Rem 5.���я��̓t�@�C�����t�H���_�ł��B
Rem
Rem     �����Ԃ�G�N�X�v���[���ŕ\������鏇���Ƃ͈قȂ�܂��B
Rem     ������A�d�l���ς�鋰�ꂪ����܂��B
Rem
Rem     ��
Rem       A001.txt
Rem       A002.txt
Rem       A01\
Rem       A01\B001.txt
Rem       A01\B002.txt
Rem       A01\B1001\
Rem       A01\B1001\C001.txt
Rem       A01\B1001\C002.txt
Rem       A01\B1001\C2001\
Rem       A01\B1001\C2001\001.txt
Rem       A01\B1001\C2001\002.txt
Rem       A01\B1001\C2002\
Rem       A01\B1001\C2002\001.txt
Rem       A01\B1001\C2002\002.txt
Rem       A01\B1002\
Rem       A01\B1002\C001.txt
Rem       A01\B1002\C002.txt
Rem       A01\B1002\C2001\
Rem       A01\B1002\C2001\001.txt
Rem       A01\B1002\C2001\002.txt
Rem       A01\B1002\C2002\
Rem       A01\B1002\C2002\001.txt
Rem       A01\B1002\C2002\002.txt
Rem
Rem Sub Sample_GetFileList_API()
Rem
Rem      Const SEARCH_PATH = "D:\test\"
Rem
Rem      Dim colPaths As Collection
Rem
Rem      Set colPaths = GetFileFolderList(SEARCH_PATH)
Rem      Debug.Print "�����擾����", colPaths.Count
Rem
Rem      Set colPaths = GetFileFolderList(SEARCH_PATH, True)
Rem      Debug.Print "�w��p�X�̃t�@�C���̂ݎ擾", colPaths.Count
Rem
Rem      Set colPaths = GetFileFolderList(SEARCH_PATH, True, True)
Rem      Debug.Print "�w��p�X�̃t�@�C���ƃt�H���_���擾", colPaths.Count
Rem
Rem      Set colPaths = GetFileFolderList(SEARCH_PATH, True, True, 2, 2)
Rem      Debug.Print "���K�w�̃t�@�C���ƃt�H���_���擾", colPaths.Count
Rem
Rem      Set colPaths = GetFileFolderList(SEARCH_PATH, True, True, 3, -1)
Rem      Debug.Print "��O�K�w�ȉ��̃t�@�C���ƃt�H���_���擾", colPaths.Count
Rem
Rem      Set colPaths = GetFileFolderList(SEARCH_PATH, True, True, -1, -1)
Rem      Debug.Print "�w��p�X�ȉ��̑S�Ẵt�@�C���ƃt�H���_���擾", colPaths.Count
Rem
Rem End Sub
Rem --------------------------------------------------------------------------------

Rem --------------------------------------------------------------------------------
Rem ��Outlook�Ń��[����M�҂����[�J���p�X���N���b�N�ł���悤�ɂ���}�N��2
Rem
Rem   �p�X��UNC�\�L�ɒu�������邱�ƂŃn�C�p�[�����N�������悤�ɂ����
Rem
Rem   �������邿��񂿂��
Rem   2019/10/22
Rem   https://www.excel-chunchun.com/entry/outlook-path-hyperlink-2
Rem
Rem --------------------------------------------------------------------------------

Rem �Q�l����

Rem �l�b�g���[�N�h���C�u����UNC���擾�����
Rem      http://dobon.net/vb/bbs/log3-14/8196.html
Rem      http://blog.livedoor.jp/shingo555jp/archives/1819741.html

Rem WNetGetConnection�ɂ���
Rem
Rem      http://www.pinvoke.net/default.aspx/advapi32/WNetGetUniversalName.html

Rem      Function mpr::WNetGetConnectionW
Rem      https://retep998.github.io/doc/mpr/fn.WNetGetConnectionW.html

Rem      Stack Overflow - Getting An Absolute Image Path
Rem      https://stackoverflow.com/questions/19079162/getting-an-absolute-image-path/19164957

Rem      Passing a LPCTSTR parameter to an API call from VBA in a PTRSAFE and UNICODE safe manner
Rem      https://stackoverflow.com/questions/10402822/passing-a-lpctstr-parameter-to-an-api-call-from-vba-in-a-ptrsafe-and-unicode-saf

Rem API��A��W�̒u�������ɂ���
Rem      RelaxTools - String�^�̒��g�͎����I��S-JIS�ɕϊ�����錏
Rem      https://software.opensquare.net/relaxtools/archives/3400/

Rem      Programming Field - Win32API�̊֐���VB�Ŏg���ɂ́c
Rem      https://www.pg-fl.jp/program/tips/vbw32api.htm

Rem      AddinBox - Tips26: MsgBox / Beep�� �� Unicode������
Rem      http://addinbox.sakura.ne.jp/Excel_Tips26.htm

Option Explicit

Rem WNetGetConnection
Rem ���[�J���f�o�C�X�Ɋ֘A�t����ꂽ�l�b�g���[�N���\�[�X�̖��O���擾���܂��B
#If VBA7 Then
    Private Declare PtrSafe Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionW" ( _
                                            ByVal lpszLocalName As LongPtr, _
                                            ByVal lpszRemoteName As LongPtr, _
                                            cbRemoteName As Long _
                                            ) As Long
#Else
    Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionW" ( _
                                            ByVal lpszLocalName As Long, _
                                            ByVal lpszRemoteName As Long, _
                                            cbRemoteName As Long _
                                            ) As Long
#End If
Rem pub unsafe extern "system" fn WNetGetConnectionW(
Rem      lpLocalName  : LPCWSTR,
Rem      lpRemoteName : LPWSTR,
Rem      lpnLength    : LPDWORD
Rem ) -> DWORD

Rem http://tokovalue.jp/function/WNetGetConnection.htm
Rem
Rem WNetGetConnection
Rem     ���[�J�����u�ɑΉ�����l�b�g���[�N�����̖��O���擾����
Rem
Rem �p�����[�^
Rem lpLocalName
Rem      �l�b�g���[�N�����K�v�ȃ��[�J�����u�̖��O��\�� NULL �ŏI��镶����ւ̃|�C���^���w�肷��B
Rem lpRemoteName
Rem      �ڑ��Ɏg���Ă��郊���[�g����\�� NULL �ŏI��镶������󂯎��o�b�t�@�ւ̃|�C���^���w�肷��B
Rem lpnLength
Rem      lpRemoteName �p�����[�^���w���o�b�t�@�̃T�C�Y�i �������j���������ϐ��ւ̃|�C���^���w�肷��B
Rem
Rem      �o�b�t�@�̃T�C�Y���s�\���Ŋ֐������s�����ꍇ�ͤ�K�v�ȃo�b�t�@�T�C�Y�����̕ϐ��Ɋi�[�����
Rem
Rem �߂�l
Rem      �֐�����������ƤNO_ERROR ���Ԃ�
Rem      �֐������s����Ƥ���̂����ꂩ�̃G���[�R�[�h���Ԃ�
Rem
Rem   �萔                      �Ӗ�
Rem   ERROR_BAD_DEVICE          lpLocalName �p�����[�^���w�������񂪖����ł���B
Rem   ERROR_NOT_CONNECTED       lpLocalName �p�����[�^�Ŏw�肵�����u�����_�C���N�g����Ă��Ȃ��B
Rem   ERROR_MORE_DATA           �o�b�t�@�̃T�C�Y���s�\���ł���B
Rem                             lpnLength �p�����[�^���w���ϐ��ɁA�K�v�ȃo�b�t�@�T�C�Y���i�[����Ă���
Rem                             ���̊֐��Ŏ擾�\�ȃG���g�����c���Ă���
Rem   ERROR_CONNECTION_UNAVAIL  ���u�͌��ݐڑ�����Ă��Ȃ�����P�v�I�Ȑڑ��Ƃ��ċL������Ă���
Rem   ERROR_NO_NETWORK          �l�b�g���[�N�ɂȂ����Ă��Ȃ��
Rem   ERROR_EXTENDED_ERROR      �l�b�g���[�N�ŗL�̃G���[�����������B�G���[�̐������擾����ɂ́AWNetGetLastError �֐����g���
Rem   ERROR_NO_NET_OR_BAD_PATH  �w�肵�����[�J�������g�����ڑ���F������v���o�C�_���Ȃ��
Rem                             ���̐ڑ����g��1�ȏ�̃v���o�C�_�̃l�b�g���[�N�ɂȂ����Ă��Ȃ��\��������

Rem   WNetGetConnection Return Result Constants
Private Const ERROR_SUCCESS As Long = 0&
Private Const ERROR_BAD_DEVICE As Long = 1200&
Private Const ERROR_NOT_CONNECTED = 2250&
Private Const ERROR_MORE_DATA = 234&
Private Const ERROR_CONNECTION_UNAVAIL = 1201&
Private Const ERROR_NO_NETWORK = 1222&
Private Const ERROR_EXTENDED_ERROR = 1208&
Private Const ERROR_NO_NET_OR_BAD_PATH = 1203&

Private Const INVALID_HANDLE_VALUE = -1

Rem FindFirstFileEx�֐����g�p���邩
Rem True�ɂ����ꍇ�ł��A���s�����玩����FindFirstFile�őΉ�����
Private Const USE_FindFirstFileEx = True

Rem --------------------------------------------------------------------------------
Rem Win32API�֐��Q��
Rem
Rem �擪�t�@�C������
#If VBA7 Then
Rem http://chokuto.ifdef.jp/urawaza/api/FindFirstFile.html
Rem https://docs.microsoft.com/ja-jp/windows/desktop/api/fileapi/nf-fileapi-findfirstfilew
Private Declare PtrSafe Function FindFirstFile Lib "Kernel32" Alias "FindFirstFileW" _
            (ByVal lpFileName As LongPtr, _
            lpFindFileData As WIN32_FIND_data1) As LongPtr
            
#Else
Private Declare Function FindFirstFile Lib "Kernel32" Alias "FindFirstFileW" _
            (ByVal lpFileName As Long, _
            lpFindFileData As WIN32_FIND_data1) As Long
#End If

Rem �擪�t�@�C������
#If VBA7 Then
Rem http://tokovalue.jp/function/FindFirstFileEx.htm
Rem https://docs.microsoft.com/ja-jp/windows/desktop/api/fileapi/nf-fileapi-findfirstfileexw
Private Declare PtrSafe Function FindFirstFileEx Lib "Kernel32" Alias "FindFirstFileExW" _
            (ByVal lpFileName As LongPtr, _
            ByVal fInfoLevelId As FINDEX_INFO_LEVELS, _
            lpFindFileData As WIN32_FIND_data1, _
            ByVal fSearchOp As FINDEX_SEARCH_OPS, _
            ByVal lpSearchFilter As LongPtr, _
            ByVal dwAdditionalFlags As Long) As LongPtr
#Else
Private Declare Function FindFirstFileEx Lib "Kernel32" Alias "FindFirstFileExW" _
            (ByVal lpFileName As Long, _
            ByVal fInfoLevelId As FINDEX_INFO_LEVELS, _
            lpFindFileData As WIN32_FIND_data1, _
            ByVal fSearchOp As FINDEX_SEARCH_OPS, _
            ByVal lpSearchFilter As Long, _
            ByVal dwAdditionalFlags As Long) As Long
#End If

Rem FindFirstFileEx�ɂ���
Rem https://blogs.yahoo.co.jp/nobuyuki_tsukasa/1059830.html
Rem https://kkamegawa.hatenablog.jp/entry/20100918/p1
Rem �u8.3�`���̒Z���t�@�C�����𐶐������Ȃ��v���ƂŁA�u81%���炢�ɍ����������v���Ⴊ������

Rem   LPCTSTR lpFileName,�@�@�@�@�@�@�@// ��������t�@�C����
Rem   FINDEX_INFO_LEVELS fInfoLevelId, // �f�[�^�̏�񃌃x��
Rem   LPVOID lpFindFileData,�@�@�@�@�@ // �Ԃ��ꂽ���ւ̃|�C���^
Rem   FINDEX_SEARCH_OPS fSearchOp,�@�@ // ���s����t�B���^�����̃^�C�v
Rem   LPVOID lpSearchFilter,�@�@�@�@�@ // ���������ւ̃|�C���^
Rem   DWORD dwAdditionalFlags�@�@�@�@�@// �⑫�I�Ȍ�������t���O

Rem https://docs.microsoft.com/ja-jp/windows/desktop/api/minwinbase/ne-minwinbase-findex_info_levels
Private Enum FINDEX_INFO_LEVELS
    FindExInfoStandard = 0&
    Rem FindFirstFile �Ɠ�������
    
    FindExInfoBasic = 1&
    Rem WIN32_FIND_DATA��cAlternateFileName�ɒZ���t�@�C�������擾���Ȃ��B
    Rem Windows Server 2008�AWindows Vista�AWindows Server 2003�AWindows XP �ł̓T�|�[�g����Ă��Ȃ��B
    Rem Windows Server 2008 R2 �� Windows 7 �ł͎g�p�\�B
    
    FindExInfoMaxInfoLevel = 2&
    '���̒l�͌��؂Ɏg�p����܂��B �T�|�[�g����Ă���l�͂��̒l�����������ł��B
End Enum

Rem FINDEX_SEARCH_OPS�񋓑�
Rem fSearchOp : ���C���h�J�[�h�Ƃ̏ƍ��ȊO�̃t�B���^�����^�C�v��\��
Private Enum FINDEX_SEARCH_OPS
    FindExSearchNameMatch = 0&
    Rem �w�肵���t�@�C�����ƈ�v����t�@�C�����������܂��B
    Rem ���̌���������g�p����Ƃ��ͤFindFirstFileEx��lpSearchFilter�p�����[�^��NULL�ɂ���K�v������܂��
    
    FindExSearchLimitToDirectories = 1&
    Rem �t�@�C���V�X�e�����f�B���N�g���t�B���^�����O���T�|�[�g���Ă���ꍇ�A�f�B���N�g�����������܂��B
    Rem ���ۂɂ̓T�|�[�g���Ă���t�@�C���V�X�e���͑��݂���?���ʂ��Ȃ��Ƃ̎��B
    Rem https://gist.github.com/kumatti1/33182de4efe99259e275
    Rem http://www.vbalab.net/vbaqa/c-board.cgi?cmd=one;no=58244;id=excel
    
    FindExSearchLimitToDevices = 2&
    Rem ���̃t�B���^�����O�^�C�v�͗��p�ł��܂���B
    
    FindExSearchMaxSearchOp = 3&
    Rem �T�|�[�g����Ă��܂���B
End Enum

Rem dwAdditionalFlags
Private Const FIND_FIRST_EX_CASE_SENSITIVE = 1&
Rem �����ł͑啶���Ə���������ʂ���܂��B

Private Const FIND_FIRST_EX_LARGE_FETCH = 2&
Rem �f�B���N�g���[�Ɖ�ɂ͂��傫�ȃo�b�t�@�[���g�p���܂��B
Rem ����ɂ��A��������̃p�t�H�[�}���X�����シ��\��������܂��B
Rem Windows Server 2008�AWindows Vista�AWindows Server 2003�A�����Windows XP�F
Rem ���̒l�́AWindows Server 2008 R2�����Windows 7�܂ł̓T�|�[�g����Ă��܂���B

Private Const FIND_FIRST_EX_ON_DISK_ENTRIES_ONLY = 4&
Rem ���ʂ𕨗��I�Ƀf�B�X�N��ɂ���t�@�C���ɐ������܂��B
Rem ���̃t���O�́A�t�@�C�����z���t�B���^�����݂���ꍇ�ɂ̂݊֌W���܂��B

Rem ���t�@�C������
Rem http://chokuto.ifdef.jp/urawaza/api/FindNextFile.html
#If VBA7 Then
Private Declare PtrSafe Function FindNextFile Lib "Kernel32" Alias "FindNextFileW" _
            (ByVal hFindFile As LongPtr, lpFindFileData As WIN32_FIND_data1) As LongPtr
#Else
Private Declare Function FindNextFile Lib "Kernel32" Alias "FindNextFileW" _
            (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_data1) As Long
#End If
            
Rem �����n���h���J��
Rem http://chokuto.ifdef.jp/urawaza/api/FindClose.html
#If VBA7 Then
Private Declare PtrSafe Function FindClose Lib "Kernel32" (ByVal hFindFile As LongPtr) As LongPtr
#Else
Private Declare Function FindClose Lib "Kernel32" (ByVal hFindFile As Long) As Long
#End If

Rem --------------------------------------------------------------------------------

Rem FILETIME�\����
Rem http://chokuto.ifdef.jp/urawaza/struct/FILETIME.html
Private Type FILETIME
     LowDateTime As Long
     HighDateTime As Long
End Type

Rem WIN32_FIND_DATA�\����
Rem http://chokuto.ifdef.jp/urawaza/struct/WIN32_FIND_data1.html
Private Type WIN32_FIND_data1
    dwFileAttributes                        As Long     ' �t�@�C������
    ftCreationTime                          As FILETIME ' �쐬��
    ftLastAccessTime                        As FILETIME ' �ŏI�A�N�Z�X��
    ftLastWriteTime                         As FILETIME ' �ŏI�X�V��
    nFileSizeHigh                           As Long     ' �t�@�C���T�C�Y�i��ʂR�Q�r�b�g�j
    nFileSizeLow                            As Long     ' �t�@�C���T�C�Y�i���ʂR�Q�r�b�g�j
    dwReserved0                             As Long     ' �\��ς݁B���p�[�X�^�O
    dwReserved1                             As Long     ' �\��ς݁B���g�p
    cFileName(260 * 2 - 1)                  As Byte     ' �t�@�C����
    cAlternateFileName(14 * 2 - 1)          As Byte     ' 8.3�`���̃t�@�C����
Rem      cFileName                               As String * MAX_PATH    ' �Ƃ������������ł���B
Rem      cAlternateFileName                      As String * 14          ' �Ƃ������������ł���B
End Type
Rem ��Unicode�Ή��̈�*2���Ă���

Rem StrCmpLogicalW�֐�
Rem �G�N�X�v���[���̃t�@�C�����ɕ��ёւ���
Rem https://docs.microsoft.com/en-us/windows/desktop/api/shlwapi/nf-shlwapi-strcmplogicalw
#If VBA7 Then
Private Declare PtrSafe Function StrCmpLogicalW Lib "shlwapi" _
                (ByVal lpStr1 As String, ByVal lpStr2 As String) As Long
#Else
Private Declare Function StrCmpLogicalW Lib "shlwapi" _
                (ByVal lpStr1 As String, ByVal lpStr2 As String) As Long
#End If

Rem wsh.SpecialFolders�v���p�e�B
Private Const SpecialFolderKey_AllUsersDesktop = "AllUsersDesktop"
Private Const SpecialFolderKey_AllUsersStartMenu = "AllUsersStartMenu"
Private Const SpecialFolderKey_AllUsersPrograms = "AllUsersPrograms"
Private Const SpecialFolderKey_AllUsersStartup = "AllUsersStartup"
Private Const SpecialFolderKey_Desktop = "Desktop"
Private Const SpecialFolderKey_Favorites = "Favorites"
Private Const SpecialFolderKey_Fonts = "Fonts"
Private Const SpecialFolderKey_MyDocuments = "MyDocuments"
Private Const SpecialFolderKey_NetHood = "NetHood"
Private Const SpecialFolderKey_PrintHood = "PrintHood"
Private Const SpecialFolderKey_Programs = "Programs"
Private Const SpecialFolderKey_Recent = "Recent"
Private Const SpecialFolderKey_SendTo = "SendTo"
Private Const SpecialFolderKey_StartMenu = "StartMenu"
Private Const SpecialFolderKey_Startup = "Startup"
Private Const SpecialFolderKey_Templates = "Templates"
Rem     ����t�H���_��     ����    ��ʓIWindows10�̋�̓I�ȃt�H���_
Rem 1   AllUsersDesktop    ���ׂẴ��[�U�[�ɋ��ʂ̃f�X�N�g�b�v        C:\Users\Public\Desktop
Rem 2   AllUsersStartMenu  ���ׂẴ��[�U�[�ɋ��ʂ̃v���O�������j���[  C:\ProgramData\Microsoft\Windows\Start Menu
Rem 3   AllUsersPrograms   ���ׂẴ��[�U�[�ɋ��ʂ̑S�Ẵv���O����    C:\ProgramData\Microsoft\Windows\Start Menu\Programs
Rem 4   AllUsersStartup    ���ׂẴ��[�U�[�ɋ��ʂ̃X�^�[�g�A�b�v      C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp
Rem 5   Desktop            �f�X�N�g�b�v                            C:\Users\[username]\Desktop
Rem 6   Favorites          ���C�ɓ���                              C:\Users\[username]\Favorites
Rem 7   Fonts              �C���X�g�[������Ă���t�H���g          C:\Windows\Fonts
Rem 8   MyDocuments        �}�C�h�L�������g                        C:\Users\[username]\Documents
Rem 9   NetHood            �l�b�g���[�N�ɕ\������鋤�L�t�H���_�̏��  C:\Users\[username]\AppData\Roaming\Microsoft\Windows\Network Shortcuts
Rem 10  PrintHood          �v�����^�t�H���_                        C:\Users\[username]\AppData\Roaming\Microsoft\Windows\Printer Shortcuts
Rem 11  Programs           ���O�C�����[�U�[�̃v���O�������j���[    C:\Users\[username]\AppData\Roaming\Microsoft\Windows\Start Menu\Programs
Rem 12  Recent             �ŋߎg�����t�@�C��                      C:\Users\[username]\AppData\Roaming\Microsoft\Windows\Recent
Rem 13  SendTo             ���郁�j���[                            C:\Users\[username]\AppData\Roaming\Microsoft\Windows\SendTo
Rem 14  StartMenu          �X�^�[�g���j���[                        C:\Users\[username]\AppData\Roaming\Microsoft\Windows\Start Menu
Rem 15  Startup            ���O�C�����[�U�[�̃X�^�[�g�A�b�v        C:\Users\[username]\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup
Rem 16  Templates          �V�K�쐬�̃e���v���[�g                  C:\Users\[username]\AppData\Roaming\Microsoft\Windows\Templates

Rem SetCurrentDirectory API �g�p����
Private Declare PtrSafe Function SetCurrentDirectory _
    Lib "Kernel32" Alias "SetCurrentDirectoryA" _
    (ByVal lpPathName As String) As Long
    
#If VBA7 Then
    Private Declare PtrSafe Function SHCreateDirectoryEx Lib "Shell32" Alias "SHCreateDirectoryExW" ( _
        ByVal hWnd As LongPtr, _
        ByVal pszPath As LongPtr, _
        ByVal psa As LongPtr) As LongPtr
        
Rem     Private Declare PtrSafe Function SHCreateDirectoryExA Lib "shell32" ( _
Rem         ByVal hwnd As LongPtr, _
Rem         ByVal pszPath As String, _
Rem         ByVal psa As LongPtr) As LongPtr
#Else
    Private Declare Function SHCreateDirectoryEx Lib "Shell32" Alias "SHCreateDirectoryExW" ( _
        ByVal hWnd As Long, _
        ByVal pszPath As Long, _
        ByVal psa As Long) As Long
#End If

Rem SHCreateDirectoryEx �̖߂�l
Const ERROR_BAD_PATHNAME = 161&         '�w�肳�ꂽ�p�X�������ł��B
Const ERROR_FILENAME_EXCED_RANGE = 206& '�t�@�C�����܂��͊g���q���������܂��B
Const ERROR_PATH_NOT_FOUND = 3&         '�w�肳�ꂽ�p�X��������܂���B
Const ERROR_FILE_EXISTS = 80&           '�f�B���N�g���͑��݂���B
Const ERROR_ALREADY_EXISTS = 183&       '�f�B���N�g���͑��݂���B
Const ERROR_CANCELLED = 0&              '���[�U�[�͑�������������B
Const ERROR_ACCESS_DENIED = 5&          '�A�N�Z�X�����ۂ���܂����B

Rem �G���[�R�[�h�\
Rem https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes

Rem --------------------------------------------------------------------------------
Rem ���ʑg�ݍ���
Private Property Get fso() As FileSystemObject
    Static xxFso As Object  'FileSystemObject
    If xxFso Is Nothing Then Set xxFso = CreateObject("Scripting.FileSystemObject")
    Set fso = xxFso
End Property
Rem --------------------------------------------------------------------------------

Rem �w�肵���p�X�̃t�H���_����C�ɍ쐬����
Rem ���s����������False��Ԃ��B���ɑ��݂����ꍇ�͖�����OK
Rem
Rem  @param folder_path �쐬�������t�H���_
Rem
Rem  @return As Boolen  �����������ǂ���
Rem                      �쐬�ɐ��� : True
Rem                      ���ɑ���   : True
Rem                      �쐬�Ɏ��s : False
Rem
Public Function CreateDirectoryEx(folder_path As String, Optional ByRef errValue) As Boolean
    errValue = SHCreateDirectoryEx(0&, StrPtr(SupportMaxPath260over(folder_path)), 0&)
    Select Case errValue
        Case 0:  CreateDirectoryEx = True '����
        Case 183: CreateDirectoryEx = True '���ɑ���
        Case Else: CreateDirectoryEx = False '���s
    End Select
End Function

Rem Public Function CreateDirectoryExA(folder_path As String) As Boolean
Rem     Select Case SHCreateDirectoryExA(0&, folder_path, 0&)
Rem         Case 0:  CreateDirectoryExA = True '����
Rem         Case 183: CreateDirectoryExA = True '���ɑ���
Rem         Case Else: CreateDirectoryExA = False '���s
Rem     End Select
Rem End Function

Rem Win32API��W�t���֐��ɂ����āA260 (MAX_PATH) �����������������ɑΉ������邽�߂̏���
Rem
Rem  @param file_folder_path �t�@�C�����t�H���_�̃p�X
Rem
Rem  @return As String       �ϊ���̃p�X
Rem
Rem  @note
Rem    �p�X�̐擪�� "\\?\"��"\\?\UNC" ��ǉ����Ă���
Rem
Rem  @example
Rem    \\SERVERNAME\    >>  \\?\UNC\SERVERNAME\
Rem    C:\DRIVE         >>  \\?\C:\DRIVE
Rem
Public Function SupportMaxPath260over(ByRef file_folder_path As String) As String
    
    '�Ώ��ς�
    If file_folder_path Like "\\?\*" Then
        SupportMaxPath260over = file_folder_path
        
    '�l�b�g���[�N�p�X
    ElseIf file_folder_path Like "\\*" Then
        SupportMaxPath260over = "\\?\UNC" & Mid$(file_folder_path, 2)
        
    '�W���h���C�u�p�X
    Else
        SupportMaxPath260over = "\\?\" & file_folder_path
    End If
End Function

Rem *******************************************************************************
Rem �w��t�H���_�ȉ��̔C�ӂ̊K�w�̃t�@�C���E�t�H���_�����X�g�A�b�v����֐�
Rem *******************************************************************************
Rem �p�����[�^           : ����l  :  �T�v  : �Ӗ�
Rem parent_folder_path   :     -   :  �K�{  : �����Ώۃt�H���_�̖�����\�ŏI���p�X
Rem add_files            :  False  : �ȗ��� : �t�@�C����ΏۂɊ܂߂邩
Rem add_folders          :  False  : �ȗ��� : �t�H���_��ΏۂɊ܂߂邩
Rem search_min_layer     :      0  : �ȗ��� : ���K�w�ȍ~��T�����邩�i0�`n�A-1�̎��͖������j
Rem search_max_layer     :      0  : �ȗ��� : ���K�w�ȑO��T�����邩�i0�`n�A-1�̎��͖������j
Rem filter_obj           : Missing : �ȗ��� : �t�B���^(RegExp,LIKE�p������,Everything�����d�l�j
Rem recursive_subfolder  :     ""  : �ċA�p : �����̃��[�g�t�H���_�ȍ~�̃p�X
Rem recursive_now_layer  :      0  : �ċA�p : ���݉��K�w�ڂ�
Rem recursive_path_list  : Nothing : �ċA�p : �p�X�ꗗ�B�ŏI�I�Ȗ߂�l�ɂ��g����
Public Function GetFileFolderList(ByVal parent_folder_path As String, _
                                    Optional ByVal add_files = False, _
                                    Optional ByVal add_folders = False, _
                                    Optional ByVal search_min_layer As Long = 0, _
                                    Optional ByVal search_max_layer As Long = 0, _
                                    Optional ByVal filter_obj As Variant, _
                                    Optional ByVal recursive_subfolder As String = "", _
                                    Optional ByVal recursive_now_layer As Long = 0, _
                                    Optional ByRef recursive_path_list As Collection = Nothing _
                                    ) As Collection
    Const PROC_NAME = "GetFileFolderList"
    
    Rem �֐��˓����̏���������
    If recursive_path_list Is Nothing Then
        Set recursive_path_list = New Collection
        
        If Len(parent_folder_path) > 0 Then
            If Right(parent_folder_path, 1) <> "\" Then
                Err.Raise 9999, PROC_NAME, "�t�H���_�p�X�̖�����\�ŏI���悤�ɂ��Ă��������B"
            End If
        End If
    End If
    
    Dim ResFolder As Collection: Set ResFolder = New Collection
    Dim ResFile As Collection: Set ResFile = New Collection
    Dim findData As WIN32_FIND_data1
    
    Dim UnicodeFolderPath As String
    UnicodeFolderPath = SupportMaxPath260over(parent_folder_path)
    
    Rem �����n���h����������Ȃ��ꍇ�́uINVALID_HANDLE_VALUE�v��Ԃ�
#If VBA7 Then
    Dim FileHandle As LongPtr
#Else
    Dim FileHandle As Long
#End If
    If USE_FindFirstFileEx Then
        Rem ��FindExInfoBasic/FIND_FIRST_EX_LARGE_FETCH�w��ɂ�荂����������
        FileHandle = FindFirstFileEx(StrPtr(UnicodeFolderPath & "*"), FindExInfoBasic, _
                                findData, FindExSearchNameMatch, 0&, FIND_FIRST_EX_LARGE_FETCH)
    End If
    
    If Not USE_FindFirstFileEx Or FileHandle = INVALID_HANDLE_VALUE Then
        FileHandle = FindFirstFile(StrPtr(UnicodeFolderPath & "*"), findData)
    End If
    
    If FileHandle = INVALID_HANDLE_VALUE Then
        Exit Function
    End If
    
    Do
        Rem FindFirstFile�ł̓t�@�C�����̌��ɁuMax_Path�v�Ŏw�肵���������܂�Null���l�܂��Ă���B
        Dim intStLen As Long
        intStLen = InStr(findData.cFileName, vbNullChar) - 1
        If intStLen > 0 Then
            Dim sFilename As String
            sFilename = Trim$(Left$(findData.cFileName, intStLen))
            
            Rem �J�����g�t�H���_�ȊO����ʃt�H���_
            If sFilename = "." Or sFilename = ".." Then
                
            Rem �t�H���_�\��
            ElseIf findData.dwFileAttributes And vbDirectory Then
                ResFolder.Add sFilename
                
            Rem �t�@�C���\��
            Else
                If add_files And _
                        (search_min_layer = -1 Or search_min_layer <= recursive_now_layer) And _
                        IsMatchPathFilter(filter_obj, folder_path:=recursive_subfolder, file_name:=sFilename) Then
                    Rem ���t�@�C���̓��[�g�p�X�����������΃p�X
                    ResFile.Add recursive_subfolder & sFilename
                End If
            End If
        End If
        Rem ���̃t�@�C����������Ȃ������ꍇ��0��Ԃ����߃��[�v�I��
    Loop Until FindNextFile(FileHandle, findData) = 0
    
    Rem �����n���h�������
    FindClose FileHandle
    
    Rem �t�@�C�����X�g���\�[�g���Ă���ǉ�
Rem     CollectionSort_StrCmpLogicalW ResFile
    Dim myFile As Variant
    For Each myFile In ResFile
        recursive_path_list.Add myFile
    Next
    
    Rem �t�H���_���X�g���\�[�g���Ă���ǉ����āA�ċA�T����
Rem     CollectionSort_StrCmpLogicalW ResFolder
    
    Dim myFolder As Variant
    For Each myFolder In ResFolder
        Rem �t�H���_�ǉ�
        If add_folders Then
            Rem ���t�H���_�̓��[�g�p�X�����������΃p�X�Ŗ����� "\"
            recursive_path_list.Add recursive_subfolder & myFolder & "\"
        End If
        Rem �T�u�t�H���_�ċA�T��
        If recursive_now_layer < search_max_layer Or search_max_layer = -1 Then
            Call GetFileFolderList( _
                parent_folder_path & myFolder & "\", _
                add_files:=add_files, _
                add_folders:=add_folders, _
                filter_obj:=filter_obj, _
                search_min_layer:=search_min_layer, _
                search_max_layer:=search_max_layer, _
                recursive_subfolder:=recursive_subfolder & myFolder & "\", _
                recursive_now_layer:=recursive_now_layer + 1, _
                recursive_path_list:=recursive_path_list)
        End If
    Next
    
    Set GetFileFolderList = recursive_path_list
    
End Function

Rem �t�@�C���E�t�H���_�t�B���^�����O�p�̌��ؗp�֐�
Rem ���K�\���AEverything�������d�l�AVBA��LIKE���Z�q���g����B
Public Function IsMatchPathFilter( _
        filter_obj As Variant, _
        Optional FullPath As String, _
        Optional folder_path As String, _
        Optional file_name As String, _
        Optional file_basename As String, _
        Optional file_extension As String) As Boolean
    
    If IsMissing(filter_obj) Then IsMatchPathFilter = True: Exit Function
    If file_name = "" Then file_name = file_basename & file_extension
    If FullPath = "" Then FullPath = folder_path & file_name
    
    '���K�\��
    If TypeName(filter_obj) = "RegExp" Then
        Dim reg As Object 'RegExp
        Set reg = filter_obj
        IsMatchPathFilter = reg.Execute(FullPath)
        
    '������w��
    ElseIf TypeName(filter_obj) = "String" Then
        'LIKE���Z�q
        If VBA.Strings.InStr(filter_obj, ":") = 0 Then
            IsMatchPathFilter = (FullPath Like filter_obj)
        
        'Everything��
        Else
            '������
            Stop
        End If
    End If
End Function

Rem 'Colection����ւ�
Rem Private Sub CollectionSwap(C As Collection, Index1 As Long, Index2 As Long)
Rem     Dim Item1 As Variant, Item2 As Variant
Rem     Item1 = C.Item(Index1)
Rem     Item2 = C.Item(Index2)
Rem
Rem     C.Add Item1, After:=Index2
Rem     C.Remove Index2
Rem     C.Add Item2, After:=Index1
Rem     C.Remove Index1
Rem End Sub
Rem
Rem 'Collection��StrCmpLogicalW�Ń\�[�g
Rem Private Sub CollectionSort_StrCmpLogicalW(C As Collection)
Rem     Dim i As Long, j As Long
Rem     For i = 1 To C.Count
Rem         For j = C.Count To i Step -1
Rem             If StrCmpLogicalW(StrConv(C(i), vbUnicode), _
Rem                               StrConv(C(j), vbUnicode)) > 0 Then
Rem                 CollectionSwap C, i, j
Rem             End If
Rem         Next
Rem     Next
Rem End Sub

Rem �ꎞ�t�@�C���̃t���p�X���擾
Public Function GetPathByTemporaryFile() As String
    GetPathByTemporaryFile = GetPathTemporary & "\" & fso.GetTempName
End Function

Rem --------------------------------------------------------------------------------
Rem   �t�H���_�̈ꊇ�쐬
Rem --------------------------------------------------------------------------------
Public Sub CreateAllFolder(ByVal strPath As String, Optional without_lastfilename As Boolean = False)

    Dim s, v, f
    Dim i As Long
    
    v = Split(strPath, "\")

    On Error Resume Next
    For i = LBound(v) To UBound(v)
        If without_lastfilename And i = UBound(v) Then Exit For
    
        If f = "" Then
            f = v(i)
            fso.CreateFolder f & "\"
        Else
            f = f & "\" & v(i)
            fso.CreateFolder f
        End If
    
    Next

End Sub

Function GetPathWSH(WSH_SpecialFolders_Keyword) As String
    On Error Resume Next
    GetPathWSH = CreateObject("Wscript.Shell").SpecialFolders(WSH_SpecialFolders_Keyword)
End Function

Rem �h�L�������g�t�H���_
Public Function GetPathMyDocument() As String: GetPathMyDocument = GetPathWSH("MyDocuments"): End Function
Rem AppData�t�H���_
Public Function GetPathAppData() As String: GetPathAppData = GetPathWSH("AppData"): End Function
Rem �f�X�N�g�b�v�t�H���_
Public Function GetPathDesktop() As String: GetPathDesktop = GetPathWSH("Desktop"): End Function

Rem �e���|��\�ꎞ�t�@�C�����t�H���_
Public Function GetPathTemporary() As String: GetPathTemporary = fso.GetSpecialFolder(TemporaryFolder): End Function

Rem �A�v�����̃T�u�t�H���_�𐶐����ă��b�v���ĕԂ�
Public Function GetAppPath(SpecialFolders_Keyword, ProjectFolderName) As String
    If VBA.IsMissing(ProjectFolderName) Then ProjectFolderName = ""
    If ProjectFolderName = "" Then ProjectFolderName = ThisWorkbook.Name
    
    GetAppPath = ""
    With CreateObject("Scripting.FileSystemObject")
        Dim strFolder As String
        strFolder = .BuildPath(GetPathAppData, ProjectFolderName)
        If .FolderExists(strFolder) Then
        Else
            On Error Resume Next
                .CreateFolder strFolder
            On Error GoTo 0
        End If
        GetAppPath = .BuildPath(strFolder, "\")
    End With

End Function

Rem AppData�t�H���_
Public Function GetAppPathAppData(Optional ProjectFolderName) As String: GetAppPathAppData = GetAppPath("AppData", ProjectFolderName): End Function

Rem �e���|�����t�H���_�擾
Rem
Rem  @return C:\Users\%USERNAME%\AppData\Local\Temp
Rem
Public Function GetAppPathTemporary() As String
    GetAppPathTemporary = ""
    With CreateObject("Scripting.FileSystemObject")
        Dim strFolder As String
        strFolder = GetPathTemporary() & "\Temp"
        If .FolderExists(strFolder) Then
        Else
            On Error Resume Next
                .CreateFolder strFolder
            On Error GoTo 0
        End If
        GetAppPathTemporary = .BuildPath(strFolder, "\")
    End With
End Function

Rem �e���|�����t�H���_�擾
Public Function CreateTempFolder(SpecialFolderKey As String, Optional folder_name_format As String = "yyyymmdd_hhmmss") As String
    CreateTempFolder = CreateObject("Wscript.Shell").SpecialFolders(CVar(SpecialFolderKey)) & "\" & Format(Now, folder_name_format)
    On Error Resume Next
    If fso.CreateFolder(CreateTempFolder) Then
        If Err Then Debug.Print "ERROR CreateTempFolder : " & Err.Description
    End If
    CreateTempFolder = CreateTempFolder & "\"
End Function

Rem �J�����g�f�B���N�g���̕ύX�@�\�@�l�b�g���[�N�p�X���J�����g�f�B���N�g���ɂ���
Rem �@ChDir�@CurDir�@�p�X�ύX�@���݂̃t�H���_
Sub SetCurrentDirectory_WScriptShell(new_path)
    CreateObject("WScript.Shell").CurrentDirectory = new_path
End Sub

Rem �l�̌ܓ��@�\�@���l��C�ӂ̗L�������Ɏl�̌ܓ�����
Rem   Round ���[�N�V�[�g�֐�
Public Function SignificantFigures(Number, l) As Double
    '���l��L������L���Ɏl�̌ܓ�����
    If Number = 0 Then
        SignificantFigures = 0
    Else
        SignificantFigures = Application.Round(Number, -Int(Application.Log(Abs(Number))) - 1 + l)
    End If
End Function

Rem �w�肵���t�@�C�������b�N����Ă��邩�`�F�b�N����B
Public Function GetFileLock(FileName As String) As Boolean
    On Error Resume Next
    Dim fn: fn = FreeFile
    Open FileName For Append As #fn
    Close #fn
    GetFileLock = (Err.Number > 0)
End Function

Rem �w�肵���t�@�C�����ǂݎ���p���`�F�b�N����B
Public Function GetFileReadonly(FileName As String) As Boolean
    'Readonly�������̔��f�̓R��
    GetFileReadonly = (GetAttr(FileName) And vbReadOnly)
End Function

Rem �l�b�g���[�N�h���C�u��UNC�p�X���擾
Rem
Rem  @param nDriveLetter     �h���C�u���^�[������i"A:"��"Z:"�j
Rem
Rem  @return As String       ������
Rem
Rem  @note �T�[�o�[�ɃA�N�Z�X�ł��邩�ۂ��͍l�����Ȃ��B
Rem
Public Function GetUNCPath(ByVal nDriveLetter As String, Optional ByVal bufLen As Long = 64) As String
    Dim UncPath As String: UncPath = String(bufLen, vbNullChar)
    Dim ret As Long
#If VBA7 Then
    ret = WNetGetConnection(StrPtr(nDriveLetter), StrPtr(UncPath), bufLen)
#Else
    ret = WNetGetConnection(nDriveLetter, UncPath, bufLen)
#End If
    Select Case ret
        Case ERROR_SUCCESS: GetUNCPath = Left(UncPath, InStr(UncPath, vbNullChar) - 1)
        Case ERROR_MORE_DATA: GetUNCPath = GetUNCPath(nDriveLetter, bufLen)
        Case Else: GetUNCPath = "GetUNCPath Error : " & ret
    End Select
End Function

Rem �����ς̃h���C�u���^�[��UNC�p�X�̃��X�g��Dictionary�ŕԂ��֐�
Rem
Rem  @return As Dictionary     dic(�h���C�u���^�[) = UNC
Rem
Public Function GetNetworkDriveAndUncByAllocated() As Object
    Dim DicDrives '  As Dictionary
    Set DicDrives = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim UncPath As String
    For i = Asc("A") To Asc("Z")
        UncPath = GetUNCPath(Chr(i) & ":")
        If UncPath <> "" Then DicDrives.Add Chr(i) & ":", UncPath
    Next
    Set GetNetworkDriveAndUncByAllocated = DicDrives
End Function

Rem �ڑ��ς̃h���C�u���^�[��UNC�p�X�̃��X�g��Dictionary�ŕԂ��֐�(WSH�o�[�W����)
Rem
Rem  @return As Dictionary     dic(�h���C�u���^�[) = UNC
Rem
Rem  @note �ڑ��Ϗ�Ԃ̃h���C�u�������o�ł��Ȃ�����
Rem         �S�Ẵh���C�u��񋓂��邱�Ƃ͂ł��Ȃ��B
Rem
Public Function GetNetworkDriveAndUncByConnected() As Object ' As Dictionary
    Dim DicDrives ' As Dictionary
    Set DicDrives = CreateObject("Scripting.Dictionary")
    
    Dim Network ' As WScript.Network
    Set Network = CreateObject("WScript.Network")
    
    Rem Network.EnumNetworkDrives
    Rem   (0):=�h���C�u���^�[1, (1):=UNC�p�X1
    Rem   (2):=�h���C�u���^�[2, (3):=UNC�p�X2
    Rem   (4):=�h���C�u���^�[3, (5):=UNC�p�X3
    Dim Drives  ' As IWshCollection
    Set Drives = Network.EnumNetworkDrives
    
    Dim i As Long
    For i = 0 To Drives.Count - 1 Step 2
        If Drives.Item(i) <> "" Then
            DicDrives.Add Drives.Item(i), Drives.Item(i + 1)
        End If
    Next
    Set GetNetworkDriveAndUncByConnected = DicDrives
End Function

Rem �ڑ��ς̃h���C�u���^�[��UNC�p�X�̃��X�g��Dictionary�ŕԂ��֐�(WMI�o�[�W����)
Rem
Rem  @return As Dictionary     dic(�h���C�u���^�[) = UNC
Rem
Rem  @note �ڑ��Ϗ�Ԃ̃h���C�u�������o�ł��Ȃ�����
Rem         �S�Ẵh���C�u��񋓂��邱�Ƃ͂ł��Ȃ��B
Rem
Public Function GetNetworkDriveAndUncByConnectedWMI() As Object
    Const WQL = _
        "SELECT Name, ProviderName " & _
        "FROM Win32_LogicalDisk " & _
        "WHERE DriveType = 4"
        
    Dim Locator As Object 'WbemScripting.SWbemLocator
    Set Locator = VBA.Interaction.CreateObject("WbemScripting.SWbemLocator")
    
    Dim NetworkDrives As Object 'WbemScripting.SWbemObjectSet
    Set NetworkDrives = Locator.ConnectServer().ExecQuery(WQL)
    
    Dim driveDic As Dictionary
    Set driveDic = VBA.Interaction.CreateObject("Scripting.Dictionary")
    
    Dim drv As Object 'WbemScripting.SWbemObject
    For Each drv In NetworkDrives
        With drv.Properties_
            driveDic.Add .Item("Name").Value, .Item("ProviderName").Value
        End With
    Next
    
    Set GetNetworkDriveAndUncByConnectedWMI = driveDic
End Function

Rem �R�}���h�v�����v�g�Ŏ擾����Q�l����
Rem
Rem C:\Users\USERNAME>net use
Rem �V�����ڑ��͋L������܂��
Rem
Rem �X�e�[�^�X  ���[�J���� �����[�g��                �l�b�g���[�N��
Rem
Rem --------------------------------------------------------------------------------
Rem ���p�s��     V:        \\192.168.11.1\Share      Microsoft Windows Network
Rem ���p�s��     W:        \\landisk\disk            Microsoft Windows Network
Rem OK           X:        \\servername-nuc\Downloads
Rem                                                   Microsoft Windows Network
Rem OK           Y:        \\servername-nuc\Server   Microsoft Windows Network
Rem ���p�s��     Z:        \\crib35nas\Share         Microsoft Windows Network
Rem
Rem �E�ڑ��ς݈ȊO���S�ė񋓂����B
Rem �E�����[�g���̕������������Ɖ��s����ďo�͂����B

Rem --------------------------------------------------------------------------------
Rem ��kccFuncString
Rem   ������ϊ��֐�
Rem --------------------------------------------------------------------------------
Rem
Rem ����
Rem
Rem --------------------------------------------------------------------------------

Rem ���͒��̃p�X�Ǝv���镶������n�C�p�[�����N�ɑΉ�������֐�
Rem
Rem  @param base_str        �ϊ���������
Rem  @param DoNetDriveToUNC �l�b�g���[�N�h���C�u��UNC�ɕϊ����邩�ۂ�
Rem                          False:=�ϊ����Ȃ�(����)
Rem                          True :=�ϊ�����
Rem
Rem  @return  As string     Outlook���n�C�p�[�����N���\�ȕ�����
Rem
Rem  @example
Rem     IN :
Rem          ���L�̃t�@�C�����䗗��������
Rem          C:\Test\hoge.xls
Rem          Z:\fuga.xls
Rem          �ȏ�
Rem
Rem    OUT :
Rem       DoNetDriveToUNC:=False
Rem          ���L�̃t�@�C�����䗗��������
Rem          <"file://C:\Test\hoge.xls">
Rem          <"file://Z:\Test\hoge.xls">
Rem          �ȏ�
Rem
Rem       DoNetDriveToUNC:=True
Rem          <"\\server\share\fuga.xls">
Rem
Rem  @note
Rem         (True�Ȃ�)�l�b�g���[�N�h���C�u�̃p�X��UNC�ɕύX���邱�ƂŃn�C�p�[�����N��
Rem         ���[�J���h���C�u�̃p�X�� <"file:// "> �ň͂����ƂŃn�C�p�[�����N��
Rem         UNC�p�X�� <" "> �ň͂����Ƃœr�؂�h�~
Rem
Rem         �p�X�͕K�����s�ŏI��邱��
Rem         Outlook�ł̓��[�����M���̎����ܕԂ���؂��Ă�������
Rem         ���[���쐬��ʂł̓����N��Ԃɂ͂Ȃ�Ȃ��B
Rem         �������玩���֑��M���ăe�X�g����悤�ɁB
Rem
Public Function ReplacePathToHyperlink(ByVal base_str, Optional DoNetDriveToUNC As Boolean = False) As String
    Const LocalPrefix = "file://"
    
    Dim pathIdx: pathIdx = 1
    Dim lfIdx: lfIdx = 1
    Dim pathData
    Dim v
    Dim i As Long
    Dim s As String

    Dim pathHeader As String
    Dim dicUncPath  As Object 'Dictionary
    Dim DriveLetter As String

    '���s(CRLF)���p�X�I���Ƃ݂Ȃ�
    Dim base_str_arr
    base_str_arr = Split(base_str, vbCrLf)

    'UNC�p�X�̕ϊ�
    Const UncPathPrefix = "\\"
    For i = LBound(base_str_arr) To UBound(base_str_arr)
        s = base_str_arr(i)

        'UNC�p�X��<"UNC�p�X">�ɕϊ�
        pathIdx = InStr(lfIdx, s, UncPathPrefix)
        If pathIdx > 0 Then
            pathData = Mid(s, pathIdx, Len(s))
            s = Replace(s, pathData, "<""" & pathHeader & pathData & """>")
            base_str_arr(i) = s
        End If
    Next

    '�h���C�u���^�[�t���p�X�̕ϊ�
    Dim pathArr(1 To 26)
    For i = 1 To 26: pathArr(i) = Chr(Asc("A") - 1 + i) & ":": Next

    For i = LBound(base_str_arr) To UBound(base_str_arr)
        s = base_str_arr(i)

        '�p�X�Ǝv���镶�͂�����
        For Each v In pathArr
            pathIdx = InStr(lfIdx, s, LocalPrefix & v)
            DriveLetter = v
            If pathIdx > 0 Then Exit For
        Next
        If pathIdx <= 0 Then
            For Each v In pathArr
                pathIdx = InStr(lfIdx, s, v)
                DriveLetter = v
                If pathIdx > 0 Then Exit For
            Next
            pathHeader = LocalPrefix
        Else
            pathHeader = ""
        End If
        
        If pathIdx > 0 Then
            Dim UncPath As String
            UncPath = GetUNCPath(DriveLetter)
            
            If UncPath <> "" And DoNetDriveToUNC Then
                '�l�b�g���[�N�h���C�u�̃p�X��<"\\ServerName\ShareName\�p�X">�ɕϊ�(������file://�͏���)
                pathData = Mid(s, pathIdx, Len(s))
                s = Replace(s, pathData, "<""" & Replace(pathData, DriveLetter, UncPath) & """>")
                s = Replace(s, LocalPrefix, "")
                base_str_arr(i) = s
            Else
                '���[�J���h���C�u�̃p�X��<"file://�p�X">�ɕϊ�
                pathData = Mid(s, pathIdx, Len(s))
                s = Replace(s, pathData, "<""" & pathHeader & pathData & """>")
                base_str_arr(i) = s
            End If
        End If
    Next

    '���ɕt�^����Ă����ꍇ�̓�d�t�^������
    For i = LBound(base_str_arr) To UBound(base_str_arr)
        s = base_str_arr(i)
        s = Replace(s, "<""<""", "<""")
        s = Replace(s, """>"">", """>")
        s = Replace(s, """<""", "<""")
        s = Replace(s, """"">", """>")
        base_str_arr(i) = s
    Next

    ReplacePathToHyperlink = Join(base_str_arr, vbCrLf)
End Function

#If DEF_OUTLOOK Then
Sub ���[���쐬��ʂ̃p�X���n�C�p�[�����N�ɕϊ�()
    Dim objItem As Outlook.MailItem
    Set objItem = ActiveInspector.CurrentItem
    objItem.body = ReplacePathToHyperlink(objItem.body)
End Sub
#End If

Rem \\�ł��n�C�p�[�����N�ɂȂ邪�Afile://����Ȃ��ƃ����N�͖���������

Rem ���łɎ�M���[���̎������s���C��������
#If NO_COMPILE Then
C:\Test\hoge.xls
#End If


Rem �Q�l�����@���̂��炢�̃R�[�h�͒��ڂ������ق���������₷��
Rem '�t�@�C�����J���_�C�A���O��\�����āA�p�X��Ԃ��B�i����EXCEL�Ή��j
Public Function OpenDialog(Path As String, Filter As String) As String
    OpenDialog = ""
    Dim fileToOpen As Variant
    If Path <> "" Then SetCurrentDirectory Path
    fileToOpen = Application _
        .GetOpenFileName(Filter)  '"�G�N�Z���t�@�C��(*.xls;*.xlsx), *.xls;*.xlsx"
    If fileToOpen <> False Then
        OpenDialog = fileToOpen
    End If
    Path = OpenDialog
End Function

Rem �ۑ��_�C�A���O��\�����āA�p�X��Ԃ��B�i����EXCEL�Ή��j
Public Function SaveDialog(Path As String, Filter As String) As String
    SaveDialog = ""
    Dim fileToSave As Variant
    fileToSave = Application _
        .GetSaveAsFilename(Path, Filter)  '"�G�N�Z���t�@�C��(*.xls;*.xlsx), *.xls;*.xlsx"
    If fileToSave <> False Then
        SaveDialog = fileToSave
    End If
    Path = SaveDialog
End Function

Rem �t�H���_�Q�ƃ_�C�A���O��\�����āA�p�X��Ԃ��B�iExcel 2000�ȍ~�j
Public Function FolderDialog(Optional DefaultFolder As String, Optional Title As String) As String
    Title = Title & " - �t�H���_��I�����Ă�������"
    On Error GoTo msoErr
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = Title
        If DefaultFolder <> "" Then
            If fso.FolderExists(DefaultFolder) Then
                .InitialFileName = DefaultFolder
            Else
                .InitialFileName = fso.GetParentFolderName(DefaultFolder)
            End If
        End If
        If .Show = -1 Then
            Dim Path: Path = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
            If Right$(Path, 1) <> "\" Then Path = Path + "\"
            FolderDialog = Path
            '�����I���̏ꍇ
Rem             For Each vrtSelectedItem In .SelectedItems
Rem                 �t�H���_�Q��Dialog = vrtSelectedItem
Rem                 Path = �t�H���_�Q��Dialog
Rem             Next vrtSelectedItem
        Else
            FolderDialog = ""
            Path = ""
        End If
    End With
    Exit Function
msoErr:
    '���o�[�W�����̂��߂ɁE�E�E
    FolderDialog = ShellFolderDialog(DefaultFolder)
End Function

Rem �����̃t�H���_�Q�ƃ_�C�A���O
Public Function ShellFolderDialog(Optional DefaultFolder As String, Optional Title As String) As String
    If DefaultFolder = "" Then DefaultFolder = "C:\"
    If Title = "" Then Title = "�t�H���_��I�����Ă�������"
    
    Dim shApp As Object
    Set shApp = CreateObject("Shell.Application") _
        .BrowseForFolder(0, Title, 0, DefaultFolder)
    If shApp Is Nothing Then
        ShellFolderDialog = ""
    Else
        ShellFolderDialog = shApp.Items.Item.Path
    End If
End Function

Rem '�A�v���P�[�V�����A�t�H���_�A�֘A�t����ꂽ�t�@�C���̋N��
Rem Public Sub Exec(Path As String)
Rem     Path = RenewalPath(Path)
Rem     If Strings.Right(Path, 1) = "\" Then
Rem         Interaction.Shell "C:\WINDOWS\explorer.exe " & Path, vbNormalFocus
Rem     Else
Rem         Interaction.Shell Path, vbNormalFocus
Rem     End If
Rem End Sub
Rem
Rem '�t�H���_�쐬�B����������True
Rem Public Function CreateFolder(Path As String) As Boolean
Rem     Dim fso As FileSystemObject
Rem     Set fso = New FileSystemObject
Rem     Path = RenewalPath(Path)
Rem     On Error GoTo CreateFolderError
Rem     If fso.FolderExists(Path) = False Then
Rem        ' MkDir Path
Rem         fso.CreateFolder Path
Rem     End If
Rem     CreateFolder = True
Rem     Exit Function
Rem CreateFolderError:
Rem     'MsgBox "�t�H���_�쐬�Ɏ��s���܂����B" + vbCrLf + "�e�t�H���_�̃p�X���Ԉ���Ă��Ȃ����m�F���Ă��������B"
Rem     CreateFolder = False
Rem End Function
Rem
Rem '�t�H���_�ړ��B����������True
Rem Public Function MoveFolder(Path1 As String, Path2 As String) As Boolean
Rem     Path1 = RenewalPath(Path1)
Rem     Path2 = RenewalPath(Path2)
Rem     'On Error GoTo MoveFolderError
Rem     Dim fso As FileSystemObject
Rem     Set fso = New FileSystemObject
Rem     fso.MoveFolder DeleteFolderLastYen(Path1), DeleteFolderLastYen(Path2)
Rem     MoveFolder = True
Rem     Exit Function
Rem MoveFolderError:
Rem     MsgBox "�t�H���_�ړ��Ɏ��s���܂����B" & vbCrLf & "�ړ����F" & Path1 & vbCrLf & "�ړ���F" & Path2
Rem     MoveFolder = False
Rem End Function
