Attribute VB_Name = "AppMain"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        AppMain
Rem
Rem  @description   VBA�J�����x������VBE�g���A�h�C��
Rem
Rem  @update        0.3.x
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem  @references
Rem    Microsoft Visual Basic for Applications Extensibility 5.3
Rem    Microsoft Scripting Runtime
Rem    Microsoft Excel 16.0 Object Library
Rem    Microsoft VBScript Regular Expressions 5.5
Rem
Rem --------------------------------------------------------------------------------
Rem  @history
Rem    2020/08/01 0.1.x �Đ���
Rem                     CustomUI�̏o�͂ɑΉ�
Rem                     Excel�ȊO�̃v���Z�X�̏o�͂ɑΉ�
Rem                     Win32API_PtrSafe.txt����WinAPI��Declare���̎��������ɑΉ�
Rem                     .kccignore�t�@�C����dev����bin�ɃR�s�[����t�@�C���w��ɑΉ�
Rem
Rem    2021/04/30 0.2.x �����̃}�N���u�b�N����\�����ꂽ�v���W�F�N�g�֑Ή�
Rem                     �t�H���_�\����kccsettings.json�Œ�`�ł���悤�ɕύX
Rem                     �o�̓p�Xsrc�̊���l�� ./src/[FILENAME]/*.vba �ɕύX
Rem
Rem    2021/06/14 0.3.x �G�N�X�|�[�g�ݒ�Ƀ��[�U�[�t�H�[�����̗p
Rem                     kccsettings.json�̍��ڂ�ǉ����ꕔ�f�[�^�\����z��ɕύX
Rem
Rem --------------------------------------------------------------------------------
Rem  @note
Rem    Public Function ParamsToString(Optional Delimiter = " ,") As String �̃R���}�K�؂Ƀp�[�X�ł��Ȃ��s�������
Rem    �Ȃ��������͂��̃u�b�N���]���r�����Ďc�邱�Ƃ�����B
Rem    Outlook��VBE�ւ̃A�N�Z�X��i�͑��݂����G�N�X�|�[�g�����邱�Ƃ��ł��Ȃ��B
Rem
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Public Const APP_NAME = "VBA�J���x���A�h�C��"
Public Const APP_CREATER = "@KotorinChunChun"
Public Const APP_VERSION = "0.3.x"
Public Const APP_SETTINGFILE = APP_NAME & ".xml"
Public Const APP_MENU_MODULE_NAME = "AppMain"
Public Const APP_URL = "https://github.com/KotorinChunChun/VbaDevelopSupport"

Public Const DEF_�啶���������t�@�C�� = "�啶������������.bas.vba"

Rem �{�A�h�C���Łu��~�v�����炱������s���čċN��������
Public Sub Reset_Addin(): Call VbeMenuItemDel: Call VbeMenuItemAdd: End Sub
Public Sub Close_Addin(): Call ThisWorkbook.Close(False): End Sub

'Public Sub Auto_Open(): Call Auto_Sub("Open"): End Sub
'Public Sub Auto_Close(): Call Auto_Sub("Close"): End Sub

Rem ���j���[�ɒǉ�����v���V�[�W��
Public Sub Group_�\�[�X�R�[�h�Ǘ�(): End Sub
Public Sub �\�[�X���G�N�X�|�[�g():                      Call VBComponents_Export_Form: End Sub
Public Sub �\�[�X���C���|�[�g():                        Call VBComponents_Import_SRC: End Sub
Public Sub CustomUI���C���|�[�g():                      Call CurrentProject_CustomUI_Import: End Sub
Public Sub �v���V�[�W���ꗗ���o��():                    Call VbeProcInfo_Output: End Sub

Public Sub Group_�R�[�f�B���O�x��(): End Sub
Public Sub Declare�̐���():                             Call OpenFormDeclareSourceGenerate: End Sub
Public Sub Declare�̕ϊ�():                             Call OpenFormDeclareSourceTo64bit: End Sub
Public Sub �啶������������e�L�X�g���J��():            Call OpenTextFileBy�啶��������: End Sub

Public Sub Group_VBE�̋@�\�g��(): End Sub
'Public Sub �v���W�F�N�g�̃p�X���[�h��1234�ɕύX����():  Call BreakPassword1234Project: End Sub

Public Sub �v���W�F�N�g�̃t�H���_���J��():              Call OpenProjectFolder: End Sub
Public Sub �v���W�F�N�g�����():                      Call CloseProject: End Sub
Public Sub �t�@�C��������Ă��Ȃ��u�b�N�S�Ă����():  Call CloseNofileWorkbook: End Sub

Public Sub �S�ẴR�[�h�E�C���h�E�����():            Call CloseCodePanes: End Sub
Public Sub �C�~�f�B�G�C�g�E�B���h�E����ɂ���():        Call ImdClearGAX: End Sub

Public Sub Group_VBA�J���x���A�h�C��(): End Sub
Public Sub �z�z��WEB�T�C�g�̃w���v������():             Call OpenWebSite(APP_URL): End Sub
Public Sub �I��():                                      Call Close_Addin: End Sub

'Public Sub �e�X�g�֐������s����():          Call TestExecute: End Sub
'Public Sub �e�X�g�֐��̏ꏊ�փW�����v����(): Call TestJump: End Sub

'Public Sub �v���V�[�W���ꗗ�𕪉�����(): Call �v���V�[�W���ꗗ��������𕪉�����: End Sub
