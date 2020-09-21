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
Rem ShellExecute関数
Rem 機能   指定ファイルを指定した動作で実行します｡
Rem 引数
Rem  hWnd           ShellExecuteを呼び出すウィンドウのハンドル
Rem  lpOperation    処理制御文字列。指定しないときは"open"になります。設定値はSHELLEXECUTEINFO構造体のlpVerbメンバを参照してください。但し、"properties"は設定できません。
Rem  lpFile         起動するファイルの名前
Rem  lpParameters   起動する実行ファイルへのパラメータ（lpFileが実行可能ファイルのとき）。lpFileがドキュメントファイルのときは設定しないで下さい。
Rem  lpDirectory    作業用ディレクトリ｡設定しないときはカレントディレクトリになります｡
Rem  nShowCmd       起動する実行可能ファイルのウィンドウの状態｡設定値はSHELLEXECUTEINFO構造体のnShowメンバ SW_****
Private Const SW_HIDE = 0            'ウィンドウを非表示にして、他のウィンドウをアクティブにします。
Private Const SW_SHOWNORMAL = 1      'ウィンドウをアクティブにして表示します。ウィンドウが最小化または最大化されている場合は、ウィンドウの位置とサイズを元に戻します。アプリケーションは、最初にウィンドウを表示させるときにこのフラグを指定するべきです。
Private Const SW_SHOWMINIMIZED = 2   'ウィンドウをアクティブにして、最小化されたウィンドウとして表示します。
Private Const SW_SHOWMAXIMIZED = 3   'ウィンドウをアクティブにして、最大化されたウィンドウとして表示します。
Private Const SW_MAXIMIZE = 3        'ウィンドウをアクティブにして、最大化されたウィンドウとして表示します。
Private Const SW_SHOWNOACTIVATE = 4  'ウィンドウをアクティブにはせずに表示します。
Private Const SW_SHOW = 5            'ウィンドウをアクティブにして、現在の位置とサイズで表示します。
Private Const SW_MINIMIZE = 6        '指定されたウィンドウを最小化して、次の Z オーダーにあるトップレベルウィンドウをアクティブにします。
Private Const SW_SHOWMINNOACTIVE = 7 'ウィンドウを最小化されたウィンドウとして表示します。ウィンドウはアクティブ化されません。
Private Const SW_SHOWNA = 8          'ウィンドウを現在の位置とサイズで表示します。ウィンドウはアクティブ化されません。
Private Const SW_RESTORE = 9         'ウィンドウをアクティブにして表示します。ウィンドウが最小化または最大化されている場合は、ウィンドウの位置とサイズを元に戻します。アプリケーションは、最小化されたウィンドウの位置とサイズを元に戻すときにこのフラグを指定するべきです。
Private Const SW_SHOWDEFAULT = 10    'アプリケーションを起動したプログラムがCreateProcess関数にパラメータとして渡したSTARTUPINFO構造体で指定されている SW_ 値に基づいて表示状態が設定されます。
Private Const SW_FORCEMINIMIZE = 11  'Windows 2000/XP：ウィンドウを所有しているスレッドがハングしている状態であっても、ウィンドウを最小化します。他のスレッドからウィンドウを最小化させる場合にのみ、このフラグを使用するべきです。
Rem 戻り値
Rem  33以上 : 開いたファイルのインスタンスハンドル
Rem  32以下 : エラーコード
Rem                                = 0    'メモリまたはリソースが不足しています。
Private Const ERROR_FILE_NOT_FOUND = 2    '指定されたファイルが見つかりませんでした。
Private Const ERROR_PATH_NOT_FOUND = 3    '指定されたパスが見つかりませんでした。
Private Const ERROR_BAD_FORMAT = 11       '.exe ファイルが無効です。Win32 の .exe ではないか、.exe イメージ内にエラーがあります。
Private Const SE_ERR_ACCESSDENIED = 5     'オペレーティングシステムが、指定されたファイルへのアクセスを拒否しました。
Private Const SE_ERR_ASSOCINCOMPLETE = 27 'ファイル名の関連付けが不完全または無効です。
Private Const SE_ERR_DDEBUSY = 30         'ほかの DDE トランザクションが現在処理中なので、DDE トランザクションを完了できませんでした。
Private Const SE_ERR_DDEFAIL = 29         'DDE トランザクションが失敗しました。
Private Const SE_ERR_DDETIMEOUT = 28      '要求がタイムアウトしたので、DDE トランザクションを完了できませんでした。
Private Const SE_ERR_DLLNOTFOUND = 32     '指定されたダイナミックリンクライブラリ（DLL）が見つかりませんでした。
Private Const SE_ERR_FNF = 2              '指定されたファイルが見つかりませんでした。
Private Const SE_ERR_NOASSOC = 31         '指定されたファイル拡張子に関連付けられたアプリケーションがありません。
                                         '印刷可能ではないファイルを印刷しようとした場合も、このエラーが返ります。
Private Const SE_ERR_OOM = 8              '操作を完了するのに十分なメモリがありません。
Private Const SE_ERR_PNF = 3              '指定されたパスが、見つかりませんでした。
Private Const SE_ERR_SHARE = 26           '共有違反が発生しました。

#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hWnd As LongPtr, _
        ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As LongPtr) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hWnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd_ As Long) As Long
#End If
Rem --------------------------------------------------------------------------------

'指定ファイルを関連付けられたアプリケーションで開く
Public Sub ShellEx(FileName As String)

    'Application.hwndが使えるのはExcel2002以降
#If VBA7 Then
    Dim ret As LongPtr
#Else
    Dim ret As Long
#End If
    ret = ShellExecute(GetActiveWindow(), "Open", FileName, _
              vbNullString, vbNullString, SW_SHOW)
              
End Sub
