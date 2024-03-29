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
Rem https://docs.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shellexecutea
Rem 機能   指定ファイルを指定した動作で実行します｡
Rem 引数
Rem  hWnd           ShellExecuteを呼び出すウィンドウのハンドル
Rem  lpOperation    処理制御文字列 open edit explore find print runas NULL
Rem                 NULLの挙動:デフォルト動作(定義されていたら)→OPEN(定義されていたら)→レジストリで列挙されている最初の動作
Rem                 設定値はSHELLEXECUTEINFO構造体のlpVerbメンバを参照してください。但し、"properties"は設定できません。
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
Rem 実行ファイルの指定が必須で、対象ファイルはパラメータとして付与しないといけない
Public Function OpenAssociationShell(ByVal ExeFileName As String)
    VBA.Shell ExeFileName, vbNormalFocus
'5
'プロシージャの呼び出し､または引数が不正です｡
End Function

Rem VBAからはバッチを生成し起動、そのバッチの中でファイルを開く方法
Rem Win10関連付け失敗せず
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
Rem Win10関連付け失敗せず
Public Function OpenAssociationShell32(ByVal FileName As String)
    Dim Sh As Object 'Shell32.Shell '参照設定「Microsoft Shell Controls And Automation」
    Set Sh = CreateObject("Shell.Application")
    Sh.ShellExecute FileName
    Set Sh = Nothing
End Function

Rem 指定ファイルを関連付けられたアプリケーションで開く(API)
Rem
Rem  ※"open"動作が設定されていないとエラーになる
Rem   ret = 31 : 指定されたファイル拡張子に関連付けられたアプリケーションがありません。
Public Function OpenAssociationAPI(ByVal FileName As String)

    'Application.hwndが使えるのはExcel2002以降
#If VBA7 Then
    Dim ret As LongPtr
#Else
    Dim ret As Long
#End If
    ret = ShellExecute(GetActiveWindow(), vbNullString, FileName, vbNullString, vbNullString, SW_SHOW)
'    ret = ShellExecute(GetActiveWindow(), "Open", FileName, vbNullString, vbNullString, SW_SHOW)
    OpenAssociationAPI = CLng(ret) 'SR_ERR
End Function

Rem 指定ファイルを関連付けられたアプリケーションで開く(WSH方式)
Rem
Rem  ※"open"動作が設定されていないとエラーになる
Rem   -2147023741
Rem   Run' メソッドは失敗しました: 'IWshShell3' オブジェクト
Public Function OpenAssociationWSH(ByVal strFileName, Optional strParam = "", Optional nMode = SW_SHOWMAXIMIZED)
    Dim strP As String: strP = strFileName
    If strParam <> "" Then strP = strP & " " & strParam
    OpenAssociationWSH = CreateObject("Wscript.Shell").Run(strP, nMode) 'SW_HIDE SW_SHOWMAXIMIZED
End Function

Rem 指定ファイルを関連付けられたアプリケーションで開く(Excelのハイパーリンク機能)
Rem
Rem 必ず確認メッセージが出る
Rem
Rem  ※"open"動作が設定されていないとエラーになる
Rem   -2147221018
Rem   このファイルを開くためのプログラムが登録されていません｡
Public Function OpenAssociationExcelHyperlink(ByVal FileName)
    ThisWorkbook.FollowHyperlink FileName
End Function

Rem ファイル・フォルダをエクスプローラで開く
Rem
Rem  @param full_path   対象ファイル・フォルダのフルパス
Rem  @param IsSelected  選択状態で開くか
Rem
Rem  @note
Rem    選択状態にすると挙動が微妙に変化するので注意
Rem
Public Sub ShellExplorer(full_path, Optional IsSelected As Boolean = False)
    If IsSelected Then
        VBA.Shell "explorer " & full_path & ",/select", vbNormalFocus
    Else
        VBA.Shell "explorer " & full_path & "", vbNormalFocus
    End If
End Sub
