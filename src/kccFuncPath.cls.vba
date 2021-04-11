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
Rem  @description   ファイル・フォルダ・パス解析関数
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
Rem    不要
Rem
Rem --------------------------------------------------------------------------------
Rem  @history
Rem    2019/06/24 モジュール整理完了
Rem    2019/09/28 FuncFileListとFuncPathを統合しFuncFileFolderPathとして再定義
Rem    2019/11/12 SpecialFolders追加
Rem    2019/12/05 CreateAllFolderを更新
Rem    2020/02/22 検索関数に汎用フィルタ引数を追加
Rem    2020/05/10 統合 ModIOStream、outlook_path_hyperlink_unc
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
Rem Unicode対応版ファイルリスト作成関数
Rem
Rem --------------------------------------------------------------------------------
Rem
Rem  @history
Rem
Rem 2019/04/24 : 初回リリース
Rem 2019/04/26 : 4/25ブログコメントの指摘を元に修正
Rem 2019/04/27 : 64bit対応。APIをExに変更。エクスプローラ順ソートに対応はできていない。
Rem
Rem --------------------------------------------------------------------------------
Rem
Rem ■使い方
Rem
Rem 低速なFileSystemObjectを使わずに、WindowsAPIのみを使用してファイルリストを
Rem 作成するための関数です。
Rem 私が必要とする独特な機能を搭載しています。
Rem
Rem 以下、注意事項です。
Rem
Rem 1.parentFolderは末尾に\を付けたフォルダパスを指定してください。
Rem     parentFolder :     -   :  必須  : 検索対象フォルダの末尾が\で終わるパス
Rem
Rem     末尾に\が無いと実行時エラーを発生させます。
Rem
Rem 2.AddFileとAddFolderを省略すると、何も取得されません。
Rem     AddFile      :  False  : 省略可 : ファイルを対象に含めるか
Rem     AddFolder    :  False  : 省略可 : フォルダを対象に含めるか
Rem
Rem     少なくとも検索したいどちらかをTrueにしてください。
Rem
Rem 3.SubMinとSubMaxを省略すると、直下のモノしか取得しません。
Rem     SubMin       :      0  : 省略可 : 何階層以降を探索するか（0～n、-1の時は無制限）
Rem     SubMax       :      0  : 省略可 : 何階層以前を探索するか（0～n、-1の時は無制限）
Rem
Rem     parentFolderで指定したパス直下を第0階層としてカウントします。
Rem     よって、SubMinの省略、0、-1は全て同義です。
Rem
Rem   配下の全てのファイルを取得したい場合は、-1,-1になります。
Rem     あるいは、0,9999としても実質的に同じ結果が得られます。
Rem
Rem   ※既定値を配下全てのファイルとすると、莫大な時間がかかる恐れがあるためです。
Rem
Rem 4.戻り値はparentFolderから見た【相対パス】になります。
Rem     絶対パスを返すようにすると、全てのモノに同一の文字列が付与されるため、
Rem     深い階層で検索を開始した時にメモリを無駄に消費するのを防ぐためです。
Rem     したがって、取り出したアイテムはparentFolderと連結してから使用します。
Rem
Rem     また、フォルダの末尾には必ず\を付与した状態で返します。
Rem     ※パス文字列からファイルとフォルダを識別できるようにするためです。
Rem
Rem 5.並び順はファイル→フォルダです。
Rem
Rem     ※たぶんエクスプローラで表示される順序とは異なります。
Rem     ※今後、仕様が変わる恐れがあります。
Rem
Rem     例
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
Rem      Debug.Print "何も取得せず", colPaths.Count
Rem
Rem      Set colPaths = GetFileFolderList(SEARCH_PATH, True)
Rem      Debug.Print "指定パスのファイルのみ取得", colPaths.Count
Rem
Rem      Set colPaths = GetFileFolderList(SEARCH_PATH, True, True)
Rem      Debug.Print "指定パスのファイルとフォルダを取得", colPaths.Count
Rem
Rem      Set colPaths = GetFileFolderList(SEARCH_PATH, True, True, 2, 2)
Rem      Debug.Print "第二階層のファイルとフォルダを取得", colPaths.Count
Rem
Rem      Set colPaths = GetFileFolderList(SEARCH_PATH, True, True, 3, -1)
Rem      Debug.Print "第三階層以下のファイルとフォルダを取得", colPaths.Count
Rem
Rem      Set colPaths = GetFileFolderList(SEARCH_PATH, True, True, -1, -1)
Rem      Debug.Print "指定パス以下の全てのファイルとフォルダを取得", colPaths.Count
Rem
Rem End Sub
Rem --------------------------------------------------------------------------------

Rem --------------------------------------------------------------------------------
Rem ■Outlookでメール受信者がローカルパスをクリックできるようにするマクロ2
Rem
Rem   パスをUNC表記に置き換えることでハイパーリンク化されるようにする案
Rem
Rem   えくせるちゅんちゅん
Rem   2019/10/22
Rem   https://www.excel-chunchun.com/entry/outlook-path-hyperlink-2
Rem
Rem --------------------------------------------------------------------------------

Rem 参考資料

Rem ネットワークドライブからUNCを取得する例
Rem      http://dobon.net/vb/bbs/log3-14/8196.html
Rem      http://blog.livedoor.jp/shingo555jp/archives/1819741.html

Rem WNetGetConnectionについて
Rem
Rem      http://www.pinvoke.net/default.aspx/advapi32/WNetGetUniversalName.html

Rem      Function mpr::WNetGetConnectionW
Rem      https://retep998.github.io/doc/mpr/fn.WNetGetConnectionW.html

Rem      Stack Overflow - Getting An Absolute Image Path
Rem      https://stackoverflow.com/questions/19079162/getting-an-absolute-image-path/19164957

Rem      Passing a LPCTSTR parameter to an API call from VBA in a PTRSAFE and UNICODE safe manner
Rem      https://stackoverflow.com/questions/10402822/passing-a-lpctstr-parameter-to-an-api-call-from-vba-in-a-ptrsafe-and-unicode-saf

Rem APIのAとWの置き換えについて
Rem      RelaxTools - String型の中身は自動的にS-JISに変換される件
Rem      https://software.opensquare.net/relaxtools/archives/3400/

Rem      Programming Field - Win32APIの関数をVBで使うには…
Rem      https://www.pg-fl.jp/program/tips/vbw32api.htm

Rem      AddinBox - Tips26: MsgBox / Beep音 と Unicode文字列
Rem      http://addinbox.sakura.ne.jp/Excel_Tips26.htm

Option Explicit

Rem WNetGetConnection
Rem ローカルデバイスに関連付けられたネットワークリソースの名前を取得します。
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
Rem     ローカル装置に対応するネットワーク資源の名前を取得する｡
Rem
Rem パラメータ
Rem lpLocalName
Rem      ネットワーク名が必要なローカル装置の名前を表す NULL で終わる文字列へのポインタを指定する。
Rem lpRemoteName
Rem      接続に使われているリモート名を表す NULL で終わる文字列を受け取るバッファへのポインタを指定する。
Rem lpnLength
Rem      lpRemoteName パラメータが指すバッファのサイズ（ 文字数）が入った変数へのポインタを指定する。
Rem
Rem      バッファのサイズが不十分で関数が失敗した場合は､必要なバッファサイズがこの変数に格納される｡
Rem
Rem 戻り値
Rem      関数が成功すると､NO_ERROR が返る｡
Rem      関数が失敗すると､次のいずれかのエラーコードが返る｡
Rem
Rem   定数                      意味
Rem   ERROR_BAD_DEVICE          lpLocalName パラメータが指す文字列が無効である。
Rem   ERROR_NOT_CONNECTED       lpLocalName パラメータで指定した装置がリダイレクトされていない。
Rem   ERROR_MORE_DATA           バッファのサイズが不十分である。
Rem                             lpnLength パラメータが指す変数に、必要なバッファサイズが格納されている｡
Rem                             この関数で取得可能なエントリが残っている｡
Rem   ERROR_CONNECTION_UNAVAIL  装置は現在接続されていないが､恒久的な接続として記憶されている｡
Rem   ERROR_NO_NETWORK          ネットワークにつながっていない｡
Rem   ERROR_EXTENDED_ERROR      ネットワーク固有のエラーが発生した。エラーの説明を取得するには、WNetGetLastError 関数を使う｡
Rem   ERROR_NO_NET_OR_BAD_PATH  指定したローカル名を使った接続を認識するプロバイダがない｡
Rem                             その接続を使う1つ以上のプロバイダのネットワークにつながっていない可能性もある｡

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

Rem FindFirstFileEx関数を使用するか
Rem Trueにした場合でも、失敗したら自動でFindFirstFileで対応する
Private Const USE_FindFirstFileEx = True

Rem --------------------------------------------------------------------------------
Rem Win32API関数参照
Rem
Rem 先頭ファイル検索
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

Rem 先頭ファイル検索
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

Rem FindFirstFileExについて
Rem https://blogs.yahoo.co.jp/nobuyuki_tsukasa/1059830.html
Rem https://kkamegawa.hatenablog.jp/entry/20100918/p1
Rem 「8.3形式の短いファイル名を生成させない」ことで、「81%くらいに高速化される」事例があった

Rem   LPCTSTR lpFileName,　　　　　　　// 検索するファイル名
Rem   FINDEX_INFO_LEVELS fInfoLevelId, // データの情報レベル
Rem   LPVOID lpFindFileData,　　　　　 // 返された情報へのポインタ
Rem   FINDEX_SEARCH_OPS fSearchOp,　　 // 実行するフィルタ処理のタイプ
Rem   LPVOID lpSearchFilter,　　　　　 // 検索条件へのポインタ
Rem   DWORD dwAdditionalFlags　　　　　// 補足的な検索制御フラグ

Rem https://docs.microsoft.com/ja-jp/windows/desktop/api/minwinbase/ne-minwinbase-findex_info_levels
Private Enum FINDEX_INFO_LEVELS
    FindExInfoStandard = 0&
    Rem FindFirstFile と同じ動作｡
    
    FindExInfoBasic = 1&
    Rem WIN32_FIND_DATAのcAlternateFileNameに短いファイル名を取得しない。
    Rem Windows Server 2008、Windows Vista、Windows Server 2003、Windows XP ではサポートされていない。
    Rem Windows Server 2008 R2 と Windows 7 では使用可能。
    
    FindExInfoMaxInfoLevel = 2&
    'この値は検証に使用されます。 サポートされている値はこの値よりも小さいです。
End Enum

Rem FINDEX_SEARCH_OPS列挙体
Rem fSearchOp : ワイルドカードとの照合以外のフィルタ処理タイプを表す
Private Enum FINDEX_SEARCH_OPS
    FindExSearchNameMatch = 0&
    Rem 指定したファイル名と一致するファイルを検索します。
    Rem この検索操作を使用するときは､FindFirstFileExのlpSearchFilterパラメータをNULLにする必要があります｡
    
    FindExSearchLimitToDirectories = 1&
    Rem ファイルシステムがディレクトリフィルタリングをサポートしている場合、ディレクトリを検索します。
    Rem 実際にはサポートしているファイルシステムは存在せず?効果がないとの事。
    Rem https://gist.github.com/kumatti1/33182de4efe99259e275
    Rem http://www.vbalab.net/vbaqa/c-board.cgi?cmd=one;no=58244;id=excel
    
    FindExSearchLimitToDevices = 2&
    Rem このフィルタリングタイプは利用できません。
    
    FindExSearchMaxSearchOp = 3&
    Rem サポートされていません。
End Enum

Rem dwAdditionalFlags
Private Const FIND_FIRST_EX_CASE_SENSITIVE = 1&
Rem 検索では大文字と小文字が区別されます。

Private Const FIND_FIRST_EX_LARGE_FETCH = 2&
Rem ディレクトリー照会にはより大きなバッファーを使用します。
Rem これにより、検索操作のパフォーマンスが向上する可能性があります。
Rem Windows Server 2008、Windows Vista、Windows Server 2003、およびWindows XP：
Rem この値は、Windows Server 2008 R2およびWindows 7まではサポートされていません。

Private Const FIND_FIRST_EX_ON_DISK_ENTRIES_ONLY = 4&
Rem 結果を物理的にディスク上にあるファイルに制限します。
Rem このフラグは、ファイル仮想化フィルタが存在する場合にのみ関係します。

Rem 次ファイル検索
Rem http://chokuto.ifdef.jp/urawaza/api/FindNextFile.html
#If VBA7 Then
Private Declare PtrSafe Function FindNextFile Lib "Kernel32" Alias "FindNextFileW" _
            (ByVal hFindFile As LongPtr, lpFindFileData As WIN32_FIND_data1) As LongPtr
#Else
Private Declare Function FindNextFile Lib "Kernel32" Alias "FindNextFileW" _
            (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_data1) As Long
#End If
            
Rem 検索ハンドル開放
Rem http://chokuto.ifdef.jp/urawaza/api/FindClose.html
#If VBA7 Then
Private Declare PtrSafe Function FindClose Lib "Kernel32" (ByVal hFindFile As LongPtr) As LongPtr
#Else
Private Declare Function FindClose Lib "Kernel32" (ByVal hFindFile As Long) As Long
#End If

Rem --------------------------------------------------------------------------------

Rem FILETIME構造体
Rem http://chokuto.ifdef.jp/urawaza/struct/FILETIME.html
Private Type FILETIME
     LowDateTime As Long
     HighDateTime As Long
End Type

Rem WIN32_FIND_DATA構造体
Rem http://chokuto.ifdef.jp/urawaza/struct/WIN32_FIND_data1.html
Private Type WIN32_FIND_data1
    dwFileAttributes                        As Long     ' ファイル属性
    ftCreationTime                          As FILETIME ' 作成日
    ftLastAccessTime                        As FILETIME ' 最終アクセス日
    ftLastWriteTime                         As FILETIME ' 最終更新日
    nFileSizeHigh                           As Long     ' ファイルサイズ（上位３２ビット）
    nFileSizeLow                            As Long     ' ファイルサイズ（下位３２ビット）
    dwReserved0                             As Long     ' 予約済み。リパースタグ
    dwReserved1                             As Long     ' 予約済み。未使用
    cFileName(260 * 2 - 1)                  As Byte     ' ファイル名
    cAlternateFileName(14 * 2 - 1)          As Byte     ' 8.3形式のファイル名
Rem      cFileName                               As String * MAX_PATH    ' という書き方もできる。
Rem      cAlternateFileName                      As String * 14          ' という書き方もできる。
End Type
Rem ※Unicode対応の為*2している

Rem StrCmpLogicalW関数
Rem エクスプローラのファイル順に並び替える
Rem https://docs.microsoft.com/en-us/windows/desktop/api/shlwapi/nf-shlwapi-strcmplogicalw
#If VBA7 Then
Private Declare PtrSafe Function StrCmpLogicalW Lib "shlwapi" _
                (ByVal lpStr1 As String, ByVal lpStr2 As String) As Long
#Else
Private Declare Function StrCmpLogicalW Lib "shlwapi" _
                (ByVal lpStr1 As String, ByVal lpStr2 As String) As Long
#End If

Rem wsh.SpecialFoldersプロパティ
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
Rem     特殊フォルダ名     説明    一般的Windows10の具体的なフォルダ
Rem 1   AllUsersDesktop    すべてのユーザーに共通のデスクトップ        C:\Users\Public\Desktop
Rem 2   AllUsersStartMenu  すべてのユーザーに共通のプログラムメニュー  C:\ProgramData\Microsoft\Windows\Start Menu
Rem 3   AllUsersPrograms   すべてのユーザーに共通の全てのプログラム    C:\ProgramData\Microsoft\Windows\Start Menu\Programs
Rem 4   AllUsersStartup    すべてのユーザーに共通のスタートアップ      C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp
Rem 5   Desktop            デスクトップ                            C:\Users\[username]\Desktop
Rem 6   Favorites          お気に入り                              C:\Users\[username]\Favorites
Rem 7   Fonts              インストールされているフォント          C:\Windows\Fonts
Rem 8   MyDocuments        マイドキュメント                        C:\Users\[username]\Documents
Rem 9   NetHood            ネットワークに表示される共有フォルダの情報  C:\Users\[username]\AppData\Roaming\Microsoft\Windows\Network Shortcuts
Rem 10  PrintHood          プリンタフォルダ                        C:\Users\[username]\AppData\Roaming\Microsoft\Windows\Printer Shortcuts
Rem 11  Programs           ログインユーザーのプログラムメニュー    C:\Users\[username]\AppData\Roaming\Microsoft\Windows\Start Menu\Programs
Rem 12  Recent             最近使ったファイル                      C:\Users\[username]\AppData\Roaming\Microsoft\Windows\Recent
Rem 13  SendTo             送るメニュー                            C:\Users\[username]\AppData\Roaming\Microsoft\Windows\SendTo
Rem 14  StartMenu          スタートメニュー                        C:\Users\[username]\AppData\Roaming\Microsoft\Windows\Start Menu
Rem 15  Startup            ログインユーザーのスタートアップ        C:\Users\[username]\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup
Rem 16  Templates          新規作成のテンプレート                  C:\Users\[username]\AppData\Roaming\Microsoft\Windows\Templates

Rem SetCurrentDirectory API 使用準備
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

Rem SHCreateDirectoryEx の戻り値
Const ERROR_BAD_PATHNAME = 161&         '指定されたパスが無効です。
Const ERROR_FILENAME_EXCED_RANGE = 206& 'ファイル名または拡張子が長すぎます。
Const ERROR_PATH_NOT_FOUND = 3&         '指定されたパスが見つかりません。
Const ERROR_FILE_EXISTS = 80&           'ディレクトリは存在する。
Const ERROR_ALREADY_EXISTS = 183&       'ディレクトリは存在する。
Const ERROR_CANCELLED = 0&              'ユーザーは操作を取り消した。
Const ERROR_ACCESS_DENIED = 5&          'アクセスが拒否されました。

Rem エラーコード表
Rem https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes

Rem --------------------------------------------------------------------------------
Rem 共通組み込み
Private Property Get fso() As FileSystemObject
    Static xxFso As Object  'FileSystemObject
    If xxFso Is Nothing Then Set xxFso = CreateObject("Scripting.FileSystemObject")
    Set fso = xxFso
End Property
Rem --------------------------------------------------------------------------------

Rem 指定したパスのフォルダを一気に作成する
Rem 失敗した時だけFalseを返す。既に存在した場合は無視でOK
Rem
Rem  @param folder_path 作成したいフォルダ
Rem
Rem  @return As Boolen  成功したかどうか
Rem                      作成に成功 : True
Rem                      既に存在   : True
Rem                      作成に失敗 : False
Rem
Public Function CreateDirectoryEx(folder_path As String, Optional ByRef errValue) As Boolean
    errValue = SHCreateDirectoryEx(0&, StrPtr(SupportMaxPath260over(folder_path)), 0&)
    Select Case errValue
        Case 0:  CreateDirectoryEx = True '成功
        Case 183: CreateDirectoryEx = True '既に存在
        Case Else: CreateDirectoryEx = False '失敗
    End Select
End Function

Rem Public Function CreateDirectoryExA(folder_path As String) As Boolean
Rem     Select Case SHCreateDirectoryExA(0&, folder_path, 0&)
Rem         Case 0:  CreateDirectoryExA = True '成功
Rem         Case 183: CreateDirectoryExA = True '既に存在
Rem         Case Else: CreateDirectoryExA = False '失敗
Rem     End Select
Rem End Function

Rem Win32APIのW付き関数において、260 (MAX_PATH) 文字よりも長い文字に対応させるための処理
Rem
Rem  @param file_folder_path ファイルかフォルダのパス
Rem
Rem  @return As String       変換後のパス
Rem
Rem  @note
Rem    パスの先頭に "\\?\"や"\\?\UNC" を追加しておく
Rem
Rem  @example
Rem    \\SERVERNAME\    >>  \\?\UNC\SERVERNAME\
Rem    C:\DRIVE         >>  \\?\C:\DRIVE
Rem
Public Function SupportMaxPath260over(ByRef file_folder_path As String) As String
    
    '対処済み
    If file_folder_path Like "\\?\*" Then
        SupportMaxPath260over = file_folder_path
        
    'ネットワークパス
    ElseIf file_folder_path Like "\\*" Then
        SupportMaxPath260over = "\\?\UNC" & Mid$(file_folder_path, 2)
        
    '標準ドライブパス
    Else
        SupportMaxPath260over = "\\?\" & file_folder_path
    End If
End Function

Rem *******************************************************************************
Rem 指定フォルダ以下の任意の階層のファイル・フォルダをリストアップする関数
Rem *******************************************************************************
Rem パラメータ           : 既定値  :  概要  : 意味
Rem parent_folder_path   :     -   :  必須  : 検索対象フォルダの末尾が\で終わるパス
Rem add_files            :  False  : 省略可 : ファイルを対象に含めるか
Rem add_folders          :  False  : 省略可 : フォルダを対象に含めるか
Rem search_min_layer     :      0  : 省略可 : 何階層以降を探索するか（0～n、-1の時は無制限）
Rem search_max_layer     :      0  : 省略可 : 何階層以前を探索するか（0～n、-1の時は無制限）
Rem filter_obj           : Missing : 省略可 : フィルタ(RegExp,LIKE用文字列,Everything検索仕様）
Rem recursive_subfolder  :     ""  : 再帰用 : 当初のルートフォルダ以降のパス
Rem recursive_now_layer  :      0  : 再帰用 : 現在何階層目か
Rem recursive_path_list  : Nothing : 再帰用 : パス一覧。最終的な戻り値にも使われる
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
    
    Rem 関数突入時の初期化処理
    If recursive_path_list Is Nothing Then
        Set recursive_path_list = New Collection
        
        If Len(parent_folder_path) > 0 Then
            If Right(parent_folder_path, 1) <> "\" Then
                Err.Raise 9999, PROC_NAME, "フォルダパスの末尾は\で終わるようにしてください。"
            End If
        End If
    End If
    
    Dim ResFolder As Collection: Set ResFolder = New Collection
    Dim ResFile As Collection: Set ResFile = New Collection
    Dim findData As WIN32_FIND_data1
    
    Dim UnicodeFolderPath As String
    UnicodeFolderPath = SupportMaxPath260over(parent_folder_path)
    
    Rem 検索ハンドルが見つからない場合は「INVALID_HANDLE_VALUE」を返す
#If VBA7 Then
    Dim FileHandle As LongPtr
#Else
    Dim FileHandle As Long
#End If
    If USE_FindFirstFileEx Then
        Rem ※FindExInfoBasic/FIND_FIRST_EX_LARGE_FETCH指定により高速化を実現
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
        Rem FindFirstFileではファイル名の後ろに「Max_Path」で指定した文字数までNullが詰まっている。
        Dim intStLen As Long
        intStLen = InStr(findData.cFileName, vbNullChar) - 1
        If intStLen > 0 Then
            Dim sFilename As String
            sFilename = Trim$(Left$(findData.cFileName, intStLen))
            
            Rem カレントフォルダ以外か上位フォルダ
            If sFilename = "." Or sFilename = ".." Then
                
            Rem フォルダ予約
            ElseIf findData.dwFileAttributes And vbDirectory Then
                ResFolder.Add sFilename
                
            Rem ファイル予約
            Else
                If add_files And _
                        (search_min_layer = -1 Or search_min_layer <= recursive_now_layer) And _
                        IsMatchPathFilter(filter_obj, folder_path:=recursive_subfolder, file_name:=sFilename) Then
                    Rem ※ファイルはルートパスを除いた相対パス
                    ResFile.Add recursive_subfolder & sFilename
                End If
            End If
        End If
        Rem 次のファイルが見つからなかった場合は0を返すためループ終了
    Loop Until FindNextFile(FileHandle, findData) = 0
    
    Rem 検索ハンドルを閉じる
    FindClose FileHandle
    
    Rem ファイルリストをソートしてから追加
Rem     CollectionSort_StrCmpLogicalW ResFile
    Dim myFile As Variant
    For Each myFile In ResFile
        recursive_path_list.Add myFile
    Next
    
    Rem フォルダリストをソートしてから追加して、再帰探索へ
Rem     CollectionSort_StrCmpLogicalW ResFolder
    
    Dim myFolder As Variant
    For Each myFolder In ResFolder
        Rem フォルダ追加
        If add_folders Then
            Rem ※フォルダはルートパスを除いた相対パスで末尾は "\"
            recursive_path_list.Add recursive_subfolder & myFolder & "\"
        End If
        Rem サブフォルダ再帰探索
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

Rem ファイル・フォルダフィルタリング用の検証用関数
Rem 正規表現、Everything式検索仕様、VBA式LIKE演算子が使える。
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
    
    '正規表現
    If TypeName(filter_obj) = "RegExp" Then
        Dim reg As Object 'RegExp
        Set reg = filter_obj
        IsMatchPathFilter = reg.Execute(FullPath)
        
    '文字列指定
    ElseIf TypeName(filter_obj) = "String" Then
        'LIKE演算子
        If VBA.Strings.InStr(filter_obj, ":") = 0 Then
            IsMatchPathFilter = (FullPath Like filter_obj)
        
        'Everything式
        Else
            '未完成
            Stop
        End If
    End If
End Function

Rem 'Colection入れ替え
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
Rem 'CollectionをStrCmpLogicalWでソート
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

Rem 一時ファイルのフルパスを取得
Public Function GetPathByTemporaryFile() As String
    GetPathByTemporaryFile = GetPathTemporary & "\" & fso.GetTempName
End Function

Rem --------------------------------------------------------------------------------
Rem   フォルダの一括作成
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

Rem ドキュメントフォルダ
Public Function GetPathMyDocument() As String: GetPathMyDocument = GetPathWSH("MyDocuments"): End Function
Rem AppDataフォルダ
Public Function GetPathAppData() As String: GetPathAppData = GetPathWSH("AppData"): End Function
Rem デスクトップフォルダ
Public Function GetPathDesktop() As String: GetPathDesktop = GetPathWSH("Desktop"): End Function

Rem テンポラ\一時ファイルリフォルダ
Public Function GetPathTemporary() As String: GetPathTemporary = fso.GetSpecialFolder(TemporaryFolder): End Function

Rem アプリ名のサブフォルダを生成してラップして返す
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

Rem AppDataフォルダ
Public Function GetAppPathAppData(Optional ProjectFolderName) As String: GetAppPathAppData = GetAppPath("AppData", ProjectFolderName): End Function

Rem テンポラリフォルダ取得
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

Rem テンポラリフォルダ取得
Public Function CreateTempFolder(SpecialFolderKey As String, Optional folder_name_format As String = "yyyymmdd_hhmmss") As String
    CreateTempFolder = CreateObject("Wscript.Shell").SpecialFolders(CVar(SpecialFolderKey)) & "\" & Format(Now, folder_name_format)
    On Error Resume Next
    If fso.CreateFolder(CreateTempFolder) Then
        If Err Then Debug.Print "ERROR CreateTempFolder : " & Err.Description
    End If
    CreateTempFolder = CreateTempFolder & "\"
End Function

Rem カレントディレクトリの変更　―　ネットワークパスをカレントディレクトリにする
Rem 　ChDir　CurDir　パス変更　現在のフォルダ
Sub SetCurrentDirectory_WScriptShell(new_path)
    CreateObject("WScript.Shell").CurrentDirectory = new_path
End Sub

Rem 四捨五入　―　数値を任意の有効桁数に四捨五入する
Rem   Round ワークシート関数
Public Function SignificantFigures(Number, l) As Double
    '数値を有効数字L桁に四捨五入する
    If Number = 0 Then
        SignificantFigures = 0
    Else
        SignificantFigures = Application.Round(Number, -Int(Application.Log(Abs(Number))) - 1 + l)
    End If
End Function

Rem 指定したファイルがロックされているかチェックする。
Public Function GetFileLock(FileName As String) As Boolean
    On Error Resume Next
    Dim fn: fn = FreeFile
    Open FileName For Append As #fn
    Close #fn
    GetFileLock = (Err.Number > 0)
End Function

Rem 指定したファイルが読み取り専用かチェックする。
Public Function GetFileReadonly(FileName As String) As Boolean
    'Readonly属性かの判断はコレ
    GetFileReadonly = (GetAttr(FileName) And vbReadOnly)
End Function

Rem ネットワークドライブのUNCパスを取得
Rem
Rem  @param nDriveLetter     ドライブレター文字列（"A:"や"Z:"）
Rem
Rem  @return As String       文字列
Rem
Rem  @note サーバーにアクセスできるか否かは考慮しない。
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

Rem 割当済のドライブレターとUNCパスのリストをDictionaryで返す関数
Rem
Rem  @return As Dictionary     dic(ドライブレター) = UNC
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

Rem 接続済のドライブレターとUNCパスのリストをDictionaryで返す関数(WSHバージョン)
Rem
Rem  @return As Dictionary     dic(ドライブレター) = UNC
Rem
Rem  @note 接続済状態のドライブしか検出できないため
Rem         全てのドライブを列挙することはできない。
Rem
Public Function GetNetworkDriveAndUncByConnected() As Object ' As Dictionary
    Dim DicDrives ' As Dictionary
    Set DicDrives = CreateObject("Scripting.Dictionary")
    
    Dim Network ' As WScript.Network
    Set Network = CreateObject("WScript.Network")
    
    Rem Network.EnumNetworkDrives
    Rem   (0):=ドライブレター1, (1):=UNCパス1
    Rem   (2):=ドライブレター2, (3):=UNCパス2
    Rem   (4):=ドライブレター3, (5):=UNCパス3
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

Rem 接続済のドライブレターとUNCパスのリストをDictionaryで返す関数(WMIバージョン)
Rem
Rem  @return As Dictionary     dic(ドライブレター) = UNC
Rem
Rem  @note 接続済状態のドライブしか検出できないため
Rem         全てのドライブを列挙することはできない。
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

Rem コマンドプロンプトで取得する参考資料
Rem
Rem C:\Users\USERNAME>net use
Rem 新しい接続は記憶されます｡
Rem
Rem ステータス  ローカル名 リモート名                ネットワーク名
Rem
Rem --------------------------------------------------------------------------------
Rem 利用不可     V:        \\192.168.11.1\Share      Microsoft Windows Network
Rem 利用不可     W:        \\landisk\disk            Microsoft Windows Network
Rem OK           X:        \\servername-nuc\Downloads
Rem                                                   Microsoft Windows Network
Rem OK           Y:        \\servername-nuc\Server   Microsoft Windows Network
Rem 利用不可     Z:        \\crib35nas\Share         Microsoft Windows Network
Rem
Rem ・接続済み以外も全て列挙される。
Rem ・リモート名の文字数が長いと改行されて出力される。

Rem --------------------------------------------------------------------------------
Rem ■kccFuncString
Rem   文字列変換関数
Rem --------------------------------------------------------------------------------
Rem
Rem 抜粋
Rem
Rem --------------------------------------------------------------------------------

Rem 文章中のパスと思われる文字列をハイパーリンクに対応させる関数
Rem
Rem  @param base_str        変換元文字列
Rem  @param DoNetDriveToUNC ネットワークドライブをUNCに変換するか否か
Rem                          False:=変換しない(既定)
Rem                          True :=変換する
Rem
Rem  @return  As string     Outlookがハイパーリンク化可能な文字列
Rem
Rem  @example
Rem     IN :
Rem          下記のファイルを御覧ください
Rem          C:\Test\hoge.xls
Rem          Z:\fuga.xls
Rem          以上
Rem
Rem    OUT :
Rem       DoNetDriveToUNC:=False
Rem          下記のファイルを御覧ください
Rem          <"file://C:\Test\hoge.xls">
Rem          <"file://Z:\Test\hoge.xls">
Rem          以上
Rem
Rem       DoNetDriveToUNC:=True
Rem          <"\\server\share\fuga.xls">
Rem
Rem  @note
Rem         (Trueなら)ネットワークドライブのパスはUNCに変更することでハイパーリンク化
Rem         ローカルドライブのパスは <"file:// "> で囲うことでハイパーリンク化
Rem         UNCパスは <" "> で囲うことで途切れ防止
Rem
Rem         パスは必ず改行で終わること
Rem         Outlookではメール送信時の自動折返しを切っておくこと
Rem         メール作成画面ではリンク状態にはならない。
Rem         自分から自分へ送信してテストするように。
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

    '改行(CRLF)をパス終了とみなす
    Dim base_str_arr
    base_str_arr = Split(base_str, vbCrLf)

    'UNCパスの変換
    Const UncPathPrefix = "\\"
    For i = LBound(base_str_arr) To UBound(base_str_arr)
        s = base_str_arr(i)

        'UNCパスを<"UNCパス">に変換
        pathIdx = InStr(lfIdx, s, UncPathPrefix)
        If pathIdx > 0 Then
            pathData = Mid(s, pathIdx, Len(s))
            s = Replace(s, pathData, "<""" & pathHeader & pathData & """>")
            base_str_arr(i) = s
        End If
    Next

    'ドライブレター付きパスの変換
    Dim pathArr(1 To 26)
    For i = 1 To 26: pathArr(i) = Chr(Asc("A") - 1 + i) & ":": Next

    For i = LBound(base_str_arr) To UBound(base_str_arr)
        s = base_str_arr(i)

        'パスと思われる文章を検索
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
                'ネットワークドライブのパスを<"\\ServerName\ShareName\パス">に変換(既存のfile://は消す)
                pathData = Mid(s, pathIdx, Len(s))
                s = Replace(s, pathData, "<""" & Replace(pathData, DriveLetter, UncPath) & """>")
                s = Replace(s, LocalPrefix, "")
                base_str_arr(i) = s
            Else
                'ローカルドライブのパスを<"file://パス">に変換
                pathData = Mid(s, pathIdx, Len(s))
                s = Replace(s, pathData, "<""" & pathHeader & pathData & """>")
                base_str_arr(i) = s
            End If
        End If
    Next

    '既に付与されていた場合の二重付与を解除
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
Sub メール作成画面のパスをハイパーリンクに変換()
    Dim objItem As Outlook.MailItem
    Set objItem = ActiveInspector.CurrentItem
    objItem.body = ReplacePathToHyperlink(objItem.body)
End Sub
#End If

Rem \\でもハイパーリンクになるが、file://じゃないとリンクは無効だった

Rem ついでに受信メールの自動改行も修復したい
#If NO_COMPILE Then
C:\Test\hoge.xls
#End If


Rem 参考資料　このくらいのコードは直接かいたほうが分かりやすい
Rem 'ファイルを開くダイアログを表示して、パスを返す。（旧式EXCEL対応）
Public Function OpenDialog(Path As String, Filter As String) As String
    OpenDialog = ""
    Dim fileToOpen As Variant
    If Path <> "" Then SetCurrentDirectory Path
    fileToOpen = Application _
        .GetOpenFileName(Filter)  '"エクセルファイル(*.xls;*.xlsx), *.xls;*.xlsx"
    If fileToOpen <> False Then
        OpenDialog = fileToOpen
    End If
    Path = OpenDialog
End Function

Rem 保存ダイアログを表示して、パスを返す。（旧式EXCEL対応）
Public Function SaveDialog(Path As String, Filter As String) As String
    SaveDialog = ""
    Dim fileToSave As Variant
    fileToSave = Application _
        .GetSaveAsFilename(Path, Filter)  '"エクセルファイル(*.xls;*.xlsx), *.xls;*.xlsx"
    If fileToSave <> False Then
        SaveDialog = fileToSave
    End If
    Path = SaveDialog
End Function

Rem フォルダ参照ダイアログを表示して、パスを返す。（Excel 2000以降）
Public Function FolderDialog(Optional DefaultFolder As String, Optional Title As String) As String
    Title = Title & " - フォルダを選択してください"
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
            '複数選択の場合
Rem             For Each vrtSelectedItem In .SelectedItems
Rem                 フォルダ参照Dialog = vrtSelectedItem
Rem                 Path = フォルダ参照Dialog
Rem             Next vrtSelectedItem
        Else
            FolderDialog = ""
            Path = ""
        End If
    End With
    Exit Function
msoErr:
    '旧バージョンのために・・・
    FolderDialog = ShellFolderDialog(DefaultFolder)
End Function

Rem 旧式のフォルダ参照ダイアログ
Public Function ShellFolderDialog(Optional DefaultFolder As String, Optional Title As String) As String
    If DefaultFolder = "" Then DefaultFolder = "C:\"
    If Title = "" Then Title = "フォルダを選択してください"
    
    Dim shApp As Object
    Set shApp = CreateObject("Shell.Application") _
        .BrowseForFolder(0, Title, 0, DefaultFolder)
    If shApp Is Nothing Then
        ShellFolderDialog = ""
    Else
        ShellFolderDialog = shApp.Items.Item.Path
    End If
End Function

Rem 'アプリケーション、フォルダ、関連付けられたファイルの起動
Rem Public Sub Exec(Path As String)
Rem     Path = RenewalPath(Path)
Rem     If Strings.Right(Path, 1) = "\" Then
Rem         Interaction.Shell "C:\WINDOWS\explorer.exe " & Path, vbNormalFocus
Rem     Else
Rem         Interaction.Shell Path, vbNormalFocus
Rem     End If
Rem End Sub
Rem
Rem 'フォルダ作成。成功したらTrue
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
Rem     'MsgBox "フォルダ作成に失敗しました。" + vbCrLf + "親フォルダのパスが間違っていないか確認してください。"
Rem     CreateFolder = False
Rem End Function
Rem
Rem 'フォルダ移動。成功したらTrue
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
Rem     MsgBox "フォルダ移動に失敗しました。" & vbCrLf & "移動元：" & Path1 & vbCrLf & "移動先：" & Path2
Rem     MoveFolder = False
Rem End Function
