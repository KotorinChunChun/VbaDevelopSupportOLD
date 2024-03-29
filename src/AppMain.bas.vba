Attribute VB_Name = "AppMain"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        AppMain
Rem
Rem  @description   VBA開発を支援するVBE拡張アドイン
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
Rem    2020/08/01 0.1.x 再整備
Rem                     CustomUIの出力に対応
Rem                     Excel以外のプロセスの出力に対応
Rem                     Win32API_PtrSafe.txtからWinAPIのDeclare文の自動生成に対応
Rem                     .kccignoreファイルでdevからbinにコピーするファイル指定に対応
Rem
Rem    2021/04/30 0.2.x 複数のマクロブックから構成されたプロジェクトへ対応
Rem                     フォルダ構成をkccsettings.jsonで定義できるように変更
Rem                     出力パスsrcの既定値が ./src/[FILENAME]/*.vba に変更
Rem
Rem    2021/06/14 0.3.x エクスポート設定にユーザーフォームを採用
Rem                     kccsettings.jsonの項目を追加し一部データ構造を配列に変更
Rem
Rem --------------------------------------------------------------------------------
Rem  @note
Rem    Public Function ParamsToString(Optional Delimiter = " ,") As String のコンマ適切にパースできない不具合がある
Rem    なぜか閉じたはずのブックがゾンビ化して残ることがある。
Rem    OutlookのVBEへのアクセス手段は存在せずエクスポートさせることができない。
Rem
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Public Const APP_NAME = "VBA開発支援アドイン"
Public Const APP_CREATER = "@KotorinChunChun"
Public Const APP_VERSION = "0.3.x"
Public Const APP_SETTINGFILE = APP_NAME & ".xml"
Public Const APP_MENU_MODULE_NAME = "AppMain"
Public Const APP_URL = "https://github.com/KotorinChunChun/VbaDevelopSupport"

Public Const DEF_大文字小文字ファイル = "大文字小文字統一.bas.vba"

Rem 本アドインで「停止」したらこれを実行して再起動させる
Public Sub Reset_Addin(): Call VbeMenuItemDel: Call VbeMenuItemAdd: End Sub
Public Sub Close_Addin(): Call ThisWorkbook.Close(False): End Sub

'Public Sub Auto_Open(): Call Auto_Sub("Open"): End Sub
'Public Sub Auto_Close(): Call Auto_Sub("Close"): End Sub

Rem メニューに追加するプロシージャ
Public Sub Group_ソースコード管理(): End Sub
Public Sub ソースをエクスポート():                      Call VBComponents_Export_Form: End Sub
'Public Sub ソースをインポート():                        Call VBComponents_Import_SRC: End Sub
'Public Sub CustomUIをインポート():                      Call CurrentProject_CustomUI_Import: End Sub
Public Sub プロシージャ一覧を出力():                    Call VbeProcInfo_Output: End Sub

Public Sub Group_コーディング支援(): End Sub
Public Sub Declareの生成():                             Call OpenFormDeclareSourceGenerate: End Sub
Public Sub Declareの変換():                             Call OpenFormDeclareSourceTo64bit: End Sub
Public Sub 大文字小文字統一テキストを開く():            Call OpenTextFileBy大文字小文字: End Sub

Public Sub Group_VBEの機能拡張(): End Sub
'Public Sub プロジェクトのパスワードを1234に変更する():  Call BreakPassword1234Project: End Sub

Public Sub プロジェクトのフォルダを開く():              Call OpenProjectFolder: End Sub
Public Sub プロジェクトをVSCodeで開く():                Call OpenVSCode: End Sub
Public Sub プロジェクトを閉じる():                      Call CloseProject: End Sub
Public Sub ファイル化されていないブック全てを閉じる():  Call CloseNofileWorkbook: End Sub

Public Sub 全てのコードウインドウを閉じる():            Call CloseCodePanes: End Sub
Public Sub イミディエイトウィンドウを空にする():        Call ImdClearGAX: End Sub

Public Sub Group_VBA開発支援アドイン(): End Sub
Public Sub 配布元WEBサイトのヘルプを見る():             Call OpenWebSite(APP_URL): End Sub
Public Sub 終了():                                      Call Close_Addin: End Sub

'Public Sub テスト関数を実行する():          Call TestExecute: End Sub
'Public Sub テスト関数の場所へジャンプする(): Call TestJump: End Sub

'Public Sub プロシージャ一覧を分解する(): Call プロシージャ一覧から引数を分解する: End Sub
