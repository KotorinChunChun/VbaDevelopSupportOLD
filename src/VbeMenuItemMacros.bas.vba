Attribute VB_Name = "VbeMenuItemMacros"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        VbeMenuItemMacros
Rem
Rem  @description   VBEのメニュー追加マクロ
Rem
Rem  @update        2020/08/01
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem  @references
Rem    不要
Rem
Rem --------------------------------------------------------------------------------
Rem  @refModules
Rem    kccFuncString
Rem    VbeMenuItemCreator
Rem    VbeMenuItemEventHandler
Rem    VbeMenuItemInstructions
Rem
Rem --------------------------------------------------------------------------------
Rem  @history
Rem    2020/08/01 再整備
Rem
Rem --------------------------------------------------------------------------------
Rem  @functions
Rem
Rem --------------------------------------------------------------------------------
Rem  @note
Rem    メモ書き
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Public Sub VbeMenuItemAdd(): Call VbeMenuUpdate("Add", APP_MENU_MODULE_NAME): End Sub
Public Sub VbeMenuItemDel(): Call VbeMenuUpdate("Del", APP_MENU_MODULE_NAME): End Sub

Rem @param AddOrDel     追加か削除か
Rem                     "Add" : アドイン起動時に使用しメニューを追加
Rem                     "Del" : アドイン終了時に使用しメニューを削除
Rem @param moduleName   追加したいコマンドの列挙されたモジュール名
Rem
Private Sub VbeMenuUpdate(AddOrDel, moduleName)
    Static vbc As VBComponent
    Static Menu As VbeMenuItemCreator
    If Menu Is Nothing Then
        Set vbc = ThisWorkbook.VBProject.VBComponents(moduleName)
        Set Menu = New VbeMenuItemCreator
        Menu.Init "VBE開発支援", "VBE開発支援(&M)", vbc
    End If
    
    If AddOrDel = "Add" Then
        Dim arr() As VbeMenuItemInstructions
        arr = GetInstructions(vbc.CodeModule)
        
        Dim i As Long
        For i = LBound(arr) To UBound(arr)
            Select Case arr(i).ProcName
                Case "Reset_Addin", "Close_Addin"
                Case "Auto_Open", "Auto_Close", "Auto_Sub"
                Case "GetInstructions", "VbeMenuItemAdd", "VbeMenuItemDel"
                    'これらは処理しない。
                Case Else
                    Menu.AddSubMenu arr(i).ProcName, arr(i).Shortcut
            End Select
        Next
    ElseIf AddOrDel = "Del" Then
        On Error Resume Next
        Menu.RemoveMenu
        On Error GoTo 0
        Set Menu = Nothing
    End If
End Sub

Rem 指定モジュール内のプロシージャの一覧を配列で返す
Private Function GetInstructions(cmod As CodeModule) As VbeMenuItemInstructions()
    Dim psl As Long, pbl As String
    Dim ret() As VbeMenuItemInstructions: ReDim ret(0)
    Dim i As Long
    For i = 1 To cmod.CountOfLines
        Dim pname As String
        pname = cmod.ProcOfLine(i, vbext_pk_Proc)
        If pname <> "" Then
            psl = cmod.ProcBodyLine(pname, vbext_pk_Proc)
            If i = psl Then
                pbl = cmod.Lines(psl, 1)
                Set ret(UBound(ret)) = New VbeMenuItemInstructions

                On Error Resume Next
                    ret(UBound(ret)).Shortcut = Split(pbl, "'")(1)
                On Error GoTo 0

                ret(UBound(ret)).ProcName = pname
                ReDim Preserve ret(UBound(ret) + 1)
            End If
        End If
    Next
    ReDim Preserve ret(UBound(ret) - 1)
    GetInstructions = ret
End Function

