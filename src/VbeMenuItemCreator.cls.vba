VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbeMenuItemCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        VbeMenuItemCreator
Rem
Rem  @description   VBEメニューバーイベントハンドラ
Rem
Rem  @update        2020/08/01
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Private MenuTag As String
Private RootMenu As CommandBarPopup
Private EventHandlers As Collection
Private MenuMacroComponentFullName As String

Public Sub Init(tag As String, rootCaption As String, vbc As VBComponent)
    MenuTag = tag
    Call Me.RemoveMenu
    Set EventHandlers = New Collection
    
    Dim VBEMenuBar As CommandBar
    Set VBEMenuBar = Application.VBE.CommandBars(1)
    
    With CreateObject("Scripting.FileSystemObject")
        MenuMacroComponentFullName _
            = "'" & .GetFileName(vbc.Collection.Parent.FileName) & "'!" & vbc.Name
    End With
    
    Set RootMenu = VBEMenuBar.Controls.Add(Type:=msoControlPopup)
    RootMenu.Caption = rootCaption
    RootMenu.tag = MenuTag
End Sub

Public Sub AddSubMenu(ProcName As String, Shortcut As String)
    Dim SubMenu As CommandBarControl
    Set SubMenu = RootMenu.Controls.Add
    
    With SubMenu
        .Caption = ProcName & "(&" & Shortcut & ")"
        .BeginGroup = False
        .OnAction = MenuMacroComponentFullName & "." & ProcName
    End With
    
    With New VbeMenuItemEventHandler
        Set .MenuEvent = Application.VBE.Events.CommandBarEvents(SubMenu)
        EventHandlers.Add .Self
    End With
End Sub

Public Sub RemoveMenu()
    'RootMenu.Deleteとする代わりに、わざわざMenuTagで検索して消すのは、
    '前回の異常終了で残ってしまったメニューも片づけるため。
    Dim MyMenu As CommandBarControl
    Set MyMenu = Application.VBE.CommandBars.FindControl(tag:=MenuTag)
    Do Until MyMenu Is Nothing
        MyMenu.Delete
        Set MyMenu = Application.VBE.CommandBars.FindControl(tag:=MenuTag)
    Loop
    Set EventHandlers = Nothing
End Sub
