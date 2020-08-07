VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbeMenuItemEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        VbeMenuItemEventHandler
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

Public WithEvents MenuEvent As VBIDE.CommandBarEvents
Attribute MenuEvent.VB_VarHelpID = -1

Public Property Get Self() As Object
    Set Self = Me
End Property

Private Sub MenuEvent_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Application.Run CommandBarControl.OnAction
    handled = True
    CancelDefault = True
End Sub
