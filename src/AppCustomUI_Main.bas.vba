Attribute VB_Name = "AppCustomUI_Main"
Option Explicit

Public Sub Rbn_VbaDevelopSupport_アドイン_VBEメニュー再構築_onAction(Control As IRibbonControl)
    Call Reset_Addin
End Sub

Public Sub Rbn_VbaDevelopSupport_アドイン_終了_onAction(Control As IRibbonControl)
    Call Close_Addin
End Sub
