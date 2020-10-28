VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormDeclareSourceTo64bit 
   Caption         =   "VBA Declare 64bit対応変換ツール"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "FormDeclareSourceTo64bit.frm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FormDeclareSourceTo64bit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Rem VBA Declare宣言 64bit対応変換ツール
Rem
Rem 新規で作成したユーザーフォームのコードに貼り付けで使用する
Rem

Option Explicit

Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000

#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function GetWindowLongPtr Lib "User32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLongPtr Lib "User32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        Private Declare PtrSafe Function GetWindowLongPtr Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLongPtr Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #End If
#Else
    Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function GetActiveWindow Lib "User32" () As LongPtr
#Else
    Private Declare Function GetActiveWindow Lib "User32" () As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function DrawMenuBar Lib "User32" (ByVal hWnd As LongPtr) As Long
#Else
    Private Declare Function DrawMenuBar Lib "User32" (ByVal hWnd As Long) As Long
#End If

Private WithEvents TextBox1 As MSForms.TextBox
Attribute TextBox1.VB_VarHelpID = -1
Private WithEvents TextBox2 As MSForms.TextBox
Attribute TextBox2.VB_VarHelpID = -1
Private WithEvents Label1 As MSForms.Label
Attribute Label1.VB_VarHelpID = -1
Private WithEvents Label2 As MSForms.Label
Attribute Label2.VB_VarHelpID = -1

Private Sub UserForm_Initialize()
    Me.Caption = "VBA Declare宣言 64bit対応変換ツール"
    
    Set Label1 = Me.Controls.Add("Forms.Label.1", "Label1", True)
    Label1.Caption = "変換したいソース"
    
    Set Label2 = Me.Controls.Add("Forms.Label.1", "Label2", True)
    Label2.Caption = "変換されたソース"
    
    Set TextBox1 = Me.Controls.Add("Forms.TextBox.1", "TextBox1", True)
    TextBox1.EnterKeyBehavior = True
    TextBox1.MultiLine = True
    TextBox1.ScrollBars = fmScrollBarsBoth
    TextBox1.WordWrap = False
    
    Set TextBox2 = Me.Controls.Add("Forms.TextBox.1", "TextBox2", True)
    TextBox2.EnterKeyBehavior = True
    TextBox2.MultiLine = True
    TextBox2.ScrollBars = fmScrollBarsBoth
    TextBox2.WordWrap = False
    TextBox2.Locked = True
    TextBox2.BackColor = &H80000004
    
    'イベント発生につき最後に実行
    Me.Width = 800
    Me.Height = 600
End Sub
 
Private Sub UserForm_Activate()
    Call FormSetting
End Sub

' フォームをリサイズ可能にするための設定
Public Sub FormSetting()
    Dim Result As LongPtr
    Dim hWnd As LongPtr
    Dim Wnd_STYLE As LongPtr
 
    hWnd = GetActiveWindow()
    Wnd_STYLE = GetWindowLongPtr(hWnd, GWL_STYLE)
    Wnd_STYLE = Wnd_STYLE Or WS_THICKFRAME Or &H30000
 
    Result = SetWindowLongPtr(hWnd, GWL_STYLE, Wnd_STYLE)
    Result = DrawMenuBar(hWnd)
End Sub

Private Sub UserForm_Resize()
    If TextBox1 Is Nothing Then Exit Sub
    On Error Resume Next
    TextBox1.Left = 10
    TextBox1.Top = 20
    TextBox1.Width = Me.InsideWidth / 2 - 20
    TextBox1.Height = Me.InsideHeight - 40
    Label1.Left = TextBox1.Left
    Label1.Top = 5
    
    TextBox2.Left = Me.InsideWidth / 2 + 10
    TextBox2.Top = 20
    TextBox2.Width = Me.InsideWidth / 2 - 20
    TextBox2.Height = Me.InsideHeight - 40
    Label2.Left = TextBox2.Left
    Label2.Top = 5
End Sub

Private Sub TextBox1_Change()
'    On Error Resume Next
    TextBox2.Text = ConvertVBACodeDeclare(TextBox1.Text, 10)
'    On Error GoTo 0
End Sub
