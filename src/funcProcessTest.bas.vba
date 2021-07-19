Attribute VB_Name = "funcProcessTest"
Rem https://exceldevelopmentplatform.blogspot.com/2019/01/vba-code-to-get-excel-word-powerpoint.html#Word
Rem ¬Œ÷
Option Explicit
Option Private Module
 
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" ( _
    ByVal hWnd As LongPtr, ByVal dwId As Long, ByRef riid As Any, ByRef ppvObject As Object) As Long
 
Private Declare PtrSafe Function FindWindowEx Lib "User32" Alias "FindWindowExA" _
    (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
 

Private Declare PtrSafe Function IIDFromString Lib "ole32" _
        (ByVal lpsz As Any, ByRef lpiid As Any) As Long

Private Type GUID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(7) As Byte
End Type

Private Const S_OK As Long = &H0
Private Const IID_IDISPATCH As String = "{00020400-0000-0000-C000-000000000046}"
Private Const OBJID_NATIVEOM As Long = &HFFFFFFF0

Sub testWord()
Dim i As Long
Dim hWinWord As LongPtr
Dim wordApp As Object
Dim doc As Object
    'Below line is finding all my Word instances
    hWinWord = FindWindowEx(0&, 0&, "OpusApp", vbNullString)
    While hWinWord > 0
        i = i + 1
        '########Successful output
        Debug.Print "Instance_" & i; hWinWord
        '########Instance_1 2034768
        '########Instance_2 3086118
        '########Instance_3 595594
        '########Instance_4 465560
        '########Below is the problem
        If GetWordapp(hWinWord, wordApp) Then
            Stop
            For Each doc In wordApp.documents
                Debug.Print , doc.Name
            Next
        End If
        hWinWord = FindWindowEx(0, hWinWord, "OpusApp", vbNullString)
    Wend
End Sub

Function GetWordapp(hWinWord As LongPtr, wordApp As Object) As Boolean
Dim hWinDesk As LongPtr, hWin7 As LongPtr
Dim obj As Object
Dim iid As GUID

    Call IIDFromString(StrPtr(IID_IDISPATCH), iid)
    hWinDesk = FindWindowEx(hWinWord, 0&, "_WwF", vbNullString)
   '########Return 0 for majority of classes; only for _WwF it returns other than 0
    hWin7 = FindWindowEx(hWinDesk, 0&, "_WwB", vbNullString)
   '########Return 0 for majority of classes; only for _WwB it returns other than 0
    If AccessibleObjectFromWindow(hWin7, OBJID_NATIVEOM, iid, obj) = S_OK Then
   '########Return -2147467259 and does not get object...
        Set wordApp = obj.Application
        GetWordapp = True
    End If
End Function



Public Function GetExcelAppObjectByIAccessible() As Object
    Dim GUID(0 To 3) As Long, acc As Object
    GUID(0) = &H20400: GUID(1) = &H0: GUID(2) = &HC0: GUID(3) = &H46000000
 
    Dim alHandles(0 To 2) As Long
    alHandles(0) = FindWindowEx(0, 0, "XLMAIN", vbNullString)
    alHandles(1) = FindWindowEx(alHandles(0), 0, "XLDESK", vbNullString)
    alHandles(2) = FindWindowEx(alHandles(1), 0, "EXCEL7", vbNullString)
    If AccessibleObjectFromWindow(alHandles(2), -16&, GUID(0), acc) = 0 Then
        Set GetExcelAppObjectByIAccessible = acc.Application
    End If
End Function
 
 
Public Function GetWordAppObjectByIAccessible() As Object
    Dim GUID(0 To 3) As Long, acc As Object
    GUID(0) = &H20400: GUID(1) = &H0: GUID(2) = &HC0: GUID(3) = &H46000000
 
    Dim alHandles(0 To 3) As Long
    alHandles(0) = FindWindowEx(0, 0, "OpusApp", vbNullString)
    alHandles(1) = FindWindowEx(alHandles(0), 0, "_WwF", vbNullString)
    alHandles(2) = FindWindowEx(alHandles(1), 0, "_WwB", vbNullString)
    alHandles(3) = FindWindowEx(alHandles(2), 0, "_WwG", vbNullString)
    If AccessibleObjectFromWindow(alHandles(3), -16&, GUID(0), acc) = 0 Then
        Set GetWordAppObjectByIAccessible = acc.Application
    End If
End Function
 
 
Public Function GetPowerPointAppObjectByIAccessible() As Object
    Dim GUID(0 To 3) As Long, acc As Object
    GUID(0) = &H20400: GUID(1) = &H0: GUID(2) = &HC0: GUID(3) = &H46000000
 
    Dim alHandles(0 To 2) As Long
    alHandles(0) = FindWindowEx(0, 0, "PPTFrameClass", vbNullString)
    alHandles(1) = FindWindowEx(alHandles(0), 0, "MDIClient", vbNullString)
    alHandles(2) = FindWindowEx(alHandles(1), 0, "mdiClass", vbNullString)
    If AccessibleObjectFromWindow(alHandles(2), -16&, GUID(0), acc) = 0 Then
        Set GetPowerPointAppObjectByIAccessible = acc.Application
    End If
End Function
 
Sub TestGetExcelAppObjectByIAccessible()
    Dim obj As Object
    Set obj = GetExcelAppObjectByIAccessible()
    Debug.Print obj.Name
End Sub
 
 
Sub TestGetWordAppObjectByIAccessible()
    Dim obj As Object
    Set obj = GetWordAppObjectByIAccessible()
    Debug.Print obj.Name
End Sub
 
 
Sub TestGetPowerPointAppObjectByIAccessible()
    Dim obj As Object
    Set obj = GetPowerPointAppObjectByIAccessible()
    Debug.Print obj.Name
End Sub
