VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kccFuncMSForms_p"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccFuncMSForms
Rem
Rem  @description   MSFormsのイケてないコントロールを、イイ感じに使うための関数群
Rem                 から抽出した一部の関数
Rem
Rem  @update        2021/07/10
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
Rem    不要
Rem
Rem --------------------------------------------------------------------------------
Rem  @history
Rem    2009/  /      過去の履歴は消失
Rem    2020/05/15    更新
Rem    2021/07/10    リストボックス関連の関数を追加
Rem
Rem --------------------------------------------------------------------------------
Rem  @functions
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function SetWindowPos Lib "User32" ( _
                                            ByVal hWnd As LongPtr, _
                                            ByVal hWndInsertAfter As LongPtr, _
                                            ByVal x As Long, _
                                            ByVal y As Long, _
                                            ByVal cx As Long, _
                                            ByVal cy As Long, _
                                            ByVal wFlags As Long _
                                            ) As Long
#Else
    Private Declare Function SetWindowPos Lib "User32" ( _
                                            ByVal hWnd As LongPtr, _
                                            ByVal hWndInsertAfter As LongPtr, _
                                            ByVal x As Long, _
                                            ByVal y As Long, _
                                            ByVal cx As Long, _
                                            ByVal cy As Long, _
                                            ByVal wFlags As Long _
                                            ) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "User32" Alias "FindWindowA" ( _
                                            ByVal lpClassName As String, _
                                            ByVal lpWindowName As String _
                                            ) As Long
#Else
    Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" ( _
                                            ByVal lpClassName As String, _
                                            ByVal lpWindowName As String _
                                            ) As Long
#End If

Private Const SWP_NOSIZE = &H1       'サイズ変更しない
Private Const SWP_NOMOVE = &H2       '位置変更しない
Private Const SWP_SHOWWINDOW = &H40  'ウィンドウを表示

Private Const hWnd_TOP = 0
Private Const hWnd_BOTTOM = 1
Private Const hWnd_TOPMOST = -1
Private Const hWnd_NOTOPMOST = -2

Private Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"

Rem Controlオブジェクトからのキャスト関数
Public Function ToCtrl(Ctrl As MSForms.Control) As MSForms.Control: Set ToCtrl = Ctrl: End Function
Public Function ToCheckBox(Ctrl As MSForms.Control) As MSForms.CheckBox: Set ToCheckBox = Ctrl: End Function
Public Function ToComboBox(Ctrl As MSForms.Control) As MSForms.ComboBox: Set ToComboBox = Ctrl: End Function
Public Function ToCommandButton(Ctrl As MSForms.Control) As MSForms.CommandButton: Set ToCommandButton = Ctrl: End Function
Public Function ToFrame(Ctrl As MSForms.Control) As MSForms.Frame: Set ToFrame = Ctrl: End Function
Public Function ToImage(Ctrl As MSForms.Control) As MSForms.Image: Set ToImage = Ctrl: End Function
Public Function ToLabel(Ctrl As MSForms.Control) As MSForms.Label: Set ToLabel = Ctrl: End Function
Public Function ToListBox(Ctrl As MSForms.Control) As MSForms.ListBox: Set ToListBox = Ctrl: End Function
Public Function ToMultiPage(Ctrl As MSForms.Control) As MSForms.MultiPage: Set ToMultiPage = Ctrl: End Function
Public Function ToOptionButton(Ctrl As MSForms.Control) As MSForms.OptionButton: Set ToOptionButton = Ctrl: End Function
Public Function ToSpinButton(Ctrl As MSForms.Control) As MSForms.SpinButton: Set ToSpinButton = Ctrl: End Function
Public Function ToTabStrip(Ctrl As MSForms.Control) As MSForms.TabStrip: Set ToTabStrip = Ctrl: End Function
Public Function ToTextBox(Ctrl As MSForms.Control) As MSForms.TextBox: Set ToTextBox = Ctrl: End Function
Public Function ToToggleButton(Ctrl As MSForms.Control) As MSForms.ToggleButton: Set ToToggleButton = Ctrl: End Function

Rem フォームを常に最前面に表示
Rem
Rem  @param  fm          ユーザーフォームオブジェクト
Rem  @param  top_most    最前面表示するか否か
Rem
Rem  @note
Rem    MSForms.UserForm型にキャストすると、fm.Captionが正しい値を返さないため禁止
Rem
Public Sub UserForm_TopMost(fm As Object, top_most As Boolean)
    Dim fmHWnd As LongPtr
    fmHWnd = FindWindow(C_VBA6_USERFORM_CLASSNAME, fm.Caption)
    If fmHWnd = 0 Then Debug.Print Err.LastDllError: Err.Raise 9999, , "FindWindow Faild"
    If top_most Then
        Call SetWindowPos(fmHWnd, hWnd_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Else
        Call SetWindowPos(fmHWnd, hWnd_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
    End If
End Sub

Rem リストボックスにアイテムを追加する
Rem   @param lb                    対象ListBox
Rem   @param insertRowData         単独の文字列 or 一次元配列
Rem   @param insertRowIndex        挿入する行インデックス（0~）（既定：-1 最後に追加）
Rem   @param isIfUnique            値がユニーク（既存と一致する行がない）なら追加
Rem   @param isSelect              追加したアイテムを選択状態にする（MultiSelectの状態に注意）
Rem
Rem   @return As Long              挿入された行インデックス or 既存の同一値の行インデックス(0~)
Rem
Rem   @note
Rem     標準のAddItemメソッドは配列に対応していないため必要
Rem     渡された配列の要素数がColumnCountを超えていても切り捨てられる
Rem
Public Function ListBox_AddItem( _
        lb As MSForms.ListBox, _
        insertRowData As Variant, _
        Optional ByVal insertRowIndex As Long = -1, _
        Optional isIfUnique As Boolean = False, _
        Optional isSelect As Boolean = False) As Long
        
    If TypeName(insertRowData) = "Collection" Then
        insertRowData = kccFuncArray.CollectionToArray(insertRowData)
    End If
    
    If isIfUnique Then
        ListBox_AddItem = ListBox_AddItem_Sub(lb, insertRowData)
    Else
        Dim isUnique As Long
        isUnique = True
        
        Dim joinedInsertRowData As String
        If IsArray(insertRowData) Then
            joinedInsertRowData = Strings.Join(insertRowData, "")
        Else
            joinedInsertRowData = insertRowData  'Join
        End If

        Dim i As Long, j As Long
        For i = 0 To lb.ListCount - 1
            Dim strRecord As String: strRecord = ""
            For j = 0 To lb.ColumnCount - 1
                strRecord = strRecord & lb.List(i, j)
            Next
            If strRecord = joinedInsertRowData Then
                isUnique = False
                Exit For
            End If
        Next
        
        If isUnique Then
            ListBox_AddItem = ListBox_AddItem_Sub(lb, insertRowData, insertRowIndex)
        Else
            ListBox_AddItem = i
        End If
    End If
    
    If isSelect Then lb.Selected(ListBox_AddItem) = True
End Function

Public Function ListBox_AddItem_Sub( _
        lb As MSForms.ListBox, _
        arrInsertRowData As Variant, _
        Optional ByVal insertRowIndex As Long = -1) As Long

    If insertRowIndex = -1 Then
        insertRowIndex = lb.ListCount
    End If

    If Not IsArray(arrInsertRowData) Then
        lb.AddItem arrInsertRowData, insertRowIndex
        ListBox_AddItem_Sub = insertRowIndex
        Exit Function
    End If

    lb.AddItem "", insertRowIndex
    Dim ColumnIndex As Long, itemIndex As Long
    itemIndex = LBound(arrInsertRowData)
    For ColumnIndex = 0 To lb.ColumnCount - 1
        If ColumnIndex <= UBound(arrInsertRowData) Then
            lb.List(insertRowIndex, ColumnIndex) = arrInsertRowData(itemIndex)
        End If
        itemIndex = itemIndex + 1
    Next
    ListBox_AddItem_Sub = insertRowIndex
End Function

Rem リストボックスの選択中アイテムを削除する
Rem   @param lb                    対象ListBox
Public Sub ListBox_RemoveSelectedItems(lb As MSForms.ListBox)
    Dim i As Long
    For i = lb.ListCount - 1 To 0 Step -1
        If lb.Selected(i) Then lb.RemoveItem i
    Next
End Sub

Rem リストボックスの選択中アイテムを1つ上に移動する
Rem   @param lb                    対象ListBox
Public Sub ListBox_MoveUpSelectedItems(lb As MSForms.ListBox)
    
    Dim MAX_INDEX As Long: MAX_INDEX = lb.ListCount - 1
    
    Dim i As Long
    For i = 0 To MAX_INDEX - 1
        If Not lb.Selected(i) And lb.Selected(i + 1) Then
            Do
                If i >= MAX_INDEX Then Exit Do
                If Not lb.Selected(i + 1) Then Exit Do
                
                Dim j As Long
                For j = 0 To lb.ColumnCount - 1
                    Dim txt1 As Variant: txt1 = lb.List(i + 0, j)
                    Dim txt2 As Variant: txt2 = lb.List(i + 1, j)
                    lb.List(i + 0, j) = IIf(IsNull(txt2), "", txt2)
                    lb.List(i + 1, j) = IIf(IsNull(txt1), "", txt1)
                Next
                lb.Selected(i + 0) = True
                lb.Selected(i + 1) = False
                
                i = i + 1
            Loop
        End If
    Next
    
End Sub

Rem リストボックスの選択中アイテムを1つ下に移動する
Rem   @param lb                    対象ListBox
Public Sub ListBox_MoveDownSelectedItems(lb As MSForms.ListBox)
    
    Dim MIN_INDEX As Long: MIN_INDEX = 0
    
    Dim i As Long
    For i = lb.ListCount - 1 To MIN_INDEX + 1 Step -1
        If Not lb.Selected(i) And lb.Selected(i - 1) Then
            Do
                If i <= MIN_INDEX Then Exit Do
                If Not lb.Selected(i - 1) Then Exit Do
                
                Dim j As Long
                For j = 0 To lb.ColumnCount - 1
                    Dim txt1 As Variant: txt1 = lb.List(i - 0, j)
                    Dim txt2 As Variant: txt2 = lb.List(i - 1, j)
                    lb.List(i - 0, j) = IIf(IsNull(txt2), "", txt2)
                    lb.List(i - 1, j) = IIf(IsNull(txt1), "", txt1)
                Next
                lb.Selected(i - 0) = True
                lb.Selected(i - 1) = False
                
                i = i - 1
            Loop
        End If
    Next
    
End Sub

Rem リストボックスの選択アイテムを行ディクショナリで取得
Rem
Rem   @param lb                     対象ListBox
Rem   @param column_index           取得する列番号0~（省略時は全ての列の一次元配列0~を取得）
Rem
Rem   @return As Dictionary(row)    選択アイテムのディクショナリ。Key:行番号0~ Val:行の値or行の配列
Rem
Public Function ListBox_GetSelectedItemsDictionary(lb As MSForms.ListBox, Optional column_index) As Dictionary
    Dim retVal As New Dictionary
    Set ListBox_GetSelectedItemsDictionary = retVal
    If lb.ListCount = 0 Then Exit Function

    Dim rowItem()
    Dim i As Long, j As Long
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) Then
            If IsMissing(column_index) Then
                ReDim rowItem(0 To lb.ColumnCount - 1)
                For j = 0 To lb.ColumnCount - 1
                    rowItem(j) = lb.List(i, j)
                Next
                retVal(i) = rowItem
            Else
                retVal(i) = "" & lb.List(i, column_index)
            End If
        End If
    Next

    Set ListBox_GetSelectedItemsDictionary = retVal
End Function

Rem リストボックスの選択アイテムの先頭列を配列で取得
Rem
Rem @param lb   リストボックスオブジェクト
Rem
Rem @return As Variant/Variant(0 to #)  選択中のアイテムの先頭列の配列
Rem                                      非選択時:要素0の配列
Rem
Rem @note
Rem      ※文字列認識なので重複アイテムは無条件に全て取得します。
Rem      ※重複を許容できない場合はIndexsの方を使用してください。
Rem
Public Function ListBox_GetSelectedItems(lb As MSForms.ListBox) As Variant
    ListBox_GetSelectedItems = VBA.Array()
    If lb.ListCount = 0 Then Exit Function

    Dim arr
    ReDim arr(0 To lb.ListCount - 1)
    Dim i As Long, nextIndex As Long

    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) Then
            arr(nextIndex) = lb.List(i)
            nextIndex = nextIndex + 1
        End If
    Next

    Dim listData
    listData = lb.List

    If nextIndex = 0 Then ListBox_GetSelectedItems = VBA.Array(): Exit Function
    ReDim Preserve arr(0 To nextIndex - 1)

    ListBox_GetSelectedItems = arr
End Function
