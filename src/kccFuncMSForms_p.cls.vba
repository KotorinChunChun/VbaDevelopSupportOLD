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
Rem  @description   MSForms�̃C�P�ĂȂ��R���g���[�����A�C�C�����Ɏg�����߂̊֐��Q
Rem                 ���璊�o�����ꕔ�̊֐�
Rem
Rem  @update        2021/07/10
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem  @references
Rem    �s�v
Rem
Rem --------------------------------------------------------------------------------
Rem  @refModules
Rem    �s�v
Rem
Rem --------------------------------------------------------------------------------
Rem  @history
Rem    2009/  /      �ߋ��̗����͏���
Rem    2020/05/15    �X�V
Rem    2021/07/10    ���X�g�{�b�N�X�֘A�̊֐���ǉ�
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

Private Const SWP_NOSIZE = &H1       '�T�C�Y�ύX���Ȃ�
Private Const SWP_NOMOVE = &H2       '�ʒu�ύX���Ȃ�
Private Const SWP_SHOWWINDOW = &H40  '�E�B���h�E��\��

Private Const hWnd_TOP = 0
Private Const hWnd_BOTTOM = 1
Private Const hWnd_TOPMOST = -1
Private Const hWnd_NOTOPMOST = -2

Private Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"

Rem Control�I�u�W�F�N�g����̃L���X�g�֐�
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

Rem �t�H�[������ɍőO�ʂɕ\��
Rem
Rem  @param  fm          ���[�U�[�t�H�[���I�u�W�F�N�g
Rem  @param  top_most    �őO�ʕ\�����邩�ۂ�
Rem
Rem  @note
Rem    MSForms.UserForm�^�ɃL���X�g����ƁAfm.Caption���������l��Ԃ��Ȃ����ߋ֎~
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

Rem ���X�g�{�b�N�X�ɃA�C�e����ǉ�����
Rem   @param lb                    �Ώ�ListBox
Rem   @param insertRowData         �P�Ƃ̕����� or �ꎟ���z��
Rem   @param insertRowIndex        �}������s�C���f�b�N�X�i0~�j�i����F-1 �Ō�ɒǉ��j
Rem   @param isIfUnique            �l�����j�[�N�i�����ƈ�v����s���Ȃ��j�Ȃ�ǉ�
Rem   @param isSelect              �ǉ������A�C�e����I����Ԃɂ���iMultiSelect�̏�Ԃɒ��Ӂj
Rem
Rem   @return As Long              �}�����ꂽ�s�C���f�b�N�X or �����̓���l�̍s�C���f�b�N�X(0~)
Rem
Rem   @note
Rem     �W����AddItem���\�b�h�͔z��ɑΉ����Ă��Ȃ����ߕK�v
Rem     �n���ꂽ�z��̗v�f����ColumnCount�𒴂��Ă��Ă��؂�̂Ă���
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

Rem ���X�g�{�b�N�X�̑I�𒆃A�C�e�����폜����
Rem   @param lb                    �Ώ�ListBox
Public Sub ListBox_RemoveSelectedItems(lb As MSForms.ListBox)
    Dim i As Long
    For i = lb.ListCount - 1 To 0 Step -1
        If lb.Selected(i) Then lb.RemoveItem i
    Next
End Sub

Rem ���X�g�{�b�N�X�̑I�𒆃A�C�e����1��Ɉړ�����
Rem   @param lb                    �Ώ�ListBox
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

Rem ���X�g�{�b�N�X�̑I�𒆃A�C�e����1���Ɉړ�����
Rem   @param lb                    �Ώ�ListBox
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

Rem ���X�g�{�b�N�X�̑I���A�C�e�����s�f�B�N�V���i���Ŏ擾
Rem
Rem   @param lb                     �Ώ�ListBox
Rem   @param column_index           �擾�����ԍ�0~�i�ȗ����͑S�Ă̗�̈ꎟ���z��0~���擾�j
Rem
Rem   @return As Dictionary(row)    �I���A�C�e���̃f�B�N�V���i���BKey:�s�ԍ�0~ Val:�s�̒lor�s�̔z��
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

Rem ���X�g�{�b�N�X�̑I���A�C�e���̐擪���z��Ŏ擾
Rem
Rem @param lb   ���X�g�{�b�N�X�I�u�W�F�N�g
Rem
Rem @return As Variant/Variant(0 to #)  �I�𒆂̃A�C�e���̐擪��̔z��
Rem                                      ��I����:�v�f0�̔z��
Rem
Rem @note
Rem      ��������F���Ȃ̂ŏd���A�C�e���͖������ɑS�Ď擾���܂��B
Rem      ���d�������e�ł��Ȃ��ꍇ��Indexs�̕����g�p���Ă��������B
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
