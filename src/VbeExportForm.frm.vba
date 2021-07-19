VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VbeExportForm 
   Caption         =   "VBA�̃G�N�X�|�[�g�ƃo�b�N�A�b�v"
   ClientHeight    =   12900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15840
   OleObjectBlob   =   "VbeExportForm.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "VbeExportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public st As kccSettings

Public ReturnProject As Object

'Public Property Get ReturnExportBackupBinFolders() As String
'    ReturnExportBackupBinFolders = st.BackupBinFile
'End Property

Private Sub btnBinAdd_Click()
    If tbBin.Text = "" Then Exit Sub
    kccFuncMSForms_p.ListBox_AddItem lbBin, tbBin.Text, isSelect:=True
    tbBin.Text = ""
    tbBin.SetFocus
End Sub

Private Sub btnCancel_Click()
    Set ReturnProject = Nothing
    Unload Me
End Sub

Private Sub btnEditKccsettings_Click()
    SelectedProjectPath.SelectPathToFilePath("kccsettings.json").OpenAssociation
End Sub

Private Sub btnOK_Click()
    If SaveKccsettingsByGui Then
        Me.Hide
        Call VBComponents_BackupAndExport_Sub( _
                ReturnProject, _
                st.ProjectFolder, _
                st.ExportBinFolder, _
                st.ExportSrcFolder, _
                st.ExportBackupBinFolders, _
                st.ExportBackupSrcFolders)
    End If
End Sub

Private Sub btnReloadKccsettings_Click()
    Call LoadGuiByKccsettings
End Sub

Private Sub chkSrcAddExt_Click()
    txtSrcAddExt.Enabled = chkSrcAddExt.Value
End Sub

Private Sub cmbExportProject_Change()
    Call LoadGuiByKccsettings
End Sub

Private Sub spinBinUpDown_SpinDown()
    Call kccFuncMSForms_p.ListBox_MoveDownSelectedItems(lbBin)
End Sub

Private Sub spinBinUpDown_SpinUp()
    Call kccFuncMSForms_p.ListBox_MoveUpSelectedItems(lbBin)
End Sub

Private Sub btnBinDel_Click()
    Call kccFuncMSForms_p.ListBox_RemoveSelectedItems(lbBin)
End Sub

Private Sub btnBinEditKccignore_Click()
    SelectedProjectPath.GetIgnoreFile.OpenAssociation
End Sub

Private Sub btnFormSmall_Click()
    Me.Zoom = Me.Zoom * 0.9
End Sub

Private Sub btnSrcAdd_Click()
    If tbSrc.Text = "" Then Exit Sub
    kccFuncMSForms_p.ListBox_AddItem lbSrc, tbSrc.Text, isSelect:=True
    tbSrc.Text = ""
    tbSrc.SetFocus
End Sub

Private Sub spinSrcUpDown_SpinDown()
    Call kccFuncMSForms_p.ListBox_MoveDownSelectedItems(lbSrc)
End Sub

Private Sub spinSrcUpDown_SpinUp()
    Call kccFuncMSForms_p.ListBox_MoveUpSelectedItems(lbSrc)
End Sub

Private Sub btnSrcDel_Click()
    Call kccFuncMSForms_p.ListBox_RemoveSelectedItems(lbSrc)
End Sub

Private Sub UserForm_Initialize()
    Dim obj As Object: Set obj = Application.VBE.ActiveVBProject
    On Error Resume Next
    Dim fn As String: fn = obj.FileName
    On Error GoTo 0
    cmbExportProject.Style = fmStyleDropDownList
    
    kccFuncMSForms_p.UserForm_TopMost Me, True
    
    Dim pj As VBProject
    For Each pj In GetVBProjects
        cmbExportProject.AddItem kccPath.Init(pj).FullPath
    Next
    If fn <> "" Then cmbExportProject.Text = obj.FileName
    
    tbSrc.AddItem ".\..\src"
    tbSrc.AddItem ".\..\src\[FILENAME]"
    tbSrc.AddItem ".\[YYYYMMDD]_[HHMMSS]_[FILENAME]"
    tbSrc.AddItem ".\..\backup\src\[YYYYMMDD]_[HHMMSS]_[FILENAME]"
    
    tbBin.AddItem ".\..\bin"
    
    Call LoadGuiByKccsettings
End Sub

Property Get SelectedProjectPath() As kccPath
    Set SelectedProjectPath = kccPath.Init(GetVBProjectByPath(cmbExportProject.Text))
End Property

Sub LoadSetting()
    Set st = kccSettings.Init(SelectedProjectPath.FullPath)
End Sub

Rem KccSettings�t�@�C����ǂݍ����GUI�R���g���[���֔��f
Sub LoadGuiByKccsettings()

    Call LoadSetting
    Call st.LoadFile

    Rem �ݒ�UI������(LoadFile�Ŋ���l��ǂ�ł邽�߂��܂�Ӗ��͂Ȃ�)
    lbSrc.MultiSelect = fmMultiSelectMulti
    lbSrc.Clear
    chkSrcIgnoreEmptyFile.Value = True
    chkSrcAddExt.Value = True
    txtSrcAddExt.Value = ".vba"
    chkSrcIncludeCustomUI.Value = False
    
    lbBin.MultiSelect = fmMultiSelectMulti
    lbBin.Clear
    chkBinBackup.Value = False

    Rem �ݒ�l�̓ǂݍ���
    If st Is Nothing Then
        'Create New File?
    ElseIf st.Path = "" Then
        'Create New File?
    Else
        kccFuncMSForms_p.ListBox_AddItem lbSrc, st.ExportSrcFolder, isSelect:=True
        kccFuncMSForms_p.ListBox_AddItem lbSrc, st.ExportBackupSrcFolders, isSelect:=False
        kccFuncMSForms_p.ListBox_AddItem lbBin, st.ExportBinFolder, isSelect:=True
        kccFuncMSForms_p.ListBox_AddItem lbBin, st.ExportBackupBinFolders, isSelect:=False
        
        chkSrcIgnoreEmptyFile.Value = st.IgnoreEmptyModule
        chkSrcAddExt.Value = st.HasExtension
        chkSrcIncludeCustomUI.Value = st.IsExportCustomUI
        chkBinBackup.Value = st.IsBackupProject
        '�{���͑I����⃊�X�g�ƁA�I���ς݃��X�g�͕ʁX�ɕK�v�Ȃ͂��B
    End If

End Sub

Rem KccSettings�t�@�C�����X�V����
Function SaveKccsettingsByGui() As Boolean
    Set ReturnProject = GetVBProjectByPath(cmbExportProject.Text)
    Call LoadSetting
    
    Dim dic As Dictionary
    
    Rem �\�[�X�R�[�h�̃G�N�X�|�[�g�ݒ�
    Rem ��ԍŏ��ɑI������Ă���t�H���_���o�͎��̍�����r�ΏۂƂ���
    Rem ��Ԗڈȍ~����Ԗڂ̕�����t�H���_�Ƃ���
    Set dic = kccFuncMSForms_p.ListBox_GetSelectedItemsDictionary(lbSrc, 0)
    If dic.Count = 0 Then MsgBox "�o�͐��1�ȏ�I�����Ă��������B": Exit Function
    st.ExportSrcFolder = dic.Items(0)
    If dic.Count > 1 Then
        dic.Remove dic.Keys(0)
        st.ExportBackupSrcFolders = dic.Items
    End If
    st.IgnoreEmptyModule = chkSrcIgnoreEmptyFile.Value
    st.HasExtension = chkSrcAddExt.Value
    st.StrExtension = txtSrcAddExt.Text
    st.IsExportCustomUI = chkSrcIncludeCustomUI.Value
    
    Rem �o�C�i���t�@�C���̃R�s�[
    Set dic = kccFuncMSForms_p.ListBox_GetSelectedItemsDictionary(lbBin, 0)
    If dic.Count = 0 Then MsgBox "�o�͐��1�ȏ�I�����Ă��������B": Exit Function
    st.ExportBinFolder = dic.Items(0)
    If dic.Count > 1 Then
        dic.Remove dic.Keys(0)
        st.ExportBackupBinFolders = dic.Items
    End If
    st.IsBackupProject = chkBinBackup.Value
    
    If chkSaveKccsettings.Value Then st.SaveFile
'    Stop
    SaveKccsettingsByGui = True
End Function
