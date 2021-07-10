VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kccFuncArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccFuncArray
Rem
Rem  @description   �z��^�R���N�V�����^��������^WSF�݊��֐�
Rem
Rem  @update        2020/08/07
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
Rem    2009/  /   start  �ߋ��̗����͏���
Rem    2019/01/28 clean  ���W���[����������
Rem    2019/01/30 fix    Transpose�o�O�C��
Rem    2019/02/08 fix    �z���O������ǉ�
Rem    2019/03/19 clean  FuncVBF��FuncString���琶��
Rem    2019/05/08 add    ArrayToCollection ��ǉ��B�֘A�֐����C��
Rem    2019/09/26 update REPT�֐����X�V�A�֐�����Wsf_��t�^
Rem    2019/09/26 clean  �O�����W���[���ւ̈ˑ������S�ɐ؂藣��
Rem    2019/09/30 clean  ���W���[����Ɨ� Excel.Application�Ɛؒf�o���Ă��炸�B
Rem    2020/02/09 fix    Join2�o�O�C��
Rem    2020/02/24 clean  ���W���[�������AJoin�֘A������
Rem    2020/02/29 fix    Wsf_Transpose��1�������񎟌��̃o�O�C��
Rem    2020/03/05 split  Wsf��؂藣��
Rem    2020/07/08 clean  ���W���[������
Rem    2020/07/18 add    SetArr,GetArr,LBT,UBT�֐�(Core����)
Rem    2020/08/07 merge  FuncWsf�֐��̑唼
Rem
Rem --------------------------------------------------------------------------------
Rem  @functions
Rem
Rem --------------------------------------------------------------------------------
Rem  @note
Rem    ��������
Rem
Rem --------------------------------------------------------------------------------

Option Explicit
Option Compare Binary   '�����ɍl������B�f�t�H���g�l

Rem --------------------------------------------------------------------------------
Rem FuncCore����

Rem �ϐ��ɒl����
Rem
Rem  @creator https://qiita.com/nukie_53/items/bde16afd9a6ca789949d
Rem
Rem  @param outVariable     �o�͐�ϐ�
Rem  @param inExpression    �������ݓ��e
Rem
Rem  @example SetVar(out) = in
Rem
Private Property Let SetVar(outVariable As Variant, inExpression As Variant)
    If VBA.IsObject(inExpression) Then
        Set outVariable = inExpression
    ElseIf VBA.VarType(inExpression) = vbDataObject Then
        Set outVariable = inExpression
    Else
        Let outVariable = inExpression
    End If
End Property

Rem --------------------------------------------------------------------------------
Rem �����{��

Rem �z��̎����������߂�
Rem
Rem  @param arr         �Ώ۔z��
Rem
Rem  @return As Long    ������
Rem
Rem  @example
Rem     Dim arr
Rem
Rem  @note
Rem    ���� GetDim
Rem    ���� GetDimension
Rem    �ʖ� ArrRank By Ariawase
Public Function GetArrayDimension_NoAPI(ByRef arr As Variant) As Long
    On Error GoTo ENDPOINT
    Dim i As Long, tmp As Long
    For i = 1 To 61
        tmp = LBound(arr, i)
    Next
    GetArrayDimension_NoAPI = 0
    Exit Function
    
ENDPOINT:
    GetArrayDimension_NoAPI = i - 1
End Function

Rem ������z��̍��E�ɕ������A��
Public Function Concat(obj, Optional left_add_str, Optional right_add_str) As Variant
    Dim itm
    If IsMissing(left_add_str) Then left_add_str = ""
    If IsMissing(right_add_str) Then right_add_str = ""
    
    Dim tn As String: tn = TypeName(obj)
    Select Case tn
        Case "Collection"
            Dim cll As Collection: Set cll = New Collection
            For Each itm In obj
                cll.Add left_add_str & itm & right_add_str
            Next
            Set Concat = cll
        Case "Dictionary"
            Dim dic As Dictionary: Set dic = New Dictionary
            For Each itm In obj.Keys
                dic.Add itm, left_add_str & obj(itm) & right_add_str
            Next
            Set Concat = dic
        Case "Variant()", "String()", "Long()"
            Dim arr, i As Long, j As Long
            Select Case GetArrayDimension_NoAPI(obj)
                Case 1
                    ReDim arr(LBound(obj, 1) To UBound(obj, 1))
                    For i = LBound(obj, 1) To UBound(obj, 1)
                        arr(i) = left_add_str & obj(i) & right_add_str
                    Next
                Case 2
                    ReDim arr(LBound(obj, 1) To UBound(obj, 1), LBound(obj, 2) To UBound(obj, 2))
                    For i = LBound(obj, 1) To UBound(obj, 1)
                        For j = LBound(obj, 2) To UBound(obj, 2)
                            arr(i, j) = left_add_str & obj(i, j) & right_add_str
                        Next
                    Next
                Case Else
                    '3�����ȏ��Ή�
                    Stop
            End Select
            Let Concat = arr
        Case Else
            If IsObject(obj) Then
                '�I�u�W�F�N�g��Ή�
                Stop
            Else
                Let Concat = left_add_str & obj & right_add_str
            End If
    End Select
End Function

Rem �R���N�V�������ꎟ���z��ɕϊ�
Rem
Rem ����
Rem  @param colTarget ���̓R���N�V����
Rem
Rem �߂�l
Rem  @return       Variant/Variant(0 To #)
Rem     �v�f����    Variant/Variant(0 To -1)
Rem     ���s        Variant/Variant(0 To -1)
Rem     ������      Variant/Variant(0 To n)
Rem
Public Function CollectionToArray(ByVal colTarget As Collection) As Variant
    
    Dim ret As Variant
    Dim i As Long
    Dim v As Variant
    
    If colTarget Is Nothing Then
        ret = VBA.Array()
        
    ElseIf colTarget.Count = 0 Then
        ret = VBA.Array()
        
    Else
        ReDim ret(1 To colTarget.Count)
        i = LBound(ret)
        For Each v In colTarget
            SetVar(ret(i)) = v
            i = i + 1
        Next
    End If
    
    CollectionToArray = ret
    
End Function
