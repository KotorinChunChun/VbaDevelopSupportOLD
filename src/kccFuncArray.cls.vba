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
Rem  @description   配列／コレクション／辞書操作／WSF互換関数
Rem
Rem  @update        2020/08/07
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
Rem    2009/  /   start  過去の履歴は消失
Rem    2019/01/28 clean  モジュール整理完了
Rem    2019/01/30 fix    Transposeバグ修正
Rem    2019/02/08 fix    配列例外処理を追加
Rem    2019/03/19 clean  FuncVBFとFuncStringから生成
Rem    2019/05/08 add    ArrayToCollection を追加。関連関数を修正
Rem    2019/09/26 update REPT関数を更新、関数名にWsf_を付与
Rem    2019/09/26 clean  外部モジュールへの依存を完全に切り離し
Rem    2019/09/30 clean  モジュールを独立 Excel.Applicationと切断出来ておらず。
Rem    2020/02/09 fix    Join2バグ修正
Rem    2020/02/24 clean  モジュール整理、Join関連見直し
Rem    2020/02/29 fix    Wsf_Transposeで1次元→二次元のバグ修正
Rem    2020/03/05 split  Wsfを切り離し
Rem    2020/07/08 clean  モジュール整理
Rem    2020/07/18 add    SetArr,GetArr,LBT,UBT関数(Core共通)
Rem    2020/08/07 merge  FuncWsf関数の大半
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
Option Compare Binary   '厳密に考慮する。デフォルト値

Rem 配列の次元数を求める
Rem
Rem  @param arr         対象配列
Rem
Rem  @return As Long    次元数
Rem
Rem  @example
Rem     Dim arr
Rem
Rem  @note
Rem    旧名 GetDim
Rem    旧名 GetDimension
Rem    別名 ArrRank By Ariawase
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

Rem 文字列配列の左右に文字列を連結
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
                    '3次元以上非対応
                    Stop
            End Select
            Let Concat = arr
        Case Else
            If IsObject(obj) Then
                'オブジェクト非対応
                Stop
            Else
                Let Concat = left_add_str & obj & right_add_str
            End If
    End Select
End Function
