VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbProcInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        VbProcInfo
Rem
Rem  @description   VBプログラムのプロシージャ情報
Rem
Rem  @update        2020/08/07
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem  @refModules
Rem    kccFuncString
Rem    VbProcParamInfo
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Private ModName_ As String
Private ProcName_ As String
Private ProcKind_ As Long
Private LineNo_ As Long
Private Source_ As String
Private Comment_ As String
Private ParamsText As String
Private Params_ As Collection
Private Return_ As String

Public Property Get ModName() As String: ModName = ModName_: End Property
Public Property Get ProcName() As String: ProcName = ProcName_: End Property
Public Property Get ProcKind() As Long: ProcKind = ProcKind_: End Property
Public Property Get LineNo() As Long: LineNo = LineNo_: End Property
Public Property Get Source() As String: Source = Source_: End Property
Public Property Get Comment() As String: Comment = Comment_: End Property

Rem プロシージャ宣言文字列からオブジェクト作成
Public Function Init(modname__, ProcName__, ProcKind__, LineNo__, comment__, proc_defined_str) As VbProcInfo
    If Me Is VbProcInfo Then
        With New VbProcInfo
            Set Init = .Init(modname__, ProcName__, ProcKind__, LineNo__, comment__, proc_defined_str)
        End With
        Exit Function
    End If
    Set Init = Me
    
    ModName_ = modname__
    ProcName_ = ProcName__
    ProcKind_ = ProcKind__
    LineNo_ = LineNo__
    Comment_ = comment__
    Source_ = proc_defined_str
    
    'ParamsText
    'Params_
    'Return_
    Call SetProcParse(proc_defined_str)
End Function

Rem プロシージャ宣言文字列からパラメータ部の文字列を取得する
Rem
Rem  @param proc_defined_str    プロシージャ宣言文字列
Rem
Rem  @return As String          パラメータ部の文字列
Rem
Private Function SetProcParse(ByVal proc_defined_str) As String
    If InStr(proc_defined_str, ":") > 0 Then proc_defined_str = Left(proc_defined_str, InStr(proc_defined_str, ":") - 1)
    If InStr(proc_defined_str, "'") > 0 Then proc_defined_str = Left(proc_defined_str, InStr(proc_defined_str, "'") - 1)

    Dim kakkos: kakkos = VBA.Array("{", "(")
    
    Dim repedText As String
    repedText = kccFuncString.ReplaceBracketsNest(proc_defined_str, "(", kakkos)
    
    'パラメータ部の文字列
    ParamsText = ""
    Dim paramsOrReturns
    paramsOrReturns = kccFuncString.SplitWithInBrackets(repedText, kakkos(0), True)
    If UBound(paramsOrReturns) >= 0 Then
        ParamsText = paramsOrReturns(0)
    End If
    'パラメータ毎のクラスオブジェクトのコレクション
    Set Params_ = CreateVbProcParamInfo(ParamsText)
    
    If repedText <> "" Then
        '関数名：「 ～{」
        Dim txt As String
        txt = Replace(repedText, "}", "{", 1, 1)
        Dim blocks
        blocks = Split(txt, "{", 3)
        ProcName_ = kccFuncString.RightStrRev(blocks(0), " ")
    
        '戻値名：「} As ～{}:」
        Return_ = Replace(blocks(UBound(blocks)), " As ", "")
        Return_ = Replace(Replace(Return_, "{", "("), "}", ")")
    End If
    
'    Debug.Print ProcName, "|", ParamsText, "|", Return_
'    Stop
End Function

Rem プロシージャ宣言文字列からパラメータ部のクラスオブジェクトコレクションを取得する
Rem
Rem  @param proc_defined_str                プロシージャ宣言文字列
Rem
Rem  @return As Collection/VbProcParamInfo  パラメータ部の文字列
Rem
Private Function CreateVbProcParamInfo(ParamsText) As Collection
    Dim i As Long
    Dim ret As New Collection
    
    'パラメータ毎の文字列配列
    Dim params: params = Split(vbNullString)
    If ParamsText <> "" Then
        params = Split(ParamsText, ",")
        For i = LBound(params) To UBound(params)
            params(i) = Trim(params(i))
        Next
    End If
    
    'パラメータ毎のクラスオブジェクトのコレクション
    If UBound(params) >= 0 Then
        For i = LBound(params) To UBound(params)
            ret.Add VbProcParamInfo.Init(params(i))
        Next
    End If
    Set CreateVbProcParamInfo = ret
End Function

Function params(idx) As VbProcParamInfo
    Set params = Me.params(idx)
End Function

Public Property Get ProcKindName() As String
    Select Case ProcKind
        Case 0
            If InStr(" " & Source, " Sub ") > 0 Then
                ProcKindName = "Sub"
            Else
                ProcKindName = "Function"
            End If
        Case 1
            ProcKindName = "Property Let"
        Case 2
            ProcKindName = "Property Set"
        Case 3
            ProcKindName = "Property Get"
    End Select
End Property

Public Property Get Scope() As String
    Select Case True
        Case Trim(Source) Like "Private *"
            Scope = "Private"
        Case Trim(Source) Like "Friend *"
            Scope = "Friend"
        Case Trim(Source) Like "Static *"
            Scope = "Static"
        Case Else
            Scope = "Public"
    End Select
End Property

Public Function ParamsToString(Optional Delimiter = " ,") As String
    If Params_.Count = 0 Then Exit Function
    Dim ps
    ReDim ps(1 To Params_.Count)
    Dim i As Long
    For i = LBound(ps) To UBound(ps)
        ps(i) = Params_(i).ToString
    Next
    ParamsToString = Join(ps, Delimiter)
End Function

Public Function ReturnToString() As String
    ReturnToString = Return_
End Function
