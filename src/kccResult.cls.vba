VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kccResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsSuccess As Boolean
Public IsAbort As Boolean
Public SuccessValue As Dictionary
Public FailureValue As Dictionary

Rem 関数の実行結果
Rem
Rem  @param isSuccess       True:成功 False:失敗
Rem  @param pSuccessValue   成功時データ
Rem  @param pFailureValue   失敗時データ
Rem
Rem  @return As kccResult 実行結果を示すオブジェクト
Rem
Function Init(pIsSuccess As Boolean, Optional pValue) As kccResult
    If Me Is kccResult Then
        With New kccResult
            Set Init = .Init(pIsSuccess, pValue)
        End With
        Exit Function
    End If
    Set Init = Me
    
    IsSuccess = pIsSuccess
    
    Set SuccessValue = New Dictionary
    Set FailureValue = New Dictionary
    
    If Not VBA.IsMissing(pValue) Then
        Me.Add IsSuccess, pValue
    End If
End Function

Rem 関数の実行結果を追記
Function Add(pIsSuccess As Boolean, pValue, Optional pAbort As Boolean) As kccResult
    Set Add = Me
    If pIsSuccess Then
        SuccessValue.Add SuccessValue.Count + 1, pValue
    Else
        FailureValue.Add FailureValue.Count + 1, pValue
        IsSuccess = False
        If pAbort Then IsAbort = True
    End If
End Function
