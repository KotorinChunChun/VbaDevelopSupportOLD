VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbProcParamInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        VbProcParamInfo
Rem
Rem  @description   VBプログラムのプロシージャのパラメータ情報
Rem
Rem  @update        2020/08/01
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem  @refModules
Rem    kccFuncString_Partial
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Public argOptional      As String
Public argBy            As String
Public argParamArray    As String
Public argVarName       As String
Public argType          As String
Public argDefaultValue  As String

Public Function ToString() As String
    ToString = kccFuncString_Partial.Trim2to1(Join(VBA.Array( _
                    argOptional, _
                    argBy, _
                    argParamArray, _
                    argVarName, _
                    IIf(argType = "", "", "As " & argType), _
                    IIf(argDefaultValue = "", "", "= " & argDefaultValue)), " "))
End Function

Public Function Init(ByVal base_str) As VbProcParamInfo
    Set Init = Me
    
    Dim txt As String
    txt = " " & base_str
    
    If InStr(txt, " Optional") > 0 Then argOptional = "Optional": txt = Replace(txt, " Optional", "")
    If InStr(txt, " ByVal") > 0 Then argBy = "ByVal": txt = Replace(txt, " ByVal", "")
    If InStr(txt, " ByRef") > 0 Then argBy = "ByRef": txt = Replace(txt, " ByRef", "")
    If InStr(txt, " ParamArray") > 0 Then argParamArray = "ParamArray": txt = Replace(txt, " ParamArray", "")
    
    If InStr(txt, " = ") > 0 Then
        argDefaultValue = Split(txt, " = ")(1)
        txt = Split(txt, " = ")(0)
    End If
    
    If InStr(txt, " As ") > 0 Then
        argType = Split(txt, " As ")(1)
        txt = Split(txt, " As ")(0)
    End If
    
    argVarName = Trim(txt)
    
'    Stop
End Function
