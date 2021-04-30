Attribute VB_Name = "kccWsFuncRegExp"
Option Explicit

Rem マッチするか
Rem
Rem  @param strSource       調査対象文字列
Rem  @param strPattern      検査パターン
Rem
Rem  @return As Boolean     True:マッチした。False:マッチしなかった
Rem
Function RegexIsMatch(strSource As String, strPattern As String) As Boolean
    Dim re As RegExp
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = strPattern       ''検索パターンを設定
        .IgnoreCase = True          ''大文字と小文字を区別しない
        .Global = True              ''文字列全体を検索
        RegexIsMatch = re.Test(strSource)
    End With
End Function

Sub Test_RegexIsMatch()
    Const src = "abc def ghi jkl abc ghi"
    Debug.Print RegexIsMatch(src, "abc")
    Debug.Print RegexIsMatch(src, "dgh")
End Sub

Rem マッチした文字列を置換
Rem
Rem  @param strSource       調査対象文字列
Rem  @param strPattern      検査パターン
Rem  @param strReplace      置換文字列
Rem
Rem  @return As String      置換後の文字列
Rem
Function RegexReplace(strSource As String, strPattern As String, strReplace As String) As String
    Dim re As RegExp
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = strPattern       ''検索パターンを設定
        .IgnoreCase = True          ''大文字と小文字を区別しない
        .Global = True              ''文字列全体を検索
        RegexReplace = re.Replace(strSource, strReplace)
    End With
End Function

Sub Test_RegexReplace()
    Const src = "abc def ghi jkl abc ghi"
    Debug.Print RegexReplace(src, "abc", "XXX")
    Debug.Print RegexReplace(src, "xyz", "XXX")
End Sub

Rem マッチした箇所を配列で返す
Rem
Rem  @param strSource       調査対象文字列
Rem  @param strPattern      検査パターン
Rem  @param strProperty     取得したいプロパティ
Rem
Rem  @return As VBScript_RegExp_55.MatchCollection
Rem                         プロパティ未指定ではmcコレクションをそのまま返す
Rem
Function RegexMatches(strSource As String, strPattern As String, strProperty As String) As Variant
    Dim re As RegExp
    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = strPattern       ''検索パターンを設定
        .IgnoreCase = True          ''大文字と小文字を区別しない
        .Global = True              ''文字列全体を検索
        
        Dim mc As VBScript_RegExp_55.MatchCollection
        Set mc = re.Execute(strSource)
        If strProperty = "" Then Set RegexMatches = mc: Exit Function
        If strProperty = "Count" Then RegexMatches = mc.Count: Exit Function
        If mc.Count = 0 Then: RegexMatches = Array(): Exit Function
        
        Dim arr()
        ReDim arr(0 To mc.Count - 1)
        Dim i As Long
        For i = 0 To mc.Count - 1
            If strProperty = "SubMatches" Then
                Dim sm As VBScript_RegExp_55.SubMatches
                Set sm = mc.Item(i).SubMatches
                Dim subarr()
                ReDim subarr(0 To sm.Count - 1)
                Dim j As Long
                For j = 0 To sm.Count - 1
                    subarr(j) = sm.Item(j)
                Next
                arr(i) = subarr
            Else
                arr(i) = CallByName(mc.Item(i), strProperty, VbGet)
            End If
        Next
        RegexMatches = arr
    End With
End Function

Rem マッチした箇所の個数
Function RegexMatchCount(strSource As String, strPattern As String)
    RegexMatchCount = RegexMatches(strSource, strPattern, "Count")
End Function

Rem マッチした箇所の開始インデックス配列
Function RegexMatchIndexs(strSource As String, strPattern As String)
    RegexMatchIndexs = RegexMatches(strSource, strPattern, "FirstIndex")
End Function

Rem マッチした箇所の文字列長配列
Function RegexMatchLengths(strSource As String, strPattern As String)
    RegexMatchLengths = RegexMatches(strSource, strPattern, "Length")
End Function

Rem マッチした箇所の値配列
Function RegexMatchValues(strSource As String, strPattern As String)
    RegexMatchValues = RegexMatches(strSource, strPattern, "Value")
End Function

Sub Test_RegexMatches()
    Const src = "aabbcc axxyyzzc ghi jkl abbaac ghi"
    Const ptn = "a.+?c" '「a」で始まり「c」で終わる文字列（最短）に一致
    Debug.Print RegexMatchCount(src, ptn)
    Debug.Print Join(RegexMatchIndexs(src, ptn), ",")
    Debug.Print Join(RegexMatchLengths(src, ptn), ",")
    Debug.Print Join(RegexMatchValues(src, ptn), ",")
End Sub

Rem マッチした箇所の配列のサブマッチ配列
Function RegexSubMatches(strSource As String, strPattern As String)
    RegexSubMatches = RegexMatches(strSource, strPattern, "SubMatches")
End Function

Sub Test_RegexSubMatches()
    Const src = "AAAAA BB001 AA202 jk345 abcde i030k X12345"
    Const ptn = "([A-Z]+)([0-9]+)" '「アルファベット大文字のグループ」「数値のグループ」に一致
    
    Dim jagArr
    jagArr = RegexSubMatches(src, ptn)
    Stop
End Sub
