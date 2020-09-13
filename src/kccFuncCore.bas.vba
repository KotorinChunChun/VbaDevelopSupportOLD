Attribute VB_Name = "kccFuncCore"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccFuncCore
Rem
Rem  @description   必須関数だけを集めたモジュール
Rem
Rem  @update        2020/09/09
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Rem バイナリ配列から指定したデータの開始する位置を返す
Rem
Rem  @param sourceBytes     検索元バイナリデータ配列
Rem  @param findData        検索対象データ
Rem  @param startIndex      書込み開始位置のインデックス 0~
Rem
Rem  @return As Long        配列のうち一致した箇所の先頭要素番号
Rem
Function IndexOfBinary(sourceBytes() As Byte, findData, Optional startIndex) As Long
    Dim findBytes() As Byte
    Select Case TypeName(findData)
        Case "String"
            'UNICODE >> JIS
            findBytes = StrConv(findData, vbFromUnicode)
        Case "Byte"
            ReDim findBytes(0 To 0)
            findBytes(0) = findData
        Case "Byte()"
            findBytes = findData
        Case Else
            Stop
    End Select
    
    Dim ret As Boolean
    Dim i As Long, j As Long
    If VBA.IsMissing(startIndex) Then startIndex = LBound(sourceBytes)
    For i = startIndex To UBound(sourceBytes)
        For j = LBound(findBytes) To UBound(findBytes)
            ret = False
            If sourceBytes(i + j) <> findBytes(j) Then Exit For
            ret = True
        Next
        If ret Then IndexOfBinary = i: Exit Function
    Next
    IndexOfBinary = -1
End Function

Rem バイナリ配列の指定したインデックス以降にバイナリ配列をコピーする
Rem
Rem  @param arrBytes()          書込み先バイト配列 0~
Rem  @param writeBytes()        書込みたい梅雨と配列
Rem  @param startIndex          書込み開始位置のインデックス 0~
Rem
Sub WriteBinary(ByRef arrBytes() As Byte, writeBytes() As Byte, startIndex)
    Dim i As Long
    For i = LBound(writeBytes) To UBound(writeBytes)
        arrBytes(startIndex + i) = writeBytes(i)
    Next
End Sub

Rem バイト配列データをデバッグ用文字列に変換
Rem
Rem  @param bData               何らかのデータ
Rem
Rem  @return As String          変換後の文字列
Rem
Function ToStringByte(bData) As String
    Select Case TypeName(bData)
        Case "Byte"
            ToStringByte = Right("   " & bData, 4) & " - 0x" & Right("00" & Hex(bData), 2) & " - Char[" & Chr(bData) & "]"
        Case "Byte()"
            Dim i As Long
            Dim arr()
            ReDim arr(LBound(bData) To UBound(bData))
            For i = LBound(bData) To UBound(bData)
                arr(i) = ToStringByte(bData(i))
            Next
            ToStringByte = Join(arr, vbLf)
        Case Else
            Stop
    End Select
End Function

Rem 1バイト読込 - デバッグ用文字列出力版
Function ReadByteToString(bfr As kccBinaryFileIO, Optional FileIndex) As String
    ReadByteToString = ToStringByte(bfr.ReadByte(FileIndex))
End Function

Rem 指定サイズをバイト配列に読込 - デバッグ用文字列出力版
Function ReadBytesToString(bfr As kccBinaryFileIO, Optional FileIndex, Optional ReadSize = 1) As Byte()
    ReadBytesToString = ToStringByte(bfr.ReadBytes(FileIndex, ReadSize))
End Function

Rem 一度にすべて読み込んで出力
Rem
Rem  @param arrBytes()          出力したいバイト配列
Rem  @param BreakCount          イミディエイトウィンドウに出力するデータ件数
Rem
Sub DebugPrintByteArray(arrBytes() As Byte, Optional BreakCount)
    
    Debug.Print "----------DebugPrintByteArray----------"
    Debug.Print "No.      - 10進 - 16進 - 文字列"
    Dim i As Long
    If VBA.IsMissing(BreakCount) Then BreakCount = UBound(arrBytes)
    For i = 0 To BreakCount
        '// 現ループの配列値を取得
        Dim bData
        bData = arrBytes(i)
        
        '// 改行コードの場合
        If bData = 10 Or bData = 13 Then
            Debug.Print "改行です"
        End If
        
        '// 出力
        Debug.Print "No." & Left(i & String(5, " "), 5) & " - " & ToStringByte(bData)
        DoEvents
    Next
    Debug.Print
End Sub
