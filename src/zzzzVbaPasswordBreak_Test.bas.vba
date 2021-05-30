Attribute VB_Name = "zzzzVbaPasswordBreak_Test"
Option Explicit
Option Private Module

Private fso As New FileSystemObject

'Const TEST_PATH = "D:\OneDrive\ExcelVBA\BOX\vbaProject.bin"
'Const TEST_PATH2 = "D:\OneDrive\ExcelVBA\BOX\vbaProject2.bin"

'【強引にエクセルのVBAのパスワードを外す方法】
'１．対象のエクセルをバイナリエディタで開く。
'２．文字列「DPB」を検索する。（後ろのほうに１つだけある）
'３．「DPB」を「ｒPB」に変更。（ｒでなくてもなんでもよい）
'４．名前を付けて保存。
'５．エクセル起動。マクロ有効。
'６．「エラーがあります。ロードを続けますか？」のようなメッセージがでる。
'７．「はい」を選択。何回か続くことがあるが全てはい。
'８．VBエディタを起動。
'９．プロジェクトのロックが外れているので、プロパティを確認して、適当なパスワードを入れて保存。
'１０．エクセルを再度開くと、６番のエラーが解決している。
'１１．VBプロジェクトのパスワード（９番で入力したもの）を入れれば、マクロを編集できるようになる。

Rem DPB用の文字列の配列を返す(1~78)
Function GetDPB1234() As Variant

    'a
    Const DPB72 = "1113BDA2DAA2DA5D26A3DADD6F4E46D78D7BC08FE39E4784E3350E9213615C86471D9812"
    Const DPB74 = "AAA8061A49374937B6C94A3786E0802E59A2C114A3763A83BECA00E8B2FB24A3E0766AA8B8"
    Const DPB78 = "1715BB935FB7ACD4ACD4532CADC4D4669508A7726A96A2284017A34DB7ECC4A8F2858EC86E3F2F"
    
    'pass
    Const DPB72_PASS = "888A242B412B41D4BF2C41D42F35B07B8CA050B9C8C160467F7F159257A33C796FD023E9"
    Const DPB74_PASS = "DBD977B3B0D0B0D04F30B1D04FA0C803EC59112508E73281B79EE07431F8025FD652B304CC"
    Const DPB76_PASS = "7C7ED01230574D574DA8B3584DA823618C47D0FC9C1D1C5DBCEADBD3B126BB3798ED8B444F5D"
    
    '1234
    Const DPB72_1234 = "6062CC13E913E9EC1714E94DB4B5BB69EAA21E41EBFCFC944A99D12883ECB411D5C4C725"
    Const DPB74_1234 = "6260CE844BA14BA1B45F4CA1150CFD13A1426AB619731444BC027169003BC46CF98D2C3F8D"
    Const DPB76_1234 = "959739565B76787678898877782825EE08AE357FA52C620973A13594A42DFC132D2E4EE37E44"
    Const DPB78_1234 = "0E0CA28EEEF2390F390FC6F13A0F679EAF816FACBCDC6B29C63A6E0CC7935ECD22FE5F1FB29113"

    Dim arr
    ReDim arr(1 To 78)
    Dim i As Long
    For i = 1 To UBound(arr)
        arr(i) = Space(i)
    Next
    arr(72) = DPB72_1234
    arr(74) = DPB74_1234
    arr(76) = DPB76_1234
    arr(78) = DPB78_1234
    GetDPB1234 = arr
End Function

Rem DPBを書き換える
Rem
Rem @param inVbaProjectPath     DPBをデータに含むファイル
Rem @param outVbaProjectPath    出力ファイル
Rem
Rem @return As Boolean          成功したらTrue
Public Function vbaProjectCrack(inVbaProjectPath As String, outVbaProjectPath As String) As Boolean
    If outVbaProjectPath = "" Then outVbaProjectPath = inVbaProjectPath
    
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    If Not fso.FileExists(inVbaProjectPath) Then
        MsgBox "[" & inVbaProjectPath & "] がありません"
        Exit Function
    End If
    
    Dim f As Scripting.File
    Set f = fso.GetFile(inVbaProjectPath)
    
    Dim ibr As kccBinaryFileIO
    Set ibr = kccBinaryFileIO.OpenFile(f.Path, 1)
'    ibr.FileSeek 182980
'    Debug.Print ibr.ReadByteToString
'    ibr.DebugPrintByteArray 10
    
'    Dim txt As String
'    txt = ibr.ReadAllToString
'    ibr.FileSeek 182981
'    txt = ibr.ReadString(, 10)
'    txt = ibr.ReadBytes(, 10)
'    Debug.Print InStrB(1, txt, "DPB=""", vbBinaryCompare)
'    Debug.Print FileLen(inVbaProjectPath)
    Dim dByte() As Byte: dByte = ibr.ReadAllToBytes()
    Dim dpbIndex As Long: dpbIndex = IndexOfBinary(dByte, "DPB=", 0)
    Dim stIndex As Long: stIndex = dpbIndex + 5
    Dim edIndex As Long: edIndex = IndexOfBinary(dByte, """", stIndex)
'    Debug.Print edIndex - stIndex
    ibr.CloseFile
    
    Dim obr As kccBinaryFileIO
    Set obr = kccBinaryFileIO.OpenFile(outVbaProjectPath, 2)
    WriteBinary dByte, StrConv(GetDPB1234(edIndex - stIndex), vbFromUnicode), stIndex
    obr.WriteByte dByte
    obr.CloseFile
'    Stop
End Function

Rem vbaProjectを書き換えるテスト
Sub Test_vbaProjectCrack()
'    Call vbaProjectCrack(TEST_PATH, TEST_PATH2)
    Call vbaProjectCrack( _
            "D:\vba\vbaProject.bin", _
            "D:\vba\vbaProject2.bin")
        'なぜか書き出しPutの時点で先頭に12バイト増加してしまう謎の現象が発生中
        '原因はVariantでPutしたのが原因でした。
        'Byte()にキャストしてからなら発生しません。
End Sub

Rem ExcelブックのVBAパスワードを1234へ置換する
Rem
Rem inFilePath      入力ファイルフルパス
Rem outFilePath     出力ファイルフルパス
Rem
Rem @return As Variant
Function BrokenVbaPassword(inFilePath, outFilePath)
    Dim vbaPath
    Select Case fso.GetExtensionName(inFilePath)
        Case "xls", "xlsb", "xla"
            vbaPath = inFilePath
            
            Call vbaProjectCrack("" & vbaPath, "" & vbaPath)
            
        Case "xlsx", "xlsm", "xlam"
            With kccFuncZip.DecompZip(inFilePath)
                Dim tempPath
                tempPath = .DecompFolder
                vbaPath = tempPath & "\xl\vbaProject.bin"
            
                Call vbaProjectCrack("" & vbaPath, "" & vbaPath) '& "2"
            
                BrokenVbaPassword = kccFuncZip.CompZip(tempPath, outFilePath)
            End With
    End Select
End Function

Sub Test_BrokenVbaPassword()
    Dim f1: f1 = "D:\vba\test.xlsm"
    Dim f2: f2 = "D:\vba\test_1234.xlsm"
    Call BrokenVbaPassword(f1, f2)
End Sub

'現在アクティブなブックを閉じてパスを統一し再度開く
Sub Test_BrokenVbaPassword2()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim fn As String: fn = wb.FullName
    wb.Close False
    Dim ext: ext = fso.GetExtensionName(fn)
    Dim fn2 As String: fn2 = Left(fn, Len(fn) - Len(ext) - 1) & "_1234." & ext
    Call BrokenVbaPassword(fn, fn2)
    
    Set wb = Workbooks.Open(fn2)
End Sub

'For Eachではコレクションの要素を書き換えられない
Sub Test_CollectionForeach()
    Dim c As Collection: Set c = New Collection
    c.Add "a\"
    c.Add "b\"
    Dim i
    For Each i In c
        i = "a"
    Next
    For Each i In c
        Debug.Print i
    Next
End Sub
