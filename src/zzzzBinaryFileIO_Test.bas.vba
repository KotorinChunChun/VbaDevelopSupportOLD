Attribute VB_Name = "zzzzBinaryFileIO_Test"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccBinaryFileIO_Test
Rem
Rem  @description   テスト
Rem
Rem  @update        2020/09/09
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem  @note
Rem
Rem    Input & Output Functions
Rem    https://bettersolutions.com/vba/functions/input-output-category.htm
Rem
Rem --------------------------------------------------------------------------------
Option Explicit

Rem バイナリファイル書き出しテスト１　バイナリ配列
Sub Test_Put1()
    Const MAX_ITEMS = 10
    Dim bData() As Byte
    ReDim bData(0 To MAX_ITEMS - 1)
    Dim i As Long
    For i = LBound(bData) To UBound(bData)
        bData(i) = i
    Next
    
    Dim ibr As kccBinaryFileIO: Set ibr = kccBinaryFileIO.OpenFile("D:\vba\binTest1.bin", 2)
    ibr.WriteByte bData
    ibr.CloseFile
End Sub

Rem バイナリファイル書き出しテスト２　文字列
Sub Test_Put2()
    Dim sData As String: sData = "1234abcdあ"
    Dim ibr As kccBinaryFileIO: Set ibr = kccBinaryFileIO.OpenFile("D:\vba\binTest2.bin", 2)
    ibr.WriteByte sData
    ibr.CloseFile
End Sub

Rem 読込テスト
Sub Test_ReadByte()
    Dim br As kccBinaryFileIO
    Set br = kccBinaryFileIO.OpenFile("D:\vba\test.bin", 1)
    
    DebugPrintByteArray br.ReadAllToBytes, 10
    
    Debug.Print br.ReadByte
    Debug.Print br.ReadByte
    Debug.Print br.ReadByte
    Debug.Print br.ReadByte(1)
    Debug.Print br.ReadByte(2)
    
    Debug.Print br.ReadBytes(, 4)
    Debug.Print br.ReadBytes(, 4)

    Debug.Print 1, br.ReadString
    Debug.Print 2, br.ReadString
    Debug.Print 3, br.ReadString
    Debug.Print 4, br.ReadString
    Debug.Print 4, br.ReadString(4)
    Debug.Print 5, br.ReadString(5, 2)
End Sub

Rem 書込みテスト
Sub Test_WriteByte()
    Dim fn As String: fn = "D:\vba\test.bin"
    
    Dim br As kccBinaryFileIO
    Set br = kccBinaryFileIO.OpenFile(fn, 2)
'     br.WriteByte "a"
'     br.WriteByte "b"
'
'     Dim b() As Byte
'     b = "b"
'     br.WriteByte b
'     br.WriteByte 1
End Sub

Rem Withの実験 効果が見られず意味分からん
Sub Sample_width()
    Open "D:\vba\width.txt" For Output As #1
    Width #1, 5     'MAX0-255
    Print #1, "ABCDEFGHIJ"
    Close #1
End Sub

Sub Sample_write()
    Open "C:\test.dat" For Output As #1
    Write #1, "ABC", 123
    Write #1,
    Write #1, Now, True
    Close #1
End Sub
