VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kccBinaryFileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        kccBinaryFileIO
Rem
Rem  @description   バイナリファイル読み書きラッパークラス
Rem
Rem  @update        2020/09/09
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
Rem  @functions
Rem
Rem --------------------------------------------------------------------------------
Rem  @note
Rem    VBAのファイル操作ステートメント一覧
Rem
Rem    基本
Rem      Seek [#]#FileNumber, position
Rem      Close #FileNumber
Rem      Reset
Rem      Lock #FileNumber, position
Rem      UnLock #FileNumber, position
Rem      Open pathname For mode [Access access] [Lock] As [#]#FileNumber [Len=recLength]
Rem        Mode
Rem          Open fn For Input As #1
Rem          Open fn For Output As #1
Rem          Open fn For Random As #1
Rem          Open fn For Append As #1
Rem          Open fn For Binary As #1
Rem        [Access access]
Rem          Open fn For Binary Access Read As #1
Rem          Open fn For Binary Access Write As #1
Rem          Open fn For Binary Access Read Write As #1
Rem        [Lock]
Rem          Open fn For Binary Shared As #1
Rem          Open fn For Binary Lock Read As #1
Rem          Open fn For Binary Lock Write As #1
Rem          Open fn For Binary Lock Read Write As #1
Rem        [Len]
Rem          Open fn For Random Shared As #1 Len = 6
Rem
Rem    読み
Rem      Get [#]FileNumber, [recnumber], varname
Rem      Line Input #FileNumber, varname
Rem      Input #FileNumber, varlist
Rem
Rem    書き
Rem      Put [#]FileNumber,[recnumber],varname
Rem      Write #FileNumber, [outputlist]
Rem      Print #FileNumber, [outputlist]
Rem      Width #FileNumber, Width
Rem
Rem    関数
Rem      FreeFile()
Rem      EOF( #FileNumber )
Rem      LOC( #FileNumber )
Rem      FileLen( FileName)
Rem      Seek( #FileNumber )
Rem      Input( size, #FileNumber )
Rem      InputB( size, #FileNumber )
Rem
Rem    関連文字列関数
Rem      TAB(n)
Rem      SPC(n)
Rem
Rem --------------------------------------------------------------------------------

Option Explicit

Private fp_ As Integer
Private fn_ As String
Private om_ As Long

Private Const ERROR_MSG_NOT_READABLE = "読み取り可能で開かれていません"
Private Const ERROR_MSG_NOT_WRITABLE = "書き込み可能で開かれていません"
Private Const ERROR_MSG_OUT_OF_INDEX = "読み書きするインデックスが範囲外です"
Private Const ERROR_MSG_OUT_OF_SIZE = "読み書きするサイズが範囲外です"
Private Const ERROR_MSG_NOT_OPEND = "ファイルが開かれていません"

Private fso As New FileSystemObject

Private Sub Class_Initialize()
    '特に何も行わない
    '正式な初期化は OpenFile で行う
End Sub

Private Sub Init()
    'CallByName対策
    '特に何も行わない
    '正式な初期化は OpenFile で行う
End Sub

Private Sub Class_Terminate()
    Call CloseFile
End Sub

Rem ファイルを開く
Rem
Rem @param mFileName        ファイル名フルパス
Rem @param R1W2RW3          1:読込専用 2:書込専用 3:読み書き可能
Rem
Rem @return As kccBinaryFileIO オブジェクトを生成
Rem
Function OpenFile(mFileName As String, R1W2RW3 As Long) As kccBinaryFileIO
    If Me Is kccBinaryFileIO Then
        With New kccBinaryFileIO
            Set OpenFile = .OpenFile(mFileName, R1W2RW3)
        End With
        Exit Function
    End If
    Set OpenFile = Me
    
    fn_ = mFileName
    fp_ = FreeFile()
    om_ = R1W2RW3
    
    If R1W2RW3 = 1 Then
        Open mFileName For Binary Access Read As fp_
    ElseIf R1W2RW3 = 2 Then
        If fso.FileExists(mFileName) Then
            fso.DeleteFile mFileName
        End If
        Open mFileName For Binary Access Write As fp_
    Else
        Open mFileName For Binary As fp_
    End If
End Function

Rem ファイルを閉じる
Sub CloseFile()
    If fp_ <> 0 Then Close fp_
    fp_ = 0
    om_ = 0
    fn_ = ""
End Sub

Rem 読み込み可能であるか
Property Get IsReadable() As Boolean
    IsReadable = (om_ And 1)
End Property

Rem 書き込み可能であるか
Property Get IsWritable() As Boolean
    IsWritable = (om_ And 2)
End Property

Rem VBAの管理するファイルポインタ
Property Get FileNumber() As Long
    If fp_ = 0 Then Err.Raise 9999, "", ERROR_MSG_NOT_OPEND
    FileNumber = fp_
End Property

Rem 開いているファイルのフルパス
Property Get FileName() As String
    FileName = fn_
End Property

Rem 全データをバイナリ配列で一括読み込み
Function ReadAllToBytes() As Byte()
    If IsReadable Then Else Err.Raise 9999, "", ERROR_MSG_NOT_READABLE
    Dim iSize: iSize = Me.FileSize()
    If iSize = 0 Then Exit Function
    
'    ReDim ReadAllToBytes(iSize)
    Seek #Me.FileNumber, 1
'    ReadAllToBytes = InputB(iSize, Me.FileNumber)
    ReDim ReadAllToBytes(0 To FileLen(Me.FileName) - 1)
    Get #Me.FileNumber, , ReadAllToBytes
    Seek #Me.FileNumber, 1
End Function

Rem 全データをバイナリ配列で一括読み込み
Function ReadAllToString() As String
    If IsReadable Then Else Err.Raise 9999, "", ERROR_MSG_NOT_READABLE
    
    Dim iSize: iSize = Me.FileSize()
    If iSize = 0 Then Exit Function
    
    Seek #Me.FileNumber, 1
    ReadAllToString = Space(iSize)
    Get #Me.FileNumber, , ReadAllToString
    Seek #Me.FileNumber, 1
End Function

Rem 1バイト読込
Rem
Rem @param SeekIndex        読込位置1~n 省略時:現在位置
Rem
Rem @return As Byte         読み込んだデータ
Rem
Function ReadByte(Optional SeekIndex) As Byte
    If IsReadable Then Else Err.Raise 9999, "", ERROR_MSG_NOT_READABLE
    
    If VBA.IsMissing(SeekIndex) Then
        Get #Me.FileNumber, , ReadByte
    Else
        Get #Me.FileNumber, SeekIndex, ReadByte
    End If
End Function

Rem 指定サイズをバイト配列に読込
Rem
Rem @param SeekIndex        読込位置1~ 省略時:現在位置
Rem @param ReadSize         読込データサイズを示すバイト数1~
Rem
Rem @return As Byte()       読み込んだデータ
Rem
Function ReadBytes(Optional SeekIndex, Optional ReadSize = 1) As Byte()
    If IsReadable Then Else Err.Raise 9999, "", ERROR_MSG_NOT_READABLE
    If ReadSize < 1 Then Err.Raise 9999, "", ERROR_MSG_OUT_OF_SIZE
    
    ReDim ReadBytes(0 To ReadSize - 1)
    If VBA.IsMissing(SeekIndex) Then
        Get #Me.FileNumber, , ReadBytes
    Else
        Get #Me.FileNumber, SeekIndex, ReadBytes
    End If
End Function

Rem 指定サイズを文字列に読込（敢えてByte配列に入れたくない場合に使用）
Rem
Rem @param SeekIndex        読込位置1~n 省略時:現在位置
Rem @param ReadSize         読込データサイズを示すバイト数1~
Rem
Rem @return As String       読み込んだデータ
Rem
Function ReadString(Optional SeekIndex, Optional ReadSize = 1) As String
    If IsReadable Then Else Err.Raise 9999, "", ERROR_MSG_NOT_READABLE
    If ReadSize < 1 Then Err.Raise 9999, "", ERROR_MSG_OUT_OF_SIZE
    
    ReadString = Space(ReadSize)
    If VBA.IsMissing(SeekIndex) Then
        Get #Me.FileNumber, , ReadString
    Else
        Get #Me.FileNumber, SeekIndex, ReadString
'        ReadString = InputB(ReadSize, #Me.FileNumber)
    End If
End Function

Rem データを書き出す
Rem
Rem @param sameData         書き出したい任意のデータ
Rem @param SeekIndex        読込位置1~n 省略時:現在位置
Rem
Rem @note
Rem   http://www016.upp.so-net.ne.jp/garger-studio/gameprog/vb0124.html
Rem   Variant型の変数のまま書き込むと、データの前に型情報や要素情報が入ってしまう。
Rem   Byte() 型情報?4byte要素数4byte  00が4byte
Rem          11 20 01 00 0A 00 00 00 00 00 00 00
Rem   String 型情報?4byte
Rem          08 00 0A 00
Rem   ユーザー定義型はVariantに抽象化出来ないので断念した。
Rem
Sub WriteByte(sameData, Optional SeekIndex)
    If Not VBA.IsMissing(SeekIndex) Then
        If SeekIndex > 0 Then
            Me.FileSeek SeekIndex
        End If
    End If
    
    Dim bData() As Byte
    Select Case TypeName(sameData)
        Case "Byte"
            ReDim bData(0 To 0)
            bData(0) = sameData
            Put #Me.FileNumber, , bData
        Case "Byte()"
            bData = sameData
            Put #Me.FileNumber, , bData
        Case "String"
            Dim sData As String
            sData = sameData
            Put #Me.FileNumber, , sData
        Case Else
            MsgBox "未定義のデータ形式のため意図しないデータで書き出される恐れがあります"
            Put #Me.FileNumber, , sameData
    End Select
End Sub

Rem ファイルの全部または一部をロック
Rem
Rem @param SeekIndex        読込位置1~n 省略時:現在位置
Rem
Rem @note
Rem   複数のプロセスが同じファイルにアクセスできる場合に使用
Rem   Open "C:\Test.dat" For Random Shared As #1 Len = 6
Rem
Sub FileLock(Optional SeekIndex As Long)
    If VBA.IsMissing(SeekIndex) Then
        Lock #Me.FileNumber
    Else
        Lock #Me.FileNumber, SeekIndex
    End If
End Sub

Rem ファイルの全部または一部をロック解除
Rem
Rem @param SeekIndex        読込位置1~n 省略時:現在位置
Rem
Sub FileUnLock(Optional SeekIndex As Long)
    If VBA.IsMissing(SeekIndex) Then
        Unlock #Me.FileNumber
    Else
        Unlock #Me.FileNumber, SeekIndex
    End If
End Sub

Rem 現在の読み書き位置を示すカーソル
Function Cursol() As LongPtr
    Cursol = VBA.Loc(Me.FileNumber)
'    Cursol = Seek(Me.FileNumber) '同義
End Function

Rem ファイルのバイト数を取得(FileLenとは違いOpen中も読める）
Function FileSize()
    FileSize = VBA.LOF(Me.FileNumber)
End Function

Rem 読み込み位置をシフト
Function FileSeek(Optional SeekIndex = 1)
    If SeekIndex < 1 Then Err.Raise 9999, "", ERROR_MSG_OUT_OF_INDEX
    Seek #Me.FileNumber, SeekIndex
End Function

Rem ファイルが末端に到達済みか
Function IsEndOfFile() As Boolean
    IsEndOfFile = VBA.EOF(Me.FileNumber)
End Function
