VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "kccFuncZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Rem ZIPファイルを解凍する
Rem
Rem  @param inFilePath      解凍したいファイルのフルパス
Rem  @param outParentPath   解凍先親フォルダ
Rem                           省略時         : 一時フォルダ
Rem                           ルート開始パス : 指定パス
Rem                           相対パス       : 元ファイル基準の相対パス
Rem                         ※指定パスに元ファイルの拡張子を除いた名前のフォルダを生成します。
Rem
Rem  @return As String      解凍されたパスの絶対パス形式
Rem
Rem  @note
Rem    内部処理の流れ
Rem    1. %temp%\VbaUnZip\FILENAME.zip へ元のファイルをコピー
Rem    2. %temp%\VbaUnZip\FILENAME\ へ解凍
Rem    3. 1のファイルを削除
Rem    4. outFolderPathへ移動 (非省略時)
Rem       ※ドライブが違うとMove出来ないためCopyからのDeleteとなる
Rem
Rem   既にフォルダが存在する場合、フォルダごと削除して置き換わります。
Rem
Rem   指定フォルダへ直接解凍することは出来ません。
Rem   将来的にオプションを追加する可能性はあります。
Rem   もし実装するとなると解凍したいファイルを個別に削除→移動するロジックに変える必要があります。
Rem
Rem   パスワード付きZIPファイルの解凍は断念した。
Rem
Function DecompZip(ByVal inFilePath, Optional outParentPath) As String
    Const PROC_NAME = "DecompZip"
    If Not fso.FileExists(inFilePath) Then
        Err.Raise 9999, PROC_NAME, "展開したいZIPファイルがありません：" & inFilePath
        Exit Function
    End If
    
    '一時フォルダの準備
    Dim tempFolderPath As String
    tempFolderPath = GetTempFolder("VbaUnZip") & fso.GetBaseName(inFilePath)
    If fso.FolderExists(tempFolderPath) Then
        If Not DeleteFolder(tempFolderPath) Then
            Err.Raise 9999, PROC_NAME, "解凍一時フォルダの初期化に失敗：" & tempFolderPath
        End If
    End If
    fso.CreateFolder tempFolderPath
    
    '出力フォルダの準備
    If VBA.IsMissing(outParentPath) Then outParentPath = ""
    Dim outFolderPath As String
    '省略時   : 一時フォルダ
    If outParentPath = "" Then
        outFolderPath = ""
    Else
        'ルート開始パス : 指定パス
        If kccFuncString.IsRootStart(outParentPath) Then
            outFolderPath = outParentPath & fso.GetBaseName(inFilePath)
        '相対パス : 元ファイル基準の相対パス
        Else
            Dim curFolderPath As String
            curFolderPath = fso.GetParentFolderName(inFilePath) & "\"
            outFolderPath = kccFuncString.AbsolutePathNameEx(curFolderPath, outParentPath) & fso.GetBaseName(inFilePath)
        End If
        If fso.FolderExists(outFolderPath) Then
            If Not DeleteFolder(outFolderPath) Then
                Err.Raise 9999, PROC_NAME, "出力先フォルダの初期化に失敗：" & outFolderPath
            End If
        End If
    End If
    
    '拡張子のチェックとZIP複製(Namespace展開に拡張子が重要)
    Dim zip_file_path As String
    Dim IsZip As Boolean: IsZip = inFilePath Like "*.zip"
    If IsZip Then
        zip_file_path = inFilePath
    Else
        zip_file_path = tempFolderPath & ".zip"
        fso.CopyFile inFilePath, zip_file_path
    End If
    
    'Namespaceには暗黙の型変換が必要なので、変数ではなく式が必要
    Dim objZip
    Set objZip = CreateObject("Shell.Application").Namespace("" & zip_file_path).Items
    
    'これだと安定性に欠けるため、パスワードZIP対応は断念
'    Application.SendKeys zipPw & "{Enter}"
    
'    DecompZip = CreateObject("Shell.Application").Namespace("" & tempFolderPath).CopyHere(objZip, &H4 Or &H10)
    DecompZip = CopyHere(tempFolderPath, objZip)
'    sha.Namespace(unzipfld_).CopyHere( sha.Namespace(zippth_).Items, &H4 Or &H10)
    
    'キャッシュのZIPファイル削除
    fso.DeleteFile zip_file_path
    
    If outFolderPath = "" Then
        DecompZip = tempFolderPath
    Else
        If fso.GetDriveName(tempFolderPath) <> fso.GetDriveName(outFolderPath) Then
            'MoveFolderはドライブ間の移動ができないためCopyFolder
            'コピー先に\を付けるとフォルダ複製。\が無いと中身複製になるらしい。
            fso.CreateFolder outFolderPath
            fso.CopyFolder tempFolderPath, outFolderPath
            fso.DeleteFolder tempFolderPath
        Else
            fso.MoveFolder tempFolderPath, outFolderPath
        End If
        DecompZip = outFolderPath
    End If
End Function

Sub Test_DecompZip()
    Dim inFile: inFile = "D:\vba\zip\test.xlsm"
    Dim outFolder: outFolder = "D:\vba\zip\temp"
'    Debug.Print DecompZip(inFile)
    Debug.Print DecompZip(inFile, outFolder)
'    Debug.Print DecompZip(inFile, outFolder, "a")
End Sub
 
Sub Test_CompZip()
    Dim inFolder: inFolder = "D:\vba\zip\temp"
'    Dim outFile: outFile = "D:\vba\zip\hoge.xlsm"
    Debug.Print CompZip(inFolder)
End Sub

Rem ファイル・フォルダをZIP形式で圧縮する
Rem
Rem  @param target_paths     圧縮元のファイル・フォルダの絶対パス。又はその配列
Rem  @param zip_file_path    圧縮後のファイルのパス
Rem                           省略時         : 元ファイルと同じフォルダで先頭のBaseName
Rem                           ルート開始パス : 指定フォルダ
Rem                           相対パス       : 元ファイル基準の相対パス
Rem                           ※指定パスに元ファイルの拡張子を除いた名前のフォルダを生成します。
Rem
Rem  @return As String      圧縮されたファイルの絶対パス
Rem
Rem  @note
Rem
Rem
Function CompZip(target_paths, Optional zip_file_path) As String
    Const PROC_NAME = "CompZip"
    
    Dim targetPaths As Collection
    Set targetPaths = ToCollection(target_paths)
    Dim i As Long
    For i = 1 To targetPaths.Count
        Dim s As String: s = targetPaths.Item(i)
        If s Like "*\" Then targetPaths(i) = Left(s, Len(s) - 1)
    Next
    Dim firstTargetPath As String: firstTargetPath = targetPaths.Item(1)
    
    '省略時   : 元ファイルと同じフォルダで先頭のBaseName
    If VBA.IsMissing(zip_file_path) Then zip_file_path = ""
    If zip_file_path = "" Then
        zip_file_path = firstTargetPath & ".zip"
    '相対パス or 絶対パス
    Else
        Dim ParentFolderPath As String
        ParentFolderPath = fso.GetParentFolderName(firstTargetPath)
        zip_file_path = kccFuncString.AbsolutePathNameEx(ParentFolderPath, zip_file_path)
    End If
    If fso.FileExists(zip_file_path) Then fso.DeleteFile zip_file_path
    If fso.FileExists(zip_file_path) Then
        Err.Raise 9999, PROC_NAME, "出力ZIPが初期化できません：" & zip_file_path
        Exit Function
    End If
    
    Dim tempZipName As String
    If zip_file_path Like "*.zip" Then
        tempZipName = zip_file_path
    Else
        tempZipName = zip_file_path & ".zip"
    End If
    
    'ZIPファイルの新規作成
    With fso.CreateTextFile(tempZipName, True)
        .Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)
        .Close
    End With
    
    'この書き方だとフォルダがルートに作成されてしまう。
'    CompZip = zipFolder.CopyHere("" & targetPaths & "\")
    
    Dim tgt As Variant
    Dim f As Object
    Dim Result As Boolean
    For Each tgt In targetPaths
        If fso.FileExists(tgt) Then
            Result = CopyHere(tempZipName, tgt)
        ElseIf fso.FolderExists(tgt) Then
            For Each f In fso.GetFolder(tgt).Files
                Result = CopyHere(tempZipName, f.Path)
            Next
            For Each f In fso.GetFolder(tgt).SubFolders
                Result = CopyHere(tempZipName, f.Path)
            Next
        Else
            Dim msg As String
            msg = PROC_NAME & " ソースがありません：" & tgt
            'Err.Raise 9999,msg
            Debug.Print msg
        End If
    Next
    
    If tempZipName <> zip_file_path Then
        fso.MoveFile tempZipName, zip_file_path
    End If
    
End Function

'非同期を解消済みのCopyHereメソッド
'※ObjectやString厳禁 shaにはVariantで渡すこと
Function CopyHere(ToObjectOrPath As Variant, FromObjectOrPath As Variant) As Boolean
    Dim toObj
    If IsObject(ToObjectOrPath) Then
        Set toObj = ToObjectOrPath
    Else
        Set toObj = CreateObject("Shell.Application").Namespace("" & ToObjectOrPath)
    End If
    
    '   : 省略可
    '4  : 指定した場合は展開時におけるダイアログが表示されなくなります。
    '16 : ????
    CopyHere = toObj.CopyHere(FromObjectOrPath, &H4 Or &H10)
    
    If IsObject(ToObjectOrPath) Then Exit Function
    If Not ToObjectOrPath Like "*.zip" Then Exit Function
    If TypeName(ToObjectOrPath) <> "String" Then Stop: Exit Function
    
    'CopyHereが非同期なので、試しにTextOpenして同期をとる
    Call WaitFileClosed("" & ToObjectOrPath)
    
End Function

Rem 指定ファイルが書き込み可能となるまで待機する
Function WaitFileClosed(fn As String, Optional max_wait_second)
    Do
        'ココに遅延が必須。そうしないとCopyHereが始まる前に検証用のOpenが動いてしまう
        Application.Wait [Now() + "00:00:00.2"]
        
        '試しに挿入OpenしてロックされてたらCopyHereが終わっていないので処理待ち継続
        On Error Resume Next
            Call fso.OpenTextFile(fn, ForAppending, False).Close
            If Err.Number = 0 Then Exit Do
'            Debug.Print ToPath, FromObjectOrPath
        On Error GoTo 0
        
        DoEvents
    Loop
    Application.Wait [Now() + "00:00:00.2"]
End Function

Public Function ToCollection(var) As Collection
    Dim Item
    If TypeName(var) = "Collection" Then
        Set ToCollection = var
    ElseIf IsArray(var) Then
        Set ToCollection = New Collection
        For Each Item In var: ToCollection.Add Item: Next
    ElseIf IsObject(var) Then
        Set ToCollection = New Collection
        On Error Resume Next
        For Each Item In var
            If Err Then Debug.Print TypeName(var), Err.Number: Stop 'オブジェクトからの変換未完成
            ToCollection.Add Item
        Next
        On Error GoTo 0
    Else
        Set ToCollection = New Collection
        ToCollection.Add var
    End If
End Function

'------------------------------------------------------------------------------------------

'IShellDispatch Folder3のParentFolderを遡ってフルパスを取得する
'この案はダメだった。物理的にありえないパスを示してしまった。
'結果
' デスクトップ\USERNAME\AppData\Roaming\VbaUnZip\BOX_sample(group_mng1)_v1.0
'正解
'    C:\Users\USERNAME\AppData\Roaming\VbaUnZip
Function GetFullPathByFolder3(obj) As String
    Dim ret As String
    ret = ""
    Dim nextObj
    Set nextObj = obj
    Do
        ret = nextObj & IIf(ret = "", "", "\") & ret
        Debug.Print ret
'        Stop
        Set nextObj = nextObj.ParentFolder
        If nextObj Is Nothing Then Exit Do
    Loop
    GetFullPathByFolder3 = ret
End Function

Function GetTempFolder(subFolder)
    GetTempFolder = VBA.CreateObject("Wscript.Shell").SpecialFolders("AppData") & "\"
    GetTempFolder = GetTempFolder & IIf(subFolder = "", "", subFolder & "\")
    On Error Resume Next
    fso.CreateFolder GetTempFolder
End Function
 
'http://excelfactory.net/excelboard/excelvba/excel.cgi?mode=all&namber=188859&rev=0
Sub 圧縮(ByVal OldFld As String, ByVal Str As String)
    
    Dim Result As Variant
    Result = Split(Str, "\") 'strを\で分割する
    Result = Result(UBound(Result)) 'フォルダ名を取り出す
    
    '空のZIPファイル作成
    Dim ts As TextStream
    Set ts = fso.CreateTextFile(Str & "\" & Result & ".zip")
    
    Dim msg As String
    msg = "PK" & Chr(5) & Chr(6) & String(18, 0)
    ts.Write (msg)
    ts.Close
    
    'フォルダオブジェクト取得
    Dim sh As Object, fol As Object
    Set sh = CreateObject("Shell.Application")
    Set fol = sh.Namespace(Str & "\" & Result & ".zip")
    
    'サンプル：フォルダ内のXLSファイルを圧縮
    Dim FName As String
    FName = Dir(OldFld & "*")
    Do While (FName <> "")
        fol.MoveHere OldFld & FName 'ZIPへファイル追加
        FName = Dir()
    Loop

End Sub

'パスワード付きZIP圧縮　Lhaplus使用版
'テストデータが0バイトの場合パスはつかない。
Sub CompPasswordZipForLhaplus()
    Dim WSH As Object
    Dim wExec As Object
    Dim comStr As String
    Set WSH = CreateObject("WScript.Shell")
    comStr = "C:\Program Files\Lhaplus\Lhaplus.exe /c:zip /p:123 /n:C:\Users\ユーザー名\Desktop\圧縮.zip C:\Users\ユーザー名\Desktop\圧縮"
    Set wExec = WSH.Exec(comStr)
    Set WSH = Nothing
    Set wExec = Nothing
End Sub
 

'7-ZIPならば、-pでパスワード指定できるようです。お試しください。

'PDF
'xpdfが良いらしい
'http://pdf-file.nnn2.com/?p=858


'https://qiita.com/RelaxTools/items/375492175ef902e59ca5

Sub Main()

    Dim Col As Collection

    Set Col = New Collection

    Col.Add "e:\README.md"

    CompressArchive Col, "E:\aaa.zip"


End Sub

'--------------------------------------------------------------
' Zip 圧縮処理
'--------------------------------------------------------------
Private Sub CompressArchive(Col As Collection, strDest As String)

    Dim strCommand As String
    Dim strPath  As String
    Dim v As Variant
    Dim First As Boolean

    'コマンド
    strCommand = "Compress-Archive"

    strCommand = strCommand & " -Path"

    First = True
    For Each v In Col

        If First Then
            strPath = """" & v & """"
            First = False
        Else
            strPath = strPath & ",""" & v & """"
        End If
    Next

    strCommand = strCommand & " " & strPath

    strCommand = strCommand & " -DestinationPath"
    strCommand = strCommand & " """ & strDest & """"

    strCommand = strCommand & " -Force"


    'PowerShell を実行する
    ExecPowerShell strCommand

End Sub

'--------------------------------------------------------------
' PowerShell 実行
'--------------------------------------------------------------
Private Sub ExecPowerShell(strCommand As String)

    Dim strTemp As String
    Dim strFile As String
    Dim strBuf As String

    With CreateObject("Scripting.FileSystemObject")

        strTemp = .GetSpecialFolder(2).Path
        strFile = .BuildPath(strTemp, .GetTempName & ".ps1")

        'テキスト出力
        With .CreateTextFile(strFile, True)
            .Write strCommand
            .Close
        End With

        strBuf = "powershell"
        strBuf = strBuf & " -ExecutionPolicy"
        strBuf = strBuf & " RemoteSigned"
        strBuf = strBuf & " -File"
        strBuf = strBuf & " """ & strFile & """"

        With CreateObject("WScript.Shell")
            Call .Run(strBuf, 0, True)
        End With

        .DeleteFile strFile

    End With

End Sub


'Scripting.FileSystemObject、Shell.Applicationを使用。
'https://qiita.com/kou_tana77/items/06f7dc897ef1a69d2ea8
Public Function zip(ArrPath() As String, zippth As String) As Boolean
    zip = False
    On Error GoTo Err
    Dim fso As Object, sha As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sha = CreateObject("Shell.Application")
    If fso.FileExists(zippth) = True Then
        fso.DeleteFile zippth
    End If
    With fso.CreateTextFile(zippth, True)
        .Write "PK" & Chr(5) & Chr(6) & String(18, 0)
        .Close
    End With
    Dim zipfld As Object
    Set zipfld = sha.Namespace(fso.GetAbsolutePathName(zippth))
    Dim idx As Long, maxidx As Long: maxidx = UBound(ArrPath, 1)
    Dim n As Long: n = 0
    Dim f As Variant
    Dim start_tim As Date, cpyflg As Boolean
    For idx = 0 To maxidx
        cpyflg = False
        f = fso.GetAbsolutePathName(ArrPath(idx))
        If fso.FolderExists(ArrPath(idx)) = True Then
            If sha.Namespace(f).Items().Count > 0 Then  '空フォルダでない？
                                                        '⇒空フォルダは圧縮できない
                cpyflg = True
            End If
        ElseIf Dir(ArrPath(idx)) <> "" Then
            cpyflg = True
        End If
        If cpyflg = True Then
            zipfld.CopyHere f, &H4 Or &H10
            n = n + 1
            'コピーが終わるのを待つ
            start_tim = Now
            Do Until zipfld.Items().Count = n
                If DateDiff("s", start_tim, Now) > 5 Then    'タイムオーバー
                    Exit Function
                End If
                Debug.Print CStr(n) & "/" & CStr(zipfld.Items().Count)
                Sleep 10
            Loop
        End If
    Next
    zip = True
Err:
    If Err.Number <> 0 Then
        Debug.Print "zip(): " & Err.Description
    End If
    Set fso = Nothing
    Set sha = Nothing
End Function

Property Get fso() As FileSystemObject: Set fso = CreateObject("Scripting.FileSystemObject"): End Property
Property Get sha() As Object: Set sha = CreateObject("Shell.Application"): End Property

'Sub Test_UnZip()
'    Dim arr: arr = VBA.Array("D:\vba\zip\test.zip")
'    Debug.Print UnZip(arr, "D:\vba\zip\test.zip", "D:\vba\zip\temp")
'End Sub
'
'Public Function UnZip(ArrPath, zippth As String, unzipfld As String) As Boolean
'    UnZip = False
'    On Error GoTo Err
'
'    If fso.FolderExists(unzipfld) Then
'        If DeleteFolder(unzipfld) Then Else Exit Function
'    ElseIf Dir(unzipfld, vbNormal) <> "" Then
'        Debug.Print "unzip(): フォルダ以外のファイル(" & unzipfld & ")が存在"
'        Exit Function
'    End If
'
'    MkDir unzipfld
'    Dim unzipfld_ As Variant, zippth_ As Variant
'    unzipfld_ = fso.GetAbsolutePathName(unzipfld)
'    zippth_ = fso.GetAbsolutePathName(zippth)
'
'    '展開 = 普通のフォルダへコピー
'    sha.Namespace(unzipfld_).CopyHere sha.Namespace(zippth_).Items, &H4 Or &H10
'
'    'サブフォルダ処理
'    UnZip = move_pth4unzip(unzipfld_, ArrPath)
'
'    Exit Function
'
'Err:
'    If Err.Number <> 0 Then
'        Debug.Print "unzip(): " & Err.Description
'    End If
'End Function
'
'Private Function move_pth4unzip(unzipfld As Variant, ArrPath) As Boolean
'    move_pth4unzip = False
''    On Error GoTo Err
'    Dim idx As Long, maxidx As Long
'    maxidx = UBound(ArrPath, 1)
'    Dim f As Variant
'    For Each f In sha.Namespace(unzipfld).Items
'        Debug.Print f.Name
'        For idx = 0 To maxidx
'            If BaseName("" & ArrPath(idx)) = f.Name Then
'                Exit For
'            End If
'        Next
'        If idx <= maxidx Then
'            If move_pth4unzip1(CStr(unzipfld), f.Name, "" & ArrPath(idx)) = False Then
'                Exit Function
'            End If
'        Else
'            Debug.Print "move_pth4unzip(): " & _
'                    "zipファイルに展開対象外ファイル(=""" & f.Name & """)が含まれていた:=>を無視"
'        End If
'    Next
'    move_pth4unzip = True
'    Exit Function
'
'Err:
'    If Err.Number <> 0 Then
'        Debug.Print "move_pth4unzip(): " & Err.Description
'    End If
'End Function
'
'Private Function move_pth4unzip1(fr_fld As String, fr_fn As String, to_Path As String) As Boolean
'    move_pth4unzip1 = False
'    On Error GoTo Err
'    If Dir(to_Path) <> "" Or fso.FolderExists(to_Path) = True Then
'        If DeleteFolder(to_Path) Then Else Exit Function
'    End If
'
'    Dim fr_Path As String: fr_Path = fr_fld & "\" & fr_fn
'    If fso.FolderExists(fr_Path) = True Then
'        fso.MoveFolder fr_Path, to_Path
'    Else
'        fso.MoveFile fr_Path, to_Path
'    End If
'    move_pth4unzip1 = True
'    Exit Function
'
'Err:
'    If Err.Number <> 0 Then
'        Debug.Print "move_pth4unzip1(): " & Err.Description
'    End If
'End Function
'
''「ファイル/フォルダをzipファイルに圧縮」の記事にあった関数と同じ関数
''⇒コメントアウトしている
''Public Function fso.FolderExists(Path As String) As Boolean
''    fso.FolderExists = CreateObject("Scripting.FileSystemObject").FolderExists(Path)
''End Function
'
'Public Function BaseName(Path As String) As String
'    Dim Path_ As String: Path_ = Trim(Path)
'    If Right(Path_, Len("\")) = "\" Then
'        Path_ = Left(Path_, Len(Path_) - Len("\"))
'    End If
'    Dim pos As Long
'    pos = InStrRev(Path_, "\")
'    If pos <> 0 Then
'        BaseName = Right(Path_, Len(Path_) - pos)
'    Else
'        BaseName = Path_
'    End If
'End Function

Rem 削除が完了してればOK
Public Function DeleteFolder(Path) As Boolean
    DeleteFolder = False
    
    If Not fso.FolderExists(Path) Then
        Debug.Print "DeleteFolder(): 対象パス(" & Path & ")が存在しない"
        Exit Function
    End If
    
    On Error Resume Next
    fso.DeleteFolder Path
    If Err.Number <> 0 Then Debug.Print "DeleteFolder(): " & Err.Description, Path
    On Error GoTo 0
    
    DeleteFolder = Not fso.FolderExists(Path)
End Function
