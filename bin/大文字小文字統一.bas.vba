'このソースコード全文を適当なモジュールにコピペした後、Ctrl+Zで戻すことで大文字小文字が統一できる。
'
'ただし、以下の点に注意が必要
'・Declare文のDLL名は、拡張子の有無で別物として扱われる
'・条件付きコンパイルによって、コンパイル対象外となっている範囲は、VBEの自動修正が働かない
'
'本コードでは「.dll無し」で統一している。
'できるだけ「先頭大文字」で統一している。（ただし定義済みの状態を優先したため一部は違う）

'Declareも踏まえた修正手順は、
'1. 置換で「.dll"」を「"」へ（ひとつずつ慎重にやること）
'2. 定義ファイルをコピペ
'3. 置換で「"user32"」を「"User32"」に置換（大文字小文字チェック有りで個別実行が望ましい）

Option Explicit

'WinAPIのDLL文字列
<<<<<<< HEAD
Declare PtrSafe Sub WinAPI01 Lib "Kernel32" ()
Declare PtrSafe Sub WinAPI02 Lib "User32" ()
Declare PtrSafe Sub WinAPI03 Lib "Gdi32" ()
Declare PtrSafe Sub WinAPI04 Lib "GDIPlus" ()
Declare PtrSafe Sub WinAPI05 Lib "Shell32" ()
Declare PtrSafe Sub WinAPI06 Lib "oleacc" ()
Declare PtrSafe Sub WinAPI07 Lib "ole32" ()
Declare PtrSafe Sub WinAPI08 Lib "oleaut32" ()
Declare PtrSafe Sub WinAPI09 Lib "mpr" ()
Declare PtrSafe Sub WinAPI10 Lib "xdwapi" ()
Declare PtrSafe Sub WinAPI11 Lib "advapi32" ()
Declare PtrSafe Sub WinAPI12 Lib "SHLWAPI" ()
Declare PtrSafe Sub WinAPI13 Lib "VBE7" ()
=======
Declare PtrSafe Sub CopyMemory Lib "Kernel32" ()
Declare PtrSafe Function GetAsyncKeyState% Lib "User32" ()
Declare PtrSafe Function CreateCompatibleDC Lib "Gdi32" (ByVal hDc As LongPtr) As Long
Declare PtrSafe Function GdipCreateSolidFill Lib "GDIPlus" ()
Declare PtrSafe Function SHCreateDirectoryEx Lib "Shell32" ()
Declare PtrSafe Function ObjectFromLresult Lib "oleacc" ()
Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As GUID) As Long
Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As PICTDESC, ByRef refiid As GUID, ByVal fPictureOwnsHandle As Long, obj As Any) As Long
Declare PtrSafe Function WNetGetConnection Lib "mpr" Alias "WNetGetConnectionW" ()
Declare PtrSafe Function XDW_Finalize Lib "xdwapi" (ByVal reserved As String) As Long
Declare PtrSafe Function RegCloseKey Lib "advapi32" (ByVal hKey As LongPtr) As Long
Declare PtrSafe Sub ColorRGBToHLS Lib "SHLWAPI" ()
Declare PtrSafe Function rtcCallByName Lib "VBE7" ()
>>>>>>> 13f244c5e163ce0185d9dd0a52abf1c44daec412

'VBA標準関数
Type KeywordUpperLowerCaseUnification_VBA_Function
        
    '数学関数 VBA.Mathメンバ
    Abs As Long
    Atn As Long
    Cos As Long
    Exp As Long
    Log As Long
    Rnd As Long
    Sgn As Long
    Sin As Long
    Sqr As Long
    Tan As Long
    Round As Long 'その他から移動
    
    'データ型変換関数　VBA.Conversion
    CBool As Long
    CByte As Long
    CCur As Long
    CDate As Long
    CDbl As Long
    CDec As Long
    CInt As Long
    CLng As Long
    CLngPtr As Long   '追加
    CSng As Long
    CStr As Long
    CVar As Long
    CVDate As Long
    CVErr As Long     '変換関数から移動
    Error As Long     '変換関数から移動
    Fix As Long       '数学関数から移動
    Hex As Long       '変換関数から移動
    Int As Long       '数学関数から移動
    MacID As Long     'その他から移動
    Oct As Long       '変換関数から移動
    Str As Long       '変換関数から移動
    Val As Long       '変換関数から移動
    
    '文字列関数 VBA.Strings
    Asc As Long       '変換関数から移動
    AscB As Long      '変換関数から移動
    AscW As Long      '変換関数から移動
    Chr As Long       '変換関数から移動
    ChrB As Long      '変換関数から移動
    ChrW As Long      '変換関数から移動
    Filter As Long    'その他から移動
    Format As Long    '変換関数から移動
    FormatCurrency As Long    'その他から移動
    FormatDateTime As Long    'その他から移動
    FormatNumber As Long      'その他から移動
    FormatPercent As Long     'その他から移動
    InStr As Long
    InStrB As Long
    InStrRev As Long
    Join As Long
    LCase As Long
    Left As Long
    LeftB As Long
    Len As Long
    LenB As Long
    Ltrim As Long
    Mid As Long
    MidB As Long
    MonthName As Long
    Replace As Long
    Right As Long
    RightB As Long
    Rtrim As Long
    Space As Long
    Split As Long
    StrComp As Long
    StrConv As Long
    String As Long
    StrReverse As Long
    Trim As Long
    UCase As Long
    WeekdayName As Long
    
    'VBA.[_HiddenModule]
    Array As Long
    Input As Long
    InputB As Long
    ObjPtr As Long    '追加
    StrPtr As Long    '追加
    VarPtr As Long    '追加
    Width As Long     '追加
    
    'VBA.Information
    Erl As Long       '追加
    Err As Long       '追加
    IMEStatus As Long
    IsArray As Long
    IsDate As Long
    IsEmpty As Long
    IsError As Long
    IsMissing As Long
    IsNull As Long
    IsNumeric As Long
    IsObject As Long
    QBColor As Long
    Rgb As Long
    TypeName As Long
    VarType As Long
    
    'VBA.Interaction
    AppActivate As Long     '追加
    Beep As Long     '追加
    CallByName As Long
    Choose As Long
    Command As Long
    CreateObject As Long
    DeleteSetting As Long     '追加
    DoEvents As Long
    Environ As Long
    GetAllSettings As Long
    GetObject As Long
    GetSetting As Long
    IIf As Long
    InputBox As Long
    MacScript As Long
    MsgBox As Long
    Partition As Long
    SaveSetting As Long   '追加
    SendKeys As Long      '追加
    Shell As Long
    Switch As Long
    
    'VBA.FileSystem
    ChDir As Long      '追加
    ChDrive As Long      '追加
    CurDir As Long
    Dir As Long
    EOF As Long
    FileAttr As Long
    FileCopy As Long      '追加
    FileDateTime As Long
    FileLen As Long
    FreeFile As Long
    GetAttr As Long
    Kill As Long      '追加
    Loc As Long
    LOF As Long
    MkDir As Long      '追加
    Reset As Long      '追加
    RmDir As Long      '追加
    Seek As Long
    SetAttr As Long      '追加
    
    'VBA.DateTime
    Calendar As Long      '追加
    Date As Long
    DateAdd As Long
    DateDiff As Long
    DatePart As Long
    DateSerial As Long
    DateValue As Long
    Day As Long
    Hour As Long
    Minute As Long
    Month As Long
    Now As Long
    Second As Long
    Time As Long
    Timer As Long
    TimeSerial As Long
    TimeValue As Long
    Weekday As Long
    Year As Long
    
    'VBA.Financial
    DDB As Long
    FV As Long
    IPmt As Long
    IRR As Long
    MIRR As Long
    NPer As Long
    NPV As Long
    Pmt As Long
    PPmt As Long
    PV As Long
    Rate As Long
    SLN As Long
    SYD As Long
    
    '該当ライブラリ無し
    LoadPicture As Long
    Spc As Long
    Tab As Long
    LBound As Long
    UBound As Long

End Type


'VBA定義済み1
Type KeywordUpperLowerCaseUnification_VBA_Property
    Size        As Long
    Color       As Long
    Destination As Long
    FileFilter  As Long
    Image       As Long
    Appearance  As Long
    Key         As Long
    Keys        As Long
    Items       As Long
    Add         As Long
    Control     As Long
    Controls    As Long
    ListIndex   As Long
    Scroll      As Long
    Pages       As Long
    Number      As Long
    Version     As Long
    Str         As Long
    Val         As Long
End Type


'VBA定義済みステートメント
Type KeywordUpperLowerCaseUnification_VBA_Statement
    Goto                    As Long
    Get                     As Long
    Set                     As Long
    Let                     As Long
    Select                  As Long
    End                     As Long
    Next                    As Long
End Type
'Application.Goto
'GoTo Label

'VBA名前付き引数
Type KeywordUpperLowerCaseUnification_VBA_NamedParam
    Delimiter               As Long
    Target                  As Long
    Sh                      As Long
    No                      As Long
    Key1                    As Long
    Key2                    As Long
    Params                  As Long
    Compare                 As Long
    Cursor                  As Long
    Output                  As Long
    SaveChanges             As Long
End Type

'Excelキーワード
Type KeywordUpperLowerCaseUnification_Excel_Method
    Activate                As Long
    AddComment              As Long
    AddCommentThreaded      As Long
    AdvancedFilter          As Long
    AllocateChanges         As Long
    ApplyNames              As Long
    ApplyOutlineStyles      As Long
    AutoComplete            As Long
    AutoFill                As Long
    AutoFilter              As Long
    AutoFit                 As Long
    AutoOutline             As Long
    BorderAround            As Long
    Calculate               As Long
    CalculateRowMajorOrder  As Long
    CheckSpelling           As Long
    Clear                   As Long
    ClearComments           As Long
    ClearContents           As Long
    ClearFormats            As Long
    ClearHyperlinks         As Long
    ClearNotes              As Long
    ClearOutline            As Long
    ColumnDifferences       As Long
    Consolidate             As Long
    ConvertToLinkedDataType As Long
    Copy                    As Long
    CopyFromRecordset       As Long
    CopyPicture             As Long
    CreateNames             As Long
    Cut                     As Long
    DataTypeToText          As Long
    DataSeries              As Long
    Delete                  As Long
    DialogBox               As Long
    Dirty                   As Long
    DiscardChanges          As Long
    EditionOptions          As Long
    ExportAsFixedFormat     As Long
    FillDown                As Long
    FillLeft                As Long
    FillRight               As Long
    FillUp                  As Long
    Find                    As Long
    FindNext                As Long
    FindPrevious            As Long
    FlashFill               As Long
    FunctionWizard          As Long
    Group                   As Long
    Insert                  As Long
    InsertIndent            As Long
    Justify                 As Long
    ListNames               As Long
    Merge                   As Long
    NavigateArrow           As Long
    NoteText                As Long
    Parse                   As Long
    PasteSpecial            As Long
    PrintOut                As Long
    PrintPreview            As Long
    RemoveDuplicates        As Long
    RemoveSubtotal          As Long
    Replace                 As Long
    RowDifferences          As Long
    Run                     As Long
    Select                  As Long
    SetCellDataTypeFromCell As Long
    SetPhonetic             As Long
    Show                    As Long
    ShowCard                As Long
    ShowDependents          As Long
    ShowErrors              As Long
    ShowPrecedents          As Long
    Sort                    As Long
    SortSpecial             As Long
    Speak                   As Long
    SpecialCells            As Long
    SubscribeTo             As Long
    Subtotal                As Long
    Table                   As Long
    TextToColumns           As Long
    Ungroup                 As Long
    UnMerge                 As Long
    Properties              As Long
    AddIndent               As Long
    Address                 As Long
    AddressLocal            As Long
    AllowEdit               As Long
    Application             As Long
    Areas                   As Long
    Borders                 As Long
    Cells                   As Long
    Characters              As Long
    Column                  As Long
    Columns                 As Long
    ColumnWidth             As Long
    Comment                 As Long
    CommentThreaded         As Long
    Count                   As Long
    CountLarge              As Long
    Creator                 As Long
    CurrentArray            As Long
    CurrentRegion           As Long
    Dependents              As Long
    DirectDependents        As Long
    DirectPrecedents        As Long
    DisplayFormat           As Long
    End                     As Long
    EntireColumn            As Long
    EntireRow               As Long
    Errors                  As Long
    Font                    As Long
    FormatConditions        As Long
    Formula                 As Long
    FormulaArray            As Long
    FormulaHidden           As Long
    FormulaLocal            As Long
    FormulaR1C1             As Long
    FormulaR1C1Local        As Long
    HasArray                As Long
    HasFormula              As Long
    HasRichDataType         As Long
    Height                  As Long
    Hidden                  As Long
    HorizontalAlignment     As Long
    Hyperlinks              As Long
    ID                      As Long
    IndentLevel             As Long
    Interior                As Long
    Item                    As Long
    Left                    As Long
    LinkedDataTypeState     As Long
    ListHeaderRows          As Long
    ListObject              As Long
    LocationInTable         As Long
    Locked                  As Long
    MDX                     As Long
    MergeArea               As Long
    MergeCells              As Long
    Name                    As Long
    Next                    As Long
    NumberFormat            As Long
    NumberFormatLocal       As Long
    Offset                  As Long
    Orientation             As Long
    OutlineLevel            As Long
    PageBreak               As Long
    Parent                  As Long
    Phonetic                As Long
    Phonetics               As Long
    PivotCell               As Long
    PivotField              As Long
    PivotItem               As Long
    PivotTable              As Long
    Precedents              As Long
    PrefixCharacter         As Long
    Previous                As Long
    QueryTable              As Long
    Range                   As Long
    ReadingOrder            As Long
    Resize                  As Long
    Row                     As Long
    RowHeight               As Long
    Rows                    As Long
    ServerActions           As Long
    ShowDetail              As Long
    ShrinkToFit             As Long
    SoundNote               As Long
    SparklineGroups         As Long
    Style                   As Long
    Summary                 As Long
    Text                    As Long
    Top                     As Long
    UseStandardHeight       As Long
    UseStandardWidth        As Long
    Validation              As Long
    Value                   As Long
    Value2                  As Long
    VerticalAlignment       As Long
    Width                   As Long
    Worksheet               As Long
    WrapText                As Long
    XPath                   As Long
End Type

Type KeywordUpperLowerCaseUnification_Excel_Other
    Selection               As Long
    Test                    As Long
    Caption                 As Long
    Col                     As Long
    Cols                    As Long
End Type

'Excel関数
Type KeywordUpperLowerCaseUnification_Excel_WorksheetFunction
    Min                     As Long
    Max                     As Long
    Sum                     As Long
    Index                   As Long
End Type

'FileSystemObject
Type KeywordUpperLowerCaseUnification_FileSystemObject
    'オブジェクト
    Collection              As Long
    Dictionary              As Long
    Drive                   As Long
    Drives                  As Long
    File                    As Long
    Files                   As Long
    FileSystemObject        As Long
    Folder                  As Long
    Folders                 As Long
    TextStream              As Long

    'メソッド
    BuildPath               As Long
    CopyFile                As Long
    CopyFolder              As Long
    CreateFolder            As Long
    CreateTextFile          As Long
    DeleteFile              As Long
    DeleteFolder            As Long
    DriveExists             As Long
    FileExists              As Long
    FolderExists            As Long
    GetAbsolutePathName     As Long
    GetBaseName             As Long
    GetDrive                As Long
    GetDriveName            As Long
    GetExtensionName        As Long
    GetFile                 As Long
    GetFileName             As Long
    GetFolder               As Long
    GetParentFolderName     As Long
    GetSpecialFolder        As Long
    GetTempName             As Long
    Move                    As Long
    MoveFile                As Long
    MoveFolder              As Long
    OpenAsTextStream        As Long
    OpenTextFile            As Long
    WriteLine               As Long

    'プロパティ
    Property                As Long
    'Drives                  As Long
    Name                    As Long
    Path                    As Long
    Size                    As Long
    Type                    As Long
End Type

'その他 APIや参照設定してよく使うライブラリの予約語
Type 他のライブラリで使われる予約語Type
    SaveToFile              As Long
    SetRequestHeader        As Long
    Page                    As Long
    Cursol                  As Long
    Send                    As Long
    Status                  As Long
    Parameter               As Long
    ResponseText            As Long
    Result                  As Long
    hWnd                    As Long
    hDc                     As Long
End Type

'オリジナル一文字変数取得
'これをソースコードに貼ると欲しい状態に変化するので
'結果をExcelに貼って、フラッシュフィルで英字部分を取り出して頭にDim を付ければOK。
'A=1
'B=1
'C=1
'D=1
'E=1
'F=1
'G=1
'H=1
'I=1
'J=1
'K=1
'L=1
'M=1
'N=1
'O=1
'P=1
'Q=1
'R=1
'S=1
'T=1
'U=1
'V=1
'W=1
'X=1
'Y=1
'Z=1

Type 一文字変数Type
    a                       As Long
    b                       As Long
    c                       As Long
    d                       As Long
    e                       As Long
    f                       As Long
    g                       As Long
    h                       As Long
    i                       As Long
    j                       As Long
    k                       As Long
    l                       As Long
    m                       As Long
    n                       As Long
    o                       As Long
    p                       As Long
    q                       As Long
    r                       As Long
    s                       As Long
    t                       As Long
    u                       As Long
    v                       As Long
    w                       As Long
    x                       As Long
    y                       As Long
    z                       As Long
End Type

Type 一文字変数と数字Type
    n1                      As Long
    n2                      As Long
    r1                      As Long
    r2                      As Long
    r3                      As Long
    r4                      As Long
End Type

Type 定数用Type
    PRJ_NAME                As Long
    PROJECT_NAME            As Long
    MOD_NAME                As Long
    MODULE_NAME             As Long
    PROC_NAME               As Long
    FUNC_NAME               As Long
    OLD_NAME                As Long
    NEW_NAME                As Long
    APP_NAME                As Long
    APP_CREATER             As Long
    APP_VERSION             As Long
    APP_SETTINGFILE         As Long
    APP_UPDATE              As Long
    APP_URL                 As Long
End Type

'イレギュラー 上記に該当するが頻繁に多用するため上書きしている名前
'ApplicationイベントではWbだがwbで上書き wsは違うが並べたほうが分かりやすいので。
Type イレギュラーType
    wb                      As Long
    ws                      As Long
End Type

Type オリジナルメンバType
    st                      As Long
    ed                      As Long
    data                    As Long
    data1                   As Long
    data2                   As Long
    data3                   As Long
    data4                   As Long
    handle                  As Long
    dKey                    As Long
    dItem                   As Long
    objHttp                 As Long
    obj                     As Long
    msg                     As Long
    cnt                     As Long
    txt                     As Long
    ret                     As Long
    buf                     As Long
    wsh                     As Long
    bkn                     As Long
    shn                     As Long
    adr                     As Long
    sw                      As Long
    dic                     As Long
    rgs                     As Long
    ir                      As Long
    ok                      As Long
    keta                    As Long
    mode                    As Long
    btn                     As Long
    token                   As Long
    app                     As Long
End Type

'過去に大文字と決めてしまって今更変えられなくなった名前
Type TypeType
    NewName                 As Long
    FName                   As Long
    ProcName                As Long
    BaseRow                 As Long
    BaseCol                 As Long
    LastRow                 As Long
    LastCol                 As Long
    ColDic                  As Long
    TableName               As Long
    URL                     As Long
    OutCol                  As Long
    Adrs                    As Long
    ColumnIndex             As Long
    ColIndex                As Long
    RowIndex                As Long
    CelIndex                As Long
    AddItem                 As Long
    AWin                    As Long
    SX                      As Long
    SY                      As Long
    ZX                      As Long
    ZY                      As Long
End Type

'変更保留メンバ（コード変更と大文字小文字変更のコミットを分離するために使用する

Type 未定Type
    未定 As Long
End Type
