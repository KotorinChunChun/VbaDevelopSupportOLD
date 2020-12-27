
'このソースコード全文を適当なモジュールにコピペした後、Ctrl+Zで戻すことで大文字小文字が統一できる。
Option Explicit

'大文字小文字反転には弱点がある。
'コンパイル制御をしている場合
'32bitでは反映される。
'64bitではもう一方には変更が反映されない？？？

'WinAPIのDLL文字列
'先頭大文字、".dll"無しで統一
Public Declare PtrSafe Sub CopyMemory Lib "Kernel32" ()
Public Declare PtrSafe Function GetAsyncKeyState% Lib "User32" ()
Public Declare PtrSafe Function CreateCompatibleDC Lib "GDI32" ( ByVal hDc As LongPtr ) As Long
Public Declare PtrSafe Function GdipCreateSolidFill Lib "GDIPlus" ()
Public Declare PtrSafe Function SHCreateDirectoryEx Lib "Shell32" ()
Public Declare PtrSafe Function ObjectFromLresult Lib "oleacc" ()
Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As GUID) As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As PICTDESC, ByRef refiid As GUID, ByVal fPictureOwnsHandle As Long, obj As Any) As Long
Private Declare PtrSafe Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionW" ()
Declare PtrSafe Function XDW_Finalize Lib "xdwapi.dll" (ByVal reserved As String) As Long
Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long
Public Declare PtrSafe Sub ColorRGBToHLS Lib "SHLWAPI.DLL" ()

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
    Goto As Long
    Get As Long
    Set As Long
    Let As Long
    Select As Long
    End As Long
    Next As Long
End Type
'Application.Goto
'GoTo Label

'VBA名前付き引数
Type KeywordUpperLowerCaseUnification_VBA_NamedParam
    Delimiter As Long
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
    Selection As Long
    Test      As Long
    Caption   As Long
    Col       As Long
    Cols      As Long
End Type

'Excel関数
Type KeywordUpperLowerCaseUnification_Excel_WorksheetFunction
    Min As Long
    Max As Long
    Sum as Long
End Type

'その他　未分類
Type KeywordUpperLowerCaseUnification_Other
    SaveToFile          As Long
    SetRequestHeader    As Long
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

'オリジナル一文字変数
Dim a
Dim b
Dim C
Dim d
Dim E
Dim f
Dim g
Dim H
Dim i
Dim j
Dim k
Dim l
Dim M
Dim n
Dim o
Dim p
Dim Q
Dim r
Dim s
Dim t
Dim u
Dim v
Dim W
Dim x
Dim y
Dim Z


'オリジナルメンバ
Dim st
Dim data
Dim data1
Dim data2
Dim data3
Dim data4
Dim handle
Dim N2
Dim hWnd
Dim hDc
Dim BaseRow
Dim BaseCol
Dim LastRow
Dim LastCol
Dim ColDic
Dim TableName
Dim URL
Dim OutCol
Dim dItem
Dim dKey

'暫定確定メンバ（上記に移動する前の一時置き場）
Dim Page
Dim Cursol
Dim Send
Dim Status
Dim NewName
Dim Parameter
Dim objHttp
Dim FName
Dim ResponseText
Dim msg
Dim PROC_NAME
Dim cnt
dim ProcName
dim Result

Type 未定Type
    Output  As Long
    Test    As Long
    OLD_NAME    As Long
    NEW_NAME    As Long
            
    Cursor  As Long
            
    wsh As Long
    wb As Long 'ApplicationイベントでWbだがwbで上書き
    ws As Long
    No  As Long
    ColumnIndex As Long
    Key1    As Long
    Key2    As Long
    Adrs    As Long
    SHN As Long
    ColIndex    As Long
    CelIndex    As Long
            
    Target  As Long
    Sh  As Long
    Index   As Long
    AddItem As Long
            
    AWin    As Long
    adr As Long
    sw  As Long
    dic As Long
    rgs As Long
    
    SX   As Long
    SY   As Long
    ZX   As Long
    ZY   As Long
    Params As Long
End Type

'変更保留メンバ（コード変更と大文字小文字変更のコミットを分離するために使用する



Dim IR
Dim R3
Dim ok
dim txt
