
'このソースコード全文を適当なモジュールにコピペした後、Ctrl+Zで戻すことで大文字小文字が統一できる。
Option Explicit


'WinAPIのDLL文字列
'先頭大文字、".dll"無しで統一
Public Declare PtrSafe Function GdipCreateSolidFill Lib "GDIPlus" ()
Public Declare PtrSafe Sub CopyMemory Lib "Kernel32" ()
Public Declare PtrSafe Function GetAsyncKeyState% Lib "User32" ()
Public Declare PtrSafe Function SHCreateDirectoryEx Lib "Shell32" ()


'VBA標準関数
Type VBAKeywordUpperLowerCaseUnification
        
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
Dim Size
Dim Color
Dim Destination
Dim FileFilter
Dim Image
Dim Appearance
Dim Key
Dim Keys
Dim Items
Dim Add
Dim Control
Dim Controls
Dim ListIndex
Dim Scroll
Dim Pages
Dim Number
Dim Version
Dim Str
Dim Val


'VBA定義済み2
Type KeywordUpperLowerCaseUnification
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


'Excelキーワード
Dim Activate
Dim AddComment
Dim AddCommentThreaded
Dim AdvancedFilter
Dim AllocateChanges
Dim ApplyNames
Dim ApplyOutlineStyles
Dim AutoComplete
Dim AutoFill
Dim AutoFilter
Dim AutoFit
Dim AutoOutline
Dim BorderAround
Dim Calculate
Dim CalculateRowMajorOrder
Dim CheckSpelling
Dim Clear
Dim ClearComments
Dim ClearContents
Dim ClearFormats
Dim ClearHyperlinks
Dim ClearNotes
Dim ClearOutline
Dim ColumnDifferences
Dim Consolidate
Dim ConvertToLinkedDataType
Dim Copy
Dim CopyFromRecordset
Dim CopyPicture
Dim CreateNames
Dim Cut
Dim DataTypeToText
Dim DataSeries
Dim Delete
Dim DialogBox
Dim Dirty
Dim DiscardChanges
Dim EditionOptions
Dim ExportAsFixedFormat
Dim FillDown
Dim FillLeft
Dim FillRight
Dim FillUp
Dim Find
Dim FindNext
Dim FindPrevious
Dim FlashFill
Dim FunctionWizard
Dim Group
Dim Insert
Dim InsertIndent
Dim Justify
Dim ListNames
Dim Merge
Dim NavigateArrow
Dim NoteText
Dim Parse
Dim PasteSpecial
Dim PrintOut
Dim PrintPreview
Dim RemoveDuplicates
Dim RemoveSubtotal
Dim Replace
Dim RowDifferences
Dim Run
Dim SetCellDataTypeFromCell
Dim SetPhonetic
Dim Show
Dim ShowCard
Dim ShowDependents
Dim ShowErrors
Dim ShowPrecedents
Dim Sort
Dim SortSpecial
Dim Speak
Dim SpecialCells
Dim SubscribeTo
Dim Subtotal
Dim Table
Dim TextToColumns
Dim Ungroup
Dim UnMerge
Dim Properties
Dim AddIndent
Dim Address
Dim AddressLocal
Dim AllowEdit
Dim Application
Dim Areas
Dim Borders
Dim Cells
Dim Characters
Dim Column
Dim Columns
Dim ColumnWidth
Dim Comment
Dim CommentThreaded
Dim Count
Dim CountLarge
Dim Creator
Dim CurrentArray
Dim CurrentRegion
Dim Dependents
Dim DirectDependents
Dim DirectPrecedents
Dim DisplayFormat
Dim EntireColumn
Dim EntireRow
Dim Errors
Dim Font
Dim FormatConditions
Dim Formula
Dim FormulaArray
Dim FormulaHidden
Dim FormulaLocal
Dim FormulaR1C1
Dim FormulaR1C1Local
Dim HasArray
Dim HasFormula
Dim HasRichDataType
Dim Height
Dim Hidden
Dim HorizontalAlignment
Dim Hyperlinks
Dim ID
Dim IndentLevel
Dim Interior
Dim Item
Dim Left
Dim LinkedDataTypeState
Dim ListHeaderRows
Dim ListObject
Dim LocationInTable
Dim Locked
Dim MDX
Dim MergeArea
Dim MergeCells
Dim Name
Dim NumberFormat
Dim NumberFormatLocal
Dim Offset
Dim Orientation
Dim OutlineLevel
Dim PageBreak
Dim Parent
Dim Phonetic
Dim Phonetics
Dim PivotCell
Dim PivotField
Dim PivotItem
Dim PivotTable
Dim Precedents
Dim PrefixCharacter
Dim Previous
Dim QueryTable
Dim Range
Dim ReadingOrder
Dim Resize
Dim Row
Dim RowHeight
Dim Rows
Dim ServerActions
Dim ShowDetail
Dim ShrinkToFit
Dim SoundNote
Dim SparklineGroups
Dim Style
Dim Summary
Dim Text
Dim Top
Dim UseStandardHeight
Dim UseStandardWidth
Dim Validation
Dim Value
Dim Value2
Dim VerticalAlignment
Dim Width
Dim Worksheet
Dim WrapText
Dim XPath
Dim Selection
Dim Test
Dim Caption


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
dim OutCol
Dim dItem
dim dKey


'変更保留メンバ（コード変更と大文字小文字変更のコミットを分離するために使用する
