'���̃\�[�X�R�[�h�S����K���ȃ��W���[���ɃR�s�y������ACtrl+Z�Ŗ߂����Ƃő啶��������������ł���B
'
'�������A�ȉ��̓_�ɒ��ӂ��K�v
'�EDeclare����DLL���́A�g���q�̗L���ŕʕ��Ƃ��Ĉ�����
'�E�����t���R���p�C���ɂ���āA�R���p�C���ΏۊO�ƂȂ��Ă���͈͂́AVBE�̎����C���������Ȃ�
'
'�{�R�[�h�ł́u.dll�����v�œ��ꂵ�Ă���B
'�ł��邾���u�擪�啶���v�œ��ꂵ�Ă���B�i��������`�ς݂̏�Ԃ�D�悵�����߈ꕔ�͈Ⴄ�j

'Declare�����܂����C���菇�́A
'1. �u���Łu.dll"�v���u"�v�ցi�ЂƂ��T�d�ɂ�邱�Ɓj
'2. ��`�t�@�C�����R�s�y
'3. �u���Łu"user32"�v���u"User32"�v�ɒu���i�啶���������`�F�b�N�L��Ōʎ��s���]�܂����j

Option Explicit

'WinAPI��DLL������
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

'VBA�W���֐�
Type KeywordUpperLowerCaseUnification_VBA_Function
        
    '���w�֐� VBA.Math�����o
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
    Round As Long '���̑�����ړ�
    
    '�f�[�^�^�ϊ��֐��@VBA.Conversion
    CBool As Long
    CByte As Long
    CCur As Long
    CDate As Long
    CDbl As Long
    CDec As Long
    CInt As Long
    CLng As Long
    CLngPtr As Long   '�ǉ�
    CSng As Long
    CStr As Long
    CVar As Long
    CVDate As Long
    CVErr As Long     '�ϊ��֐�����ړ�
    Error As Long     '�ϊ��֐�����ړ�
    Fix As Long       '���w�֐�����ړ�
    Hex As Long       '�ϊ��֐�����ړ�
    Int As Long       '���w�֐�����ړ�
    MacID As Long     '���̑�����ړ�
    Oct As Long       '�ϊ��֐�����ړ�
    Str As Long       '�ϊ��֐�����ړ�
    Val As Long       '�ϊ��֐�����ړ�
    
    '������֐� VBA.Strings
    Asc As Long       '�ϊ��֐�����ړ�
    AscB As Long      '�ϊ��֐�����ړ�
    AscW As Long      '�ϊ��֐�����ړ�
    Chr As Long       '�ϊ��֐�����ړ�
    ChrB As Long      '�ϊ��֐�����ړ�
    ChrW As Long      '�ϊ��֐�����ړ�
    Filter As Long    '���̑�����ړ�
    Format As Long    '�ϊ��֐�����ړ�
    FormatCurrency As Long    '���̑�����ړ�
    FormatDateTime As Long    '���̑�����ړ�
    FormatNumber As Long      '���̑�����ړ�
    FormatPercent As Long     '���̑�����ړ�
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
    ObjPtr As Long    '�ǉ�
    StrPtr As Long    '�ǉ�
    VarPtr As Long    '�ǉ�
    Width As Long     '�ǉ�
    
    'VBA.Information
    Erl As Long       '�ǉ�
    Err As Long       '�ǉ�
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
    AppActivate As Long     '�ǉ�
    Beep As Long     '�ǉ�
    CallByName As Long
    Choose As Long
    Command As Long
    CreateObject As Long
    DeleteSetting As Long     '�ǉ�
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
    SaveSetting As Long   '�ǉ�
    SendKeys As Long      '�ǉ�
    Shell As Long
    Switch As Long
    
    'VBA.FileSystem
    ChDir As Long      '�ǉ�
    ChDrive As Long      '�ǉ�
    CurDir As Long
    Dir As Long
    EOF As Long
    FileAttr As Long
    FileCopy As Long      '�ǉ�
    FileDateTime As Long
    FileLen As Long
    FreeFile As Long
    GetAttr As Long
    Kill As Long      '�ǉ�
    Loc As Long
    LOF As Long
    MkDir As Long      '�ǉ�
    Reset As Long      '�ǉ�
    RmDir As Long      '�ǉ�
    Seek As Long
    SetAttr As Long      '�ǉ�
    
    'VBA.DateTime
    Calendar As Long      '�ǉ�
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
    
    '�Y�����C�u��������
    LoadPicture As Long
    Spc As Long
    Tab As Long
    LBound As Long
    UBound As Long

End Type


'VBA��`�ς�1
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


'VBA��`�ς݃X�e�[�g�����g
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

'VBA���O�t������
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

'Excel�L�[���[�h
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

'Excel�֐�
Type KeywordUpperLowerCaseUnification_Excel_WorksheetFunction
    Min                     As Long
    Max                     As Long
    Sum                     As Long
    Index                   As Long
End Type

'FileSystemObject
Type KeywordUpperLowerCaseUnification_FileSystemObject
    '�I�u�W�F�N�g
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

    '���\�b�h
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

    '�v���p�e�B
    Property                As Long
    'Drives                  As Long
    Name                    As Long
    Path                    As Long
    Size                    As Long
    Type                    As Long
End Type

'���̑� API��Q�Ɛݒ肵�Ă悭�g�����C�u�����̗\���
Type ���̃��C�u�����Ŏg����\���Type
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

'�I���W�i���ꕶ���ϐ��擾
'������\�[�X�R�[�h�ɓ\��Ɨ~������Ԃɕω�����̂�
'���ʂ�Excel�ɓ\���āA�t���b�V���t�B���ŉp�����������o���ē���Dim ��t�����OK�B
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

Type �ꕶ���ϐ�Type
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

Type �ꕶ���ϐ��Ɛ���Type
    n1                      As Long
    n2                      As Long
    r1                      As Long
    r2                      As Long
    r3                      As Long
    r4                      As Long
End Type

Type �萔�pType
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

'�C���M�����[ ��L�ɊY�����邪�p�ɂɑ��p���邽�ߏ㏑�����Ă��閼�O
'Application�C�x���g�ł�Wb����wb�ŏ㏑�� ws�͈Ⴄ�����ׂ��ق���������₷���̂ŁB
Type �C���M�����[Type
    wb                      As Long
    ws                      As Long
End Type

Type �I���W�i�������oType
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

'�ߋ��ɑ啶���ƌ��߂Ă��܂��č��X�ς����Ȃ��Ȃ������O
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

'�ύX�ۗ������o�i�R�[�h�ύX�Ƒ啶���������ύX�̃R�~�b�g�𕪗����邽�߂Ɏg�p����

Type ����Type
    ���� As Long
End Type
