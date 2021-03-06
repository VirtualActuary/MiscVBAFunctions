Attribute VB_Name = "MiscErrorEnums"
' Use a* so this is on top of the fn. MiscF library

Option Explicit
Enum ErrNr
    '********************************************
    'These are internal error codes collected from
    'https://bettersolutions.com/vba/error-handling/error-codes.htm
    '********************************************
    ReturnWithoutGoSub = 3
    InvalidProcedureCall = 5
    Overflow = 6
    OutOfMemory_ = 7
    SubscriptOutOfRange = 9
    ThisArrayIsFixedOrTemporarilyLocked = 10
    DivisionByZero = 11
    TypeMismatch = 13
    OutOfStringSpace = 14
    ExpressionTooComplex = 16
    CantPerformRequestedOperation = 17
    UserInterruptOccurred = 18
    ResumeWithoutError = 20
    OutOfStackSpace = 28
    SubFunctionOrPropertyNotDefined = 35
    TooManyDLLApplicationClients = 47
    ErrorInLoadingDLL = 48
    BadDLLCallingConvention = 49
    InternalError = 51
    BadFileNameOrNumber = 52
    FileNotFound = 53
    BadFileMode = 54
    FileAlreadyOpen = 55
    DeviceIOError = 57
    FileAlreadyExists = 58
    BadRecordLength = 59
    DiskFull = 61
    InputPastEndOfFile = 62
    BadRecordNumber = 63
    TooManyFiles = 67
    DeviceUnavailable = 68
    PermissionDenied = 70
    DiskNotReady = 71
    CantRenameWithDifferentDrive = 74
    PathFileAccessError = 75
    PathNotFound = 76
    ObjectVariableOrWithBlockVariableNotSet = 91
    ForLoopNotInitialized = 92
    InvalidPatternString = 93
    InvalidUseOfNull = 94
    CantCallFriendProcedureOnAnObjectThatIsNotAnInstanceOfTheDefiningClass = 97
    SystemDLLCouldNotBeLoaded = 298
    CantUseCharacterDeviceNamesInSpecifiedFileNames = 320
    InvalidFileFormat = 321
    CantCreateNecessaryTemporaryFile = 322
    InvalidFormatInResourceFile = 325
    DataValueNamedNotFound = 327
    IllegalParameterCantWriteArrays = 328
    CouldNotAccessSystemRegistry = 335
    ActiveXComponentNotCorrectlyRegistered = 336
    ActiveXComponentNotFound = 337
    ActiveXComponentDidNotRunCorrectly = 338
    ObjectAlreadyLoaded = 360
    CantLoadOrUnloadThisObject = 361
    ActiveXControlSpecifiedNotFound = 363
    ObjectWasUnloaded = 364
    UnableToUnloadWithinThisContext = 365
    TheSpecifiedFileIsOutOfDateThisProgramRequiresALaterVersion = 368
    TheSpecifiedObjectCantBeUsedAsAnOwnerFormForShow = 371
    InvalidPropertyValue = 380
    InvalidPropertyarrayIndex = 381
    PropertySetCantBeExecutedAtRunTime = 382
    PropertySetCantBeUsedWithAReadonlyProperty = 383
    NeedPropertyArrayIndex = 385
    PropertySetNotPermitted = 387
    PropertyGetCantBeExecutedAtRunTime = 393
    PropertyGetCantBeExecutedOnWriteonlyProperty = 394
    FormAlreadyDisplayedCantShowModally = 400
    CodeMustCloseTopmostModalFormFirst = 402
    PermissionToUseObjectDenied = 419
    PropertyNotFound = 422
    PropertyOrMethodNotFound = 423
    ObjectRequired = 424
    InvalidObjectUse = 425
    ActiveXComponentCantCreateObjectOrReturnReferenceToThisObject = 429
    ClassDoesntSupportAutomation = 430
    FileNameOrClassNameNotFoundDuringAutomationOperation = 432
    ObjectDoesntSupportThisPropertyOrMethod = 438
    AutomationError = 440
    ConnectionToTypeLibraryOrObjectLibraryForRemoteProcessHasBeenLost = 442
    AutomationObjectDoesntHaveADefaultValue = 443
    ObjectDoesntSupportThisAction = 445
    ObjectDoesntSupportNamedArguments = 446
    ObjectDoesntSupportCurrentLocaleSetting = 447
    NamedArgumentNotFound = 448
    ArgumentNotOptionalOrInvalidPropertyAssignment = 449
    WrongNumberOfArgumentsOrInvalidPropertyAssignment = 450
    ObjectNotACollection = 451
    InvalidOrdinal = 452
    SpecifiedDLLFunctionNotFound = 453
    CodeResourceNotFound = 454
    CodeResourceLockError = 455
    ThisKeyIsAlreadyAssociatedWithAnElementOfThisCollection = 457
    VariableUsesATypeNotSupportedInVisualBasic = 458
    ThisComponentDoesntSupportEvents = 459
    InvalidClipboardFormat = 460
    SpecifiedFormatDoesntMatchFormatOfData = 461
    CantCreateAutoRedrawImage = 480
    InvalidPicture = 481
    PrinterError = 482
    PrinterDriverDoesNotSupportSpecifiedProperty = 483
    ProblemGettingPrinterInformationFromTheSystemMakeSureThePrinterIsSetUpCorrectly = 484
    InvalidPictureType = 485
    CantPrintFormImageToThisTypeOfPrinter = 486
    CantEmptyClipboard = 520
    CantOpenClipboard = 521
    CantSaveFileToTEMPDirectory = 735
    SearchTextNotFound = 744
    ReplacementsTooLong = 746
    OutOfMemory = 31001
    NoObject = 31004
    ClassIsNotSet = 31018
    UnableToActivateObject = 31027
    UnableToCreateEmbeddedObject = 31032
    ErrorSavingToFile = 31036
End Enum
