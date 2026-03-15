Attribute VB_Name = "modTypes"

'== modTypes.bas ==
Option Explicit

Public Enum eImportSource
    srcFEC = 1
    srcBalance = 2
End Enum

Public Type tImportBalanceParams
    SourceN As eImportSource
    pathN As String
    SourceN1 As eImportSource
    pathN1 As String
    DropZero As Boolean
    RoundEuro As Boolean
    ToKE As Boolean
    KeepAN As Boolean
End Type

Public Enum eBalance4ColsMode
    b4None = 0
    b4NN1 = 1
    b4DebitCredit = 2
    b4NN1_ColD = 3
End Enum

Public Enum eExportMode
    emAll = 1
    emFS = 2
    emLeads = 3
End Enum


Public Enum eBalanceFormat
    bfUnknown = 0
    bf3ColsSolde = 1
    bf4ColsDebitCredit = 2
    bfHeuristicMapped = 3
End Enum

' ============================================================
' Type session : remplace les 16 variables globales g*
' ============================================================
Public Type tLeadsSession
    ' Import
    BalancePath     As String
    pathN           As String
    pathN1          As String
    arrN()          As Variant
    arrN1()         As Variant
    ArrCompiled()   As Variant
    FullData()      As Variant
    MaxAcctLen      As Long
    skipRows        As Long
    ImportedN       As Boolean
    ImportedN1      As Boolean
    ImportParamsReady As Boolean
    ImportParams    As tImportBalanceParams
    ' Controles
    ControlRows     As Collection
    OkToGenerate    As Boolean
    HasNonBlockingIssue As Boolean
    ' Metadonnees
    client          As String
    Exercice        As String
    Version         As String
    ' Options generation
    GenerateInKE    As Boolean
    exportMode      As eExportMode
    KScalingApplied As Boolean
    MetaAppliedToParam As Boolean
    EntryPointWasYes As Boolean
    ' Runtime
    Step            As String
    PreviewData()   As Variant
End Type

' Type utilitaire detection colonnes balance (utilise par modBalanceCreator.DetectColumns)
Public Type tBalanceDetectResult
    Format      As eBalanceFormat
    hasHeader   As Boolean
    accountCol  As Long
    labelCol    As Long
    soldeCol    As Long
    debitCol    As Long
    creditCol   As Long
    lastRow     As Long
    lastCol     As Long
End Type
