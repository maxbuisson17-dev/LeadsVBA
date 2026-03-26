Attribute VB_Name = "modGlobals"
' ============================================================
' modGlobals.bas
' Constantes et variables globales Public g*
' Ce module doit etre charge AVANT modLeadsWizard
' Toutes les variables g* sont ici pour compatibilite forms
' ============================================================
Option Explicit

'========================
' CONSTANTES
'========================
Public Const EPS_BALANCE_EUR    As Double = 1#
Public Const SH_HOME            As String = "ACCUEIL"
Public Const SH_BG              As String = "BG"
Public Const SH_BS              As String = "BS"
Public Const SH_MAP             As String = "Mapping"
Public Const SH_SIG             As String = "SIG"
Public Const BG_FIRST_ROW       As Long = 2
Public Const COL_A              As Long = 1
Public Const COL_B              As Long = 2
Public Const COL_C              As Long = 3
Public Const COL_D              As Long = 4
Public Const COL_F              As Long = 6
Public Const COL_G              As Long = 7
Public Const COL_H              As Long = 8
Public Const COL_I              As Long = 9

'========================
' VARIABLES GLOBALES
'========================
Public gStep                As String
Public gBalancePath         As String
Public gSkipRows            As Long
Public gPreviewData         As Variant
Public gFullData            As Variant
Public gMaxAcctLen          As Long
Public gControlRows         As Collection
Public gOkToGenerate        As Boolean
Public gHasNonBlockingIssue As Boolean
Public gClient              As String
Public gExercice            As String
Public gVersion             As String
Public gGenerateInKE        As Boolean
Public gExportMode          As eExportMode
Public gImportParams        As tImportBalanceParams
Public gImportParamsReady   As Boolean
Public gPathN               As String
Public gPathN1              As String
Public gArrN                As Variant
Public gArrN1               As Variant
Public gArrCompiled         As Variant
Public gImportedN           As Boolean
Public gImportedN1          As Boolean
Public gEntryPointWasYes    As Boolean

Public gDetailRowRanges    As Object
Public gKScalingApplied    As Boolean
Public gMetaAppliedToParam As Boolean
Public gExportPDF          As Boolean
Public gLastExportSucceeded As Boolean
Public gLastExportedWorkbook As Workbook

' ============================================================
' GESTABLE : onglets de gestion centralises
' SIG, SIG_detail, CAF, BFR, TFT
' Utiliser cette fonction partout au lieu de tester les noms
' en dur pour faciliter la maintenance.
' ============================================================
Public Function IsGestTableSheet(ByVal sheetName As String) As Boolean
    Select Case LCase$(Trim$(sheetName))
        Case "sig", "sig_detail", "caf", "bfr", "tft"
            IsGestTableSheet = True
    End Select
End Function
