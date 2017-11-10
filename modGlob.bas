Attribute VB_Name = "modGlob"
Option Explicit

'====================================================================================
'   definizione costanti globali
'====================================================================================
Public Const Debug_PROGRAMMA = 0
Public LottoSede As String
Public strCodGruppo As String

Public Const PesoBancale = 20
Public Const PesoBobina = 0.9

Public Const C_SCH_TESTA = 0
Public Const C_SCH_RIGHE = 1
Public Const C_SCH_PIEDE = 2
Public Const C_SCH_SITUAZ = 3
Public Const C_SCH_CTRLALLEST = 4
Public Const C_SCH_NOTIFICADOC = 5
Public Const C_SCH_PALLET_INFO = 6
Public Const C_SCH_SELCLIFOR = 7
Public Const C_SCH_PALLET_ASSIGN = 8
Public Const C_SCH_ARTICOLI = 9
Public Const C_SCH_SPLASH = 10
Public Const C_SCH_REGPALLET = 11
Public Const C_SCH_TMP = 12
Public Const C_SCH_CODARTTMP = 13
Public Const C_SCH_LOTTI = 14
Public Const C_SCH_OF = 15
Public Const C_SCH_INV = 16
Public Const C_SCH_RMP = 17

Public Const C_IND_CATTURATESTO = 0
Public Const C_IND_IDDOC = 1
Public Const C_IND_CODART = 2
Public Const C_IND_QTA = 3
Public Const C_IND_DSCART = 4
Public Const C_IND_QTARES = 5
Public Const C_IND_MATRICOLA = 6
Public Const C_IND_DATACONS = 7
Public Const C_IND_IDTESTA = 8
Public Const C_IND_IDRIGA = 9
Public Const C_IND_NRRIGA = 10
Public Const C_IND_CODCLIFOR = 11
Public Const C_IND_DSCCLIFOR = 12
Public Const C_IND_QTAORIG = 13
Public Const C_IND_BARCODE = 14
Public Const C_IND_NRACQUISIZ = 15
Public Const C_IND_STATORIGA = 16
Public Const C_IND_IDTESTAALLEST = 17
Public Const C_IND_COLLI = 0 '18
Public Const C_IND_PESO = 8 '19
Public Const C_IND_VOLUME = 20
Public Const C_IND_LOTTO = 21
Public Const C_IND_PALLET = 22
Public Const C_IND_TIPOCF = 23
Public Const C_IND_PALLET_BC = 24
Public Const C_IND_QTAPERPALLET = 25
Public Const C_IND_PESONETTO = 46
Public Const C_IND_DATATRASPORTO = 47
Public Const C_IND_ORATRASPORTO = 48
Public Const C_IND_NUMRIFDOC = 10
Public Const C_IND_DATARIFDOC = 9
Public Const C_IND_REGPALLET = 11
Public Const C_IND_F_PALLET = 1


Public Const C_CMB_SPED = 10 '0
Public Const C_CMB_ASPBENI = 9 '1
Public Const C_CMB_DOCPREL = 2
Public Const C_CMB_ALLEST = 3
Public Const C_CMB_PALLET = 4
Public Const C_CMB_PALLET_INFO = 5

Public Const C_CMB_PORTO = 6
Public Const C_CMB_CAUSALE = 7
Public Const C_CMB_TRASP = 8

Public Const C_MOV_PRIMO = 0
Public Const C_MOV_PREC = 1
Public Const C_MOV_SUCC = 2
Public Const C_MOV_ULTIMO = 3

Public Const C_SCH_OK = 0
Public Const C_SCH_CANC = 1

Public Const C_SCH_SPED = 0
Public Const C_SCH_ASPBENI = 1


Public Const C_BCTYP_NONE = 0
Public Const C_BCTYP_BARCODE = 1
Public Const C_BCTYP_BARCODE_NEW = 2
Public Const C_BCTYP_MATRLIB = 3
Public Const C_BCTYP_MATRIMP = 4
Public Const C_BCTYP_MATRACQ = 5

Public Const C_PALLET_START = 0
Public Const C_PALLET_STOP = 1
Public Const C_PALLET_INFO = 2
Public Const C_PALLET_ASSIGN = 3

Public Const C_CONS_COLLI = 0
Public Const C_CONS_PESO = 1
Public Const C_CONS_VOLUME = 2

'====================================================================================
'   definizione variabili globali
'====================================================================================
Public strSql As String
Public strCodiceScelto As String
Public dblQtaTmp As Double
Public strCodForRP As String

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Type Inventario
    Codart As String
    DescArt As String
    Codlotto As String
    Giacenza As Double
    QtaTerminale As Double
    QtaTerminale2 As Double
    QtaPallet As Double
    QtaBobine As Double
End Type

Public Inv() As Inventario
