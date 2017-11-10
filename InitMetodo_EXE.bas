Attribute VB_Name = "InitMetodo"
Option Explicit
DefLng A-Z


'####################################################################################################################
'MObjKitBus
'####################################################################################################################

'dichiarazioni di metodo.bas condivise con mwserver
Global MXNU As MXNucleo.XNucleo
Global MXDB As MXKit.XODBC

Global hndDBArchivi As MXKit.CConnessione

Global MXCREP As MXKit.CAmbCRW
Global MXAA As MXKit.CAmbAgenti
Global MXCT As MXKit.CAmbTab
Global MXVI As MXKit.CAmbVisioni
Global MXVA As MXKit.CAmbValid
Global MXFT As MXKit.CAmbFiltri
Global MXWKF As MXKit.CAmbWorkFlow

Global MXSC As MXBusiness.CAmbScad
Global MXART As MXBusiness.CAmbVArt
Global MXSM As MXBusiness.CAmbStMag       'movimentazione storico magazzino
Global MXDBA As MXBusiness.CAmbDba        'gestione distinta base
Global MXGD As MXBusiness.CAmbGestDoc
Global MXPIAN As MXBusiness.CAmbPian
Global MXPN As MXBusiness.CAmbPN          'Prima Nota e Cespiti
Global MXPROD As MXBusiness.CAmbProd      'ambiente ordini di produzione
Global MXCICLI As MXBusiness.CAmbCicliLav 'ambiente cicli lavorazione
Global MXCC As MXBusiness.CAmbCommCli     'ambiente commesse clienti
Global MXRIS As MXBusiness.CAmbRisorse    'ambiente gestione risorse
Global MXSCH As MXBusiness.cAmbSched      'ambiente schedulazione
Global NETFX As Object 'ambiente dot net
Global GBolWorkflow As Boolean
'REMIND: modifiche per MXConsole
'Global MXALL As MXConsole.CAmbConsole
Global MXALL As Object

'REMIND: modifiche per Quality
Global MXQM As Object

'Modifiche per Wizard
Global MXWIZARD As Object

Private Enum setModuliRunTime
    MD32_KIT = 150
    MD32_BUSINESS_DBA = 160
    MD32_BUSINESS_PRIMANOTA = 161
    MD32_BUSINESS_SCADENZE = 162
    MD32_BUSINESS_STORICO = 163
    MD32_BUSINESS_DOCUMENTI = 164
    MD32_BUSINESS_PIANIFICAZIONE = 165
    MD32_BUSINESS_CTRLCODARTICOLO = 166
    MD32_BUSINESS_PRODUZIONE = 167
    MD32_BUSINESS_CICLILAVORAZIONE = 168
    MD32_BUSINESS_COMMESSECLIENTI = 169
    MD32_BUSINESS_GESTIONERISORSE = 170
    MD32_BUSINESS_SCHEDULAZIONE = 171
End Enum

'####################################################################################################################
'MGlobal
'####################################################################################################################

Public MbolInChiusura As Boolean

Global istanzecreateclassi As Integer


'####################################################################################################################
'MObjKitBus
'####################################################################################################################


Private mBolCambioUtente As Boolean

Private MBolSaltaMessaggiConnessione As Boolean  'Per evitare i messaggi di riconnessione in caso di annullamento della selezione Ditta su Evolus (segnalazione Evolus Nr. 27)

Public InSelezioneDitta As Boolean

'Rif anomalia #3160
Global bolTrustedConnection As Boolean

'####################################################################################################################
'MCostanti
'####################################################################################################################




' nome degli oggetti plug-in
Global Const OGGETTO_ESTENSIONE = "objExt"
Global Const OGGETTO_WRAPPER_ESTENSIONE = "objExtWrapper"

'DATABASE
Global Const NOME_DB_ABICAB = "abicab"
'Global Const NOME_DB_DITTE = "metditte"
'Global Const NOME_DB_OGG = "metogg"
Global Const NOME_DB_ARCHIVI = "metxxxx"

Global Const NOME_FILE_MENU_TMP = "menu.ini"
Global Const NOME_FILE_MENUTOOLS_TMP = "menutools.ini"
Global Const NOME_FILE_MENUPERS_TMP = "menup.ini"
Global Const NOME_FILE_MENUPERSDITTA_TMP = "menud.ini"
Global Const NOME_FILE_MENUHIDDEN_TMP = "menuh.xml"

'errori
Global Const ERR_UTENTE = 3059
Global Const ERR_DBVARIAZIONE = 20000



Global Const KEY_F1 = &H53    '&H70 -> sostituito con S
Global Const HOURGLASS = 11     ' 11 - Hourglass

Global Const SHIFT_MASK = 1
Global Const CTRL_MASK = 2
Global Const ALT_MASK = 4

Global Const MB_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
Global Const IDYES = 6                 ' Yes button pressed
Global Const IDNO = 7                  ' No button pressed

Global Const OLE_ACTIVATE = 7

'Common Dialog Control
'Action Property
Global Const DLG_FILE_OPEN = 1
Global Const DLG_FILE_SAVE = 2
Global Const DLG_COLOR = 3
Global Const DLG_FONT = 4
Global Const DLG_PRINT = 5
Global Const DLG_HELP = 6

'tasti corrispondenti ai bottoni della toolbox
Global Const BTN_INS = 73 'I
Global Const BTN_MOD = 68 '"D"
Global Const BTN_REG = 82 '"R"
Global Const BTN_PRIMO = 49 '"1"
Global Const BTN_PREC = 50 '"2"
Global Const BTN_SUCC = 51 '"3"
Global Const BTN_ULTIMO = 52 '"4"
Global Const BTN_ANN = 65 '"A"
' S#3040 - rimossa la Gestione Accessi da Evolus
'Global Const BTN_DEF_ACC = 85 '"U"
Global Const BTN_STP = 80 '"P"
Global Const BTN_DUP = 48 '"0"
Global Const BTN_VISUTMOD = 77 '"M"
Global Const BTN_ALLINONE = 78 '"N"
Global Const BTN_ZOOM = 90 '"N"
Global Const BTN_TROVA = 84 '"T"
Global Const BTN_DESIGNER = 87 '"W"


'maschere dei tasti
Global Const BTN_INS_MASK = &H1
Global Const BTN_MOD_MASK = &H2
Global Const BTN_REG_MASK = &H4
Global Const BTN_PRIMO_MASK = &H8
Global Const BTN_PREC_MASK = &H10
Global Const BTN_SUCC_MASK = &H20
Global Const BTN_ULTIMO_MASK = &H40
Global Const BTN_ANN_MASK = &H80
Global Const BTN_TUTTI_MASK = &H1 + &H2 + &H4 + &H8 + &H10 + &H20 + &H40 + &H80
Global Const BTN_STP_MASK = &H100

Global Const SQL_SUCCESS As Long = 0
Global Const SQL_FETCH_NEXT As Long = 1
Global Const ODBC_ADD_SYS_DSN = 4
Global Const ODBC_REMOVE_SYS_DSN = 6

'formule preventivo commessa
Global MobjScriptFormuleCC As Object

'--------- costanti gestione movimenti ------------------------
Public Enum setGestioneMovimentiAzione
    REC_ANNULLA = 1
    REC_INSERISCI = 2
    REC_MODIFICA = 3
End Enum

'Public Enum setGestioneMovimentiOperazione
'    MOV_AGGIORNA = 1
'    MOV_STORNO = -1
'End Enum

'costanti per TIPOMOV:
'Public Enum setStoricoTipoMovimento
'    ST_MOV_MANUALE = 0
'    ST_MOV_RIGADOC = 1
'    ST_MOV_RIGADOC_COLL = 2
'    ST_MOV_COMP = 3
'    ST_MOV_COMP_COLL = 4
'    ST_MOV_COMPCOMM = 5           'Componenti Commessa Prod.
'    ST_MOV_COMPCOMM_COLL = 6      'Componenti Commessa Prod. Collegati
'End Enum

'*** DEFINIZIONI TIPI DI DATI ***
'*** METODO.BAS ***
Type Proprieta_Aggiuntive
     'Tipo As String * 1 'vedere sotto le costanti
     Tipo As setTipoInput
     DataF As String
     frmt As String
     dflt As Variant
     ValCorrente As Variant 'per gestire le modifiche del campo
End Type

Type SS_Prop_Aggiuntive
    Row As Long
    Col As Long
    Tipo As String * 1 'vedere sotto le costanti
    DataF As String
    'frmt As String
    dflt As Variant
    'ValCorrente As Variant 'per gestire le modifiche del campo
End Type

'*** MAGAZZIN.BAS ***
Type Dati_tipologia
    cod As String * 2     'da TabTipologie e da TipologieArticoli
    des As String * 25
    CTRLEs As Integer
    lngvar As Integer
    SelVar As String * 1
    aggDes As Integer
    varcar As Integer
    nr As Integer
    hsnapvar As Integer 'indice dello snapshot delle varianti usato solo per la generazione automatica
End Type

Type StrDisponibilita
    Giacenze1UM(1 To 10) As Currency    'Giacenze Prima Unità di Misura
    TotGiacenze1UM As Currency          'Totale Giacenze 1UM
    Giacenze2UM(1 To 10) As Currency    'Giacenze Seconda Unità di Misura
    TotGiacenze2UM As Currency          'Totale Giacenze 2UM
    Ordinato(1 To 2) As Currency        'Ordinato 1UM e 2UM
    Impegnato(1 To 2) As Currency       'Impegnato 1UM e 2UM
    GiacenzaIniziale(1 To 2) As Currency 'Giacenza Iniziale 1UM e 2UM
End Type

'Struttura per l'aggiornamento Prezzo/Sconto nella tabella GestionePrezzi
Type ParPrezziSconto
    ProgV As Long
    ProgN As Long
    CliFor As String
    CodArt As String
    Listino As Integer
    prez As Variant
    DataInizioVal As String
    tipocampo As Integer  '1=Prezzo 2=Sconto
End Type

'Struttura per la ricerca del Magazzino
Type DatiRicMag
    TipoRicerca As Integer     '0=Ricerca su Parametri Doc; 1=Ricerca su Parametri Ord. Prod.
    CodiceDoc As String
    CodiceArt As String
    CodConto As String
    NumDestDiv As Integer
    TipoMag As Integer
    CodMagRP As String   'Codice Magazzino della Riga Prodotto
End Type

'struttura per la stampa differita documenti
Type Parametri_Stampa_Documento
    NrTerminale As Integer
    AnnoDoc As Integer '(rif 10)
    TipoDoc As String
    NumeroDoc As Long
    Bis As String
    DataDoc As String
    CodConto As String
    Lingua As Integer
    StampaVar As String
    'StampaDscLingua As Integer
    OpzioniStampa As Integer
    StampaDistBase As String
    SaltoPag    As Integer
    StampaInfo  As Integer
    DEVStampa As String
    DEVStampaInfo As String
    ModuloStampaDist As String
    DEVStampaEtic As String
    ModuloStampaEtic As String
    TipoStampaEtic As Integer
End Type


'       COSTANTI
Global Const LISTA_TABELLE_STD = 0
Global Const LISTA_VALIDAZIONI_STD = 1
Global Const LISTA_VISIONI_STD = 2
Global Const LISTA_SITUAZIONI_STD = 3
Global Const LISTA_ANAGRAFICHE_STD = 4
Global Const LISTA_MULTIANAGRAFICHE_STD = 5
Global Const LISTA_TABELLE_PERS = 6
Global Const LISTA_VALIDAZIONI_PERS = 7
Global Const LISTA_VISIONI_PERS = 8
Global Const LISTA_SITUAZIONI_PERS = 9
Global Const LISTA_ANAGRAFICHE_PERS = 10
Global Const LISTA_MULTIANAGRAFICHE_PERS = 11
Global Const LISTA_TABELLE_PERSDITTA = 12
Global Const LISTA_VALIDAZIONI_PERSDITTA = 13
Global Const LISTA_VISIONI_PERSDITTA = 14
Global Const LISTA_SITUAZIONI_PERSDITTA = 15
Global Const LISTA_ANAGRAFICHE_PERSDITTA = 16
Global Const LISTA_MULTIANAGRAFICHE_PERSDITTA = 17
Global Const LISTA_RISORSE_TOOLBAR = 18
Global Const LISTA_RISORSE_MSGBOX = 19
Global Const LISTA_RISORSE_ETICHETTE = 20
Global Const LISTA_RISORSE_LINGUETTE = 21
Global Const LISTA_RISORSE_CAPTIONFORM = 22
Global Const LISTA_RISORSE_VARIE = 23
Global Const LISTA_RISORSE_BOTTONI = 24
Global Const LISTA_RISORSE_FOGLI = 25
Global Const LISTA_RISORSE_POPUP = 26
Global Const LISTA_RISORSE_CHECK = 27
Global Const LISTA_RISORSE_OPTION = 28
Global Const LISTA_RISORSE_STATUS = 29
Global Const LISTA_RISORSE_COMBO = 30
Global Const LISTA_RISORSE_ERRORI = 31
Global Const LISTA_RISORSE_TITOLO_VIS = 32
Global Const LISTA_RISORSE_TIPO_VIS = 33
Global Const LISTA_RISORSE_TITOLO_SIT = 34
Global Const LISTA_RISORSE_FILTROVEL = 35
Global Const LISTA_RISORSE_TITOLOTOT = 36
Global Const LISTA_RISORSE_NOMECOL = 37
Global Const LISTA_RISORSE_STAMPE = 38
Global Const LISTA_RISORSE_NOMISTAMPE = 39
Global Const LISTA_RISORSE_TEMA = 40
Global Const LISTA_RISORSE_ITALCOM = 41
Global Const ODBCDAT = 42

Global Const PROCESS_QUERY_INFORMATION = &H400

Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4

Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259

Global Const KEY_ALL_ACCESS = &H3F



'####################################################################################################################
'MMETODO
'####################################################################################################################

Public Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
'Private Declare Function GetClipCursor Lib "user32" (lprc As RECT) As Long

Const MOUSEEVENTF_ABSOLUTE = &H8000   'spostamento assoluto
Const MOUSEEVENTF_LEFTDOWN = &H2    'pulsante sinistro premuto
Const MOUSEEVENTF_LEFTUP = &H4      'pulsante sinistro rilasciato

'=======================
'   tipi enumerativi
'=======================
Public Enum setBottoneAgente
    ageCollegamenti = 0
    ageMostraRiferimenti = 1
    ageNascondiRiferimenti = 2
    ageMostraCampiDB = 3
    ageNascondiCampiDB = 4
    ageDipendenze = 5
End Enum

Public Enum setTipoOpAn
    enmCompila = 0
    enmCarica = 1
End Enum

'=======================
'   costanti
'=======================

'VINCOLI GENERALI
'Global Const RS_MASTRO_CI = "MaCliIta"
'Global Const RS_MASTRO_CE = "MaCliEst"
'Global Const RS_MASTRO_FI = "MaForIta"
'Global Const RS_MASTRO_FE = "MaForEst"
Global Const RS_MASTRO_CLI = 1
Global Const RS_MASTRO_FOR = 2

Global Const GA_CreaDitta = 1
Global Const GA_CreaAnno = 2
Global Const GA_CopiaArchivi = 3
Global Const GA_CancellaAnno = 4
Global Const GA_TRASFSALDI = 5
Global Const GA_TRASFSCAD = 6
Global Const GA_TRASFPART = 7

Global Const SEL_DITTE = 0
Global Const SEL_ANNI = 1

'=============================================
'   dichiarazione costanti
'=============================================
Enum enmTestSalva
    tsnessuno = 0
    tsSalvato = 1
    tsNonSalvato = 2
    tsritorna = 3
End Enum
'=============================================
'=============================================
'   dichiarazione tipi di dati
'=============================================
'Gestione dei tasti della toolBox
'Type Metodo_Form_Attiva 'struttura che contiene informazioni sulla mdichild attiva
'     hwnd As Long   'handle
'     Tool_Mask As Long 'maschera dei tasti della toolbox
'End Type


'=============================================
'   dichiarazione variabili
'=============================================
'Global UltimoErr As Integer
'Global HlpAttivo As Integer
'Global SelAttiva  As Integer
'Global FormAttiva As Metodo_Form_Attiva

Global strinitexe As String
Global commitparziale As Integer
Global GTestMode As Boolean

'Globali per la stampa CRW
Global StpAVideo As Integer
'Global InStampa As Integer

Global hVinCfg As Integer      'handle del DYNASET dei S/Conti Generici Vincolati

Global frmModuli As Form

'*** DESIGNER ***
Global Designer As MXDesigner.cDesigner


'Flag per sapere se l'utente ha selezionato gli Extra Articoli o gli Extra Depositi
'per la Copia Archivi
Dim ExtraArtSel%
Dim ExtraDepSel%
Dim ExtraGiacDepSel%

'##### PER  METODO 2005 #################################################################
Dim MIdxBotAgentiAttuale As Long
Dim MIdxBotDesignerAttuale As Long
Dim MIdxBotZoomAttuale As Long
Dim MIdxBotTemaAttuale As Long

'per vedere se ci sono cambi ditta in corso
Global CmbDittaBusy As Boolean

#If IsMetodo2005 = 1 Then
  Global mMessagingEngine As Object
  Global mMetodoInterop As CMetodoInterop
  Global mMetodoBrowser As Object 'MxBrowser.CBrowserEngine
  Global GTemaAttivo As String
#End If
'########################################################################################

Global GBolNoMsgConfermaUscita As Boolean

'####################################################################################################################
'MMAGAZZINO
'####################################################################################################################


'================================
'   definizione costanti
'================================
Const MAG_DFLT_LenTiplogia = 3 'lunghezza tipologia
Const MAG_DFLT_LenVariante = 8 'lunghezza variante
Const MAG_DFLT_LenArticolo = 50 'lunghezza articolo
Const MAG_DFLT_LenPartita = 15 'lunghezza partita
Const MAG_DFLT_LenUM = 3 'lunghezza unità di misura
Const MAG_DFLT_DecFC = 9 'numero decimali fattore conversione
Global Const MAG_TUTTE_LE_PARTITE = "¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡"
Global Const MAG_TUTTE_LE_UBICAZIONI = "¡¡¡¡¡¡¡¡¡¡¡¡¡¡¡"
'================================
'   definizione tipi enumerativi
'================================
'provenienza dell'articolo
Public Enum setProvenienzaArticolo
    PA_daAcquisto = 0
    PA_daProduzione = 1
    PA_daContoLavoro = 2
End Enum
'arrotondamento lead time
Public Enum setArrotondaLeadTime
    arrLTProporzionale = 0
    arrLTMultiplo = 1
End Enum
'dati articolo
Public Enum setDatiArticolo
    artNessuna = 0
    artAnagrafici = 1
    artCommerciali = 2
    artProduzione = 3
    artInformazioni = 4
    artLIFO = 5
    artExtra = 6
End Enum

'flag sui vincoli per la generazione della partita
Public Enum setFlagGeneraPartita
    partitaRichiedi = 0
    partitaNonInserire = 1
    partitaInserisciDefault = 2
End Enum
'================================
'   definizione variabili
'================================
'####################################################################################################################
'####################################################################################################################

Public Function InizializzaMetodo(Utente As String, Password As String, Ditta As String, Preferiti As String, CTLXKit1 As CTLXKit, CTLXBus1 As CTLXBus, blnMessaggio As Boolean) As Boolean
On Error GoTo Err_InizializzaMetodo
    Dim intRis As Boolean
    Dim strUtente As String
    Dim bolConfInt As Boolean
    Dim strLineErr As String
    Dim strDateLayout As String

    'On Local Error GoTo err_InizializzaMetodo
    InizializzaMetodo = True

    strLineErr = "Creazione Oggetto Nucleo"
    Set MXNU = New MXNucleo.XNucleo

    'Load frmmInit
    '>>> INIZIALIZZAZIONE NUCLEO
    #If TOOLS = 1 Then
        MXNU.FileDatLocali = False
    #End If
    strLineErr = "Inizializzazione Oggetto Nucleo"

    If (MXNU.Inizializza(Preferiti, strUtente)) Then

        If Mid(strDateLayout, 3, 1) <> "/" Then
            'Call MXNU.MsgBoxEX("Impostare la data nel formato dd/MM/yyyy", vbOKOnly + vbCritical, 1007, strDateLayout)
            'GoTo err_objInit
        End If


        MXNU.VersioneMetodo = App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "00")
        MXNU.EXEName = App.EXEName

        strLineErr = "Inizializzazione Oggetto Spread"

'        Call InizializzaSpread(MXNU.MetodoXP)

        #If ISKEY Then
            funzione1
        #End If
            strLineErr = "Creazione Oggetti KIT - BUSINESSS"

            '>>> CREAZIONE OGGETTI KIT - BUSINESS
        #If SOLOKIT = 1 Then
            Call CreateObjKitBus(CTLXKit1, Nothing)
        #Else
            Call CreateObjKitBus(CTLXKit1, CTLXBus1)
        #End If

        strLineErr = "Inizializzazione Libreria Database"

        '>>> INIZIALIZZAZIONE LIBRERIA DATABASE
        If Not (MXDB.dbInizializza(MXNU)) Then
            If blnMessaggio Then Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Libreria ODBC"))
        End If

        strLineErr = "Apertura Ditta"

        '>>> APERTURA DITTA
        #If ISKEY <> 1 Then
            If Not ApriDitta(Utente, Password, Ditta, blnMessaggio) Then
                Call DropObjKitBus
            End If
        #End If

        strLineErr = "Inizializzazione Oggetti KIT - BUSINESS"
        '>>> INIZIALIZZAZIONE OGGETTI KIT - BUSINESS
        If Not InitObjKitBus(hndDBArchivi) Then
            Call DropObjKitBus
        End If

        strLineErr = "Apertura Anno"
        '>>> SELEZIONE ANNO
        #If ISKEY <> 1 Then
            Dim NuovoAnno As Integer
            'If SelezioneAnno( False, NuovoAnno) Then
                NuovoAnno = Year(Now)
                MXNU.AnnoAttivo = NuovoAnno
                Call ApriAnno(False)
            'Else
            '    Call DropObjKitBus
            'End If
        #Else

        #End If
    End If
Err_InizializzaMetodo:
If Err <> 0 Then

    InizializzaMetodo = False
End If
End Function

Public Sub TerminaMetodo()
On Error Resume Next
    DoEvents
    
        'Termino la visione e l'interfaccia dell'anno corrente
'    If (Not mCVisione Is Nothing) Then
'        Call mCVisione.TerminaVisione
'        Set mCVisione = Nothing
'        Call ctlImpFiltro.Termina
'    End If
    
'    Set mCInterfaccia = Nothing
    
    'Set mFunzioniM98 = Nothing
    Set MXNU.FrmMetodo = Nothing
    'Set frmIntro = Nothing
    'MbolInChiusura = True
    Call ChiudiDitta
    Call ChiudiMetodo
    'Set McolFormsInNavBar = Nothing
    'Set metodo = Nothing

End Sub

Private Sub ChiudiMetodo()
    Dim q As Integer

    If Not (hndDBArchivi Is Nothing) Then q = MXDB.dbChiudiDB(hndDBArchivi)
    q = MXDB.dbDisattiva()

    Call DropObjKitBus
End Sub



'####################################################################################################################
'MMAGAZZINO
'####################################################################################################################


'NOME           : LeggiVincoliMagazzino
'DESCRIZIONE    : legge i vincoli e le dimensioni dei campi del magazzino
Sub LeggiVincoliMagazzino()
Dim hSS As CRecordSet
Dim strsql As String
    With MXDB
        'leggo le dimensioni del campo tipologia
        strsql = "SELECT Tipologia FROM TabTipologie WHERE Tipologia=''"
        Set hSS = .dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
        MXNU.MAG_LenTiplogia = .dbGetLenCampo(hSS, TIPO_SNAPSHOT, 0)
        If (MXNU.MAG_LenTiplogia = 0) Then MXNU.MAG_LenTiplogia = MAG_DFLT_LenTiplogia
        Call .dbChiudiSS(hSS)
        'leggo le dimensioni del campo variante
        strsql = "SELECT Variante FROM TabVarianti WHERE Variante=''"
        Set hSS = .dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
        MXNU.MAG_LenVariante = .dbGetLenCampo(hSS, TIPO_SNAPSHOT, 0)
        If (MXNU.MAG_LenVariante = 0) Then MXNU.MAG_LenVariante = MAG_DFLT_LenVariante
        Call .dbChiudiSS(hSS)
        'leggo le dimensioni del campo codice articolo
        strsql = "SELECT Codice FROM AnagraficaArticoli WHERE Codice=''"
        Set hSS = .dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
        MXNU.MAG_LenArticolo = .dbGetLenCampo(hSS, TIPO_SNAPSHOT, 0)
        If (MXNU.MAG_LenArticolo = 0) Then MXNU.MAG_LenArticolo = MAG_DFLT_LenArticolo
        Call .dbChiudiSS(hSS)
        'leggo le dimensioni del campo partita
        strsql = "SELECT CodLotto FROM AnagraficaLotti WHERE CodArticolo=''"
        Set hSS = .dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
        MXNU.MAG_LenPartita = .dbGetLenCampo(hSS, TIPO_SNAPSHOT, 0)
        If (MXNU.MAG_LenPartita = 0) Then MXNU.MAG_LenPartita = MAG_DFLT_LenPartita
        Call .dbChiudiSS(hSS)
        'leggo le dimensioni del campo unità di misura
        strsql = "SELECT Codice FROM TabUnitaMisura WHERE Codice=''"
        Set hSS = .dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
        MXNU.MAG_LenUM = .dbGetLenCampo(hSS, TIPO_SNAPSHOT, 0)
        If (MXNU.MAG_LenUM = 0) Then MXNU.MAG_LenUM = MAG_DFLT_LenUM
        Call .dbChiudiSS(hSS)
        'leggo il numero decimali del campo fattore conversione
        strsql = "SELECT Fattore FROM ArticoliFattoriConversione WHERE CodArt=''"
        Set hSS = .dbCreaSS(hndDBArchivi, strsql)
        Call .dbGetLenCampo(hSS, TIPO_SNAPSHOT, "Fattore", MXNU.MAG_DecFC)
        'ATTENZIONE: mettendo 10 come decimali fattore conversione fdec tronca ad una cifra decimale
        If (MXNU.MAG_DecFC = 0 Or MXNU.MAG_DecFC > 9) Then MXNU.MAG_DecFC = MAG_DFLT_DecFC
        Call .dbChiudiSS(hSS)
    End With
End Sub


'NOME           : ArticoloMovimentato
'DESCRIZIONE    : controlla se ci sono movimenti di storico che fanno riferimento all'articolo
'PARAMETRO 1    : articolo da controllare
Function ArticoloMovimentato(ByVal strCodArt As String) As Boolean
Dim strsql As String
Dim hSS As MXKit.CRecordSet

    With MXDB
        strsql = "SELECT TOP 1 Progressivo" _
            & " FROM StoricoMag" _
            & " WHERE CodArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT)
        Set hSS = .dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
        ArticoloMovimentato = (Not .dbFineTab(hSS, TIPO_SNAPSHOT))
        Call .dbChiudiSS(hSS)
    End With

End Function

'NOME           : ArticoloGeneratore
'DESCRIZIONE    : controlla se per un articolo con tipologie ci sono articoli a varianti generati
'PARAMETRO 1    : articolo con tipologie da controllare
Function ArticoloGeneratore(ByVal strCodArt As String) As Boolean
Dim strsql As String
Dim hSS As MXKit.CRecordSet

    With MXDB
        strsql = "SELECT Codice" _
            & " FROM AnagraficaArticoli" _
            & " WHERE CodicePrimario=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT)
        Set hSS = .dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
        ArticoloGeneratore = (Not .dbFineTab(hSS, TIPO_SNAPSHOT))
        Call .dbChiudiSS(hSS)
    End With

End Function

Function CaricaDatiArticolo(ByVal vntArticolo As Variant, _
                            ByVal strListaCampi As String, _
                            ByVal enmTipoDato As setDatiArticolo, _
                            colValoriRitorno As Collection, _
                            Optional bolLettiDatiPadre As Boolean = False) As Boolean

    If (MXDB.SupportEnhancements) Then
        CaricaDatiArticolo = CaricaDatiArticoloExt(vntArticolo, strListaCampi, enmTipoDato, colValoriRitorno, bolLettiDatiPadre)
    Else
        CaricaDatiArticolo = CaricaDatiArticoloOld(vntArticolo, strListaCampi, enmTipoDato, colValoriRitorno, bolLettiDatiPadre)
    End If

End Function

'NOME           : CaricaDatiArticolo
'DESCRIZIONE    : carica i dati di un articolo e, se non generato, dall'articolo con tipologia
'PARAMETRO 1    : codice articolo
'PARAMETRO 2    : lista campi da leggere
'PARAMETRO 3    : collection campi ritorno
'PARAMETRO 4    : flag carica i dati dell'articolo padre (nel caso di art. varianti non generato) si/no
'RISULTATO      : esito del caricamento
'ATTENZIONE     : la funzione non carica i campi CODICE e DESCRIZIONE
Private Function CaricaDatiArticoloOld(ByVal vntArticolo As Variant, _
                            ByVal strListaCampi As String, _
                            ByVal enmTipoDato As setDatiArticolo, _
                            colValoriRitorno As Collection, _
                            Optional bolLettiDatiPadre As Boolean = False) As Boolean

Dim strsql As String
Dim strFrom As String, strWhr As String
Dim hSS As CRecordSet
Dim intPosSep As Integer
Dim cnt As Integer, intNC As Integer
Dim vetCampi() As String
Dim strCodArt As String

    CaricaDatiArticoloOld = True
    bolLettiDatiPadre = False
    'leggo i campi aggiuntivi
    strCodArt = CStr(vntArticolo)
    If (Len(strCodArt) = 0 Or Len(strListaCampi) = 0) Then
        'rif.A-3707 - NECESSARIO in quanto se non passo il codice o la lista campi utilizzo la collection
        '             dei valori di ritorno che risulterà senza elementi e pertanto l'accesso va in errore
        CaricaDatiArticoloOld = False
    Else
        GoSub componiQuery
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
        If MXDB.dbFineTab(hSS, TIPO_SNAPSHOT) Then
            bolLettiDatiPadre = True
            'cerco su codice padre
            intPosSep = InStr(vntArticolo, MXNU.SepVar)
            If (intPosSep <> 0) Then
                Call MXDB.dbChiudiSS(hSS)
                strCodArt = Left$(vntArticolo, intPosSep - 1)
                GoSub componiQuery
                Set hSS = MXDB.dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
            End If
        End If
        If (MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)) Then
            CaricaDatiArticoloOld = False
        Else
            ReDim vetCampi(0) As String
            intNC = slice(strListaCampi, ",", vetCampi())
            For cnt = 0 To intNC - 1
                If (StrComp(vetCampi(cnt), "codice", vbTextCompare) <> 0) Then
                    On Local Error Resume Next
                    Call colValoriRitorno.Add(MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, vetCampi(cnt), Empty), vetCampi(cnt))
                    On Local Error GoTo 0
                End If
            Next cnt
        End If
    End If

fine_CaricaDatiArticolo:
    Call MXDB.dbChiudiSS(hSS)
    Exit Function

componiQuery:
    Select Case enmTipoDato
        Case artNessuna
            'OTTIMIZZAZIONE: risulta migliore che utilizzare VISTAANAGRAFICAARTICOLI
            strFrom = "{oj ANAGRAFICAARTICOLI ART inner join" _
                    & " ANAGRAFICAARTICOLICOMM COMM on ART.CODICE = COMM.CODICEART inner join" _
                    & " ANAGRAFICAARTICOLIPROD PROD on ART.CODICE = PROD.CODICEART}"
            strWhr = "ART.CODICE=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT) _
                    & " and COMM.ESERCIZIO=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER) _
                    & " and PROD.ESERCIZIO=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER)
        Case artAnagrafici
            strFrom = "AnagraficaArticoli"
            strWhr = "Codice=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT)
        Case artCommerciali
            strFrom = "AnagraficaArticoliComm"
            strWhr = "CodiceArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT) _
                    & " AND Esercizio=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER)
        Case artInformazioni
            strFrom = "DescrArticoli"
            strWhr = "CodiceArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT)
        Case artProduzione
            strFrom = "AnagraficaArticoliProd"
            strWhr = "CodiceArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT) _
                    & " AND Esercizio=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER)
        Case artLIFO
            strFrom = "LifoArticoli"
            strWhr = "CodiceArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT) _
                    & " AND Esercizio=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER)
        Case artExtra
            strFrom = "ExtraMag"
            strWhr = "CodArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT)
    End Select
    strsql = "SELECT " & strListaCampi _
            & " FROM " & strFrom _
            & " WHERE " & strWhr
Return
End Function

'NOME           : CaricaDatiArticolo
'DESCRIZIONE    : carica i dati di un articolo e, se non generato, dall'articolo con tipologia
'PARAMETRO 1    : codice articolo
'PARAMETRO 2    : lista campi da leggere
'PARAMETRO 3    : collection campi ritorno
'PARAMETRO 4    : flag carica i dati dell'articolo padre (nel caso di art. varianti non generato) si/no
'RISULTATO      : esito del caricamento
'ATTENZIONE     : la funzione non carica i campi CODICE e DESCRIZIONE
Private Function CaricaDatiArticoloExt(ByVal vntArticolo As Variant, _
                            ByVal strListaCampi As String, _
                            ByVal enmTipoDato As setDatiArticolo, _
                            colValoriRitorno As Collection, _
                            Optional bolLettiDatiPadre As Boolean = False) As Boolean

Dim strsql As MXKit.StatementFragment
Dim strFrom As MXKit.StatementFragment, strWhr As MXKit.StatementFragment
Dim hSS As CRecordSet
Dim intPosSep As Integer
Dim cnt As Integer, intNC As Integer
Dim vetCampi() As String
Dim strCodArt As String

    CaricaDatiArticoloExt = True
    bolLettiDatiPadre = False
    'leggo i campi aggiuntivi
    strCodArt = CStr(vntArticolo)
    If (Len(strCodArt) = 0 Or Len(strListaCampi) = 0) Then
        'rif.A-3707 - NECESSARIO in quanto se non passo il codice o la lista campi utilizzo la collection
        '             dei valori di ritorno che risulterà senza elementi e pertanto l'accesso va in errore
        CaricaDatiArticoloExt = False
    Else
        GoSub componiQuery
        Set hSS = MXDB.dbCreaSSEx(hndDBArchivi, strsql, TIPO_TABELLA)
        If MXDB.dbFineTab(hSS, TIPO_SNAPSHOT) Then
            bolLettiDatiPadre = True
            'cerco su codice padre
            intPosSep = InStr(vntArticolo, MXNU.SepVar)
            If (intPosSep <> 0) Then
                Call MXDB.dbChiudiSS(hSS)
                strCodArt = Left$(vntArticolo, intPosSep - 1)
                GoSub componiQuery
                Set hSS = MXDB.dbCreaSSEx(hndDBArchivi, strsql, TIPO_TABELLA)
            End If
        End If
        If (MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)) Then
            CaricaDatiArticoloExt = False
        Else
            ReDim vetCampi(0) As String
            intNC = slice(strListaCampi, ",", vetCampi())
            For cnt = 0 To intNC - 1
                If (StrComp(vetCampi(cnt), "codice", vbTextCompare) <> 0) Then
                    On Local Error Resume Next
                    Call colValoriRitorno.Add(MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, vetCampi(cnt), Empty), vetCampi(cnt))
                    On Local Error GoTo 0
                End If
            Next cnt
        End If
    End If

fine_CaricaDatiArticolo:
    Set strsql = Nothing
    Call MXDB.dbChiudiSS(hSS)
    Exit Function

componiQuery:
    Set strFrom = New StatementFragment
    Set strWhr = New StatementFragment
    Select Case enmTipoDato
        Case artNessuna
            'OTTIMIZZAZIONE: risulta migliore che utilizzare VISTAANAGRAFICAARTICOLI
            strFrom.Statement = "{oj ANAGRAFICAARTICOLI ART inner join" _
                    & " ANAGRAFICAARTICOLICOMM COMM on ART.CODICE = COMM.CODICEART inner join" _
                    & " ANAGRAFICAARTICOLIPROD PROD on ART.CODICE = PROD.CODICEART}"
            strWhr.Statement = "ART.CODICE=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT, strWhr, "CODICE", adVarChar, MXDB.dbGetLength(Articolo)) _
                    & " and COMM.ESERCIZIO=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER, strWhr, "ESERCIZIO1", adDecimal, 5, adParamInput, 5, 0) _
                    & " and PROD.ESERCIZIO=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER, strWhr, "ESERCIZIO2", adDecimal, 5, adParamInput, 5, 0)
        Case artAnagrafici
            strFrom.Statement = "AnagraficaArticoli"
            strWhr.Statement = "Codice=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT, strWhr, "CODICE", adVarChar, MXDB.dbGetLength(Articolo))
        Case artCommerciali
            strFrom.Statement = "AnagraficaArticoliComm"
            strWhr.Statement = "CodiceArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT, strWhr, "CODICE", adVarChar, MXDB.dbGetLength(Articolo)) _
                    & " AND Esercizio=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER, strWhr, "ESERCIZIO1", adDecimal, 5, adParamInput, 5, 0)
        Case artInformazioni
            strFrom.Statement = "DescrArticoli"
            strWhr.Statement = "CodiceArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT, strWhr, "CODICE", adVarChar, MXDB.dbGetLength(Articolo))
        Case artProduzione
            strFrom.Statement = "AnagraficaArticoliProd"
            strWhr.Statement = "CodiceArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT, strWhr, "CODICE", adVarChar, MXDB.dbGetLength(Articolo)) _
                    & " AND Esercizio=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER, strWhr, "ESERCIZIO1", adDecimal, 5, adParamInput, 5, 0)
        Case artLIFO
            strFrom.Statement = "LifoArticoli"
            strWhr.Statement = "CodiceArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT, strWhr, "CODICE", adVarChar, MXDB.dbGetLength(Articolo)) _
                    & " AND Esercizio=" & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_INTEGER, strWhr, "ESERCIZIO1", adDecimal, 5, adParamInput, 5, 0)
        Case artExtra
            strFrom.Statement = "ExtraMag"
            strWhr.Statement = "CodArt=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT, strWhr, "CODICE", adVarChar, MXDB.dbGetLength(Articolo))
    End Select
    Set strsql = Nothing
    Set strsql = New MXKit.StatementFragment
    strsql.AppendFragments "SELECT " & strListaCampi & " FROM ", strFrom, " WHERE ", strWhr

    Set strFrom = Nothing
    Set strWhr = Nothing
Return
End Function

Function GeneraArticoloVarianti(ByVal strCodArt As String, _
    Optional strDscArt As String = "", _
    Optional strVarEspl As String = "", _
    Optional bolAggMag As Boolean = False, _
    Optional bolCopiaExtra As Boolean = True) As Boolean

Dim intq As Integer
Dim xCArt As MXBusiness.CVArt
Dim strsql As String
Dim hSS As MXKit.CRecordSet

    GeneraArticoloVarianti = True
    bolAggMag = False
    'inizializzo le classi
    strsql = "SELECT AggiornaMag" _
        & " FROM AnagraficaArticoli" _
        & " WHERE Codice=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT)
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
    If (MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)) Then
        Call MXDB.dbChiudiSS(hSS)
        'il codice non esiste -> lo genero
        Set xCArt = MXART.CreaCVArt()
        With xCArt
            .Codice = strCodArt
            If (Not .Valida(CHIEDIVAR_NESSUNA, False, , 0, False)) Then
                GeneraArticoloVarianti = False
                GoTo fine_GeneraArticoloVarianti
            Else
                GeneraArticoloVarianti = .Genera(bolCopiaExtra)
                'rileggo il flag aggiorna magazzino
                Call MXDB.dbChiudiSS(hSS)
                strsql = "SELECT AggiornaMag" _
                    & " FROM AnagraficaArticoli" _
                    & " WHERE Codice=" & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT)
                Set hSS = MXDB.dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
            End If
        End With
    End If
    'restituisco il flag aggiorna magazzino
    bolAggMag = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "AggiornaMag", True)

fine_GeneraArticoloVarianti:
    On Local Error GoTo 0
    GoSub disalloca_GeneraArticoloVarianti
Exit Function

disalloca_GeneraArticoloVarianti:
    'disalloco variabili
    If Not (xCArt Is Nothing) Then
        Call xCArt.Termina
    End If
    Set xCArt = Nothing
    If Not (hSS Is Nothing) Then Call MXDB.dbChiudiDY(hSS)
Return

err_GeneraArticoloVarianti:
    GeneraArticoloVarianti = False
    Call MXNU.MsgBoxEX(1866, vbCritical, 1007, Array(Err.Number, Err.Description, strCodArt))
    Resume fine_GeneraArticoloVarianti:

End Function

'************************************************************************
'NOME           : ArticoloCancellabile
'DESCRIZIONE    : controlla se un articolo è o meno cancellabile
'PARAMETRO 1    : codice articolo
'PARAMETRO 2    : flag articolo tipologia
'************************************************************************
Function ArticoloCancellabile(strCodArt As String, bolArtTipologia As Boolean) As Boolean
Dim strsql As String
Dim hSS As CRecordSet
Dim strMsg As String

    ArticoloCancellabile = True
    strMsg = ""
    If bolArtTipologia Then
        'controllo se ci sono articoli generati
        strsql = "SELECT Codice" _
                & " FROM AnagraficaArticoli" _
                & " WHERE (CodicePrimario='" & strCodArt & "')"
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
        If (Not MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)) Then
            strMsg = MXNU.CaricaStringaRes(1852, Array("", strCodArt))
            GoTo err_ArticoloCancellabile
        End If
        Call MXDB.dbChiudiSS(hSS)
    Else
        'controllo movimenti di magazzino
        strsql = "SELECT TOP 1 Progressivo" _
                & " FROM StoricoMag" _
                & " WHERE CodArt='" & strCodArt & "'"
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
        If (Not MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)) Then
            strMsg = MXNU.CaricaStringaRes(1853, Array("", strCodArt))
            Call MXDB.dbChiudiSS(hSS)
            GoTo err_ArticoloCancellabile
        End If
        Call MXDB.dbChiudiSS(hSS)
    End If
    'controllo esistenza distinta
'    If (InStr(strCodArt, "#") > 0 Or bolArtTipologia) Then
'        strSQL = "SELECT Progressivo" _
'                & " FROM DistintaArtComposti" _
'                & " WHERE (ArtComposto = '" & strCodArt & "')"
'    Else
'        strSQL = "SELECT Progressivo" _
'                & " FROM DistintaArtComposti" _
'                & " WHERE (ArtComposto = '" & strCodArt & "') OR (ArtComposto = '" & strCodArt & "#')"
'    End If
'    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strSQL, TIPO_TABELLA)
'    If (Not MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)) Then
'        strMsg = MXNU.CaricaStringaRes(1854, Array("", strCodArt))
'        GoTo err_ArticoloCancellabile
'    End If
'    Call MXDB.dbChiudiSS(hSS)

fine_ArticoloCancellabile:
    Exit Function

err_ArticoloCancellabile:
    ArticoloCancellabile = False
    Call MXNU.MsgBoxEX(strMsg, vbExclamation, 1007)
    GoTo fine_ArticoloCancellabile

End Function

'NOME           : LeggiMagPrincipale
'DESCRIZIONE    : legge il magazzino principale
'PARAMETRO 1    : (ritorno) codice magazzino principale
'PARAMETRO 2    : (ritorno) descrizione magazzino principale
'RITORNO        : True se il magazzino principale esiste, False altrimenti
Function LeggiMagPrincipale(strCodMagP As String, strDscMagP As String) As Boolean
Dim strsql As String
Dim hSS As CRecordSet

    strsql = "SELECT Codice,Descrizione" _
            & " FROM AnagraficaDepositi" _
            & " WHERE Principale <> 0"
    Set hSS = MXDB.dbCreaDY(hndDBArchivi, strsql, TIPO_TABELLA)
    LeggiMagPrincipale = Not MXDB.dbFineTab(hSS, TIPO_DYNASET)
    strCodMagP = MXDB.dbGetCampo(hSS, TIPO_DYNASET, "Codice", "")
    strDscMagP = MXDB.dbGetCampo(hSS, TIPO_DYNASET, "Descrizione", "")
    Call MXDB.dbChiudiDY(hSS)
End Function

'NOME           : CaricaComboRaggruppaProd
'DESCRIZIONE    : carica il combo dei raggruppamento di produzione e/o assegna il valore a tale combo
'PARAMETRO 1    : oggetto combo box da caricare
'PARAMETRO 2    : codice articolo
'PARAMETR0 3    : true per caricare i dati del combo; false per assegnare solo il valore
'PARAMETRO 4    : valore da assegnare al combo
Sub CaricaComboRaggruppaProd(ByVal objCombo As ComboBox, _
                                ByVal strCodArt As String, _
                                bolCarica As Boolean, _
                                strValSalva As String, _
                                Optional vntValCombo As Variant)
Dim bolEnd As Boolean
Dim intAus As Integer
Dim strsql As String
Dim hSS As CRecordSet

    If (bolCarica) Then
        'carico i valori del combo
        intAus = InStr(strCodArt, MXNU.SepVar)
        If (intAus = 0) Then intAus = Len(strCodArt) + 1
        strsql = "SELECT CodTipologia,NumeroTip" _
                & " FROM TipologieArticoli" _
                & " WHERE CodiceArt=" & hndDBArchivi.FormatoSQL(Left$(strCodArt, intAus - 1), DB_TEXT) _
                & " ORDER BY NumeroTip"
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
        Call objCombo.Clear
        Call objCombo.AddItem("")
        Call objCombo.AddItem(MXNU.CaricaStringaRes(75058))
        strValSalva = " R"
        bolEnd = MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)
        Do While (Not bolEnd)
            Call objCombo.AddItem(MXNU.CaricaStringaRes(75059) & " " & MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "CodTipologia", ""))
            strValSalva = strValSalva & CStr(MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "NumeroTip", 0))
            bolEnd = (Not MXDB.dbSuccessivo(hSS, TIPO_SNAPSHOT))
        Loop
        Call MXDB.dbChiudiSS(hSS)
    End If
    'assegna il valore al combo
    If Not IsMissing(vntValCombo) Then
        If (Trim$(vntValCombo) = " ") Then
            intAus = 0
        ElseIf (vntValCombo = "R") Then
            intAus = 1
        Else
            intAus = 2 + (Asc(vntValCombo) - 49)
        End If
        If (intAus < 0) Then
            intAus = 0
        ElseIf (intAus > objCombo.ListCount) Then
            intAus = objCombo.ListCount
        End If
        If (objCombo.ListCount > 0) Then objCombo.ListIndex = intAus
    End If
End Sub

'NOME           : MovimentaArticolo
'DESCRIZIONE    : restituisce il flag di movimentazione dell'articolo
'PARAMETRO 1    : codice articolo
Function MovimentaArticolo(ByVal strCodArt As String) As Boolean
Dim strsql As String
Dim hSS As CRecordSet

    strsql = "SELECT AggiornaMag" _
            & " FROM AnagraficaArticoli" _
            & " WHERE Codice = " & hndDBArchivi.FormatoSQL(strCodArt, DB_TEXT)
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
    MovimentaArticolo = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "AggiornaMag", True)
    Call MXDB.dbChiudiSS(hSS)
End Function

Public Function LeggiListinoArticolo(ByVal strArt As String, ByVal lngListino, Prezzo As Variant, PrezzoEuro As Variant)

    Dim hSS As MXKit.CRecordSet, q

    With MXDB
        Set hSS = .dbCreaSS(hndDBArchivi, "SELECT CODART, NRLISTINO,PREZZO,PREZZOEURO FROM LISTINIARTICOLI WHERE CODART=" & hndDBArchivi.FormatoSQL(strArt, DB_TEXT) & " AND NRLISTINO =" & lngListino)
        If .dbFineTab(hSS) Then
            Prezzo = 0
            PrezzoEuro = 0
            LeggiListinoArticolo = False
        Else
            Prezzo = .dbGetCampo(hSS, TIPO_SNAPSHOT, "PREZZO", 0)
            PrezzoEuro = .dbGetCampo(hSS, TIPO_SNAPSHOT, "PREZZOEURO", 0)
            LeggiListinoArticolo = True
        End If
        q = .dbChiudiSS(hSS)
    End With

End Function

Function LeggiContropartitaArticolo(ByVal CodArt As String, ByVal NrControPCont As Long, ByVal TipoConto As String, ByVal Nazione As Long) As String

    Dim Found As Boolean
    Dim hSS As MXKit.CRecordSet
    Dim coll As Collection
    Dim res As String

    Found = False
    res = ""
    If NrControPCont <> 0 Then
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT CodArt,Numero,SCGen FROM ControPartArticoli WHERE CodArt=" & _
            hndDBArchivi.FormatoSQL(CodArt, DB_TEXT) & " AND Esercizio = " & MXNU.AnnoAttivo & " AND Numero=" & NrControPCont)

        Found = Not MXDB.dbFineTab(hSS)
        If Found Then
            res = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "SCGen", "")
        End If
        Call MXDB.dbChiudiSS(hSS)
    End If
    If Not Found Then
        Set coll = New Collection
        If CaricaDatiArticolo(CodArt, "SCGenVenditeIta,SCGenVenditeEst,SCGenAcquistiIta,SCGenAcquistiEst", artCommerciali, coll) Then
            If TipoConto = "C" Then
                If Nazione = 0 Then
                    res = coll("SCGenVenditeIta")
                Else
                    res = coll("SCGenVenditeEst")
                End If
            ElseIf TipoConto = "F" Then
                If Nazione = 0 Then
                    res = coll("SCGenAcquistiIta")
                Else
                    res = coll("SCGenAcquistiEst")
                End If
            Else
                res = coll("SCGenVenditeIta")
            End If
        End If
        Set coll = Nothing
    End If
    LeggiContropartitaArticolo = res
End Function

Sub ScomponiCodiceArticolo(ByVal strArticolo As String, _
    Optional strCodiceNeutro As String, _
    Optional intPosSeparatore As Integer, _
    Optional strVarianti As String, _
    Optional bolAVarianti As Boolean)

    intPosSeparatore = InStr(strArticolo, MXNU.SepVar)
    bolAVarianti = (intPosSeparatore <> 0)
    If (bolAVarianti) Then
        strCodiceNeutro = Left$(strArticolo, intPosSeparatore - 1)
        strVarianti = Mid$(strArticolo, intPosSeparatore + 1)
    Else
        strCodiceNeutro = strArticolo
        strVarianti = ""
    End If
End Sub

Function VincolaUM(varListino As Variant, Optional bolListinoTrasformazione As Boolean = False) As Boolean
    Dim hSS As MXKit.CRecordSet
    Dim intq As Integer
    Dim strsql As String

    If bolListinoTrasformazione Then
        strsql = "SELECT VincolaUM FROM TabListiniTrasformazione WHERE NrListino=" & varListino
    Else
        strsql = "SELECT VincolaUM FROM TabListini WHERE NrListino=" & varListino
    End If

    VincolaUM = False
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strsql)
    If MXDB.dbGetCampo(hSS, NO_REPOSITION, "VincolaUM", 0) = 1 Then
        VincolaUM = True
    End If
    intq = MXDB.dbChiudiSS(hSS)

End Function

Public Function LeggiVariantiArticolo(ByVal strCodArt As String, colPar As Collection) As Boolean
Dim cArt As MXBusiness.CVArt
Dim vntTipVar As Variant
Dim strVar As String
Dim inti As Integer

    Set cArt = MXART.CreaCVArt()
    LeggiVariantiArticolo = cArt.Valida(CHIEDIVAR_NESSUNA, False, strCodArt)
    If LeggiVariantiArticolo Then
        LeggiVariantiArticolo = (Len(cArt.VariantiEsplicite) > 0)
        If LeggiVariantiArticolo Then
            For Each vntTipVar In Split(Left$(cArt.VariantiEsplicite, Len(cArt.VariantiEsplicite) - 1), ";")
                strVar = Split(vntTipVar, "=")(1)
                inti = inti + 1
                colPar.Add strVar, CStr(inti)
            Next vntTipVar
        End If
    End If
    Call cArt.Termina
    Set cArt = Nothing

End Function

'------------------------------------------------------------
'nome:          Data2Esercizio
'descrizione:   restituisce l'esercizio di pertinenza della data passata
'parametri:     (in) data da controllare
'               (out) esercizio
'ritorno:       esito dell'operazione; se false la data è fuori
'               da tutti gli esercizi attualmente inseriti in tabella e viene restituito 0
'annotazioni:   rif.A-4767
'------------------------------------------------------------
Public Function i_Data2Esercizio(ByVal dteData As Variant, intEsercizio As Integer) As Boolean
Dim bolRes As Boolean
Dim strQuery As String
Dim hRSData As MXKit.CRecordSet

    bolRes = True
    On Local Error GoTo Data2Esercizio_ERR
    intEsercizio = 0
    With MXDB
        strQuery = "select CODICE" _
            & " from TABESERCIZI" _
            & " where DATAINIMAG<=" & hndDBArchivi.FormatoSQL(dteData, DB_DATE) _
            & " and DATAFINEMAG>=" & hndDBArchivi.FormatoSQL(dteData, DB_DATE)
        Set hRSData = .dbCreaSS(hndDBArchivi, strQuery)
        bolRes = Not .dbFineTab(hRSData)
        If (bolRes) Then
            intEsercizio = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "CODICE", 0)
            bolRes = (intEsercizio <> 0)
        End If
    End With

Data2Esercizio_END:
    Call MXDB.dbChiudiSS(hRSData)
    i_Data2Esercizio = bolRes
    On Local Error GoTo 0
    Exit Function

Data2Esercizio_ERR:
Dim lngErrCod As Long
Dim strErrDsc As String
    bolRes = False
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("Data2Esercizio", lngErrCod, strErrDsc))
    Resume Data2Esercizio_END
End Function


'------------------------------------------------------------
'nome:          LeggiDatiRiordino
'descrizione:   lettura dei dati di riordino di un dato articolo
'parametri:     Articolo: (in) codice articolo da considerare
'               Fornitore: (in/out) se passato cerca il fornitore fra quelli preferenziale e alternativi
'               Provenienza: (out) restituisce la provenienza dell'articolo
'               GGApprontamento: (out) restituisce i giorni di approntamento dell'articolo
'               GGApprovvigionamento: (out) restituisce i giornii di approvvigionamento dell'articolo
'               LottoRiferimento: (out) restituisce il lotto di riferimento per il tempo di approvvigionamento dell'articolo
'               TipoArrotondamento: (out) restituisce la modalità di arrotondamento del tempo di approvvigionamento rispetto al lotto
'ritorno:       esito dell'operazione
'annotazioni:   rif.A#5292
'------------------------------------------------------------
Public Function LeggiDatiRiordino(ByVal Articolo As String, ByRef fornitore As String, _
    Optional ByRef Provenienza As setProvenienzaArticolo, _
    Optional ByRef GGApprontamento As Long, _
    Optional ByRef GGApprovvigionamento As Long, _
    Optional ByRef LottoRiferimento As Variant, _
    Optional ByRef UmLottoRif As String, _
    Optional ByRef TipoArrotondamento As setArrotondaLeadTime) As Boolean

Dim bolRes As Boolean
Dim strQuery As String
Dim hRSData As MXKit.CRecordSet
Dim strFornitoreIn As String
Dim strSuffisso As String
Dim sCodicePadre As String
Dim sVarianti As String
Dim bArtVarianti As Boolean
Const SEGNAPOSTO_ARTICOLO = "%ARTICOLO%"

    bolRes = True
    On Local Error GoTo LeggiDatiRiordino_ERR
    strFornitoreIn = fornitore

    If (Len(Articolo) = 0) Then
        fornitore = ""
        GGApprontamento = 0
        GGApprovvigionamento = 0
        LottoRiferimento = CDec(0)
        UmLottoRif = ""
        TipoArrotondamento = arrLTProporzionale
    Else
        'RIF.A#6559 - determino il codice padre
        Call SeparaVarianti_i(Articolo, sCodicePadre, sVarianti)
        bArtVarianti = (Articolo <> sCodicePadre)

        With MXDB
            'determino i dati dell'articolo
            strQuery = "select PROVENIENZA," _
                & "FORNPREFACQ,TAPPRONTACQ,TAPPROVVACQ,LOTTORIFACQ,UMLOTTOACQ,ARROTLOTTOACQ," _
                & "(select top 1 CODFOR from TABVINCOLIPRODUZIONE) as FORNPREFPROD,TAPPRONTPROD,TAPPROVVPROD,LOTTORIFPROD,UMLOTTOPROD,ARROTLOTTOPROD," _
                & "FORNPREFLAV,TAPPRONTLAV,TAPPROVVLAV,LOTTORIFLAV,UMLOTTOLAV,ARROTLOTTOLAV" _
                & " from ANAGRAFICAARTICOLIPROD" _
                & " where CODICEART=" & SEGNAPOSTO_ARTICOLO & " and ESERCIZIO=" & MXNU.AnnoAttivo
            Set hRSData = .dbCreaSS(hndDBArchivi, Replace(strQuery, SEGNAPOSTO_ARTICOLO, hndDBArchivi.FormatoSQL(Articolo, DB_TEXT)))    'RIF.A#6559 - prima lettura: codice articolo
            'RIF.A#6559 - lettura dati dal padre se codice non generato
            If (.dbFineTab(hRSData) And bArtVarianti) Then
                Call .dbChiudiSS(hRSData)
                Set hRSData = .dbCreaSS(hndDBArchivi, Replace(strQuery, SEGNAPOSTO_ARTICOLO, hndDBArchivi.FormatoSQL(sCodicePadre, DB_TEXT)))
            End If
            If (.dbFineTab(hRSData)) Then
                bolRes = False
                GoTo LeggiDatiRiordino_END
            Else
                Provenienza = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "PROVENIENZA", PA_daAcquisto)
                Select Case Provenienza
                    Case PA_daAcquisto: strSuffisso = "ACQ"
                    Case PA_daProduzione: strSuffisso = "PROD"
                    Case PA_daContoLavoro: strSuffisso = "LAV"
                End Select
                fornitore = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "FORNPREF" & strSuffisso, "")
                bolRes = (Len(fornitore) > 0)
                'confronto con il fornitore passato
                If (bolRes And Len(strFornitoreIn) > 0) Then
                    bolRes = (fornitore = strFornitoreIn)
                End If
                'se trovato il fornitore principale => leggo i dati
                'RIF.A#9901 - i dati di riordino generali devono essere letti indipenentemente dalla presenza del fornitore preferenziale
                'If (bolRes) Then
                    GGApprontamento = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "TAPPRONT" & strSuffisso, 0)
                    GGApprovvigionamento = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "TAPPROVV" & strSuffisso, 0)
                    LottoRiferimento = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "LOTTORIF" & strSuffisso, 0)
                    UmLottoRif = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "UMLOTTO" & strSuffisso, 0)
                    TipoArrotondamento = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "ARROTLOTTO" & strSuffisso, 0)
                'End If
            End If
            Call .dbChiudiSS(hRSData)
            'se non trovato il fornitore principale => leggo i dati dei fornitori alternativi
            If (Not bolRes) Then
                strQuery = "select top 1 CODFOR,GGAPPRONT,GGAPPROVV,LOTTORIF,UM,ARROTLOTTO" _
                    & " from TABLOTTIRIORDINO" _
                    & " where CODART=" & SEGNAPOSTO_ARTICOLO _
                    & " and TIPORIORD=" & Provenienza
                If (Len(strFornitoreIn) > 0) Then
                    strQuery = strQuery & " and CODFOR=" & hndDBArchivi.FormatoSQL(strFornitoreIn, DB_TEXT)
                End If
                'RIF.A#8830 - l'ordinamento va fatto per percentuale di ripartizione e, a parità di percentuale per posizione
                strQuery = strQuery & " order by PRCRIPART desc, NUMERO asc"

                Set hRSData = .dbCreaSS(hndDBArchivi, Replace(strQuery, SEGNAPOSTO_ARTICOLO, hndDBArchivi.FormatoSQL(Articolo, DB_TEXT))) 'RIF.A#6559 - prima lettura: codice articolo
                'RIF.A#6559 - lettura dati dal padre se codice non generato
                If (.dbFineTab(hRSData) And bArtVarianti) Then
                    Call .dbChiudiSS(hRSData)
                    Set hRSData = .dbCreaSS(hndDBArchivi, Replace(strQuery, SEGNAPOSTO_ARTICOLO, hndDBArchivi.FormatoSQL(sCodicePadre, DB_TEXT)))
                End If
                If (.dbFineTab(hRSData)) Then
                    bolRes = False
                    GoTo LeggiDatiRiordino_END
                Else
                    fornitore = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "CODFOR", 0)
                    GGApprontamento = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "GGAPPRONT", 0)
                    GGApprovvigionamento = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "GGAPPROVV", 0)
                    LottoRiferimento = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "LOTTORIF", 0)
                    UmLottoRif = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "UM", 0)
                    TipoArrotondamento = .dbGetCampo(hRSData, TIPO_SNAPSHOT, "ARROTLOTTO", 0)
                End If
                Call .dbChiudiSS(hRSData)
            End If
        End With
    End If

LeggiDatiRiordino_END:
    Call MXDB.dbChiudiSS(hRSData)
    LeggiDatiRiordino = bolRes
    On Local Error GoTo 0
    Exit Function

LeggiDatiRiordino_ERR:
Dim lngErrCod As Long
Dim strErrDsc As String
    bolRes = False
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("LeggiDatiRiordino", lngErrCod, strErrDsc))
    Resume LeggiDatiRiordino_END
End Function

'------------------------------------------------------------
'nome:          i_GeneraPartita
'descrizione:   genera una partita se non presente
'parametri:     codice articolo
'               codice partita
'ritorno:       esito dell'operazione
'annotazioni:
'------------------------------------------------------------
Public Function i_GeneraPartita(ByVal strArticolo As String, ByVal strPartita As String) As Boolean
Dim bolRes As Boolean
Dim hRSPart As MXKit.CRecordSet
Dim hRSCar As MXKit.CRecordSet
Dim strsql As String
Dim lngNrRiga As Long
Dim vntValue As Variant

    bolRes = True
    On Local Error GoTo GeneraPartita_ERR
    If ((Len(strArticolo) > 0) And (Len(strPartita) > 0)) Then
        With MXDB
            'controlla se la partita è già stata generata
            strsql = "select CODLOTTO" _
                & " from ANAGRAFICALOTTI" _
                & " where CODARTICOLO=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT) _
                & " and CODLOTTO=" & hndDBArchivi.FormatoSQL(strPartita, DB_TEXT)
            Set hRSPart = .dbCreaSS(hndDBArchivi, strsql)
            If (.dbFineTab(hRSPart)) Then
                'se la partita non c'è la genero
                strsql = "insert into ANAGRAFICALOTTI (CODARTICOLO,CODLOTTO,BLOCCATO,UTENTEMODIFICA,DATAMODIFICA)" _
                    & " values (" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT) & "," & hndDBArchivi.FormatoSQL(strPartita, DB_TEXT) & ",0," & hndDBArchivi.FormatoSQL(MXNU.UtenteAttivo, DB_TEXT) & "," & hndDBArchivi.FormatoSQL(Now, DB_DATETIME) & ")"
                Call .dbEseguiSQL(hndDBArchivi, strsql)
    '            Call .dbInserisci(hDYPart)
    '            Call .dbSetCampo(hDYPart, TIPO_DYNASET, "CODARTICOLO", strArticolo)
    '            Call .dbSetCampo(hDYPart, TIPO_DYNASET, "CODLOTTO", strPartita)
    '            Call .dbSetCampo(hDYPart, TIPO_DYNASET, "BLOCCATO", CStr(vbUnchecked))
    '            Call .dbRegistra(hDYPart)
                'genero le caratteristiche in base al default
                strsql = "select NRRIGA,CARATTDEFAULT" _
                    & " from TABCARATTLOTTI" _
                    & " where CODICE=(select CATEGORIA from ANAGRAFICAARTICOLI where CODICE=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT) & ")"
                Set hRSCar = .dbCreaSS(hndDBArchivi, strsql)
                Do While Not (.dbFineTab(hRSCar))
                    lngNrRiga = .dbGetCampo(hRSCar, TIPO_SNAPSHOT, "NRRIGA", 0)
                    vntValue = .dbGetCampo(hRSCar, TIPO_SNAPSHOT, "CARATTDEFAULT", "")
                    strsql = "insert into ANAGRCARLOTTI (CODARTICOLO,CODLOTTO,NRRIGA,VALORE,UTENTEMODIFICA,DATAMODIFICA)" _
                        & " values (" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT) & "," & hndDBArchivi.FormatoSQL(strPartita, DB_TEXT) _
                        & "," & lngNrRiga & "," & hndDBArchivi.FormatoSQL(vntValue, DB_TEXT) & "," & hndDBArchivi.FormatoSQL(MXNU.UtenteAttivo, DB_TEXT) & "," & hndDBArchivi.FormatoSQL(Now, DB_DATETIME) & ")"
                    Call .dbEseguiSQL(hndDBArchivi, strsql)

                    Call .dbSuccessivo(hRSCar)
                Loop
                Call .dbChiudiSS(hRSPart)
            End If
        End With
    End If

GeneraPartita_END:
    Call MXDB.dbChiudiSS(hRSCar)
    Call MXDB.dbChiudiDY(hRSPart)
    i_GeneraPartita = bolRes
    On Local Error GoTo 0
    Exit Function

GeneraPartita_ERR:
Dim lngErrCod As Long
Dim strErrDsc As String
    bolRes = False
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("GeneraPartita", lngErrCod, strErrDsc))
    Resume GeneraPartita_END
Resume
End Function

'------------------------------------------------------------
'nome:          LeggiFlagGeneraPartita
'descrizione:   legge per l'anno attivo il valore del flagPartita dalla tabella TABVINCOLIGIC ovvero l'option button Partita dalla form Vincoli
'parametri:
'ritorno:       valore del flag per l'esercizio attivo
'annotazioni:   RIF.A#5948
'------------------------------------------------------------
Public Function LeggiFlagGeneraPartita() As setFlagGeneraPartita
    Dim strQuery As String
    Dim hRS As MXKit.CRecordSet

    On Local Error GoTo ERR_LeggiFlagGeneraPartita

    strQuery = "select FLGPARTITA" _
        & " from TABVINCOLIGIC" _
        & " where ESERCIZIO = " & MXNU.AnnoAttivo
    With MXDB
        Set hRS = .dbCreaSS(hndDBArchivi, strQuery)
        LeggiFlagGeneraPartita = .dbGetCampo(hRS, TIPO_SNAPSHOT, "FLGPARTITA", 0)
    End With

END_LeggiFlagGeneraPartita:
    Call MXDB.dbChiudiSS(hRS)
    Set hRS = Nothing
    On Local Error GoTo 0
    Exit Function

ERR_LeggiFlagGeneraPartita:
    Dim lngErrCod As Long
    Dim strErrDsc As String
    lngErrCod = Err.Number
    strErrDsc = Err.Description
    On Local Error GoTo 0
    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("LeggiFlagGeneraPartita", lngErrCod, strErrDsc))
    Resume END_LeggiFlagGeneraPartita

End Function

Public Function SeparaVarianti_i(ByVal strCod As String, strArtTip As String, strVar As String) As Boolean

    Dim psep As Integer

    psep = InStr(strCod, MXNU.SepVar)
    SeparaVarianti_i = psep > 0
    If psep > 0 Then
        strArtTip = Left$(strCod, psep - 1)
        strVar = Mid$(strCod, psep + 1)
    Else
        strArtTip = ""
        strVar = ""
    End If
End Function

'RIF.A#6234 - restituisce un valore booleano che indica se l'articolo movimenta o meno le matricole
Public Function ArticoloMovimentaMatricole(ByVal strArticolo As String) As Boolean
Dim strQuery As String
Dim hRSData As MXKit.CRecordSet

    With MXDB
        strQuery = "select MOVIMENTAMATRICOLE from ANAGRAFICAARTICOLI where CODICE=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
        Set hRSData = .dbCreaSS(hndDBArchivi, strQuery)
        ArticoloMovimentaMatricole = (.dbGetCampo(hRSData, TIPO_SNAPSHOT, "MOVIMENTAMATRICOLE", 0) <> 0)
        Call .dbChiudiSS(hRSData)
    End With
End Function

Public Function ArticoloFloorStock(ByVal strArticolo As String) As Boolean
Dim strQuery As String
Dim hRSData As MXKit.CRecordSet
Dim strCodicePadre As String

    With MXDB
        strQuery = "select FLOORSTOCK" _
            & " from ANAGRAFICAARTICOLIPROD" _
            & " where CODICEART=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT) _
            & " and ESERCIZIO=" & MXNU.AnnoAttivo
        Set hRSData = .dbCreaSS(hndDBArchivi, strQuery)
        If (.dbFineTab(hRSData)) Then
            'se articolo a varianti non generato => leggo il dato dall'articolo padre
            Call .dbChiudiSS(hRSData)
            Call ScomponiCodiceArticolo(strArticolo, strCodicePadre)
            If (strCodicePadre <> strArticolo) Then
                strQuery = "select FLOORSTOCK" _
                    & " from ANAGRAFICAARTICOLIPROD" _
                    & " where CODICEART=" & hndDBArchivi.FormatoSQL(strCodicePadre, DB_TEXT) _
                    & " and ESERCIZIO=" & MXNU.AnnoAttivo
                Set hRSData = .dbCreaSS(hndDBArchivi, strQuery)
            End If
        End If
        'leggo e restituisco il risultato
        If (.dbFineTab(hRSData)) Then
            ArticoloFloorStock = False
        Else
            ArticoloFloorStock = (.dbGetCampo(hRSData, TIPO_SNAPSHOT, "FLOORSTOCK", 0) <> 0)
        End If

        Call .dbChiudiSS(hRSData)
    End With
End Function

''####################################################################################################################
''MMETODO
''####################################################################################################################
'
'Public Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
'
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
''Private Declare Function GetClipCursor Lib "user32" (lprc As RECT) As Long
'
'Const MOUSEEVENTF_ABSOLUTE = &H8000   'spostamento assoluto
'Const MOUSEEVENTF_LEFTDOWN = &H2    'pulsante sinistro premuto
'Const MOUSEEVENTF_LEFTUP = &H4      'pulsante sinistro rilasciato
'
''=======================
''   tipi enumerativi
''=======================
'Public Enum setBottoneAgente
'    ageCollegamenti = 0
'    ageMostraRiferimenti = 1
'    ageNascondiRiferimenti = 2
'    ageMostraCampiDB = 3
'    ageNascondiCampiDB = 4
'    ageDipendenze = 5
'End Enum
'
'Public Enum setTipoOpAn
'    enmCompila = 0
'    enmCarica = 1
'End Enum
'
''=======================
''   costanti
''=======================
'
''VINCOLI GENERALI
''Global Const RS_MASTRO_CI = "MaCliIta"
''Global Const RS_MASTRO_CE = "MaCliEst"
''Global Const RS_MASTRO_FI = "MaForIta"
''Global Const RS_MASTRO_FE = "MaForEst"
'Global Const RS_MASTRO_CLI = 1
'Global Const RS_MASTRO_FOR = 2
'
'Global Const GA_CreaDitta = 1
'Global Const GA_CreaAnno = 2
'Global Const GA_CopiaArchivi = 3
'Global Const GA_CancellaAnno = 4
'Global Const GA_TRASFSALDI = 5
'Global Const GA_TRASFSCAD = 6
'Global Const GA_TRASFPART = 7
'
'Global Const SEL_DITTE = 0
'Global Const SEL_ANNI = 1
'
''=============================================
''   dichiarazione costanti
''=============================================
'Enum enmTestSalva
'    tsnessuno = 0
'    tsSalvato = 1
'    tsNonSalvato = 2
'    tsritorna = 3
'End Enum
''=============================================
''=============================================
''   dichiarazione tipi di dati
''=============================================
''Gestione dei tasti della toolBox
''Type Metodo_Form_Attiva 'struttura che contiene informazioni sulla mdichild attiva
''     hwnd As Long   'handle
''     Tool_Mask As Long 'maschera dei tasti della toolbox
''End Type
'
'
''=============================================
''   dichiarazione variabili
''=============================================
''Global UltimoErr As Integer
''Global HlpAttivo As Integer
''Global SelAttiva  As Integer
''Global FormAttiva As Metodo_Form_Attiva
'
'Global strinitexe As String
'Global commitparziale As Integer
'Global GTestMode As Boolean
'
''Globali per la stampa CRW
'Global StpAVideo As Integer
''Global InStampa As Integer
'
'Global hVinCfg As Integer      'handle del DYNASET dei S/Conti Generici Vincolati
'
'Global frmModuli As Form
'
''*** DESIGNER ***
'Global Designer As MXDesigner.cDesigner
'
'
''Flag per sapere se l'utente ha selezionato gli Extra Articoli o gli Extra Depositi
''per la Copia Archivi
'Dim ExtraArtSel%
'Dim ExtraDepSel%
'Dim ExtraGiacDepSel%
'
''##### PER  METODO 2005 #################################################################
'Dim MIdxBotAgentiAttuale As Long
'Dim MIdxBotDesignerAttuale As Long
'Dim MIdxBotZoomAttuale As Long
'Dim MIdxBotTemaAttuale As Long
'
''per vedere se ci sono cambi ditta in corso
'Global CmbDittaBusy As Boolean
'
'#If IsMetodo2005 = 1 Then
'  Global mMessagingEngine As Object
'  Global mMetodoInterop As CMetodoInterop
'  Global mMetodoBrowser As Object 'MxBrowser.CBrowserEngine
'  Global GTemaAttivo As String
'#End If
''########################################################################################
'
'Global GBolNoMsgConfermaUscita As Boolean


#If IsMetodo2005 = 1 Then

    Public Sub InitToolbarForm(frmHost As Object)
        Dim tlbForm As XtremeCommandBars.CommandBar
        Dim Btn As CommandBarControl
        Dim btnPopup As CommandBarPopup

        'CommandBarsGlobalSettings.App = App
        'Risorse in lingua per i componenti Codejock
        CommandBarsGlobalSettings.ResourceFile = MXNU.PercorsoPgm & "\LanguageResources\XTPResource" & MXNU.LinguaAttiva & ".dll"

        If TypeOf frmHost Is Form Then
            frmHost.CommandBars.EnableCustomization False
            frmHost.CommandBars.ActiveMenuBar.Visible = False
            Set tlbForm = frmHost.CommandBars.Add("BarraForm", xtpBarTop)
            frmHost.CommandBars.Options.KeyboardCuesUse = xtpKeyboardCuesUseNone   'Anomalia 8691
            frmHost.CommandBars.KeyBindings.DeleteAll   'Necessario altrimenti in caso di caricamento di layout vecchi vengono reimpostate le combinazioni di tasti vecchie (es. F12 per aprire le Opzioni Generali)
        Else
            frmHost.Controls("CommandBars").EnableCustomization False
            frmHost.Controls("CommandBars").ActiveMenuBar.Visible = False
            Set tlbForm = frmHost.Controls("CommandBars").Add("BarraForm", xtpBarTop)
            frmHost.Controls("CommandBars").Options.KeyboardCuesUse = xtpKeyboardCuesUseNone   'Anomalia 8691
            frmHost.Controls("CommandBars").KeyBindings.DeleteAll   'Necessario altrimenti in caso di caricamento di layout vecchi vengono reimpostate le combinazioni di tasti vecchie (es. F12 per aprire le Opzioni Generali)
        End If
        tlbForm.BarID = ID_TLB_PRINC
        tlbForm.ContextMenuPresent = False

        tlbForm.Closeable = False
        tlbForm.Controls.DeleteAll

        With tlbForm.Controls
            Set Btn = .Add(xtpControlButton, idxBottoneInserisci, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "inserimento")
            Btn.ToolTipText = MXNU.CaricaStringaRes(2)
            Set Btn = .Add(xtpControlButton, idxBottoneDettaglio, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "dettagli")
            Btn.ToolTipText = MXNU.CaricaStringaRes(3)
            Set Btn = .Add(xtpControlButton, idxBottoneRegistra, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "registra")
            Btn.ToolTipText = MXNU.CaricaStringaRes(4)
            Set Btn = .Add(xtpControlButton, idxBottoneAnnulla, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "annulla")
            Btn.ToolTipText = MXNU.CaricaStringaRes(5)
            Set Btn = .Add(xtpControlButton, idxBottonePrimo, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "primo")
            Btn.ToolTipText = MXNU.CaricaStringaRes(6)
            Btn.BeginGroup = True
            Set Btn = .Add(xtpControlButton, idxBottonePrecedente, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "precedente")
            Btn.ToolTipText = MXNU.CaricaStringaRes(7)
            Set Btn = .Add(xtpControlButton, idxBottoneSuccessivo, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "succesivo")
            Btn.ToolTipText = MXNU.CaricaStringaRes(8)
            Set Btn = .Add(xtpControlButton, idxBottoneUltimo, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "ultimo")
            Btn.ToolTipText = MXNU.CaricaStringaRes(9)
            Set Btn = .Add(xtpControlButton, idxBottoneTrova, "")
            Btn.IconId = ImgListKey2ImgListIdx(metodo.ImglistBottoniXP, "trova")
            Btn.ToolTipText = MXNU.CaricaStringaRes(10)
            Btn.BeginGroup = True
        End With
        tlbForm.ModifyStyle XTP_CBRS_GRIPPER, 0
        tlbForm.EnableDocking xtpFlagAlignTop
        If TypeOf frmHost Is Form Then
            frmHost.CommandBars.AddImageList metodo.ImglistBottoniXP
            If MXCtrl.TemaAttivo = "Office2007" Then
                frmHost.CommandBars.GlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2007.dll", "Office2007Blue.Ini"
                frmHost.CommandBars.VisualTheme = xtpThemeResource
            ElseIf MXCtrl.TemaAttivo = "Office2010" Then
                frmHost.CommandBars.GlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2010.dll", "Office2010Blue.Ini"
                frmHost.CommandBars.VisualTheme = xtpThemeRibbon    'xtpThemeResource
            Else
                frmHost.CommandBars.VisualTheme = MXNU.FrmMetodo.CommandBars.VisualTheme
            End If
            frmHost.CommandBars.Options.ShowExpandButtonAlways = False
            frmHost.CommandBars.DockToolBar tlbForm, 0, 0, xtpBarTop
        Else
            frmHost.Controls("CommandBars").AddImageList metodo.ImglistBottoniXP
            If MXCtrl.TemaAttivo = "Office2007" Then
                frmHost.Controls("CommandBars").GlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2007.dll", "Office2007Blue.Ini"
                frmHost.Controls("CommandBars").VisualTheme = xtpThemeResource
            ElseIf MXCtrl.TemaAttivo = "Office2010" Then
                frmHost.Controls("CommandBars").GlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2010.dll", "Office2010Blue.Ini"
                frmHost.Controls("CommandBars").VisualTheme = xtpThemeRibbon 'xtpThemeResource
            Else
                frmHost.Controls("CommandBars").VisualTheme = MXNU.FrmMetodo.CommandBars.VisualTheme
            End If
            frmHost.Controls("CommandBars").Options.ShowExpandButtonAlways = False
            frmHost.Controls("CommandBars").DockToolBar tlbForm, 0, 0, xtpBarTop
        End If
    End Sub

    Public Sub CambiaTema(NomeSkin As String)
        On Local Error Resume Next
        Dim bolPanelNascosto As Boolean
        metodo.MousePointer = vbHourglass
        bolPanelNascosto = metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hidden
        If Not bolPanelNascosto Then metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Close
        DoEvents

        If NomeSkin <> "" Then
            GTemaAttivo = NomeSkin  'Left(NomeSkin, InStrRev(NomeSkin, ".") - 1)
        Else
            GTemaAttivo = ""
        End If
        DoEvents
        Dim SysColor As OLE_COLOR

        CommandBarsGlobalSettings.ColorManager.EnableLunaBlueForRoyaleTheme = False
        DockingPaneGlobalSettings.ColorManager.EnableLunaBlueForRoyaleTheme = False
        SuiteControlsGlobalSettings.ColorManager.EnableLunaBlueForRoyaleTheme = False
        ShortcutBarGlobalSettings.ColorManager.EnableLunaBlueForRoyaleTheme = False

        Select Case LCase(GTemaAttivo)
            Case "office2010"
                SysColor = RGB(187, 206, 230)
                'CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
                'DockingPaneGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
                'SuiteControlsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
                'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
            Case "office2007"
                SysColor = RGB(191, 219, 255)
                'CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
                'DockingPaneGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
                'SuiteControlsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
                'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
            Case "winxp.luna"
                SysColor = RGB(230, 227, 210)
                'CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
                'DockingPaneGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
                'SuiteControlsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
                'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
            Case "winxp.royale"
                SysColor = RGB(228, 226, 230)
                'CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
                'DockingPaneGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
                'SuiteControlsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
                'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
            Case "vista"
                SysColor = RGB(232, 232, 232)
                'CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAero
                'DockingPaneGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAero
                'SuiteControlsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAero
                'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAero
            Case Else
                SysColor = 0
                'CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
                'DockingPaneGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
                'SuiteControlsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
                'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
        End Select

        If InStr(1, NomeSkin, "Office2010", vbTextCompare) Then
            metodo.CommandBars.EnableOffice2007Frame True
            metodo.CommandBars.GlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2010.dll", "Office2010Blue.Ini"
            metodo.CommandBars.VisualTheme = xtpThemeRibbon   'xtpThemeOffice2007
            DockingPaneGlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2010.dll", "Office2010Blue.Ini"
            metodo.DockingPaneManager.VisualTheme = ThemeResource
            ShortcutBarGlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2010.dll", "Office2010Blue.Ini"
            frmModuli2005.ShortcutBar1.VisualTheme = xtpShortcutThemeResource
            SuiteControlsGlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2010.dll", "Office2010Blue.Ini"
            CalendarGlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2010.dll", "Office2010Blue.Ini"
        ElseIf InStr(1, NomeSkin, "Office2007", vbTextCompare) Then
            metodo.CommandBars.EnableOffice2007Frame True
            metodo.CommandBars.GlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2007.dll", "Office2007Blue.Ini"
            metodo.CommandBars.VisualTheme = xtpThemeResource
            DockingPaneGlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2007.dll", "Office2007Blue.Ini"
            metodo.DockingPaneManager.VisualTheme = ThemeResource
            ShortcutBarGlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2007.dll", "Office2007Blue.Ini"
            frmModuli2005.ShortcutBar1.VisualTheme = xtpShortcutThemeResource
            SuiteControlsGlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2007.dll", "Office2007Blue.Ini"
            CalendarGlobalSettings.ResourceImages.LoadFromFile MXNU.PercorsoPgm & "\Themes\Office2007.dll", "Office2007Blue.Ini"
        ElseIf InStr(1, NomeSkin, "Luna", vbTextCompare) Then
            metodo.CommandBars.EnableOffice2007Frame False
            metodo.CommandBars.VisualTheme = xtpThemeOffice2003
            metodo.DockingPaneManager.VisualTheme = ThemeOffice2003
            frmModuli2005.ShortcutBar1.VisualTheme = xtpShortcutThemeOffice2003
        Else
            metodo.CommandBars.EnableOffice2007Frame False
            metodo.CommandBars.VisualTheme = xtpThemeVisualStudio2008
            metodo.DockingPaneManager.VisualTheme = ThemeExplorer
            frmModuli2005.ShortcutBar1.VisualTheme = xtpShortcutThemeNativeWinXP
        End If

        frmModuli.CommandBars.VisualTheme = metodo.CommandBars.VisualTheme
        frmModuli.ShortcutCaption1.VisualTheme = frmModuli2005.ShortcutBar1.VisualTheme
        frmModuli.ShortcutCaption2.VisualTheme = frmModuli2005.ShortcutBar1.VisualTheme

        'SetTreeViewBackColor frmModuli.TrwModuli, SysColor
        'SetTreeViewBackColor frmModuli.TrVSearch, SysColor
        'SetTreeViewBackColor frmModuli.TwPreferiti, SysColor

        frmMenu.MWSplitter1.BackColor = metodo.CommandBars.GetSpecialColor(XPCOLOR_TOOLBAR_FACE)
        If NomeSkin <> "" Then
            Call CambiaSchemaColori(True, SysColor, SysColor)
        Else
            Call CambiaSchemaColori(False, SysColor, SysColor)
        End If
        Call InizializzaSpread(True, True, SysGradientColor1)
        'Inizializzo lo spread anche per MXKit e per MXBusiness
        Call MXVI.InitSpreadEvolus
        Call MXBusiness.InitSpreadEvolus
        DoEvents
        Call metodo.CambiaColoriFormAttive(True)
        DoEvents

        If Not bolPanelNascosto Then metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Select

        DoEvents
        metodo.MousePointer = vbDefault
        On Local Error GoTo 0
        MXCtrl.TemaAttivo = GTemaAttivo

Uscita:
        Exit Sub

'CtrlSkinMenus:
'    Dim oDocument As MSXML2.DOMDocument
'    Dim UserNode As MSXML2.IXMLDOMNode
'    Dim objNode As MSXML2.IXMLDOMNode
'
'    On Local Error Resume Next
'    'In alcuni sistemi operativi (es. Windows XP) applicando la skin anche al menu non funziona più l'attivazione della barra dei menu della MDI tramite tastiera (tasto ALT oppure F10)
'    'Leggo se attivare la skin dal file di configurazione
'    If Dir$(MXNU.PercorsoPreferenze & "\SkinFrameworkExclude.xml", vbNormal) <> "" Then
'        Set oDocument = New MSXML2.DOMDocument
'        If oDocument.Load(MXNU.PercorsoPreferenze & "\SkinFrameworkExclude.xml") Then
'            Set UserNode = oDocument.selectSingleNode("//user[@name='" & LCase(MXNU.NomeComputer) & "']")
'            If (UserNode Is Nothing) Then
'                Set UserNode = oDocument.selectSingleNode("//user[@name='*']")
'            End If
'            If Not (UserNode Is Nothing) Then
'                Set objNode = UserNode.selectSingleNode("skinmenus")
'                If Not (objNode Is Nothing) Then
'                    If objNode.text <> "0" Then
'                        metodo.SkinFramework.ApplyOptions = metodo.SkinFramework.ApplyOptions Or xtpSkinApplyMenus
'                    End If
'                End If
'            End If
'        End If
'        Set objNode = Nothing
'        Set UserNode = Nothing
'        Set oDocument = Nothing
'    End If
'Return
    End Sub

    Public Sub GestioneToolBut2005(ByVal btnId As Long)

        Dim frmAttiva As Form
        Dim bolCancellaAzione As Boolean
        Dim intAzioneEseguita As Integer
        Dim bolApriNomiCtrl As Boolean
        Dim bolAllInOne As Boolean
        'Dim DesignButton As MSComctlLib.Button ' DESIGNER
        Dim DesignButton As XtremeCommandBars.CommandBarControl
        Dim frmDesign As Form

        Dim intForm As Integer
        Dim bolFrmZoomAperta As Boolean
        Dim objExt As Object
        Dim Button As Object  'CommandBarButton
        Dim bolChkEnabled As Boolean, bolValido As Boolean

        bolChkEnabled = True
        Select Case btnId
            Case idxBottoneAttivaDesigner To idxBottoneGrigliaDesigner
                Set Button = metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneDesigner).CommandBar.FindControl(, btnId)
            '**** Bottone Zoom tolto su 9.00.00 *******************************************************
            'Case idxBottoneZoom100 To idxBottoneZoom200
            '    Set Button = metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneZoom).CommandBar.FindControl(, btnId)
            Case idxBottoneDesigner
                If MIdxBotDesignerAttuale = 0 Then MIdxBotDesignerAttuale = idxBottoneAttivaDesigner
                Set Button = metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneDesigner).CommandBar.FindControl(, MIdxBotDesignerAttuale)
                btnId = MIdxBotDesignerAttuale
            '**** Bottone Zoom tolto su 9.00.00 *******************************************************
            'Case idxBottoneZoom
            '    If MIdxBotZoomAttuale = 0 Then MIdxBotZoomAttuale = idxBottoneZoom100
            '    Set Button = metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneZoom).CommandBar.FindControl(, MIdxBotZoomAttuale)
            '    btnId = MIdxBotZoomAttuale
            Case ID_TLBITEM_CHANGETHEME
                If MIdxBotTemaAttuale = 0 Then MIdxBotTemaAttuale = ID_TLBITEM_THEME_OFFICE2010   'ID_TLBITEM_THEME_OFFICE2007
                bolChkEnabled = False
            Case ID_TLBITEM_THEME_OFFICE2007 To ID_TLBITEM_THEME_VISTA, ID_TLBITEM_THEME_SYSTEM, ID_TLBITEM_THEME_OFFICE2010
                Set Button = metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, btnId)
            Case Else
                Set Button = metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, btnId)
        End Select
        If bolChkEnabled Then
            If Not (Button Is Nothing) Then
                bolValido = Button.Enabled
            Else
                bolValido = False
            End If
        Else
            bolValido = True
        End If
        'If Button.Enabled Then
        If bolValido Then
            On Local Error Resume Next
            Set frmAttiva = metodo.FormAttiva
            metodo!helpTimer.Interval = 500
            '***If TypeName(Button) = "IButtonMenu" Then
                Select Case btnId
                ' Gestione Zoom su foglio scheda sviluppo n.ro 1128 + Designer
                '***If metodo.Barra.Buttons(idxBottoneDefAgenti).Value = tbrUnpressed And LCase$(Left$(Button.key, 4)) <> "zoom" And LCase$(Left$(Button.key, 6)) <> "design" Then
                '*** Select Case Button.Index
                        Case idxBottoneDefAgenti    'Def. Agenti
                            ' Rif. anomalia n.ro 4824
                            If MXNU.ModuloRegole And Not (MXNU.CtrlAccessi) Then
                                Call FormImpostaAgenti(frmAttiva)
                            End If
                        Case idxBottoneNomiCtrlCmp  'Nome Controlli - campi
                            If (metodo.FormsCount > 2) Then
                                bolApriNomiCtrl = True
                                Call frmAttiva.AzioniMetodo(MetFMostraCampiDBAnagr, bolApriNomiCtrl)
                                If bolApriNomiCtrl Then
                                    Set FrmNomiControlli.frmDef = frmAttiva
                                    FrmNomiControlli.Show
                                End If
                            End If
                        Case idxBottoneSituazAnagr   'Situazione Anagrafica
                            'mostro le dipendenze dell'anagrafica
                            Dim ListaCol As New Collection
                            Call frmAttiva.AzioniMetodo(MetFVisDipendenze, ListaCol)
                            If ListaCol.Count > 0 Then
                                Call MXVA.VisualizzaDipendenze(ListaCol)
                            End If
                            Set ListaCol = Nothing
                        Case idxBottoneAttivaFileLog   'Attiva\Disattiva Log
                            'If Button.Caption = "Disattiva File Log" Then
                            If Button.Checked Then
                                Call MXDB.DisattivaLog
                                Button.ToolTipText = "Attiva File Log"
                                Call frmLog.MostraFileLog(MXNU.GetTempDir & "MWSqlLog" & MXNU.NTerminale & ".sql")
                                Button.Checked = False
                            Else
                                Call MXDB.AttivaLog
                                Button.ToolTipText = "Disattiva File Log"
                                Button.Checked = True
                            End If
                        Case idxBottoneRicProfili   'ricarica profili
                            metodo.MousePointer = vbHourglass
                            MXNU.MostraMsgInfo 70058

                            '>>> profili anagrafiche
                            If Not MXVA Is Nothing Then
                                If MXVA.ChiudiDyTRAnagraf() Then
                                    Call MXVA.ApriDyTRAnagraf
                                End If
                                If MXVA.ChiudiDyTRValidazione() Then
                                    Call MXVA.ApriDyTRValidazione
                                End If
                            End If
                            '>>> profili tabelle
                            If Not MXCT Is Nothing Then
                                If MXCT.ChiudiDyTRTabelle() Then
                                    Call MXCT.ApriDyTRTabelle
                                End If
                            End If
                            '>>> profili visioni/situazioni
                            If Not MXVI Is Nothing Then
                                If MXVI.ChiudiDyTRSituazioni() Then
                                    Call MXVI.ApriDyTRSituazioni
                                End If
                                If MXVI.ChiudiDyTRVisioni() Then
                                    Call MXVI.ApriDyTRVisioni
                                End If
                            End If
                            MXNU.MostraMsgInfo ""
                            metodo.MousePointer = vbDefault
                            Button.DefaultItem = True

                            '>>> forzo ricaricamente regole synapse
                            Dim synexe As Object
                            If (Not NETFX Is Nothing) Then
                                Set synexe = NETFX.GetHostSynapse
                                synexe.Dispose
                                synexe = Nothing
                            End If
                '***ElseIf LCase$(Left$(Button.key, 6)) = "design" Then ' DESIGNER
                        Case idxBottoneAttivaDesigner To idxBottoneGrigliaDesigner
                '***End If
            '***Else
                Err.Clear
                    Case ID_TLBITEM_FRMORIGINALSIZE
                        If frmAttiva.Name <> "EmptyForm" And Not EsisteElementoCollection(McolFormsInNavBar, frmAttiva.Name) Then
                            Call frmAttiva.Frm_SetOriginalSize
                            If Err.Number <> 0 Then
                                Err.Clear
                                Call frmAttiva.mResize.Frm_SetOriginalSize
                            End If
                        End If
                    Case idxBottoneModuli
                        If metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hidden Then
                            'il menu moduli non e' aperto ne lockato quindi lo apro e lo locko
                            metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Close
                            metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Select
                            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, btnId).Checked = True
                            If frmModuli.ShortcutBar1.Selected.ID = ID_BAR_PROGMODULES Then
                                Call DaiFocusAlberoModuli
                            End If
                        Else
                            'il menu moduli e' gia' aperto e lockato, lo chiudo
                            metodo.DockingPaneManager.FindPane(ID_PANE_NAVBAR).Hide
                            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, btnId).Checked = False
                        End If

                    Case idxBottoneAllInOne
                    Case idxBottoneInserisci
                        Call EseguiAgenteAzioneForm(frmAttiva, "before_insert", bolCancellaAzione)
                        If Not bolCancellaAzione Then
                            Call DoLostFocus(frmAttiva.ActiveControl)
                            intAzioneEseguita = CInt(frmAttiva.AzioniMetodo(MetFInserisci))
                            'call frmAttiva.AzioniMetodo(MetFInserisci)
                            Call EseguiAgenteAzioneForm(frmAttiva, "after_insert", bolCancellaAzione, intAzioneEseguita)
                        End If
                    Case idxBottoneDettaglio
                        Call DoLostFocus(frmAttiva.ActiveControl)
                        Call frmAttiva.AzioniMetodo(MetFDettagli)

                    Case idxBottoneRegistra
                        Call EseguiAgenteAzioneForm(frmAttiva, "before_save", bolCancellaAzione)
                        If Not bolCancellaAzione Then
                            Call DoLostFocus(frmAttiva.ActiveControl)
                            intAzioneEseguita = CInt(frmAttiva.AzioniMetodo(MetFRegistra))
                            Call EseguiAgenteAzioneForm(frmAttiva, "after_save", bolCancellaAzione, intAzioneEseguita)
                        End If

                    Case idxBottoneAnnulla
                        Call EseguiAgenteAzioneForm(frmAttiva, "before_delete", bolCancellaAzione)
                        If Not bolCancellaAzione Then
                            Call DoLostFocus(frmAttiva.ActiveControl)
                            intAzioneEseguita = CInt(frmAttiva.AzioniMetodo(MetFAnnulla))
                            'Call frmAttiva.AzioniMetodo(MetFAnnulla)
                            Call EseguiAgenteAzioneForm(frmAttiva, "after_delete", bolCancellaAzione, intAzioneEseguita)
                        End If

                    Case idxBottonePrimo
                        Call EseguiAgenteAzioneForm(frmAttiva, "before_first", bolCancellaAzione)
                        If Not bolCancellaAzione Then
                            Call DoLostFocus(frmAttiva.ActiveControl)
                            Call frmAttiva.AzioniMetodo(MetFPrimo)
                        End If
                        Call EseguiAgenteAzioneForm(frmAttiva, "after_first", bolCancellaAzione)

                    Case idxBottonePrecedente
                        Call EseguiAgenteAzioneForm(frmAttiva, "before_previous", bolCancellaAzione)
                        If Not bolCancellaAzione Then
                            Call DoLostFocus(frmAttiva.ActiveControl)
                            Call frmAttiva.AzioniMetodo(MetFPrecedente)
                        End If
                        Call EseguiAgenteAzioneForm(frmAttiva, "after_previous", bolCancellaAzione)

                    Case idxBottoneSuccessivo
                        Call EseguiAgenteAzioneForm(frmAttiva, "before_next", bolCancellaAzione)
                        If Not bolCancellaAzione Then
                            Call DoLostFocus(frmAttiva.ActiveControl)
                            Call frmAttiva.AzioniMetodo(MetFSuccessivo)
                        End If
                        Call EseguiAgenteAzioneForm(frmAttiva, "after_next", bolCancellaAzione)
                    Case idxBottoneUltimo
                        Call EseguiAgenteAzioneForm(frmAttiva, "before_last", bolCancellaAzione)
                        If Not bolCancellaAzione Then
                            Call DoLostFocus(frmAttiva.ActiveControl)
                            Call frmAttiva.AzioniMetodo(MetFUltimo)
                        End If
                        Call EseguiAgenteAzioneForm(frmAttiva, "after_last", bolCancellaAzione)
                    Case idxBottoneStampa
                        Call frmAttiva.AzioniMetodo(MetFStampa)
                    Case idxBottoneTrova
                        If Not MXCT.TTrova(FrmTrovaGen) Then
                            Call frmAttiva.AzioniMetodo(MetFTrova)
                        End If
                    ' S#3040 - rimossa la Gestione Accessi da Evolus
'                    Case idxBottoneDefAccessi
'                        Call FormDefinisciAccessi(frmAttiva)
                    Case idxBottoneHelp
                        '#If ISMETODOXP = 1 Then
                        '    If MXNU.MetodoXP = True Then
                                    #If TOOLS = 1 Then
                                        If MXNU.DammiFormAttiva.Name = "frmModuli" Then
                                            Call EseguiAppAssociata(MXNU.PercorsoPgm & "\HELP\TECNICO", "INDICE.HTM")
                                        Else
                                            Call MXNU.ApriHelp(False)
                                        End If
                                    #Else
                                        'posso caricare l'help direttamente su il browser della form tooltip
                                        If MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", "METODOW", "HelpCompilato", 0) = 1 Then
                                            Call MXNU.ApriHelp(False)
                                        End If
                                    #End If
                            'Else
                            '    #If TOOLS = 1 Then
                            '        If MXNU.DammiFormAttiva.Name = "frmModuli" Then
                            '            Call EseguiAppAssociata(MXNU.PercorsoPgm & "\HELP\TECNICO", "INDICE.HTM")
                            '        Else
                            '            Call MXNU.ApriHelp(False)
                            '        End If
                            '    #Else
                            '        Call MXNU.ApriHelp(False)
                            '    #End If
                            'End If
                        '#Else
                        '    #If TOOLS = 1 Then
                        '        If MXNU.DammiFormAttiva.Name = "frmModuli" Then
                        '            Call EseguiAppAssociata(MXNU.PercorsoPgm & "\HELP\TECNICO", "INDICE.HTM")
                        '        Else
                        '            Call MXNU.ApriHelp(False)
                        '        End If
                        '    #Else
                        '        Call MXNU.ApriHelp(False)
                        '    #End If
                        '#End If
                    Case idxBottoneUtenteModifica
                        Call frmAttiva.AzioniMetodo(MetFVisUtenteModifica)
                    Case idxBottoneSchedula
                        Call frmAttiva.AzioniMetodo(MetFSchedulaOperazione)
                    'Case idxBottoneSpostaBarra
                    '    metodo!Barra.Align = (metodo!Barra.Align) Mod 2 + 1
                    '    metodo!Barra.Buttons(Button.Index).Image = metodo!Barra.Align + 10
                    '    metodo!tlbDesigner.Align = (metodo!tlbDesigner.Align) Mod 2 + 1 ' Designer
'**** Bottone Zoom tolto su 8.04.00 **************************************************************************************************
'                    Case idxBottoneZoom
'                        '#If ISMETODOXP = 1 Then
'                        '  If MXNU.MetodoXP Then
'                            Dim intZoomType As Integer
'                            intZoomType = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\Mw.ini", MXNU.UtenteAttivo, "ZoomType", 1)
'                            ' Gestione dell'apertura di una singola form di zoom
'                            bolFrmZoomAperta = False
'                            intForm = 0
'                            While (intForm < VB.Forms.Count) And (Not (bolFrmZoomAperta))
'                                If UCase(VB.Forms(intForm).NAME) = "FRMZOOM" Then
'                                    bolFrmZoomAperta = True
'                                End If
'                                intForm = intForm + 1
'                            Wend
'
'                            If Not bolFrmZoomAperta Then
'                                ' Rif. anomalia ISV n.ro 6 / Anomalia Metodo XP n.ro 6084 / Quesito ISV 294
'                                If LCase(frmAttiva.NAME) = "frmextchild" Then
'                                    If frmAttiva.ActiveControl.object.Controls.Count > 1 Then
'                                        Call frmZoom.DoZoom(frmAttiva, frmAttiva.ActiveControl.object.Controls(1).object.ExtActiveControl, intZoomType, frmAttiva.ActiveControl.object.Controls(1).object.ExtActiveControl.ActiveCol, frmAttiva.ActiveControl.object.Controls(1).object.ExtActiveControl.ActiveRow, MXNU.FrmMetodo.Barra.Buttons(idxBottoneTrova).Enabled)
'                                    Else
'                                        Call frmZoom.DoZoom(frmAttiva, frmAttiva.ActiveControl.object.Controls(0).object.ExtActiveControl, intZoomType, frmAttiva.ActiveControl.object.Controls(0).object.ExtActiveControl.ActiveCol, frmAttiva.ActiveControl.object.Controls(0).object.ExtActiveControl.ActiveRow, MXNU.FrmMetodo.Barra.Buttons(idxBottoneTrova).Enabled)
'                                    End If
'                                ElseIf TypeName(frmAttiva.ActiveControl) = "MWSchedaBox" Then
'                                    ' Sono in una scheda
'                                    If frmAttiva.ActiveControl.ControlsEx.Count > 0 Then
'                                        ' La scheda contiene delle estensioni
'                                        ' Setto lo UserControl dentro il wrapper in objExt (ExtWrapper deve avere una public property get come tutti gli UserControl estensioni)
'                                        Set objExt = frmAttiva.ActiveControl.ControlsEx(0).object.Controls
'                                        'Rif. anomalia #8088
'                                        If TypeName(objExt(0)) = "CTLXBus" Then
'                                            Set objActiveControl = objExt(1).object.ExtActiveControl
'                                        Else
'                                            Set objActiveControl = objExt(0).object.ExtActiveControl
'                                        End If
'                                        Set objExt = Nothing
'                                        Call frmZoom.DoZoom(frmAttiva, objActiveControl, intZoomType, objActiveControl.ActiveCol, objActiveControl.ActiveRow, MXNU.FrmMetodo.Barra.Buttons(idxBottoneTrova).Enabled)
'                                        Set objActiveControl = Nothing
'                                    End If
'                                Else
'                                    Call frmZoom.DoZoom(frmAttiva, frmAttiva.ActiveControl, intZoomType, frmAttiva.ActiveControl.ActiveCol, frmAttiva.ActiveControl.ActiveRow, MXNU.FrmMetodo.Barra.Buttons(idxBottoneTrova).Enabled)
'                                End If
'                            Else
'                                Call MXNU.MsgBoxEX(2738, vbExclamation, 23481)
'                            End If
'
'                        '  End If
'                        '#End If
'*****************************************************************************************************************************************
                    Case idxBottoneDesigner
                    Case ID_TLBITEM_INFO
                      Call ShellExecute(metodo.hwnd, "Open", "www.metodo.it/MetodoPortal/DesktopPortalMain.aspx?action=NEWS", "", App.Path, 1)
                    Case ID_TLBITEM_PREFERENCES
                    Case ID_TLBITEM_THEME_OFFICE2007 To ID_TLBITEM_THEME_VISTA, ID_TLBITEM_THEME_SYSTEM, ID_TLBITEM_THEME_OFFICE2010
                        Dim NomeFileTema As String
                        Select Case btnId
                            Case ID_TLBITEM_THEME_OFFICE2010: NomeFileTema = "Office2010"
                            Case ID_TLBITEM_THEME_OFFICE2007: NomeFileTema = "Office2007"
                            Case ID_TLBITEM_THEME_WINXPLUNA: NomeFileTema = "WinXP.Luna"
                            Case ID_TLBITEM_THEME_WINXPROYALE: NomeFileTema = "WinXP.Royale"
                            Case ID_TLBITEM_THEME_VISTA: NomeFileTema = "Vista"
                            Case ID_TLBITEM_THEME_SYSTEM: NomeFileTema = ""
                        End Select
                        Call CambiaTema(NomeFileTema)


                        If MIdxBotTemaAttuale = 0 Then
                            Dim i%
                            For i = ID_TLBITEM_THEME_OFFICE2007 To ID_TLBITEM_THEME_VISTA
                                metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, i).DefaultItem = False
                            Next i
                            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, ID_TLBITEM_THEME_OFFICE2010).DefaultItem = False
                        Else
                            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, MIdxBotTemaAttuale).DefaultItem = False
                        End If
                        metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, ID_TLBITEM_CHANGETHEME).CommandBar.FindControl(, btnId).DefaultItem = True
                        MIdxBotTemaAttuale = btnId
                End Select
            '***End If
            Set frmAttiva = Nothing
        End If

    End Sub

    'Nel caso si docki la navigation bar, il focus non è posizionato sull'albero dei moduli per un problema degli oggetti Codejock;
    'aggiro il problema simulando un click del mouse.
    Public Sub DaiFocusAlberoModuli()
        Dim pt As POINTAPI, OldPt As POINTAPI
        Dim oldClipCur As RECT
        Dim r As RECT
        Dim bolCalcolaPt As Boolean
        On Local Error Resume Next

        Call GetCursorPos(OldPt)   'Memorizzo l'attuale posizione del mouse
        Call GetClipCursor(oldClipCur)
        DoEvents
        Sleep 30   'Attenzione: Sleep necessario altrimenti và in errore di protezione generale!!

        'Determino le coordinate di un punto all'interno dell'albero dei moduli...
        pt.x = (frmMenu.TrwModuli.Left + frmMenu.TrwModuli.Width / 2) \ Screen.TwipsPerPixelX
        'pt.y = (frmMenu.TrwModuli.Top + frmMenu.TrwModuli.Height / 2) \ Screen.TwipsPerPixelY
        pt.y = 150

        'Confino il cursore all'interno dell'albero dei moduli
        Call GetWindowRect(frmMenu.TrwModuli.hwnd, r)
        Call ClipCursor(r)

        '...e ci posiziono il mouse
        SetCursorPos pt.x, pt.y
        DoEvents
        'Simulo il click del mouse
        mouse_event MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP, pt.x, pt.y, 0, 0
        DoEvents
        'Riposiziono il mouse alle coordinate originali
        Call ClipCursor(oldClipCur)

        SetCursorPos OldPt.x, OldPt.y
        On Local Error GoTo 0
    End Sub


#End If


'Sub AbilitaDisabilitaMetodo(ByVal abilita As Boolean)
'   Dim q%, ad%
'   'Salva la Maschera dei Bottoni Attivi quando premo il bottone delle Regole
'   If Not abilita Then
'        MXNU.MascheraAttiva = 0
'        Call MXNU.DammiMascheraAttiva
'   End If
'   'Abilita/Diabilita tutte le Form Attive
'   For q = 0 To Forms.Count - 1
'        Forms(q).Enabled = abilita
'   Next q
'   metodo.Enabled = True
'   'Abilita/Disabilita tutti i controlli della form Principale
'    For q = 0 To metodo.Controls.Count - 1
'        If TypeOf metodo.Controls(q) Is Menu Then
'            If metodo.Controls(q).Caption <> "-" Then metodo.Controls(q).Enabled = abilita
'        ElseIf TypeOf metodo.Controls(q) Is Toolbar Then
'            If metodo.Controls(q).Index <> 19 Then
'                metodo.Controls(q).Enabled = abilita
'            End If
'        End If
'    Next q
'    'Ripristina i Bottoni Attivi quando premo per la seconda volta il bottone delle Regole
'    If abilita Then
'        Call MXNU.SetMascheraAttiva
'    End If
'End Sub


'Sub ChiudiFormAttive()
'    Dim q As Integer
'    Dim strNome As String
'    Dim intLimitemax As Integer
'    Dim bolUnload As Boolean
'
'    #If IsMetodo2005 = 1 Then
'        If MbolInChiusura Then
'            intLimitemax = 1
'        Else
'            intLimitemax = McolFormsInNavBar.Count + 1
'        End If
'    #Else
'        'Rif.sch. 8198 - (form di log con error 91)
'        If MbolInChiusura Then
'            intLimitemax = 1
'        Else
'            intLimitemax = 2
'        End If
'    #End If
'    Do
'        If Forms.Count <= intLimitemax Then
'            Exit Do
'        Else
'            For q = 0 To Forms.Count - 1
'                strNome = Forms(q).Name
'                bolUnload = False
'                If (StrComp(strNome, "frmModuli98", vbTextCompare) <> 0) _
'                    And (StrComp(strNome, "frmModuliXP", vbTextCompare) <> 0) _
'                    And (StrComp(strNome, "metodo", vbTextCompare) <> 0) _
'                    And (StrComp(strNome, "frmToolTip", vbTextCompare) <> 0) Then
'                    bolUnload = True
'                    #If IsMetodo2005 Then
'                        If Not MbolInChiusura Then
'                            bolUnload = Not EsisteElementoCollection(McolFormsInNavBar, strNome)
'                        End If
'                    #End If
'                    If bolUnload Then
'                        Dim hwnd As Long
'                        hwnd = Forms(q).hwnd
'                        Unload Forms(q)
'                        If IsWindow(hwnd) Then CmbDittaBusy = True
'                        While IsWindow(hwnd)
'                            DoEvents
'                            Sleep 100
'                        Wend
'                        CmbDittaBusy = False
'                        Exit For
'                    End If
'                End If
'            Next q
'        End If
'    Loop
'
''    q = 1
''    Do While Forms.Count > intLimitemax
''        strNome = Forms(Forms.Count - q).Name
''        'Rif. Anomalie98 Nr. 3536: la frmModuli ha cambiato nome; è frmModuli98 oppure frmModuliXP
''        If (StrComp(strNome, "frmModuli98", vbTextCompare) <> 0) _
''            And (StrComp(strNome, "frmModuliXP", vbTextCompare) <> 0) _
''            And (StrComp(strNome, "metodo", vbTextCompare) <> 0) _
''            And (StrComp(strNome, "frmToolTip", vbTextCompare) <> 0) Then
''            Unload Forms(Forms.Count - q)
''            q = 1
''        Else
''            If StrComp(strNome, "frmModuli98", vbTextCompare) = 0 _
''               Or StrComp(strNome, "frmModuliXP", vbTextCompare) = 0 _
''               Or StrComp(strNome, "frmToolTip", vbTextCompare) = 0 Then
''                q = 2
''            ElseIf StrComp(strNome, "metodo", vbTextCompare) = 0 Then
''                q = 3
''            End If
''        End If
''    Loop
'End Sub

'Chiude tutte i database attivi,tutte le form ed esce dal programma.
'Sub ChiudiMetodo()
'    Dim q As Integer
'
'    If VB.Forms.Count > 1 Then
'        If Not (frmSelDitta Is Nothing) Then
'            Unload frmSelDitta
'            Set frmSelDitta = Nothing
'        End If
'    End If
'
'    #If USAM98SERVER Then
'    If Not GobjM98Server Is Nothing Then
'        'Call GobjM98Server.Termina
'        Set GobjM98Server = Nothing
'    End If
'    #End If
'
'    'On Local Error Resume Next
'    'chiusura recordset calendari
'    If MXNU.CalendariLocale Then
'        MXCICLI.TerminaRecSetCalendari
'    End If
'    'On Local Error GoTo 0
'
'    'RIF.A#7602 - Si assicura che anche tutte le form di AIOT siano chiuse
'    If (Not MXALL Is Nothing) Then
'        Call MXALL.CloseAllWindows
'    End If
'
'    If Not (hndDBArchivi Is Nothing) Then q = MXDB.dbChiudiDB(hndDBArchivi)
'    q = MXDB.dbDisattiva()
'
'    Call DropObjKitBus
'
'#If IsMetodo2005 = 1 Then
'    On Local Error Resume Next
'    'chiusura mailing system
'    If (Not mMessagingEngine Is Nothing) Then
'        mMessagingEngine.Dispose
'        Set mMessagingEngine = Nothing
'    End If
'
'    'distruzione dell'hosting Metodo
'    Set mMetodoInterop = Nothing
'#End If
'
'    'While RegolaBloccata()
'    '    DoEvents
'    'Wend
'    'ChiudiRegola
'    'EndAmbiente
'
'End Sub

Sub AggiornaStatusBar()
    Dim strStato As String

    'aggiorno la caption dell'applicazione
    #If TOOLS = 1 Then
        'metodo.Caption = App.Title & " - [" & MXNU.PercorsoPreferenze & "]"
    #Else
        'metodo.Caption = App.Title & IIf(MXNU.ControlloModuliChiave(modChiaveDemo) = 0, " - Versione dimostrativa non destinata alla vendita", vbNullString)
    #End If
    'aggiorno la barra di stato
    Select Case MXNU.StatoEsercizioCont
'        Case 0, 2
'            strStato = " [T]"
'        Case 1
'            strStato = ""
'        Case 3
'            strStato = " [C]"
        Case 0
            strStato = ""
        Case 1
            strStato = " [C]"
    End Select
    'metodo.BarraStato.Panels("dittaanno").Text = MXNU.Dsc_Breve_Ditta & " - " & MXNU.AnnoAttivo & " " & strStato
    'metodo.BarraStato.Panels("utente").Text = MXNU.UtenteAttivo & " (" & MXNU.NTerminale & ")"
    'metodo.BarraStato.Panels("Lingua").Text = MXNU.LinguaAttiva
    #If ISMETODOXP = 1 Then
        If MXNU.MetodoXP Then
            'metodo.BarraStato.Panels("Designer").Text = MXNU.VersioneAttiva
        End If
    #End If
    #If IsMetodo2005 = 1 Then
        metodo.BarraStato.Panels("dittaanno").Width = 0  'AutoSize
        metodo.BarraStato.Panels("dittaanno").Width = metodo.BarraStato.Panels("dittaanno").Width + 20
        metodo.BarraStato.Panels("utente").Width = 0  'AutoSize
        metodo.BarraStato.Panels("utente").Width = metodo.BarraStato.Panels("utente").Width + 20
        metodo.BarraStato.Panels("Designer").Width = 0 'AutoSize
        metodo.BarraStato.Panels("Designer").Width = metodo.BarraStato.Panels("Designer").Width + 35

        'Reimposto le immagini a run-time, altrimenti perde il maskcolor (?!?)
        'metodo.BarraStato.Panels("utente").Picture = metodo.ImglistBottoniXP.ListImages("utente").Picture
        'metodo.BarraStato.Panels("dittaanno").Picture = metodo.ImglistBottoniXP.ListImages("dbase").Picture
        'metodo.BarraStato.Panels("Lingua").Picture = metodo.ImglistBottoniXP.ListImages("Euro").Picture
        On Local Error Resume Next
        Dim pct As StdPicture
        If Dir$(MXNU.PercorsoPgm & "\LanguageResources\Flags\Flag" & MXNU.LinguaAttiva & ".ico", vbNormal) <> "" Then
            'metodo.BarraStato.Panels("Lingua").Picture = LoadPicture(MXNU.PercorsoPgm & "\LanguageResources\Flags\Flag" & MXNU.LinguaAttiva & ".ico")

            Set pct = LoadPicture(MXNU.PercorsoPgm & "\LanguageResources\Flags\Flag" & MXNU.LinguaAttiva & ".ico")

            metodo.CommandBars.AddIconHandle pct.Handle, STATUSBAR_ID_PANELLINGUA, 0, False
            metodo.BarraStato.Panels("Lingua").IconIndex = STATUSBAR_ID_PANELLINGUA
        Else
            'metodo.BarraStato.Panels("Lingua").Picture = LoadPicture(MXNU.PercorsoPgm & "\LanguageResources\Flags\FlagEuro.ico")
            Set pct = LoadPicture(MXNU.PercorsoPgm & "\LanguageResources\Flags\FlagEuro.ico")
            metodo.CommandBars.AddIconHandle pct.Handle, STATUSBAR_ID_PANELLINGUA, 0, False
            metodo.BarraStato.Panels("Lingua").IconIndex = STATUSBAR_ID_PANELLINGUA

        End If
        On Local Error GoTo 0
        'metodo.BarraStato.Panels("Designer").Picture = metodo.ImglistBottoniXP.ListImages("VersDesigner").Picture
    #End If
End Sub

Function ErroreVB(IDErrore%, Msg$) As Integer
    Dim strErrore$, q%
    strErrore = "Errore " & IDErrore & ": "
    ErroreVB = True
    Select Case IDErrore
        Case 5 'Illegal Function Call
            strErrore = strErrore & "Chiamata a funzione non valida."
        Case 6 'OverFlow
            strErrore = strErrore & "Valore fuori limite."
        Case 9 'SubScript Out of Range
            strErrore = strErrore & "Indice di vettore fuori limite."
        Case 11 'Division by Zero
            strErrore = strErrore & "Divisione per zero."
        Case 13 'Type Mismatch
            strErrore = strErrore & "Errore di conversione di dati."
        Case 94 'Invalid use of NULL
            strErrore = strErrore & "Uso non valido dell'operatore NULL."
        Case 380 'Invalid property Value
            strErrore = strErrore & "Settaggio di Proprietà non valido."
        Case ERR_UTENTE
            If Msg = "" Then
                strErrore = "Operazione Fallita."
            Else
                strErrore = Msg
            End If
            ErroreVB = False
        Case Else
            strErrore = "Errore generico: " & Error$
    End Select
    q = MXNU.MsgBoxEX(strErrore, vbExclamation, "ATTENZIONE!")
End Function

'Sub Focus_Oggetto(ByVal Desthwnd%, ByVal hwnd&)
'    Dim crct As RECT
'    Dim rct As RECT
'    Dim destDC%, res%
'
'    GetWindowRect Desthwnd%, crct
'    GetWindowRect hwnd&, rct
'    'differenza
'    rct.Left = rct.Left - crct.Left
'    rct.Right = rct.Right - crct.Left
'    rct.Top = rct.Top - crct.Top
'    rct.Bottom = rct.Bottom - crct.Top
'
'    MsgBox "InflateRect rct, 3, 3"
'
'    destDC% = GetDC(Desthwnd%)
'    MsgBox "DrawFocusRect destDC, rct"
'    res = ReleaseDC(Desthwnd%, destDC)
'End Sub


'Function FormLoader(FRM As Form, HelpContextID As Long) As Integer
'    FormLoader = False
'    #If IsMetodo2005 <> 1 Then
'        If metodo.mnuMenu.Enabled = False Then
'            FormLoader = True
'            Exit Function
'        End If
'    #End If
'    On Local Error Resume Next
'    FRM.FormProp.FormID = HelpContextID
'    On Local Error GoTo out_of_memory
'    Load FRM
'    Call AssegnaToolTip(FRM)
'
'    'FRM.HelpContextID = HelpContextID
'
''    'REMIND: da decommentare a sviluppo terminato
''    '*** modifica ExtensionLoader ***
''    If (Not Designer Is Nothing) Then
''        Dim colObj As Collection
''        Dim colAmb As Collection
''
''        On Local Error Resume Next
''        Set colObj = New Collection
''        colObj.Add hndDBArchivi
''        Set colAmb = Ambienti2Collection(False)
''        Call Designer.LoadExtension(frm, colAmb, colObj)
''    End If
'
'    With FRM
'        .ZOrder vbBringToFront
'        .WindowState = vbNormal
'    End With
'    On Local Error Resume Next
'    Call FRM.MWAgt1.RegistraAgenteFrm(FRM)
'    FormLoader = True
'    metodo.MousePointer = Default
'    On Local Error GoTo 0
'fine_Load:
'    Exit Function
'out_of_memory:
'     If Err = 7 Then Resume Next Else Resume fine_Load
'
'End Function

'Sviluppo nr. 1570
Private Sub AssegnaToolTip(frm As Form)
    Dim objCtrl As Control

    On Local Error Resume Next
    If MXNU.LinguaAttiva <> "IT" Then
        For Each objCtrl In frm.Controls
            Select Case TypeName(objCtrl)
                Case "Label"
                    objCtrl.ToolTipText = Replace(objCtrl.Caption, "&", "")
                Case "MWEtichetta", "MWLinguetta"
                    objCtrl.Lingua = MXNU.LinguaAttiva
            End Select
        Next
    End If
    On Local Error GoTo 0

End Sub

#If IsMetodo2005 <> 1 Then
'Fatta Globale per tutti i moduli
'Sub GestioneToolBut(ByVal Button As Object)
'
'    Dim frmAttiva As Form
'    Dim bolCancellaAzione As Boolean
'    Dim intAzioneEseguita As Integer
'    Dim bolApriNomiCtrl As Boolean
'    Dim bolAllInOne As Boolean
'    Dim DesignButton As MSComctlLib.Button ' DESIGNER
'    Dim frmDesign As Form
'
'    Dim intForm As Integer
'    Dim bolFrmZoomAperta As Boolean
'    Dim objExt As Object
'
'
'    If Button.Enabled Then
'        On Local Error Resume Next
'        Set frmAttiva = metodo.FormAttiva
'        metodo!helpTimer.Interval = 500
'        If TypeName(Button) = "IButtonMenu" Then
'            ' Gestione Zoom su foglio scheda sviluppo n.ro 1128 + Designer
'            If metodo.Barra.Buttons(idxBottoneDefAgenti).Value = tbrUnpressed And LCase$(Left$(Button.Key, 4)) <> "zoom" And LCase$(Left$(Button.Key, 6)) <> "design" Then
'            'If metodo.Barra.Buttons(idxBottoneDefAgenti).Value = tbrUnpressed Then
'                Select Case Button.Index
'                    Case 1 'Def. Agenti
'                        ' Rif. anomalia n.ro 4824 + rif. anomalia 6882
'                        If MXNU.ModuloRegole And Not (MXNU.CtrlAccessi) Then
'                            Call FormImpostaAgenti(frmAttiva)
'                        End If
'                    Case 2 'Nome Controlli - campi
'                        If (metodo.FormsCount > 2) Then
'                            bolApriNomiCtrl = True
'                            Call frmAttiva.AzioniMetodo(MetFMostraCampiDBAnagr, bolApriNomiCtrl)
'                            If bolApriNomiCtrl Then
'                                Set FrmNomiControlli.frmDef = frmAttiva
'                                FrmNomiControlli.Show
'                            End If
'                        End If
'                    Case 3 'Situazione Anagrafica
'                        'mostro le dipendenze dell'anagrafica
'                        Dim ListaCol As New Collection
'                        Call frmAttiva.AzioniMetodo(MetFVisDipendenze, ListaCol)
'                        If ListaCol.Count > 0 Then
'                            Call MXVA.VisualizzaDipendenze(ListaCol)
'                        End If
'                        Set ListaCol = Nothing
'                    Case 4 'Attiva\Disattiva Log
'                        If Button.Tag = "A" Then
'                            Call MXDB.DisattivaLog
'                            Button.Tag = ""
'                            Button.Text = "Attiva File Log"
'                            Call frmLog.MostraFileLog(MXNU.GetTempDir & "MWSqlLog" & MXNU.NTerminale & ".sql")
'                        Else
'                            Button.Tag = "A"
'                            Call MXDB.AttivaLog
'                            Button.Text = "Disattiva File Log"
'                        End If
'                    Case 5 'ricarica profili
'                        metodo.MousePointer = vbHourglass
'                        MXNU.MostraMsgInfo 70058
'                        '>>> profili anagrafiche
'                        If Not MXVA Is Nothing Then
'                            If MXVA.ChiudiDyTRAnagraf() Then
'                                Call MXVA.ApriDyTRAnagraf
'                            End If
'                            If MXVA.ChiudiDyTRValidazione() Then
'                                Call MXVA.ApriDyTRValidazione
'                            End If
'                        End If
'                        '>>> profili tabelle
'                        If Not MXCT Is Nothing Then
'                            If MXCT.ChiudiDyTRTabelle() Then
'                                Call MXCT.ApriDyTRTabelle
'                            End If
'                        End If
'                        '>>> profili visioni/situazioni
'                        If Not MXVI Is Nothing Then
'                            If MXVI.ChiudiDyTRSituazioni() Then
'                                Call MXVI.ApriDyTRSituazioni
'                            End If
'                            If MXVI.ChiudiDyTRVisioni() Then
'                                Call MXVI.ApriDyTRVisioni
'                            End If
'                        End If
'                        MXNU.MostraMsgInfo ""
'                        metodo.MousePointer = vbDefault
'                End Select
'            ElseIf LCase$(Left$(Button.Key, 4)) = "zoom" Then ' MYERP
'                 #If ISMETODOXP = 1 Then
'                    If MXNU.MetodoXP Then
'                        ' Rif. anomalia ISV n.ro 6 / Anomalia Metodo XP n.ro 6084 / Quesito ISV 294
'                        Dim objActiveControl As Object
'                        If LCase(frmAttiva.Name) = "frmextchild" Then
'                            If frmAttiva.ActiveControl.object.Controls.Count > 1 Then
'                                Set objActiveControl = frmAttiva.ActiveControl.object.Controls(1).object.ExtActiveControl
'                            Else
'                                Set objActiveControl = frmAttiva.ActiveControl.object.Controls(0).object.ExtActiveControl
'                            End If
'                        ElseIf TypeName(frmAttiva.ActiveControl) = "MWSchedaBox" Then
'                            ' Sono in una scheda
'                            If frmAttiva.ActiveControl.ControlsEx.Count > 0 Then
'                                ' La scheda contiene delle estensioni
'                                ' Setto lo UserControl dentro il wrapper in objExt (ExtWrapper deve avere una public property get come tutti gli UserControl estensioni)
'                                'Rif. anomalia #8088
'                                Set objExt = frmAttiva.ActiveControl.ControlsEx(0).object.Controls
'                                If TypeName(objExt(0)) = "CTLXBus" Then
'                                    Set objActiveControl = objExt(1).object.ExtActiveControl
'                                Else
'                                    Set objActiveControl = objExt(0).object.ExtActiveControl
'                                End If
'                                Set objExt = Nothing
'                            End If
'                        Else
'                            Set objActiveControl = frmAttiva.ActiveControl
'                        End If
'                        If TypeName(objActiveControl) = "fpSpread" Then
'                            ' Gestione dell'apertura di una singola form di zoom
'                            bolFrmZoomAperta = False
'                            intForm = 0
'                            While (intForm < VB.Forms.Count) And (Not (bolFrmZoomAperta))
'                                If UCase(VB.Forms(intForm).Name) = "FRMZOOM" Then
'                                    bolFrmZoomAperta = True
'                                End If
'                                intForm = intForm + 1
'                            Wend
'                            If Not bolFrmZoomAperta Then
'                                Call frmZoom.DoZoom(frmAttiva, objActiveControl, Button.Index, objActiveControl.ActiveCol, objActiveControl.ActiveRow, MXNU.FrmMetodo.Barra.Buttons(idxBottoneTrova).Enabled)
'                            Else
'                                Call MXNU.MsgBoxEX(2738, vbExclamation, 23481)
'                            End If
'                        End If
'                        Set objActiveControl = Nothing
'                    End If
'                 #End If
'            ElseIf LCase$(Left$(Button.Key, 6)) = "design" Then ' DESIGNER
'                #If ISMETODOXP = 1 Then
'                  If MXNU.MetodoXP Then
'                     Select Case LCase(Button.Key)
'                       Case "design", "designer"
'                         If Not (frmAttiva Is Nothing) Then
'                            ' Rif. anomalia n.ro 6082
'                            If frmAttiva.Name <> "frmVisioni" And frmAttiva.Name <> "frmTabelle" And _
'                            frmAttiva.Name <> "frmModuli98" And frmAttiva.Name <> "metodo" And _
'                             frmAttiva.Name <> "frmToolTip" And frmAttiva.Name <> "frmModuliXP" And frmAttiva.Name <> "frmZoom" Then
'                           'If frmAttiva.Name <> "frmExtChild" And frmAttiva.Name <> "frmTabelle" And frmAttiva.Name <> "frmModuli98" And frmAttiva.Name <> "metodo" And _
'                             frmAttiva.Name <> "frmToolTip" And frmAttiva.Name <> "frmModuliXP" Then
'                                If metodo!tlbDesigner.Visible = False Then
'                                    'Se premuto visualizza la barra del designer (l'utente è sicuramente amministratore)
'                                    metodo!cmbDesignType.AddItem MXNU.CaricaStringaRes(75474)
'                                    metodo!cmbDesignType.AddItem MXNU.CaricaStringaRes(75475)
'                                    metodo!tlbDesigner.Visible = True
'                                    For Each DesignButton In metodo!tlbDesigner.Buttons
'                                        DesignButton.Enabled = False
'                                    Next
'
'                                    'Rimane sempre attivo il bottone "showdiff"
'                                    metodo!tlbDesigner.Buttons.Item("showdiff").Enabled = True
'                                    ' "Spengo" i bottoni ...
'                                    metodo!tlbDesigner.Buttons.Item("showdiff").Value = tbrUnpressed
'                                    metodo!tlbDesigner.Buttons.Item("unvisible").Value = tbrUnpressed
'                                    metodo!tlbDesigner.Buttons.Item("disable").Value = tbrUnpressed
'                                    metodo!cmbDesignType.Enabled = True
'                                    ' Quando attivo la barra dovrò al primo salvataggio scegliere la versione
'                                    Designer.bolSalvato = False
'                                ElseIf metodo!tlbDesigner.Visible = True Then
'                                    ' Nascondo la barra del designer
'                                    If Designer.bolModify Then
'                                        If MXNU.MsgBoxEX(2631, vbYesNo + vbDefaultButton2, 2632) = vbYes Then
'                                            Call Designer.TerminateEditMode
'                                            Call Designer.ExportDesign(False)
'                                        Else
'                                            Call Designer.TerminateEditMode
'                                        End If
'                                        Call Designer.TerminateDesigner
'                                        Designer.bolModify = False
'                                    End If
'                                    metodo!cmbDesignType.Clear
'                                    metodo!tlbDesigner.Visible = False
'
'                                    For Each frmDesign In VB.Forms
'                                        If LCase(frmDesign.Name) = "frmdesignerprop" Then
'                                            Unload (FrmDesignerProp)
'                                        End If
'                                    Next
'                                End If
'                           End If
'                         End If
'                       Case "designdelete"
'                         ' Cancellazione delle versioni
'                         Load frmDeleteVersions
'                         frmDeleteVersions.Show
'                       Case "designversions"
'                         Load frmUsersVersions
'                         frmUsersVersions.Show
'                       Case "designgrid"
'                         Load FrmDesignerGrid
'                         FrmDesignerGrid.Show
'                     End Select
'                  End If
'                #End If
'            End If
'        Else
'            Err.Clear
'            Select Case Button.Index
'                Case idxBottoneModuli
'                    #If ISMETODOXP = 1 Then
'                        If MXNU.MetodoXP Then
'
'                            If frmModuli.StatoMenuAlbero(0) = False Then
'                                'il menu moduli non e' aperto ne lockato quindi lo apro e lo locko
'                                frmModuli.StatoMenuModuli = MnuModATTIVO
'                                Call frmModuli.AlberoBloccatoMenu(True, 0)
'                            Else
'                                'il menu moduli e' gia' aperto e lockato, lo chiudo
'                                'frmModuli.StatoMenuModuli = MnuModNONATTIVO
'                                Call frmModuli.AlberoBloccatoMenu(False, 0)
'                            End If
'                        Else
'                            Call AttivaMenuMetodo
'                        End If
'                    #Else
'                        Call AttivaMenuMetodo
'                    #End If
'
'                Case idxBottoneAllInOne
'                    'If MXNU.MetodoXP Then
'                        bolAllInOne = True
'                        Call frmAttiva.AzioniMetodo(MetFAllInOne, bolAllInOne)
'                        If bolAllInOne Then
'                            Call AllInOneManager(MXNU.FrmMetodo.FormAttiva)
'                        End If
'                    'End If
'
'                Case idxBottoneInserisci
'                    Call EseguiAgenteAzioneForm(frmAttiva, "before_insert", bolCancellaAzione)
'                    If Not bolCancellaAzione Then
'                        Call DoLostFocus(frmAttiva.ActiveControl)
'                        intAzioneEseguita = CInt(frmAttiva.AzioniMetodo(MetFInserisci))
'                        'call frmAttiva.AzioniMetodo(MetFInserisci)
'                        Call EseguiAgenteAzioneForm(frmAttiva, "after_insert", bolCancellaAzione, intAzioneEseguita)
'                    End If
'
'                Case idxBottoneDettaglio
'                    Call DoLostFocus(frmAttiva.ActiveControl)
'                    Call frmAttiva.AzioniMetodo(MetFDettagli)
'
'                Case idxBottoneRegistra
'                    Call EseguiAgenteAzioneForm(frmAttiva, "before_save", bolCancellaAzione)
'                    If Not bolCancellaAzione Then
'                        Call DoLostFocus(frmAttiva.ActiveControl)
'                        intAzioneEseguita = CInt(frmAttiva.AzioniMetodo(MetFRegistra))
'                        Call EseguiAgenteAzioneForm(frmAttiva, "after_save", bolCancellaAzione, intAzioneEseguita)
'                    End If
'
'                Case idxBottoneAnnulla
'                    Call EseguiAgenteAzioneForm(frmAttiva, "before_delete", bolCancellaAzione)
'                    If Not bolCancellaAzione Then
'                        Call DoLostFocus(frmAttiva.ActiveControl)
'                        intAzioneEseguita = CInt(frmAttiva.AzioniMetodo(MetFAnnulla))
'                        'Call frmAttiva.AzioniMetodo(MetFAnnulla)
'                        Call EseguiAgenteAzioneForm(frmAttiva, "after_delete", bolCancellaAzione, intAzioneEseguita)
'                    End If
'
'                Case idxBottonePrimo
'                    Call EseguiAgenteAzioneForm(frmAttiva, "before_first", bolCancellaAzione)
'                    If Not bolCancellaAzione Then
'                        Call DoLostFocus(frmAttiva.ActiveControl)
'                        Call frmAttiva.AzioniMetodo(MetFPrimo)
'                    End If
'                    Call EseguiAgenteAzioneForm(frmAttiva, "after_first", bolCancellaAzione)
'
'                Case idxBottonePrecedente
'                    Call EseguiAgenteAzioneForm(frmAttiva, "before_previous", bolCancellaAzione)
'                    If Not bolCancellaAzione Then
'                        Call DoLostFocus(frmAttiva.ActiveControl)
'                        Call frmAttiva.AzioniMetodo(MetFPrecedente)
'                    End If
'                    Call EseguiAgenteAzioneForm(frmAttiva, "after_previous", bolCancellaAzione)
'
'                Case idxBottoneSuccessivo
'                    Call EseguiAgenteAzioneForm(frmAttiva, "before_next", bolCancellaAzione)
'                    If Not bolCancellaAzione Then
'                        Call DoLostFocus(frmAttiva.ActiveControl)
'                        Call frmAttiva.AzioniMetodo(MetFSuccessivo)
'                    End If
'                    Call EseguiAgenteAzioneForm(frmAttiva, "after_next", bolCancellaAzione)
'                Case idxBottoneUltimo
'                    Call EseguiAgenteAzioneForm(frmAttiva, "before_last", bolCancellaAzione)
'                    If Not bolCancellaAzione Then
'                        Call DoLostFocus(frmAttiva.ActiveControl)
'                        Call frmAttiva.AzioniMetodo(MetFUltimo)
'                    End If
'                    Call EseguiAgenteAzioneForm(frmAttiva, "after_last", bolCancellaAzione)
'                Case idxBottoneStampa
'                    Call frmAttiva.AzioniMetodo(MetFStampa)
'                Case idxBottoneTrova
'                    If Not MXCT.TTrova(FrmTrovaGen) Then
'                        Call frmAttiva.AzioniMetodo(MetFTrova)
'                    End If
'                Case idxBottoneDefAccessi
'                    Call FormDefinisciAccessi(frmAttiva)
'                Case idxBottoneHelp
'                    #If ISMETODOXP = 1 Then
'                        If MXNU.MetodoXP = True Then
'                                #If TOOLS = 1 Then
'                                    If MXNU.DammiFormAttiva.Name = "frmModuli" Then
'                                        Call EseguiAppAssociata(MXNU.PercorsoPgm & "\HELP\TECNICO", "INDICE.HTM")
'                                    Else
'                                        Call MXNU.ApriHelp(False)
'                                    End If
'                                #Else
'                                    'posso caricare l'help direttamente su il browser della form tooltip
'                                    If MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", "METODOW", "HelpCompilato", 0) = 1 Then
'                                        Call MXNU.ApriHelp(False)
'                                    End If
'                                #End If
'                        Else
'                            #If TOOLS = 1 Then
'                                If MXNU.DammiFormAttiva.Name = "frmModuli" Then
'                                    Call EseguiAppAssociata(MXNU.PercorsoPgm & "\HELP\TECNICO", "INDICE.HTM")
'                                Else
'                                    Call MXNU.ApriHelp(False)
'                                End If
'                            #Else
'                                Call MXNU.ApriHelp(False)
'                            #End If
'                        End If
'                    #Else
'                        #If TOOLS = 1 Then
'                            If MXNU.DammiFormAttiva.Name = "frmModuli" Then
'                                Call EseguiAppAssociata(MXNU.PercorsoPgm & "\HELP\TECNICO", "INDICE.HTM")
'                            Else
'                                Call MXNU.ApriHelp(False)
'                            End If
'                        #Else
'                            Call MXNU.ApriHelp(False)
'                        #End If
'                    #End If
'                Case idxBottoneUtenteModifica
'                    Call frmAttiva.AzioniMetodo(MetFVisUtenteModifica)
'                Case idxBottoneDefAgenti 'Def Agenti predefiniti
'                    If (metodo.pBottoniAgenti = ageCollegamenti) Then
'                        'imposto agenti nella form
'                        If MXNU.ModuloRegole And Not (MXNU.CtrlAccessi) Then   ' rif. anomalia n.ro 4824+ n.ro 6882
'                            Call FormImpostaAgenti(frmAttiva)
'                        End If
'                    End If
'
'                Case idxBottoneSchedula
'                    Call frmAttiva.AzioniMetodo(MetFSchedulaOperazione)
'                Case idxBottoneSpostaBarra
'                    metodo!Barra.Align = (metodo!Barra.Align) Mod 2 + 1
'                    metodo!Barra.Buttons(Button.Index).Image = metodo!Barra.Align + 10
'                    metodo!tlbDesigner.Align = (metodo!tlbDesigner.Align) Mod 2 + 1 ' Designer
'                Case idxBottoneZoom
'                    #If ISMETODOXP = 1 Then
'                      If MXNU.MetodoXP Then
'                        Dim intZoomType As Integer
'                        intZoomType = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\Mw.ini", MXNU.UtenteAttivo, "ZoomType", 1)
'                        ' Gestione dell'apertura di una singola form di zoom
'                        bolFrmZoomAperta = False
'                        intForm = 0
'                        While (intForm < VB.Forms.Count) And (Not (bolFrmZoomAperta))
'                            If UCase(VB.Forms(intForm).Name) = "FRMZOOM" Then
'                                bolFrmZoomAperta = True
'                            End If
'                            intForm = intForm + 1
'                        Wend
'
'                        If Not bolFrmZoomAperta Then
'                            ' Rif. anomalia ISV n.ro 6 / Anomalia Metodo XP n.ro 6084 / Quesito ISV 294
'                            If LCase(frmAttiva.Name) = "frmextchild" Then
'                                If frmAttiva.ActiveControl.object.Controls.Count > 1 Then
'                                    Call frmZoom.DoZoom(frmAttiva, frmAttiva.ActiveControl.object.Controls(1).object.ExtActiveControl, intZoomType, frmAttiva.ActiveControl.object.Controls(1).object.ExtActiveControl.ActiveCol, frmAttiva.ActiveControl.object.Controls(1).object.ExtActiveControl.ActiveRow, MXNU.FrmMetodo!Barra.Buttons(idxBottoneTrova).Enabled)
'                                Else
'                                    Call frmZoom.DoZoom(frmAttiva, frmAttiva.ActiveControl.object.Controls(0).object.ExtActiveControl, intZoomType, frmAttiva.ActiveControl.object.Controls(0).object.ExtActiveControl.ActiveCol, frmAttiva.ActiveControl.object.Controls(0).object.ExtActiveControl.ActiveRow, MXNU.FrmMetodo!Barra.Buttons(idxBottoneTrova).Enabled)
'                                End If
'                            ElseIf TypeName(frmAttiva.ActiveControl) = "MWSchedaBox" Then
'                                ' Sono in una scheda
'                                If frmAttiva.ActiveControl.ControlsEx.Count > 0 Then
'                                    ' La scheda contiene delle estensioni
'                                    ' Setto lo UserControl dentro il wrapper in objExt (ExtWrapper deve avere una public property get come tutti gli UserControl estensioni)
'                                    Set objExt = frmAttiva.ActiveControl.ControlsEx(0).object.Controls
'                                    'Rif anomalia #8088
'                                    If TypeName(objExt(0)) = "CTLXBus" Then
'                                        Set objActiveControl = objExt(1).object.ExtActiveControl
'                                    Else
'                                        Set objActiveControl = objExt(0).object.ExtActiveControl
'                                    End If
'                                    Set objExt = Nothing
'                                    Call frmZoom.DoZoom(frmAttiva, objActiveControl, intZoomType, objActiveControl.ActiveCol, objActiveControl.ActiveRow, MXNU.FrmMetodo!Barra.Buttons(idxBottoneTrova).Enabled)
'                                    Set objActiveControl = Nothing
'                                End If
'                            Else
'                                Call frmZoom.DoZoom(frmAttiva, frmAttiva.ActiveControl, intZoomType, frmAttiva.ActiveControl.ActiveCol, frmAttiva.ActiveControl.ActiveRow, MXNU.FrmMetodo!Barra.Buttons(idxBottoneTrova).Enabled)
'                            End If
'                        Else
'                            Call MXNU.MsgBoxEX(2738, vbExclamation, 23481)
'                        End If
'
'                      End If
'                    #End If
'                Case idxBottoneDesigner
'                  #If ISMETODOXP = 1 Then
'                    If MXNU.MetodoXP Then
'                      If Not (frmAttiva Is Nothing) Then
'                        ' Rif. anomalia n.ro 6082
'                        If frmAttiva.Name <> "frmVisioni" And frmAttiva.Name <> "frmTabelle" And frmAttiva.Name <> "frmModuli98" And _
'                            frmAttiva.Name <> "metodo" And frmAttiva.Name <> "frmZoom" And _
'                            frmAttiva.Name <> "frmToolTip" And frmAttiva.Name <> "frmMetodoXP" And frmAttiva.Name <> "frmModuliXP" Then
'                            If metodo!tlbDesigner.Visible = False Then
'                                'Se premuto visualizza la barra del designer (utente amministratore)
'                                metodo!cmbDesignType.AddItem MXNU.CaricaStringaRes(75474)
'                                metodo!cmbDesignType.AddItem MXNU.CaricaStringaRes(75475)
'                                metodo!tlbDesigner.Visible = True
'                                For Each DesignButton In metodo!tlbDesigner.Buttons
'                                    DesignButton.Enabled = False
'                                Next
'                                'Rimane sempre attivo il bottone "showdiff" (e non premuto)
'                                metodo!tlbDesigner.Buttons.Item("showdiff").Enabled = True
'                                metodo!tlbDesigner.Buttons.Item("showdiff").Value = tbrUnpressed
'
'                                metodo!cmbDesignType.Enabled = True
'                                ' Quando attivo la barra dovrò al primo salvataggio scegliere la versione
'                                Designer.bolSalvato = False
'                            ElseIf metodo!tlbDesigner.Visible = True Then
'                                'Nascondo la barra del designer
'                                If Designer.bolModify Then
'                                    If MXNU.MsgBoxEX(2631, vbYesNo + vbDefaultButton2, 2632) = vbYes Then
'                                        Call Designer.TerminateEditMode
'                                        Call Designer.ExportDesign(False)
'                                    Else
'                                        Call Designer.TerminateEditMode
'                                    End If
'                                    Call Designer.TerminateDesigner
'                                    Designer.bolModify = False
'                                End If
'                                metodo!cmbDesignType.Clear
'                                metodo!tlbDesigner.Visible = False
'                                For Each frmDesign In VB.Forms
'                                    If LCase(frmDesign.Name) = "frmdesignerprop" Then
'                                        Unload (FrmDesignerProp)
'                                    End If
'                                Next
'                            End If
'                        End If
'                      End If
'                    End If
'                  #End If
'               Case 30
'                  Call ShellExecute(metodo.hwnd, "Open", "www.metodo.it/MetodoPortal/DesktopPortalMain.aspx?action=NEWS", "", App.Path, 1)
'            End Select
'        End If
'        Set frmAttiva = Nothing
'    End If
'
'
'End Sub
#End If

Function GetUltimoAnnoUsato() As String
    Dim riga$
    riga$ = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", "TRM" & MXNU.NTerminale, MXNU.DittaAttiva, MXNU.AnnoDflt)
    If riga$ <> "" Then
    GetUltimoAnnoUsato$ = riga
    Else
    GetUltimoAnnoUsato$ = ""
    End If
    'Dim hSS%, q%, sql$
    'sql = "SELECT * FROM TabUsoDitte WHERE ((NumTerminale =" & Val(NTerminale$) & ") AND (Ditta ='" & MXNU.DittaAttiva & "'))"
    'hSS = MXDB.dbCreaSS(hndDBDitte, sql, TIPO_TABELLA)
    'If MXDB.dbFineTab(hSS, TIPO_SNAPSHOT) Then
    '   GetUltimoAnnoUsato$ = ""
    'Else
    '    GetUltimoAnnoUsato$ = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "Anno", "")
    'End If
    'q = MXDB.dbChiudiSS(hSS)
End Function
Function LeggiDscBreve(strDitta As String) As String
    LeggiDscBreve = LeggiDscBreve_(strDitta, "")
End Function

Function LeggiDscBreve_(Ditta As String, dflt As String) As String
Dim intq As Integer
Dim hSS As MXKit.CRecordSet
Dim strsql As String

'    Select Case intMetExe
'        Case EXE_METODO95
            strsql = "SELECT DataCostituzione,DesBreve FROM  TabDitte"
'        Case EXE_RITENUTE, EXE_CESPITI
'            sqlt = "SELECT DesBreve FROM  AnagraficaDitte WHERE Ditta ='" + Ditta + "'"
'    End Select
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
    intq = Not MXDB.dbFineTab(hSS, TIPO_SNAPSHOT)
    If intq Then
        LeggiDscBreve_ = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "DesBreve", dflt)
        MXNU.DataCostituzione = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "DataCostituzione", MXNU.Default_Data)
    Else
        LeggiDscBreve_ = dflt
    End If
    intq = MXDB.dbChiudiSS(hSS)
End Function

'Sub Main()
'    Dim intRis As Boolean
'    Dim strUtente As String
'    Dim bolConfInt As Boolean
'    Dim strLineErr As String
'    Dim strDateLayout As String
'
'    On Local Error GoTo err_Main
'
'    InitCommonControls   'Necessario per il corretto rendering degli oggetti in 3D col Manifest
'
'    strLineErr = "Creazione Oggetto Nucleo"
'    Set MXNU = New MXNucleo.XNucleo
'
'    Load frmIntro
'    '>>> INIZIALIZZAZIONE NUCLEO
'    #If TOOLS = 1 Then
'        MXNU.FileDatLocali = False
'    #End If
'    strLineErr = "Inizializzazione Oggetto Nucleo"
'
'    strinitexe = Command$
'    If (Not MXNU.Inizializza(Command$, strUtente)) Then GoSub err_objInit
'
'    ' Inizializzazione del charset per la gestione delle lingue (polacco, slovacco, ecc.)
'    Call InitCharSet(MXNU.PercorsoPgm & "\" & "CharSetManager.config")
'
'    ' Spostato i controlli su carattere separatore dei decimali e del controllo della data
'    '<rif anomalia 816 RZ>
'    If QueryValue(HKEY_CURRENT_USER, "Control Panel\International", "sMonDecimalSep") <> "," Then 'And MXNU.LinguaAttiva = "IT"
'        Call MXNU.MsgBoxEX("Impostare il carattere ',' (Virgola) nelle Impostazioni Internazionali - Valuta - Separatore Decimali da Pannello di Controllo !", vbOKOnly + vbCritical, 1007)
'        GoTo err_objInit
'    End If
'
'    strDateLayout = QueryValue(HKEY_CURRENT_USER, "Control Panel\International", "sShortDate")
'    If Len(strDateLayout) < 10 Then
'        Call MXNU.MsgBoxEX(3065, vbOKOnly + vbCritical, 1007, strDateLayout)
'        GoTo err_objInit
'    End If
'
'    '<rif anomalia 9303 RZ>
'    If Mid(strDateLayout, 3, 1) <> "/" Then
'        Call MXNU.MsgBoxEX("Impostare la data nel formato dd/MM/yyyy", vbOKOnly + vbCritical, 1007, strDateLayout)
'        GoTo err_objInit
'    End If
'
'
'    MXNU.VersioneMetodo = App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "00")
'    MXNU.EXEName = App.EXEName
'
'    '[16/06/2011] Rimozione Chiave Hardware
''    '>>> RICERCA CHIAVE HARDWARE
''    strLineErr = "Ricerca Chiave Hardware"
''    frmIntro.MostraMessaggioOperazione (9008)
'
''>>>>>>>>>>> SPOSTATO SU MDIFORM_LOAD, ALTRIMENTI I COLORI DEL GRADIENTE NON SI INIZIALIZZANO CORRETTAMENTE <<<<<<<<<<<<<<<<<<<<<
''    'imposto la variabile di controllo ISMETODO2005 sui controlli
''    strLineErr = "Inizializzazione sistema Metodo 2005"
''    Dim oM2005 As MXCtrl.M2005Setup
''    Set oM2005 = New MXCtrl.M2005Setup
''#If ISMETODO2005 = 1 Then
''    oM2005.ISMETODO2005 = True
''#Else
''    oM2005.ISMETODO2005 = False
''#End If
''    Set oM2005 = Nothing
''>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'
'    strLineErr = "Inizializzazione Oggetto Spread"
''>>>>>>>>>>> SPOSTATO SU MDIFORM_LOAD, VEDI SOPRA <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
''#If ISMETODO2005 = 1 Then
''    Call InizializzaSpread(MXNU.MetodoXP, True, SysGradientColor1)
''#Else
'    Call InizializzaSpread(MXNU.MetodoXP)
''#End If
'
'#If ISKEY Then
'    funzione1
'#End If
'    strLineErr = "Creazione Oggetti KIT - BUSINESSS"
'
'    '>>> CREAZIONE OGGETTI KIT - BUSINESS
'#If SOLOKIT = 1 Then
'    If Not CreateObjKitBus(frmIntro.CTLXKit1, Nothing) Then GoSub err_objInit
'#Else
'    If Not CreateObjKitBus(frmIntro.CTLXKit1, frmIntro.CTLXBus1) Then GoSub err_objInit
'#End If
'    strLineErr = "Inizializzazione Libreria Database"
'
'    '>>> INIZIALIZZAZIONE LIBRERIA DATABASE
'    If Not (MXDB.dbInizializza(MXNU)) Then
'        Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Libreria ODBC"))
'        GoSub err_objInit
'    End If
'
'    strLineErr = "Apertura Ditta"
'
'    '>>> APERTURA DITTA
'    #If ISKEY <> 1 Then
'        If Not ApriDitta(MXNU.UtenteDB, MXNU.PasswordDB) Then
'            Call DropObjKitBus
'            GoSub err_objInit
'        End If
'    #End If
'
'    strLineErr = "Inizializzazione Oggetti KIT - BUSINESS"
'    '>>> INIZIALIZZAZIONE OGGETTI KIT - BUSINESS
'    If Not InitObjKitBus(hndDBArchivi) Then
'        Call DropObjKitBus
'        GoSub err_objInit
'    End If
'
'    strLineErr = "Apertura Anno"
'    '>>> SELEZIONE ANNO
'    #If ISKEY <> 1 Then
'        Dim NuovoAnno As Integer
'        If SelezioneAnno(False, NuovoAnno) Then
'            MXNU.AnnoAttivo = NuovoAnno
'            Call ApriAnno(False)
'        Else
'            Call DropObjKitBus
'            GoSub err_objInit
'        End If
'    #Else
'
'    #End If
'
'    strLineErr = "Connessione MetServer"
'    #If USAM98SERVER Then
'    Call ConnectM98Server(GobjM98Server)
'    #End If
'
'    strLineErr = "Supporto Script"
'    'supporto Script
'    AddAmbienti2Script
'
'    strLineErr = "Inizializzazione Designer"
'    '*** DESIGNER ***
'    ' controlla il flag in MW.INI
'    ' inizio rif.sch. A5714
'    Dim strEntryMw As String
'    Dim bolDesignerAbilitato As Boolean
'    bolDesignerAbilitato = True
'    strEntryMw = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\MW.INI", "METODOW", "ATTIVADESIGNER", "")
'    If (Len(strEntryMw) > 0) Then
'        bolDesignerAbilitato = (StrComp(strEntryMw, "DESIGNER", vbTextCompare) = 0)
'    End If
'    If bolDesignerAbilitato Then
'        'Controllo esistenza modulo runtime nella chiave
'        If ((MXNU.ControlloModuliChiave(modMyErpRunTime) = 0) _
'            Or (MXNU.ControlloModuliChiave(modMetodoXPEvolution) = 0)) Then 'modMyErpRunTime = 118
'
'            Set Designer = New MXDesigner.cDesigner
'            Set Designer.MyMXNU = MXNU
'            If Not (Designer.Attiva) Then
'                Set Designer = Nothing
'            End If
'        Else
'            Set Designer = Nothing
'        End If
'    Else
'        Set Designer = Nothing
'    End If
'    ' fine rif.sch. A5714
'
'    '>>>inizializzazione sistema di messaggistica
'#If IsMetodo2005 = 1 Then
'    On Local Error GoTo err_Main_NoBlock
'
'    'inizializzazione hosting metodo
'    strLineErr = "Inizializzazione Metodo Interop"
'    Set mMetodoInterop = New CMetodoInterop
'
'    'innizializzazione browser
'    strLineErr = "Inizializzazione Metodo Browser"
'    Set mMetodoBrowser = CreateObject("MxBrowser.CBrowserEngine")
'    If (Not mMetodoBrowser Is Nothing) Then
'            If (Not mMetodoBrowser.Initialize(mMetodoInterop, metodo.DockingPaneManager)) Then
'                    Set mMetodoBrowser = Nothing
'            End If
'    End If
'
'    'gestione messaggistica da file ini <rif anomalia #8620 RZ>
'    Dim msgingInstalled As String
'    msgingInstalled = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", MXNU.UtenteSistema, "Messaggistica", "")
'    If msgingInstalled = Empty Then
'        msgingInstalled = MXNU.LeggiProfilo(MXNU.PercorsoPreferenze & "\mw.ini", "METODOW", "Messaggistica", 1)
'    End If
'
'    If msgingInstalled > 0 Then
'        strLineErr = "Inizializzazione sistema di messaggistica"
'        On Local Error Resume Next
'        Set mMessagingEngine = CreateObject("MxMailing.MailingEngine")
'        Dim myerr As String
'        Dim myerrdescription As String
'        Dim myerrconfig As String
'        myerr = Err.Number
'        myerrdescription = Err.Description
'        On Local Error GoTo 0
'        If mMessagingEngine Is Nothing And Err = 429 Then 'il sistema di messaggistica non è correttamente registrato
'            Call MsgBox("Errore [" & myerr & " " & myerrdescription & "] nella funzione [main] durante l'operazione " & strLineErr, vbCritical, "Attenzione!")
'        Else
'            If (mMessagingEngine.Initialize(MXNU, MXNU.UtenteSistema, MXNU.PercorsoPgm, myerrconfig)) Then
'                Call mMessagingEngine.InitializeActionExecutor(Ambienti2Collection(False), hndDBArchivi)
'            End If
'            If myerrconfig <> Empty Then
'                Call MXNU.MsgBoxEX(myerrconfig, vbCritical, "Metodo Evolus Messaging")
'            End If
'        End If
'    End If
'
'    On Local Error GoTo err_Main
'#End If
'
'    strLineErr = "Load Metodo"
'    '>>> APERTURA METODO
'    Load metodo
'
'    #If ISMETODOXP = 1 Then
'        If MXNU.MetodoXP Then
'            metodo.BarraStato.Panels("Designer").Text = MXNU.VersioneAttiva
'        End If
'    #End If
'
'    'Sviluppo 2033: Agente Fisso in partenza di Metodo
'    If MXAA.AgentiFissi Then
'        Dim objAgt As MXKit.CAgenteAuto
'        Set objAgt = MXAA.CreaCAgenteAuto()
'
'        Call MXAA.EseguiAgt(objAgt, "FISSI\STARTUPMETODO")
'
'        Set objAgt = Nothing
'    End If
'
'    'Anomalia 8516 - Spostato chiamata ad AttivaMenu dopo il load della MDI
'    #If IsMetodo2005 = 1 Then
'        frmModuli.ModuloAttivo = "*DA_INI*"    'Leggo il modulo attivo dall'mw.ini. Vedi funzione AttivaMenu su MetodoXP
'    #Else
'        frmModuli.AttivaMenu
'    #End If
'
'fine_Main:
'    On Local Error GoTo 0
'    Exit Sub
'
'err_Main:
'
'
'    If Not (MXNU Is Nothing) Then
'        Call MsgBox("Errore [" & Err.Number & " " & Err.Description & "] nella funzione [main] durante l'operazione " & strLineErr, vbCritical, "Attenzione!")
'        Set MXNU = Nothing
'    Else
'        Call MsgBox("Errore [" & Err.Number & " " & Err.Description & "] nella funzione [main]", vbCritical, "Attenzione!")
'    End If
'    On Local Error Resume Next
'    Unload frmIntro
'    Resume fine_Main
'    Resume
'err_objInit:
'    If Not (MXNU Is Nothing) Then Set MXNU = Nothing
'    On Local Error Resume Next
'    Unload frmIntro
'    End
'
'err_Main_NoBlock:
'    If Not (MXNU Is Nothing) Then
'        Call MsgBox("Errore [" & Err.Number & " " & Err.Description & "] nella funzione [main] durante l'operazione " & strLineErr, vbCritical, "Attenzione!")
'        'Set MXNU = Nothing   'Remmato altrimenti và in errore successivamente bloccando l'esecuzione di Metodo
'    Else
'        Call MsgBox("Errore [" & Err.Number & " " & Err.Description & "] nella funzione [main]", vbCritical, "Attenzione!")
'    End If
'    Resume Next
'
'End Sub

'(rif 10)
Function FormaNumeroDoc$(ByVal anno%, ByVal NrDoc&, Bis$)
    FormaNumeroDoc$ = Format$(anno, "0000") & "/" & Format$(CStr(NrDoc&), "00000000") & "/" & UCase$(Bis$)
End Function

Sub ScomponiNumeroDoc(AnnoNrBis$, anno%, NrDoc&, Bis$)
    anno% = Val(Left$(AnnoNrBis$, 4))
    NrDoc& = Val(Mid(AnnoNrBis$, 6, 8))
    Bis$ = Right$(AnnoNrBis$, 1)
End Sub


'Restituisce le quantità LIFO dell'Articolo passato
Function LeggiGiacenzaInizialeLIFO(CodArt As String) As Currency
Dim q As Integer
Dim Sql As String
Dim hndSS As CRecordSet
Dim GILIFO As Currency

    Sql = "SELECT Quantita1,Quantita2,Quantita3,Quantita4,Quantita5,Quantita6,Quantita7,Quantita8,Quantita9,Quantita10,Quantita11,Quantita12,Quantita13,Quantita14,Quantita15 FROM LIFOArticoli WHERE (CodiceArt = '" & CodArt$ & "')"
    Set hndSS = MXDB.dbCreaSS(hndDBArchivi, Sql$, TIPO_TABELLA)
    If MXDB.dbFineTab(hndSS, TIPO_SNAPSHOT) Then
        GILIFO@ = 0
    Else
        For q% = 1 To 15
            GILIFO@ = GILIFO@ + MXDB.dbGetCampo(hndSS, TIPO_SNAPSHOT, "Quantita" & q%, 0@)
        Next q%
    End If
    q% = MXDB.dbChiudiSS(hndSS)
    LeggiGiacenzaInizialeLIFO = GILIFO@
End Function

'Calcola il Valore LIFO TOTALE dell'Articolo Passato
'PARAMETRI:
'   CodArt$     Codice Articolo
'   Qta@        Quantità Carico - Quantità Scarico dell'Articolo
'   PMA@        Valore Medio (=Valore Carico/Qtà Carico)
'   LIFOTot@    Valore LIFO Totale
'   LIFOUn@     Valore LIFO Unitario
Sub ValorizzazioneLIFO(CodArt$, ByVal qta@, ByVal PMA@, LIFOTot@, LIFOUn@)
    Dim hSS As CRecordSet, q%, Sql$, qt@
    Sql$ = "SELECT * FROM LIFOArticoli WHERE (CodiceArt = '" & CodArt$ & "')"
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, Sql, TIPO_TABELLA)
    LIFOTot@ = 0@
    If MXDB.dbFineTab(hSS, TIPO_SNAPSHOT) Then
        LIFOTot@ = qta@ * PMA@
    Else
        Dim i%, qtl@, vl@, qtLifo@
        qtLifo@ = 0
        For i% = 1 To 15
            qtLifo@ = qtLifo@ + MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "Quantita" & i, 0@)
        Next i%
        qt@ = qta@ + qtLifo@

        For i = 15 To 1 Step -1
            qtl@ = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "Quantita" & i, 0@)
            vl@ = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "ValoreLIFO" & i, 0@)
            If qtl < qt Then
                LIFOTot@ = LIFOTot@ + qtl * vl@
                qt@ = qt@ - qtl@
            Else
                LIFOTot@ = LIFOTot@ + (qt@ * vl@)
                qt@ = 0@
            Exit For
            End If
        Next i
        If qt@ > 0@ Then LIFOTot@ = LIFOTot@ + (qt@ * PMA@)
    End If
    q% = MXDB.dbChiudiSS(hSS)
    qtLifo@ = (qta@ + qtLifo@)
    If qtLifo@ = 0 Then
        LIFOUn@ = 0
    Else
        LIFOUn@ = LIFOTot@ / qtLifo@
    End If
End Sub

Function CaricaDescrLingue(Foglio As Object, lngPrimaColonna As Long)
    Dim lngCodice As Long, hSSLingue As CRecordSet, varDescrizione As Variant
    Dim intq As Integer, strsql As String

    strsql = "SELECT Codice,Descrizione FROM TabLingue WHERE Codice > 0"
    Set hSSLingue = MXDB.dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
    intq = Not MXDB.dbFineTab(hSSLingue, TIPO_SNAPSHOT)
    Do While intq
        lngCodice = MXDB.dbGetCampo(hSSLingue, TIPO_SNAPSHOT, "Codice", 0)
        varDescrizione = MXDB.dbGetCampo(hSSLingue, TIPO_SNAPSHOT, "Descrizione", 0)
        intq = MXDB.dbSuccessivo(hSSLingue, TIPO_SNAPSHOT)

        Call ssButtonSetPicture(Foglio, lngCodice + lngPrimaColonna, 0, Nothing, varDescrizione)
    Loop
    intq = MXDB.dbChiudiSS(hSSLingue)

End Function

'Stampa il contenuto di un foglio su stampante.
'Parametri:
'   foglio = foglio da stampare
'   PrtColHeaders = True per stampare l'intestazione di colonna,False altrimenti.
'   PrtRowHeaders = True per stampare l'intestazione di riga,False altrimenti.
'   TitoloStampa = Stringa da usare come titolo per la coda di stampa.
'   TitoloTesta = Stringa da usare come testo da stampare nell'intestazione della pagina.
'Sub StampaFoglio(Foglio As FPSpreadADO.fpSpread, ByVal PrtColHeaders%, ByVal PrtRowHeaders%, ByVal TitoloStampa$, TitoloTesta$)
'    Dim objCRW As MXKit.CCrw
'    Set objCRW = MXCREP.CreaCCrw()
'    objCRW.ClearOpzioniStp
'    objCRW.OpzioniForm = 0&
'    objCRW.MostraFrmStampa
'    If Not objCRW.Stampa_Annullata Then
'        If objCRW.Periferica = "Stampante" Then
'
'            'Foglio.hDCPrinter = objCRW.Stampante.hdc
'            Foglio.hDCPrinter = 0  'Rif. Anomalie98 Nr. 2267: altrimenti non stampa niente; rimane un solo problema: si stampa solo sulla stampante di default
'            If objCRW.Stampante.nMinPage <> 0 And objCRW.Stampante.nMaxPage <> 0 Then
'                Foglio.PrintPageStart = objCRW.Stampante.nMinPage
'                Foglio.PrintPageEnd = objCRW.Stampante.nMaxPage
'                Foglio.PrintType = PrintTypePageRange
'            Else
'                Foglio.PrintType = PrintTypeAll
'            End If
'            Foglio.PrintBorder = True
'            Foglio.PrintColHeaders = PrtColHeaders
'            Foglio.PrintRowHeaders = PrtRowHeaders
'            Foglio.PrintColor = True
'            Foglio.PrintHeader = "/fn""Arial"" /fz""10"" /fb1   " & Trim$(MXNU.Dsc_Breve_Ditta) & " - " & MXNU.AnnoAttivo & "        " & TitoloTesta & "  /n /fb0 " & Now
'            Foglio.PrintFooter = "/fn""Arial"" /fz""10"" /fb0 /r Pag. /p"
'            Foglio.PrintGrid = False
'            Foglio.PrintJobName = TitoloStampa
'            Foglio.PrintShadows = True
'            Foglio.PrintUseDataMax = True
'
'            metodo.MousePointer = vbHourglass
'            Foglio.Action = ActionPrint
'            metodo.MousePointer = vbDefault
'        End If
'    End If
'    Set objCRW = Nothing
'End Sub



'Sub CaricaTabella(strNomeDbTabella As String, vntDesTabella As Variant, lngHelp As Long, Optional ChiaveAgg As Variant, Optional WHEAgg As Variant)
'    Dim intq As Boolean, FrmTab As New frmTabelle, lngHwndForm As Long, ints As Integer
'    Dim strChiaveAgg As String
'    Dim StrWheAgg As String
'    Dim lngDesTabella As Long ' per compatibilità
'
'    FrmTab.NOMETABELLA = strNomeDbTabella
'    If IsNumeric(vntDesTabella) Then
'        vntDesTabella = "{" & Val(vntDesTabella) & "}"
'    End If
'    FrmTab.DesTabella = MXNU.CaricaCaptionInLingua(vntDesTabella)
'    FrmTab.MlngHlpTabella = lngHelp
'    If IsMissing(ChiaveAgg) Then
'        strChiaveAgg = ""
'    Else
'        strChiaveAgg = ChiaveAgg
'    End If
'    If IsMissing(WHEAgg) Then
'        StrWheAgg = ""
'    Else
'        StrWheAgg = WHEAgg
'    End If
'
'    FrmTab.ChiaveAgg = strChiaveAgg
'    FrmTab.StrWheAgg = StrWheAgg
'
'    lngHwndForm = MXCT.TabellaCaricata(strNomeDbTabella)
'    If lngHwndForm > 0 Then
'        For ints = 0 To Forms.Count - 1
'            If Forms(ints).hwnd = lngHwndForm Then
'                Forms(ints).WindowState = vbNormal
'                On Local Error Resume Next
'                Forms(ints).SetFocus
'                On Local Error GoTo 0
'                Exit For
'            End If
'        Next
'        Exit Sub
'    End If
'
'    FrmTab.Show
'
'
'End Sub

'Sub RiallineaProgressivi()
'    Dim q%, Sql$, tB$, num As Variant
'    Dim hDY As MXKit.CRecordSet
'    Dim hdtb As MXKit.CRecordSet
'    Dim strMsg As String
'    Dim strDes As String
'
'    metodo.MousePointer = vbHourglass
'    On Local Error GoTo RPT_Err
'
'    Sql = "SELECT * FROM TabProgressivi"
'    Set hDY = MXDB.dbCreaSS(hndDBArchivi, Sql, TIPO_TABELLA)
'    q = Not MXDB.dbFineTab(hDY, TIPO_DYNASET)
'    strMsg = MXNU.CaricaStringaRes(2262) & Chr(10)
'    Do While q
'        tB$ = MXDB.dbGetCampo(hDY, TIPO_DYNASET, "NomeTabella", "")
'        strDes = MXNU.CaricaCaptionInLingua(MXNU.LeggiProfilo(MXNU.File_ini_comune, "DESCRIZIONETABELLE", tB, ""))
'        If strDes <> "" Then
'            strMsg = strMsg & strDes & ", "
'        End If
'        q = MXDB.dbSuccessivo(hDY, TIPO_DYNASET)
'    Loop
'    q = MXDB.dbChiudiSS(hDY)
'    strMsg = Left(strMsg, Len(strMsg) - 2)
'
'    If MXNU.MsgBoxEX(strMsg, vbInformation + vbYesNo, 1007) = vbYes Then
'        Set hDY = MXDB.dbCreaSS(hndDBArchivi, Sql, TIPO_TABELLA)
'        q = Not MXDB.dbFineTab(hDY, TIPO_DYNASET)
'        MXDB.dbBeginTrans hndDBArchivi
'        Do While q
'            tB$ = MXDB.dbGetCampo(hDY, TIPO_DYNASET, "NomeTabella", "")
'            If tB <> "" Then
'                If (Left(UCase(tB), 13) = "PACKINGCFGIMB") Or (UCase(tB) = "CONTRATTI_RCFV") Then
'                    Sql = ""
'                Else
'                    Select Case UCase(tB)
'                        Case "GESTIONEPREZZIRIGHE", "GESTIONEPREZZIRIGHETRASF"
'                            Sql = "SELECT MAX(IDRiga) FROM " & tB
'                        Case "TESTEENASARCO"
'                            Sql = "SELECT MAX(NrBozza) FROM " & tB
'                        Case "IDRIFMOVMANCDC"
'                            'tB = "MovimentiCDC"
'                            Sql = "SELECT MAX(IdRiferimento) FROM MovimentiCDC WHERE TipoMov=0"
'                        Case "M98ARTICOLIPROD"      ' inserito per l'estensione di CALCE DEL BRENTA
'                            Sql = "SELECT MAX(ID) FROM " & tB
'                        Case Else
'                            Sql = "SELECT MAX(Progressivo) FROM " & tB
'                    End Select
'                End If
'                If Sql <> "" Then
'                    Set hdtb = MXDB.dbCreaSS(hndDBArchivi, Sql, TIPO_TABELLA)
'                    num = MXDB.dbGetCampo(hdtb, TIPO_DYNASET, 0, 0)
'                    q = MXDB.dbChiudiSS(hdtb)
'                    Call MXDB.dbEseguiSQL(hndDBArchivi, "UPDATE TabProgressivi SET Progr=" & num & " WHERE NomeTabella=" & hndDBArchivi.FormatoSQL(tB, DB_TEXT))
'                End If
'            End If
'            q = MXDB.dbSuccessivo(hDY, TIPO_DYNASET)
'        Loop
'        MXDB.dbCommitTrans hndDBArchivi
'        Call MXNU.MsgBoxEX(2121, vbInformation, 1007)
'    End If
'fine_rpt:
'    metodo.MousePointer = vbDefault
'    On Local Error GoTo 0
'    q = MXDB.dbChiudiSS(hdtb)
'    q = MXDB.dbChiudiSS(hDY)
'    Exit Sub
'
'RPT_Err:
'    MXDB.dbRollBack hndDBArchivi
'    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("RiallineaProgressivi", Err.Number, Err.Description))
'    Resume fine_rpt
'
'End Sub

'#If IsMetodo2005 <> 1 Then
'Private Sub AbDisButtonMenu(Index As Integer, bolValue As Boolean)
'    Dim i As Integer
'
'    For i = 1 To metodo.Barra.Buttons(idxBottoneDefAgenti).ButtonMenus.Count
'        If i = Index Then
'            If bolValue Then
'                metodo.Barra.Buttons(idxBottoneDefAgenti).ButtonMenus(i).Text = "> " & metodo.Barra.Buttons(idxBottoneDefAgenti).ButtonMenus(i).Text
'            Else
'                metodo.Barra.Buttons(idxBottoneDefAgenti).ButtonMenus(i).Text = Mid(metodo.Barra.Buttons(idxBottoneDefAgenti).ButtonMenus(i).Text, 3)
'            End If
'        End If
'    Next i
'
'End Sub
'#End If

Public Sub Anagrafica_File(enmTipoOp As setTipoOpAn, MForm As Object, MAnagrafica As MXKit.Anagrafica, strNomeFileDef As String)
    Dim p As New PropertyBag
    Dim hf As Long
    Dim contenuto As Variant

    hf = FreeFile
    Open MXNU.PercorsoPgm & "\DEFCAMPI\" & MAnagrafica.NomeFileDef & ".DAT" For Binary As hf
    If enmTipoOp = enmCompila Then
        'Call mAnagrafica.LeggiDefAnagrafica(mForm)
        Call p.WriteProperty(strNomeFileDef, MAnagrafica)
        contenuto = p.Contents
        Put hf, 1, contenuto
    Else
        Get hf, 1, contenuto
        p.Contents = contenuto
        Set MAnagrafica = Nothing
        Set MAnagrafica = p.ReadProperty(strNomeFileDef)
        Set MAnagrafica.FormAnagr = MForm
    End If
    Close hf
    Set p = Nothing

End Sub

Private Function EseguiAgenteAzioneForm(frm As Form, ByVal strAzione As String, bolCancellaAzione As Boolean, Optional ByVal Param As Variant) As Boolean
    bolCancellaAzione = False
    EseguiAgenteAzioneForm = True
    On Local Error Resume Next
    With frm.MWAgt1
        If Err = 0 And .CatturaEventiForm Then
            If Not IsMissing(Param) Then .Param = Param
            EseguiAgenteAzioneForm = .Esegui(frm, strAzione)
            bolCancellaAzione = .Cancel
        Else
            Err.Clear
        End If
    End With
    On Local Error GoTo 0
End Function

'=========================================================================================================
'           INIZIO Gestione Accessi
'Sub MostraMessaggioAccessi(vntStringID As Variant, Optional vntParam As Variant)
'    If (MXNU.FrmMetodo.LoadingMetodo) Then
'        'sto caricando metodo -> redireziono il messaggio sulla form intro
'        Call frmIntro.MostraMessaggioOperazione(vntStringID, vntParam)
'    Else
'        'metodo già caricato -> redireziono il messaggio sulla status bar
'        Call MXNU.MostraMsgInfo(vntStringID, vntParam)
'    End If
'    DoEvents
'End Sub

'Public Function FormDefinisciAccessi(frmMyDef As Form) As Long
'    If Not (frmMyDef Is Nothing) Then
'        'definizione accessi form
'        Set frmDefAcc.frmDef = frmMyDef
'        frmDefAcc.Show vbModal
'    End If
'End Function
'       FINE Gestione Accessi
'=========================================================================================================

Private Function SostituisciSegnaposto(ByVal strPar$) As String
    Dim i&, strStringa$, strCar$, bolTrovataGraffa As Boolean
    bolTrovataGraffa = False
    For i = 1 To Len(strPar)
        strCar = Mid(strPar, i, 1)
        Select Case strCar
            Case ";"
                If Not bolTrovataGraffa Then
                    strStringa = strStringa & vbNullChar
                Else
                    strStringa = strStringa & strCar
                End If
            Case "{"
                bolTrovataGraffa = True
                strStringa = strStringa & strCar
            Case "}"
                bolTrovataGraffa = False
                strStringa = strStringa & strCar
            Case Else
                strStringa = strStringa & strCar
        End Select
    Next i
    SostituisciSegnaposto = strStringa
End Function

'RIF.A#9621 - gestione controllo attivo considerando le estensioni
Private Function GetActiveControl(FormAttiva As Form) As Object
Dim objActive As Object
Dim strName As String

    On Local Error Resume Next
    Set objActive = FormAttiva.ActiveControl
    strName = objActive.Name
    'è una estensione o il wrapper estensioni?
    If (strName = OGGETTO_WRAPPER_ESTENSIONE) Then
        Set objActive = objActive.object.Controls(OGGETTO_ESTENSIONE).ExtActiveControl
    ElseIf (strName = OGGETTO_ESTENSIONE) Then
        Set objActive = objActive.object.ExtActiveControl
    End If
END_GetActiveControl:
    Set GetActiveControl = objActive
    Exit Function

End Function



'RIF.A#4568 - aggiunto parametro ditta origine
Sub CopiaTabella(holddb As MXKit.CConnessione, hnewdb As MXKit.CConnessione, Sql$, strBlob$, NOMETABELLA$, ByVal commitparziale As Boolean, strDescrizione As String, Prc As Object, testo As Object, LungLog As Long, Optional ByVal strDittaOrigine As String = "")
    Dim hDYOld As CRecordSet, hDYNew As CRecordSet, hndSS As CRecordSet, numcampi%, nrec&, Rec&, NomeTab$, nometab1$
    Dim q%, i%, Valore As Variant, nco$, ncn$, Azz$, s$
    Dim intrans%
    Dim strNomeCampo As String
    Dim bolCampiMod As Boolean
    Dim bolOldCampiMod As Boolean
    Dim strQuery As String

    If Not (Prc Is Nothing) Then Prc.Value = 0
    s = LTrim$(Mid$(Sql, InStr(UCase$(Sql), "FROM") + 4))
    If InStr(s, " ") Then
        nometab1 = Left$(s, InStr(s, " ") - 1)
    Else
        nometab1 = s
    End If
    If NOMETABELLA = "" Then
        NomeTab = nometab1
    Else
        NomeTab = NOMETABELLA
    End If
    If Not (testo Is Nothing) Then
        testo.Caption = MXNU.CaricaStringaRes(1808, strDescrizione)
    End If
    s = Sql
    If InStr(UCase$(s), "WHERE") Then s = Left$(s, InStr(UCase$(s), "WHERE") - 1)
    Set hDYOld = MXDB.dbCreaDY(holddb, Sql, TIPO_TABELLA)
    'Casi Particolari di copia archivi
    Select Case UCase$(nometab1)
        Case "VISTACONTATORI"
            'Il campo Progr della VistaContatori non è aggiornabile quindi vado a scrivere
            'direttamente nella tabella contatori.
            s = swapp(s, nometab1, "TabContatori")
        Case "EXTRAMAG"
            ExtraArtSel = True
        Case "EXTRADEPOSITI"
            ExtraDepSel = True
        Case "EXTRAGIACDEPOSITI"
            ExtraGiacDepSel = True
    End Select
    Call MXNU.MsgBoxEX(MXNU.CaricaStringaRes(1812, NomeTab), vbCritical, "")
    LungLog = LungLog + Len(MXNU.CaricaStringaRes(1812, NomeTab) & ":") + 1
    DoEvents
    'LungLog = MXNU.LOFLog()
    Set hDYNew = MXDB.dbCreaDY(hnewdb, s, TIPO_TABELLA)
    q = Not MXDB.dbFineTab(hDYOld, TIPO_DYNASET)
    If q Then
        Rec = 0
        nrec = MXDB.dbNumeroRecord(hDYOld, TIPO_DYNASET)
        numcampi = MXDB.dbGetNumeroColonne(hDYOld, TIPO_DYNASET)

        ReDim mapping(0 To numcampi - 1) As Integer
        ReDim Blob(0 To numcampi - 1) As Integer
        ReDim CampiBlob(0 To numcampi - 1) As String
        Dim trovatoBlob%, numBlob%, curBlob%, tmpfile$, flen&
        For i = 0 To numcampi - 1
            nco = MXDB.dbGetNomeCampo(hDYOld, TIPO_DYNASET, i)
            mapping(i%) = MXDB.dbGetNumeroCampo(hDYNew, TIPO_DYNASET, nco)
            Blob(i%) = False
        Next i

        If Trim$(strBlob$) <> "" Then
            numBlob% = slice(strBlob$, "|", CampiBlob$())
            trovatoBlob% = (numBlob% <> 0)
            For curBlob% = 0 To numBlob% - 1
                i% = MXDB.dbGetNumeroCampo(hDYOld, TIPO_DYNASET, CampiBlob$(curBlob%))
                Blob(i%) = True
            Next curBlob%
        End If

        On Local Error GoTo err_CopiaTabella
        intrans% = False
        bolCampiMod = False
        If Not commitparziale Then MXDB.dbBeginTrans hnewdb
        While q
            If commitparziale Then MXDB.dbBeginTrans hnewdb
            intrans% = True
            q = MXDB.dbInserisci(hDYNew, TIPO_DYNASET)
            For i = 0 To numcampi - 1
                If Not Blob(i%) Then
                    strNomeCampo = MXDB.dbGetNomeCampo(hDYOld, TIPO_DYNASET, i)
                    If StrComp(strNomeCampo, "UtenteModifica", vbTextCompare) <> 0 And StrComp(strNomeCampo, "DataModifica", vbTextCompare) <> 0 Then
                        Valore = MXDB.dbGetCampo(hDYOld, TIPO_DYNASET, i, Null)
                        q = MXDB.dbSetCampo(hDYNew, TIPO_DYNASET, mapping(i), Valore)
                    Else
                        bolCampiMod = True
                    End If
                End If
            Next i
            If Not bolCampiMod Then
                bolOldCampiMod = hndDBArchivi.SettaCampiModifica
                hndDBArchivi.SettaCampiModifica = False
            End If
            If trovatoBlob% Then
                q = MXDB.dbRegistra(hDYNew, TIPO_DYNASET, "1=1")
            Else
                q = MXDB.dbRegistra(hDYNew, NO_REPOSITION, "1=1")
            End If
            If Not bolCampiMod Then
                hndDBArchivi.SettaCampiModifica = bolOldCampiMod
            End If

            'Registrazione dei campi BLOB (Rif.738)
            If trovatoBlob% Then
                tmpfile$ = MXNU.GetTempFile()
                For i% = 0 To numcampi - 1
                    If Blob(i%) Then
                        q% = MXDB.dbGetBlobIntoFile(hDYOld, TIPO_DYNASET, i%, tmpfile$, flen&)
                        q% = MXDB.dbSetBlobFromFile(hDYNew, TIPO_DYNASET, mapping(i%), tmpfile$, flen&)
                    End If
                Next i%
                Kill tmpfile$
            End If
            If commitparziale Then MXDB.dbCommitTrans hnewdb
            intrans% = False

            q = MXDB.dbSuccessivo(hDYOld, TIPO_DYNASET)
            Rec = Rec + 1
            i = Fix(Rec / nrec * 100)
            If Not (Prc Is Nothing) Then
                If i > Prc.Value Then Prc.Value = i
            End If
            DoEvents
        Wend

        If UCase$(nometab1) = "TIPOREGISTROIVA" Then
            'Copiando la tabella Registri Fiscali è necessario copiare anche la tabella dei
            'progressivi di stampa in quanto deve sempre corrispondere (con il codice) alla tabella Registri Fiscali.
            Dim hSS As MXKit.CRecordSet
            Dim intq As Integer
            Dim vntCodice As Variant
            Dim vntData As Variant

            Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT Codice,DataIniCont FROM TabEsercizi")
            intq = Not MXDB.dbFineTab(hSS)
            While intq
                vntCodice = MXDB.dbGetCampo(hSS, NO_REPOSITION, "Codice", 0)
                vntData = MXDB.dbGetCampo(hSS, NO_REPOSITION, "DataIniCont", "")
                MXDB.dbEseguiSQL hndDBArchivi, "INSERT INTO ProgressiviStampa (Esercizio,NrRegistro,DataFin,UtenteModifica,DataModifica) SELECT " & vntCodice & ",Codice,{d'" & Format$(vntData, "yyyy-mm-dd") & "'},UtenteModifica,DataModifica FROM TipoRegistroIva WHERE Codice NOT IN (SELECT NrRegistro FROM ProgressiviStampa WHERE Esercizio= " & vntCodice & ")"

                intq = MXDB.dbSuccessivo(hSS)
            Wend
            intq = MXDB.dbChiudiSS(hSS)
        ElseIf UCase$(nometab1) = "TABPROVVIGIONI" Then '(Rif.860)
            Set hndSS = MXDB.dbCreaSS(holddb, "SELECT CFGProvvigioni FROM ParametriGPrezzi WHERE NrRecord=1", TIPO_TABELLA)
            Valore = MXDB.dbGetCampo(hndSS, TIPO_DYNASET, "CFGProvvigioni", 0)
            q% = MXDB.dbChiudiSS(hndSS)
            MXDB.dbEseguiSQL hndDBArchivi, "UPDATE ParametriGPrezzi SET CFGProvvigioni = " & Valore & " WHERE NrRecord=1"
        ElseIf UCase$(nometab1) = "GESTIONEPREZZI" Then '(Rif.860)
            Set hndSS = MXDB.dbCreaSS(holddb, "SELECT CFGPrezzi FROM ParametriGPrezzi WHERE NrRecord=1", TIPO_TABELLA)
            Valore = MXDB.dbGetCampo(hndSS, TIPO_DYNASET, "CFGPrezzi", 0)
            q% = MXDB.dbChiudiSS(hndSS)
            MXDB.dbEseguiSQL hndDBArchivi, "UPDATE ParametriGPrezzi SET CFGPrezzi = " & Valore & " WHERE NrRecord=1"
        ElseIf (UCase$(nometab1) = "CALPRODUZIONE") Then 'RIF.A#4568 - aggiorno codice ditta sul calendario
            strQuery = "update CALPRODUZIONE" _
                & " set CODDITTA=" & hndDBArchivi.FormatoSQL(MXNU.DittaAttiva, DB_TEXT) _
                & " where CODDITTA=" & hndDBArchivi.FormatoSQL(strDittaOrigine, DB_TEXT)
            Call MXDB.dbEseguiSQL(hndDBArchivi, strQuery)
        ElseIf (UCase$(nometab1) = "PROFILOORARIOSTANDARD") Then 'RIF.A#4568 - aggiorno codice ditta sul profilo orario standard
            strQuery = "update PROFILOORARIOSTANDARD" _
                & " set CODICE=" & hndDBArchivi.FormatoSQL(MXNU.DittaAttiva, DB_TEXT) _
                & " where TIPOPO=" & MXBusiness.setTipoCalendario.tcalAziendale _
                & " and CODICE=" & hndDBArchivi.FormatoSQL(strDittaOrigine, DB_TEXT)
            Call MXDB.dbEseguiSQL(hndDBArchivi, strQuery)
        End If
        If Not commitparziale Then MXDB.dbCommitTrans hnewdb
        On Local Error GoTo 0
    End If
fine_CopiaTabella:
    On Local Error GoTo 0
    q = MXDB.dbChiudiDY(hDYNew)
    q = MXDB.dbChiudiDY(hDYOld)
Exit Sub

err_CopiaTabella:
    Dim errDescription As String
    Select Case Err.Number
       Case -2147217900  ' record già esistente
            Resume Next
       Case Else
            errDescription = Err.Description
            Call MXNU.MsgBoxEX(MXNU.CaricaStringaRes(1814, NomeTab) & "[" & errDescription & "]", vbCritical, 1007)
            If commitparziale Then
                If intrans% Then MXDB.dbRollBack hnewdb
                Resume Next
            Else
                MXDB.dbRollBack hnewdb
                Resume fine_CopiaTabella
            End If
    End Select

End Sub

Function EsisteCampo(strNomeTab As String, strNomeCampo As String) As Boolean
    Dim hCol As MXKit.CRecordSet
    Dim intq As Integer

    Set hCol = MXDB.dbColonne(hndDBArchivi, strNomeTab)
    Call hCol.RecSet.Find("COLUMN_NAME='" & strNomeCampo & "'", , adSearchForward, 1)
    If hCol.RecSet.EOF Or hCol.RecSet.BOF Then
        EsisteCampo = False
    Else
        EsisteCampo = True
    End If
    intq = MXDB.dbChiudiSS(hCol)

End Function


Function dbSetTextBox1(nfile As MXKit.CRecordSet, Tipo As MXKit.setTipoQuery, txtb As TextBox, prop As Proprieta_Aggiuntive) As Integer
    'eventuali restrizioni al tipo di campo vanno
    'impostate dopo la chiamata a questa funzione.
    On Local Error Resume Next
    Dim tipocampo As setTipiDatiDB, dimcampo&
    'per compatibilità
    If prop.DataF = "" Then
        prop.DataF = txtb.DataField
    End If
    If prop.DataF = "" Then dbSetTextBox1 = False: Exit Function

    tipocampo = MXDB.dbGetTipoCampo(nfile, Tipo, prop.DataF)
    dimcampo = MXDB.dbGetLenCampo(nfile, Tipo, prop.DataF)
    If Err Then
       dbSetTextBox1 = False
       MsgBox "Errore " & Err & ": " & Error$ & " nella funzione 'DBSetTextBox'", vbExclamation, "Errore"
    Else
        Select Case tipocampo
           Case DB_TEXT, DB_LONGVARCHAR, DB_LONGBINARY
               txtb.MaxLength = dimcampo
               prop.Tipo = CKEY_CARASCII
               prop.frmt = "": prop.dflt = ""
           Case DB_CURRENCY
               txtb.MaxLength = 16
               'txtb.MultiLine = True
               'txtb.Alignment = 1
               prop.Tipo = CKEY_NUMFLOAT
               prop.frmt = MXNU.FORMATO_QUANTITA
               prop.dflt = 0
           Case DB_DOUBLE, DB_QUANTITA, DB_DECIMAL
               txtb.MaxLength = 16
               'txtb.MultiLine = True
               'txtb.Alignment = 1
               prop.Tipo = CKEY_NUMFLOAT
               prop.frmt = MXNU.FORMATO_QUANTITA
               prop.dflt = 0
           Case DB_SINGLE
               prop.Tipo = CKEY_NUMFLOAT
               prop.dflt = 0
           Case DB_BYTE
               txtb.MaxLength = 3
               ' txtb.MultiLine = True
               ' txtb.Alignment = 1
               prop.Tipo = CKEY_NUMINT
               prop.dflt = 0
           Case DB_INTEGER
               txtb.MaxLength = 4
               'txtb.MultiLine = True
               'txtb.Alignment = 1
               prop.Tipo = CKEY_NUMINT
               prop.dflt = 0
           Case DB_LONG
               txtb.MaxLength = 8
               'txtb.MultiLine = True
               'txtb.Alignment = 1
               prop.Tipo = CKEY_NUMINT
               prop.dflt = 0
       End Select
       dbSetTextBox1 = True
    End If
End Function


Function SpostaRec(strNomeTabella As String, strNomeCampoCod As String, ByVal Pos As Integer, varCodiceAttuale As Variant, varNuovoCodice As Variant, strWHE As String) As Boolean
    Dim intOP As Integer

    SpostaRec = False
    varNuovoCodice = varCodiceAttuale
    Select Case Pos
         Case BTN_PRIMO
             intOP = MXKit.dbFIND_FIRST
         Case BTN_PREC
             intOP = MXKit.dbFIND_PREVIOUS
         Case BTN_SUCC
             intOP = MXKit.dbFIND_NEXT
         Case BTN_ULTIMO
             intOP = MXKit.dbFIND_LAST
    End Select
    If MXDB.dbFind(hndDBArchivi, strNomeTabella, strWHE, strNomeCampoCod, varNuovoCodice, intOP) Then
       SpostaRec = True
    End If

End Function

Function Formatora(ByVal ora$, ByVal frmt$) As String

   Dim q%
   ReDim tm(5) As String
   If InStr(ora, MXNU.Sep_time) Then
      q = slice(ora, MXNU.Sep_time, tm())
   Else
      q = slice(ora, ":", tm())
   End If
   If tm(0) = "" Then tm(0) = "00"
   If tm(1) = "" Then tm(1) = "00"
   If tm(2) = "" Then tm(2) = "00"
   If frmt = MXNU.Formato_HHMMSS Then
      Formatora = Format$(tm(0), "00") & MXNU.Sep_time & Format$(tm(1), "00") & MXNU.Sep_time & Format$(tm(2), "00")
   Else
      Formatora = Format$(tm(0), "00") & MXNU.Sep_time & Format$(tm(1), "00")
   End If

End Function

'####################################################################################################################
'MObjKitBus
'####################################################################################################################



'*** modifica ExtensionLoader ***
'ATTENZIONE: l'uso di una cache per la collezione ambienti crea una GROSSA falla di sicurezza. Agendo come segue, infatti
'non viene fatto il controllo sui moduli runtime isv:
'   1. Lancio un'estensione compilata da Metodo (Es. Estensione contatti su AnaCF)
'   2. Ambienti2Collection con bolSkipKey = true => mColAmb contiene TUTTI gli ambienti
'   3. Lancio un'estensione compilata da Rivenditore
'   4. Ambienti2Collection con bolSkipKey = false => dovrebbe fare il controllo dei moduli ISV
'   5. In realtà il controllo NON viene fatto perchè la funzione utilizza la cache fatta nel punto 2
'Private mColAmb As Collection

Public Function CreateObjKitBus(CTLXKit As Control, CTLXBus As Control) As Boolean

    CreateObjKitBus = True
    On Local Error GoTo CreateObjKitBus_Err

    If Not (CTLXKit Is Nothing) Then
        Set MXDB = CTLXKit.CreaXODBC()
        Set MXCREP = CTLXKit.CreaXCREP()
        Set MXVI = CTLXKit.CreaXVis()
        Set MXAA = CTLXKit.CreaXAgenti()
        Set MXCT = CTLXKit.CreaXTab()
        Set MXFT = CTLXKit.CreaXFT()
        Set MXVA = New MXKit.CAmbValid
        If GBolWorkflow Then
            Set MXWKF = CTLXKit.CreaWorkFlow
        End If


    End If
    If Not (CTLXBus Is Nothing) Then
        Set MXSC = CTLXBus.CreaXScad()
        Set MXART = CTLXBus.CreaXVArt()
        Set MXSM = CTLXBus.CreaXStMag()
        Set MXGD = CTLXBus.CreaXGestDoc()
        Set MXDBA = CTLXBus.CreaXDba()
        Set MXPIAN = CTLXBus.CreaXPianif()
        Set MXPN = CTLXBus.CreaXPrimaNota()
        Set MXPROD = CTLXBus.CreaXProduzione()
        Set MXCICLI = CTLXBus.CreaXCicliLavorazione()
        Set MXCC = CTLXBus.CreaXCommCli()
        Set MXRIS = CTLXBus.CreaXRisorse()
        Set MXSCH = CTLXBus.CreaXSchedulazione()
    End If



    'Rif. anomalia #7649
#If ISM98SERVER = 0 And ISTOOLS = 0 Then
        On Local Error Resume Next
        '[21/04/2011] Rimozione Chiave Hardware
        Set MXALL = CreateObject("MXConsole.CAmbConsole")
        Set MXQM = CreateObject("M98quality.cAmbQuality")
        Set MXWIZARD = CreateObject("MXWizard.cWizard")
        Set NETFX = CreateObject("MxHostNetFX.HostSynapseExecutor")

'        'REMIND: modifiche per MXConsole
'        'Set MXALL = New MXConsole.CAmbConsole
'        If ((MXNU.ControlloModulichiave(modAllInOneRuntime) = 0) _
'            Or MXNU.ControlloModulichiave(modMetodoXPEvolution) = 0) Then
'
'            Set MXALL = CreateObject("MXConsole.CAmbConsole")
'        End If
'
'        'REMIND: modifiche per Quality
'        If (MXNU.ControlloModulichiave(modQualityMenagement) = 0) Or (MXNU.ControlloModulichiave(modOfficeUser) = 0) Then
'            Set MXQM = CreateObject("M98quality.cAmbQuality")
'        End If
'
'        'Modifiche per Wizard
'        If (MXNU.ControlloModulichiave(modMetodoXPEvolution) = 0) Then
'            Set MXWIZARD = CreateObject("MXWizard.cWizard")
'        End If
        On Local Error GoTo CreateObjKitBus_Err
#End If

#If ISTOOLS <> 0 Then
    '[09/06/2011] Rimozione Chiave Hardware
    'If (MXNU.ControlloModulichiave(modMetodoXPEvolution) = 0) Then
        Set MXWIZARD = CreateObject("MXWizard.cWizard")
    'End If
#End If

CreateObjKitBus_Fine:
    On Local Error GoTo 0
    Exit Function

CreateObjKitBus_Err:
    Call MXNU.MsgBoxEX(9010, vbCritical, 1007, Array("CreateObjKitBus", Err.Number, Err.Description))
    CreateObjKitBus = False
    On Local Error GoTo 0
    Resume CreateObjKitBus_Fine
Resume
End Function

Public Function DropObjKitBus() As Boolean
Dim bolRes As Boolean

    bolRes = True

    'RIF.A#8908
    'Set mColAmb = Nothing

    'Rif. anomalia #7649
    #If ISM98SERVER = 0 Then
        If (Not MXWIZARD Is Nothing) Then
            Call MXWIZARD.Termina
            Set MXWIZARD = Nothing
        End If
    #End If

    'supporto scripting
    If Not MXAA Is Nothing Then MXAA.ResetAmbienti


    If Not MXSCH Is Nothing Then If MXSCH.Termina() Then Set MXSCH = Nothing Else bolRes = False
    If Not MXRIS Is Nothing Then If MXRIS.Termina() Then Set MXRIS = Nothing Else bolRes = False
    If Not MXCC Is Nothing Then If MXCC.Termina() Then Set MXCC = Nothing Else bolRes = False
    If Not MXCICLI Is Nothing Then If MXCICLI.Termina() Then Set MXCICLI = Nothing Else bolRes = False
    If Not MXPROD Is Nothing Then If MXPROD.Termina() Then Set MXPROD = Nothing Else bolRes = False
    If Not MXPIAN Is Nothing Then If MXPIAN.Termina() Then Set MXPIAN = Nothing Else bolRes = False
    If Not MXGD Is Nothing Then If MXGD.Termina() Then Set MXGD = Nothing Else bolRes = False
    If Not MXPN Is Nothing Then If MXPN.Termina() Then Set MXPN = Nothing Else bolRes = False
    If Not MXSM Is Nothing Then If MXSM.Termina() Then Set MXSM = Nothing Else bolRes = False
    If Not MXDBA Is Nothing Then If MXDBA.Termina() Then Set MXDBA = Nothing Else bolRes = False
    If Not MXART Is Nothing Then If MXART.Termina() Then Set MXART = Nothing Else bolRes = False
    If Not MXSC Is Nothing Then If MXSC.Termina() Then Set MXSC = Nothing Else bolRes = False
    If Not MXCT Is Nothing Then If MXCT.Termina() Then Set MXCT = Nothing Else bolRes = False
    If Not MXVA Is Nothing Then If MXVA.Termina() Then Set MXVA = Nothing Else bolRes = False
    If Not MXAA Is Nothing Then If MXAA.Termina() Then Set MXAA = Nothing Else bolRes = False
    If Not MXVI Is Nothing Then If MXVI.Termina() Then Set MXVI = Nothing Else bolRes = False
    If Not MXFT Is Nothing Then If MXFT.Termina() Then Set MXFT = Nothing Else bolRes = False
    If Not MXCREP Is Nothing Then If MXCREP.Termina() Then Set MXCREP = Nothing Else bolRes = False
    If Not NETFX Is Nothing Then If NETFX.Termina() Then Set NETFX = Nothing Else bolRes = False
    If Not MXWKF Is Nothing Then If MXWKF.Termina() Then Set MXWKF = Nothing Else bolRes = False

    'Rif. anomalia #7649
    #If ISM98SERVER <> 1 Then
        'REMIND: modifiche per MXConsole
        If (Not MXALL Is Nothing) Then
            Call MXALL.Terminate
            Set MXALL = Nothing
        End If
        'REMIND: modifiche per Quality
        If (Not MXQM Is Nothing) Then
            Call MXQM.Termina
            Set MXQM = Nothing
        End If
    #End If

    Set MXDB = Nothing
    Set MXNU = Nothing

    DropObjKitBus = bolRes
End Function

Public Function InitObjKitBus(hndDbArch As MXKit.CConnessione) As Boolean
    Dim bolWarning As Boolean
    Dim sLineErr As String

    InitObjKitBus = True
    bolWarning = False
    On Local Error GoTo InitObjKitBus_Err

    sLineErr = "INZIZIALIZZAZIONE INTERFACCIA CRYSTAL REPORTS"
    If Not (MXCREP Is Nothing) Then
        If Not MXCREP.Inizializza(MXNU) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Crystal Reports"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA FILTRI DI STAMPA"
    If Not (MXFT Is Nothing) Then
        If Not MXFT.Inizializza(MXNU, MXVI, MXDB, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Filtri di Stampa"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    sLineErr = "INIZIALIZZAZIONE INTERFACCIA VISIONI"
    If Not (MXVI Is Nothing) Then
        If Not MXVI.Inizializza(MXNU, MXDB, MXFT, MXCREP, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Visioni"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA AGENTI"
    If MXNU.ModuloRegole Then
        'Anomalia interna (inutile esposizione della proprietà ModuloRegole del nucleo in modifica/scrittura)
        ' La proprietà viene inizializzata in ChiavePresente() del nucleo e solo lì....
        'MXNU.ModuloRegole = MXAA.Inizializza(MXNU, MXDB, MXVI, MXCREP, hndDbArch) '<-- vecchia riga
        Call MXAA.Inizializza(MXNU, MXDB, MXVI, MXCREP, hndDbArch)
    End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''''GESTIONE BPM''''''''''''''''''''''''''''''''''''''''
    If GBolWorkflow Then
        sLineErr = "INIZIALIZZAZIONE INTERFACCIA WORKFLOW"
        If Not (MXWKF Is Nothing) Then
            If Not MXWKF.Inizializza(MXNU, MXDB, hndDbArch) Then
                Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Workflow"))
                InitObjKitBus = False
                GoTo InitObjKitBus_Fine
            End If
        End If
        Dim loadwkfbusiness As New MXBusiness.LoadWorkFlow
        Dim loadwkfkit As New MXKit.LoadWorkFlow
        Call loadwkfbusiness.SetMXWKF(MXWKF)
        Call loadwkfkit.SetMXWKF(MXWKF)
        Set loadwkfbusiness = Nothing
        Set loadwkfkit = Nothing
    End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''''GESTIONE BPM''''''''''''''''''''''''''''''''''''''''

#If IsMetodo2005 = 1 Then
    sLineErr = "INIZIALIZZAZIONE INTERFACCIA NETFX"
      If Not (NETFX Is Nothing) Then
           Call NETFX.Inizializza(MXNU, mMetodoInterop)
    End If
#End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA VALIDAZIONI"
    If Not (MXVA Is Nothing) Then
        If Not MXVA.Inizializza(MXNU, MXDB, MXVI, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Validazioni"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA SCADENZE"
    If Not (MXSC Is Nothing) Then
        If Not MXSC.Inizializza(MXNU, MXDB, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Scadenze"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA TABELLE"
    If Not (MXCT Is Nothing) Then
        If Not MXCT.Inizializza(MXNU, MXDB, MXVI, MXAA, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Tabelle"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA VALIDAZIONE ARTICOLI"
    If Not (MXART Is Nothing) Then
        If Not MXART.Inizializza(MXNU, MXDB, MXAA, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Validazione Articoli"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA MOVIMENTAZIONE STORICO"
    If Not (MXSM Is Nothing) Then
        If Not MXSM.Inizializza(MXNU, MXDB, MXAA, MXART, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Movimentazione Magazzino"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA Prima Nota"
    If Not (MXPN Is Nothing) Then
        If Not MXPN.Inizializza(MXNU, MXDB, MXAA, MXCT, MXSC, MXVI, MXVA, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Prima Nota"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA Documenti"
    If Not (MXGD Is Nothing) Then
        If Not MXGD.Inizializza(MXNU, MXDB, MXAA, MXART, MXSM, MXCT, MXSC, MXVI, MXPN, MXFT, MXCREP, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Gestione Documenti"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

#If ISTOOLS = 0 Then
    sLineErr = "INIZIALIZZAZIONE INTERFACCIA DISTINTA BASE"
    If Not (MXDBA Is Nothing) Then
        If Not MXDBA.Inizializza(MXNU, MXDB, MXART, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Distinta Base"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA PIANIFICAZIONE"
    If Not (MXPIAN Is Nothing) Then
        If Not MXPIAN.Inizializza(MXNU, MXDB, MXART, MXDBA, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Pianificazione"))
            InitObjKitBus = False
        End If
    End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA ORDINI DI PRODUZIONE"
    If Not (MXPROD Is Nothing) Then
        'RIF.A.ISV.#9 - aggiunto ambiente MXVA
        If Not MXPROD.Inizializza(MXNU, MXDB, MXAA, MXART, MXSM, MXCT, MXVI, MXDBA, MXPIAN, MXVA, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Produzione"))
            InitObjKitBus = False
        End If
    End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA CICLI DI LAVORAZIONE"
    If Not (MXCICLI Is Nothing) Then
        If Not MXCICLI.Inizializza(MXNU, MXDB, MXART, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Cicli Lavorazione"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAZIONE INTERFACCIA COMMESSE CLIENTI"
    If Not (MXCC Is Nothing) Then
        If Not MXCC.Inizializza(MXNU, MXDB, MXAA, MXART, MXVI, MXDBA, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Commesse Clienti"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAIZONE INTERFACCIA GESTIONE RISORSE"
    If Not (MXRIS Is Nothing) Then
        If Not MXRIS.Inizializza(MXNU, MXDB, MXAA, MXPROD, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Gestione Risorse"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    sLineErr = "INIZIALIZZAIZONE INTERFACCIA SCHEDULAZIONE"
    If Not (MXSCH Is Nothing) Then
        If Not MXSCH.Inizializza(MXNU, MXDB, MXAA, MXART, MXCT, MXVI, MXPROD, MXCICLI, MXRIS, hndDbArch) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Schedulazione"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    bolWarning = True


    sLineErr = "INIZIALIZZAZIONE AMBIENTE ALLINONE"
    #If ISM98SERVER = 0 Then
        '[21/04/2011] Rimozione Chiave Hardware - Il controllo è stato spostato da CreateObjKitBus a InitObjKitBus
'        If ((MXNU.ControlloModuliChiave(modAllInOneRuntime) = 0) _
'            Or MXNU.ControlloModuliChiave(modMetodoXPEvolution) = 0) Then
'
'            Dim colObjs As Collection
'            Dim colAmbs As Collection
'
'            Set colAmbs = Ambienti2Collection(True)
'            Set colObjs = New Collection
'            colObjs.Add hndDBArchivi
'            Call MXALL.Initialize(MXNU.PercorsoPgm & "\AllInOne", colAmbs, colObjs)
'        Else
'            Set MXALL = Nothing
'        End If
    #End If

    sLineErr = "INIZIALIZZAZIONE AMBIENTE QUALITY"
    #If ISM98SERVER = 0 Then
        '[21/04/2011] Rimozione Chiave Hardware - Il controllo è stato spostato da CreateObjKitBus a InitObjKitBus
        If (MXNU.ControlloModuliChiave(modQualityMenagement) = 0) Or (MXNU.ControlloModuliChiave(modOfficeUser) = 0) Then
            If Not MXQM Is Nothing Then Call MXQM.Inizializza(MXNU)
        Else
            Set MXQM = Nothing
        End If
    #End If
#End If

    sLineErr = "INIZIALIZZAZIONE AMBIENTE WIZARD"
    #If ISM98SERVER = 0 Then
        '[21/04/2011] Rimozione Chiave Hardware - Il controllo è stato spostato da CreateObjKitBus a InitObjKitBus
'        If (MXNU.ControlloModuliChiave(modMetodoXPEvolution) = 0) Then
'            Call MXWIZARD.Inizializza(MXNU, MXDB, MXVI, MXVA, MXFT, MXCT, hndDBArchivi)
'        Else
'            Set MXWIZARD = Nothing
'        End If
    #End If


InitObjKitBus_Fine:
    On Local Error GoTo 0
    Exit Function

InitObjKitBus_Err:
    Call MXNU.MsgBoxEX(9010, vbCritical, 1007, Array("InitObjKitBus", Err.Number, Err.Description & " [" & sLineErr & "]"))
    If Not bolWarning Then InitObjKitBus = False
    On Local Error GoTo 0
    Resume InitObjKitBus_Fine

End Function


Public Function Ambienti2Collection(Optional ByVal bolSkipKey As Boolean = False) As Collection
Dim colAmb As Collection

    'creo la collezione degli ambienti
    Set colAmb = New Collection
    With colAmb
        .Add MXNU, "MXNU"
        'NOTA: MXBROWSER non deve essere controllato da moduli RUNTIME ma dal solo modulo
        '      @METODO che viene controllato in fase di creazione dell'oggetto mMetodoBrowser
#If IsMetodo2005 = 1 Then
            .Add mMetodoBrowser, "MXBROWSER"
#End If

        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_KIT) = 0) Then
            .Add MXDB, "MXDB"
            .Add MXCREP, "MXCREP"
            .Add MXCT, "MXCT"
            .Add MXVI, "MXVI"
            .Add MXVA, "MXVA"
            .Add MXFT, "MXFT"
            If MXNU.ControlloModuliChiave(modAgentiRunTime) = 0 Then .Add MXAA, "MXAA"
            .Add MXALL, "MXALL"
            .Add MXQM, "MXQM"
            .Add NETFX, "NETFX"
        End If
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_SCADENZE) = 0) Then .Add MXSC, "MXSC"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_CTRLCODARTICOLO) = 0) Then .Add MXART, "MXART"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_STORICO) = 0) Then .Add MXSM, "MXSM"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_DBA) = 0) Then .Add MXDBA, "MXDBA"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_DOCUMENTI) = 0) Then .Add MXGD, "MXGD"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_PIANIFICAZIONE) = 0) Then .Add MXPIAN, "MXPIAN"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_PRIMANOTA) = 0) Then .Add MXPN, "MXPN"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_PRODUZIONE) = 0) Then .Add MXPROD, "MXPROD"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_CICLILAVORAZIONE) = 0) Then .Add MXCICLI, "MXCICLI"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_COMMESSECLIENTI) = 0) Then .Add MXCC, "MXCC"
        If bolSkipKey Or (MXNU.ControlloModuliChiave(MD32_BUSINESS_GESTIONERISORSE) = 0) Then .Add MXRIS, "MXRIS"
    End With

    Set Ambienti2Collection = colAmb
    Set colAmb = Nothing

End Function

Public Function AddAmbienti2Script()

    With MXAA
        .AddAmbiente "MXRIS", MXRIS
        .AddAmbiente "MXCC", MXCC
        .AddAmbiente "MXCICLI", MXCICLI
        .AddAmbiente "MXPROD", MXPROD
        .AddAmbiente "MXPIAN", MXPIAN
        .AddAmbiente "MXGD", MXGD
        .AddAmbiente "MXPN", MXPN
        .AddAmbiente "MXSM", MXSM
        .AddAmbiente "MXDBA", MXDBA
        .AddAmbiente "MXART", MXART
        .AddAmbiente "MXSC", MXSC
        .AddAmbiente "MXCT", MXCT
        .AddAmbiente "MXVA", MXVA
        .AddAmbiente "MXVI", MXVI
        .AddAmbiente "MXFT", MXFT
        .AddAmbiente "MXCREP", MXCREP
        .AddAmbiente "NETFX", NETFX
        'aggiunta ambiente AIOT
        If (Not MXALL Is Nothing) Then
            .AddAmbiente "MXALL", MXALL
        End If

        '********* già presenti nella liberia ************
        '.AddAmbiente MXAA, "MXAA"
        '.AddAmbiente MXDB, "MXDB"
        '.AddAmbiente MXNU, "MXNU"
    End With
End Function

'funzione aggiunta per modulo acquisizione dati
Public Function Globals2Collection() As Collection
Dim colGlobs As Collection

    Set colGlobs = New Collection
    colGlobs.Add hndDBArchivi, "HNDDBARCHIVI"

    Set Globals2Collection = colGlobs
    Set colGlobs = Nothing
End Function

'####################################################################################################################
'MObjKitBus
'####################################################################################################################



#If ISM98SERVER <> 1 Then

'    Public Function CambioUtenteAttivo() As Boolean
'        Dim strUtente As String
'        Dim strPassword As String
'
'        mBolCambioUtente = True
'        #If IsMetodo2005 <> 1 Then
'            'scarico la form moduli affinche' il menu di rapido di metodo xp si possa salvare correttamente
'            Unload frmModuli
'        #End If
'        'CRISTIAN: Cancellazione del nomemacchina dalla TabUtenti in cambiamento dell'utente attivo
'        Select Case UCase(MXNU.EXEName)
'            Case "METODO98", "METODOXP", "METODOEVOLUS"
'                Call MXDB.dbAggiornaTabUtenti(hndDBArchivi, False)
'        End Select
'
'        #If IsMetodo2005 = 1 Then   'Spostato prima di ApriLoginUtente per Anomalia 10232
'            Call frmModuli.SalvaPreferitiUtente
'            Call frmShortcuts.SalvaShortcuts
'            '<rif anomalia 8930 RZ>
'            Set frmLog = Nothing
'            Call metodo.SalvaLayout
'        #Else
'            'Anomalia 7868
'            metodo.Barra.Buttons(idxBottoneModuli).Enabled = False
'        #End If
'
'        If (MXDB.ApriLoginUtente(strUtente, strPassword)) Then
'            MbolInChiusura = False
'
'            Call ChiudiDitta
'
'            If Not ApriDittaAnno(True, strUtente, strPassword) Then
'                MbolInChiusura = True
'                'Rif. Anomalia Nr. 7688
'                GBolNoMsgConfermaUscita = True
'                Unload metodo
'                Exit Function
'            Else
'                metodo.MousePointer = vbHourglass
'                Call CaricaModuliMetodo(True)
'                #If IsMetodo2005 = 1 Then
'                    Call frmShortcuts.CaricaShortcuts
'                    Call metodo.CaricaLayout
'                    'Anomalia 8948
'                    DoEvents
'                    ' S#3040 - rimossa la Gestione Accessi da Evolus
'                    'metodo.Barra.Buttons(idxBottoneDefAccessi).Enabled = (Not MXNU.CtrlAccessi)
'                #Else
'                    'Anomalia 7868
'                    metodo.Barra.Buttons(idxBottoneModuli).Enabled = True
'                #End If
'                frmModuli.ModuloAttivo = "Metodo98"
'                Call AggiornaStatusBar
'                metodo.MousePointer = vbNormal
'            End If
'            CambioUtenteAttivo = True
'        Else
'            'Call ChiudiMetodo
'            'End
'            'Rif. Anomalia Nr. 7688
'            GBolNoMsgConfermaUscita = True
'            Unload metodo
'            Exit Function
'        End If
'        DoEvents
'        mBolCambioUtente = False
'
'    End Function


'    Public Function ApriDittaAnno(bolAperturaMetodo As Boolean, strUtente As String, strPassword As String) As Boolean
'        Dim strDitta As String
'        Dim bolOk As Boolean
'        Dim bolVersione As Boolean
'
''        If CmbDittaBusy Then
''            Call MXNU.MsgBoxEX(3178, vbExclamation, 1007)
''            ApriDittaAnno = False
''            Exit Function
''        End If
'
'        If Not bolAperturaMetodo Then
'            If Not MXNU.IsMetodo2005 Then
'                'frmModuli.Enabled = False
'                'metodo.Enabled = False
'                If SelezioneDitta(strDitta) Then
'                    Call ChiudiDitta
'                    MXNU.DittaAttiva = strDitta
'                    ' Anomalia n.ro 5242 e n.ro 5329
'                    #If ISNUCLEO = 0 Then
'                        If Not (Designer Is Nothing) Then
'                            bolVersione = Designer.AttivaVersione
'                        End If
'                    #End If
'                Else
'                    ApriDittaAnno = False
'                    frmModuli.Enabled = True
'                    metodo.Enabled = True
'                    Exit Function
'                End If
'            Else
'                #If IsMetodo2005 = 1 Then
'                    Call frmModuli.SalvaPreferitiUtente
'                    Call frmShortcuts.SalvaShortcuts
'                    '<rif anomalia #8930 RZ>
'                    Set frmLog = Nothing
'                    Call metodo.SalvaLayout
'                #End If
'                ' Rif. anomalia #3160 per Metodo Evolus
'                If MXNU.LoginIntegrato Then
'                    If SelezioneDitta(strDitta) Then
'                        Call ChiudiDitta
'                        MXNU.DittaAttiva = strDitta
'                        ' Anomalia n.ro 5242 e n.ro 5329
'                        #If ISNUCLEO = 0 Then
'                            If Not (Designer Is Nothing) Then
'                                bolVersione = Designer.AttivaVersione
'                            End If
'                        #End If
'                    Else
'                        Call ChiudiDitta
'                        ApriDittaAnno = False
'                        frmModuli.Enabled = True
'                        metodo.Enabled = True
'                        Exit Function
'                    End If
'                Else
'                    Call ChiudiDitta
'                End If
'                ' Fine rif. anomalia #3160 per Metodo Evolus
'            End If
'        End If
'
'        bolOk = True
'        Do
'            If ApriDitta(strUtente, strPassword) Then
'                Dim NuovoAnno As Integer
'                If SelezioneAnno(False, NuovoAnno) Then
'                    MXNU.AnnoAttivo = NuovoAnno
'                    Call ApriAnno(Not bolAperturaMetodo)
'                    bolOk = True
'
'                    #If USAM98SERVER Then
'                    If MXNU.LeggiProfilo(MXNU.DirAvvio & "\mw.ini", "METODOW", "SERVIZIOMET", 0) = 0 Then
'                        Call SottoponiSessione(GobjM98Server)
'                    Else
'                        Set GobjM98Server = Nothing
'                        Call ConnectM98Server(GobjM98Server)
'                    End If
'                    #End If
'
'                    If Not bolAperturaMetodo Then
'                        On Local Error Resume Next
'                        Call MXNU.RicaricaRisorseDitta ' Rif scheda n.ro 3092 (Rif. Scheda n.ro 1 Anomalia ISV)
'                        frmModuli.ModuloAttivo = "Metodo98"
'                        If Not MXNU.IsMetodo2005 Then
'                            Unload frmModuli
'                            Call CaricaModuliMetodo
'                        Else
'                            Call CaricaModuliMetodo(True)
'                            #If IsMetodo2005 = 1 Then
'                                Call metodo.CaricaLayout
'
'                            #End If
'                        End If
'                        frmModuli.ModuloAttivo = "Metodo98"
'                        On Local Error GoTo 0
'                    End If
'                Else
'                    bolOk = False
'                End If
'            Else
'                bolOk = False
'            End If
'            If Not bolOk Then
'                If MXNU.MsgBoxEX(9011, vbCritical + vbYesNo, 1007) = vbNo Then
'                    ApriDittaAnno = False
'                    Exit Do
'                Else
'                    Call ChiudiDitta
'                    MXNU.DittaAttiva = ""
'                    ApriDittaAnno = ApriDitta(MXNU.UtenteAttivo, MXNU.PasswordUtente)
'                    If ApriDittaAnno Then
'                        metodo.MousePointer = vbHourglass
'                        Call CaricaModuliMetodo(True)
'                        #If IsMetodo2005 = 1 Then
'                            Call frmShortcuts.CaricaShortcuts
'                            Call metodo.CaricaLayout
'                            'Anomalia 8948
'                            DoEvents
'                            ' S#3040 - rimossa la Gestione Accessi da Evolus
'                            'metodo.Barra.Buttons(idxBottoneDefAccessi).Enabled = (Not MXNU.CtrlAccessi)
'                        #End If
'                        frmModuli.ModuloAttivo = "Metodo98"
'                        Call AggiornaStatusBar
'                        metodo.MousePointer = vbNormal
'                    End If
'                    Exit Do
'                End If
'            Else
'                ApriDittaAnno = True
'                ' Anomalia n.ro 5242 e n.ro 5329
'#If ISNUCLEO = 0 Then
'                If Not (Designer Is Nothing) Then
'                    bolVersione = Designer.AttivaVersione
'                End If
'#End If
'                Exit Do
'            End If
'        Loop
'
'        ' Rif. anomalia n.ro 4824
'        #If IsMetodo2005 = 1 Then
'            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneDefAgenti).Enabled = MXNU.ModuloRegole And Not (MXNU.CtrlAccessi)
'            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneNomiCtrlCmp).Enabled = MXNU.ModuloRegole And Not (MXNU.CtrlAccessi)
'            metodo.FindCommandBar(ID_TLB_PRINC).FindControl(, idxBottoneSituazAnagr).Enabled = Not (MXNU.CtrlAccessi)
'        #Else
'            With metodo.Barra.Buttons
'                .Item(idxBottoneDefAgenti).ButtonMenus.Item(1).Enabled = MXNU.ModuloRegole And Not (MXNU.CtrlAccessi)
'                .Item(idxBottoneDefAgenti).ButtonMenus.Item(2).Enabled = MXNU.ModuloRegole And Not (MXNU.CtrlAccessi)
'                .Item(idxBottoneDefAgenti).ButtonMenus.Item(3).Enabled = Not (MXNU.CtrlAccessi)
'            End With
'        #End If
'        ' Fine rif. anomalia n.ro 4824
'
'        If Not bolAperturaMetodo Then
'            frmModuli.Enabled = True
'            metodo.Enabled = True
'        End If
'
'    End Function


'    Public Function SelezioneAnno(bolSel As Boolean, intNuovoAnno As Integer) As Boolean
'        Dim colRisultati As Collection, intq As Integer, bolEseguiSel As Boolean
'        SelezioneAnno = True
'        If bolSel Then
'            bolEseguiSel = True
'        Else
'            bolEseguiSel = (MXNU.AnnoAttivo = 0) Or Not EsisteAnno()
'        End If
'        If bolEseguiSel Then
'            If MXNU.FrmMetodo Is Nothing Then Set MXNU.FrmMetodo = frmInit
'            If (MXVI.Selezione("TABESE", "CODICE", "", False, Nothing, colRisultati)) Then
'                'MXNU.AnnoAttivo = colRisultati(1)("Codice")  '<=== Remmato per sk anomalie 5113: AnnoAttivo deve essere modificato DOPO ChiudiFormAttive
'                intNuovoAnno = colRisultati(1)("Codice")
'            Else
'                intNuovoAnno = MXNU.AnnoAttivo
'                SelezioneAnno = False
'            End If
'        Else
'            intNuovoAnno = MXNU.AnnoAttivo
'            intq = MXDB.dbEseguiSQL(hndDBArchivi, "UPDATE TabUtenti SET EsercizioAttivo=" & MXNU.AnnoAttivo & " WHERE UserID='" & MXNU.UtenteAttivo & "' AND NOMEMACCHINA=" & hndDBArchivi.FormatoSQL(MXNU.NomeComputer, DB_TEXT))
'        End If
'        Set colRisultati = Nothing
'
'    End Function

'    Public Function SelezioneDitta(strDitta As String) As Boolean
'
'        SelezioneDitta = False
'        If (frmSelDitta.SelezioneDitta(strDitta)) Then
'            SelezioneDitta = True
'        End If
'
'    End Function
'
'
'    Sub CancellaDitta()
'        Dim strDitta As String
'
'        On Local Error GoTo CancellaDitta_Err
'        If SelezioneDitta(strDitta) Then
'            If strDitta <> MXNU.DittaAttiva Then
'                If MXNU.MsgBoxEX(1860, vbQuestion + vbYesNo, 1007, strDitta) = vbYes Then
'                     MXNU.ScriviProfilo MXNU.PercorsoLocal$ & "\ditte.ini", "DITTE", strDitta, 0&
'                     MXNU.ScriviProfilo MXNU.PercorsoLocal$ & "\ditte.ini", "CONNESSIONE", strDitta, 0&
'                     Call MXNU.MsgBoxEX(2368, vbInformation, 1007)
'                End If
'            End If
'        End If
'CancellaDitta_Fine:
'        On Local Error GoTo 0
'        Exit Sub
'
'CancellaDitta_Err:
'        Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("CancellaDitta", Err.Number, Err.Description))
'        Resume CancellaDitta_Fine
'    End Sub


#End If

Public Sub ApriAnno(bolAggSts As Boolean)
    Dim intq As Integer
    Dim strsql As String
    Dim hSS As CRecordSet

    strsql = "SELECT * FROM TabEsercizi WHERE Codice=" & MXNU.AnnoAttivo
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strsql, TIPO_SNAPSHOT)

    MXNU.DescrizioneAnnoAttivo = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "Descrizione", 0)
    MXNU.DataIniCont = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "DataIniCont", 0)
    MXNU.DataFineCont = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "DataFineCont", 0)
    MXNU.DataIniMag = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "DataIniMag", 0)
    MXNU.DataFineMag = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "DataFineMag", 0)
    MXNU.DataIniIva = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "DataIniIva", 0)
    MXNU.DataFineIva = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "DataFineIva", 0)
    MXNU.UsaEuro = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "UsaEuro", 0)
    MXNU.StatoEsercizioCont = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "StatoCont", 0)
    MXNU.StatoEsercizioMag = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "StatoMag", 0)
    MXNU.LiqIva = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "LiqIva", 0)
    MXNU.IntraStatAcq = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "IntraStatAcq", 0)
    MXNU.IntraStatVend = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "IntraStatVend", 0)
    MXNU.IntraRegimeAcq = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "IntraRegimeAcq", 0)
    MXNU.IntraRegimeVend = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "IntraRegimeVend", 0)

    intq = MXDB.dbChiudiSS(hSS)

    #If ISM98SERVER <> 1 Then
        intq = MXDB.dbEseguiSQL(hndDBArchivi, "UPDATE TabUtenti SET EsercizioAttivo=" & MXNU.AnnoAttivo & " WHERE UserID='" & MXNU.UtenteAttivo & "' AND NOMEMACCHINA=" & hndDBArchivi.FormatoSQL(MXNU.NomeComputer, DB_TEXT))
    #Else
        intq = MXDB.dbEseguiSQL(hndDBArchivi, "UPDATE TabUtenti SET EsercizioAttivo=" & MXNU.AnnoAttivo & " WHERE UserID='" & MXNU.UtenteAttivo & "'")
    #End If

    Call MXDB.DBSetConnProperty(hndDBArchivi.ConnessioneW)

    #If ISKEY <> 1 Then
        Call LeggiVincoli
        Call LeggiVincoliMagazzino
    #End If
    'resetto i vincoli di produzione (rif.sch.1830)
    MXDBA.ResettaVincoliDisinta
    MXCICLI.ResettaVincoliCiclo
    MXRIS.ResettaVincoliRisorse
    MXPROD.ResettaVincoliProduzione

    #If ISM98SERVER <> 1 Then
        'If (bolAggSts) Then Call AggiornaStatusBar
        'Call MXNU.SalvaImpostazioniUtente(MXNU.UtenteSistema)
    #End If

End Sub

Sub ChiudiDitta()
    Dim bAggUtenti As Boolean
'    If Not CmbDittaBusy Then
        #If ISM98SERVER <> 1 Then
            'Call ChiudiFormAttive
        #End If
        If (Left(UCase(MXNU.EXEName), 6) = "METODO") And Not mBolCambioUtente Then
            bAggUtenti = False
            If Not (hndDBArchivi Is Nothing) Then
                bAggUtenti = (hndDBArchivi.ConnessioneR.State = adStateOpen)
            End If
            If bAggUtenti Then Call MXDB.dbAggiornaTabUtenti(hndDBArchivi, False)
        End If
        Call MXVA.ChiudiDyTRAnagraf
        Call MXVA.ChiudiDyTRValidazione
        Call MXCT.ChiudiDyTRTabelle
        Call MXVI.ChiudiDyTRVisioni
        Call MXVI.ChiudiDyTRSituazioni
        #If ISM98SERVER <> 1 Then
            Call MXNU.SalvaImpostazioniUtente(MXNU.UtenteSistema)
        #End If
 '   Else
  '      MsgBox "Cambio ditta in corso, chiudere la form"
  '  End If
End Sub

Function ApriDitta(strUtente As String, strPassword As String, strDitta As String, blnMessaggio As Boolean) As Boolean
    'Dim strDitta As String
    Dim intTentativi As Integer
    Dim bolConnesso As Boolean
    Dim hSS As MXKit.CRecordSet
    Dim lngVersione As Long
    Dim intq As Integer
    Dim strLog As String
    Dim strDesErr As String
    Dim strLineErr As String
    On Local Error GoTo ApriDitta_Err

    ApriDitta = True
RitentaConnessione:
    'rif. anomalia #3160
    
    MXNU.DittaAttiva = strDitta
    
    #If ISM98SERVER <> 1 Then
        If strUtente <> "" Then
            Set hndDBArchivi = MXDB.dbApriDB(NOME_DB_ARCHIVI, strUtente, strPassword, hndDBArchivi)
        Else
            Set hndDBArchivi = MXDB.dbApriDB(NOME_DB_ARCHIVI, , , hndDBArchivi)
        End If
        bolConnesso = (Not (hndDBArchivi Is Nothing))
    #Else
        If strUtente <> "" And Not (bolTrustedConnection) Then ' <-- per rif. anomalia #3160 aggiunto variabile booleana di controllo
            Set hndDBArchivi = MXDB.dbApriDB(NOME_DB_ARCHIVI, strUtente, strPassword, hndDBArchivi)
        Else
            Set hndDBArchivi = MXDB.dbApriDB(NOME_DB_ARCHIVI, , , hndDBArchivi)
        End If
        bolConnesso = (Not (hndDBArchivi Is Nothing))
    #End If

    If bolConnesso Then bolConnesso = (hndDBArchivi.ConnessioneR.State <> adStateClosed)
    #If ISM98SERVER <> 1 Then

        If Not bolConnesso Then
            'If (MXNU.MsgBoxEX(9011, vbYesNo + vbQuestion + vbDefaultButton1, 1007) = vbYes) Then
            'Utilizzo MsgBox standard di vb perchè in alcuni casi non visualizza il msgbox e risponde automaticamente no
'            Dim r As VbMsgBoxResult
'            #If IsMetodo2005 = 1 Then
'                If InSelezioneDitta Then
'                    r = vbNo
'                Else
'                    r = MsgBox(MXNU.CaricaStringaRes(9011), vbYesNo + vbQuestion + vbDefaultButton1, MXNU.CaricaStringaRes(1007))
'                End If
'            #Else
'                r = MsgBox(MXNU.CaricaStringaRes(9011), vbYesNo + vbQuestion + vbDefaultButton1, MXNU.CaricaStringaRes(1007))
'            #End If
            'If r = vbYes Then
            '    If MXNU.IsMetodo2005 Then
            '        GoTo RitentaConnessione
            '    Else
                    'If (SelezioneDitta(strDitta)) Then
                        Call MXNU.SalvaImpostazioniUtente(MXNU.UtenteSistema)
                        GoTo RitentaConnessione
'                    Else
'                        If MXNU.FrmMetodo Is Nothing Then
'                            Call ChiudiMetodo
'                        Else
'                            GoTo RitentaConnessione
'                        End If
'                        ApriDitta = False
'                        Exit Function
'                    End If
'                End If
'            Else
''                If MXNU.FrmMetodo Is Nothing Then
''                    Call ChiudiMetodo
''                Else
''                    If Not MXNU.IsMetodo2005 Then
''                        GoTo RitentaConnessione
''                    End If
''                End If
'                ApriDitta = False
'                Exit Function
'            End If
        End If
        '>>> LOGIN UTENTE
        If MXNU.UtenteDB = "" Then
            ApriDitta = False
            GoTo ApriDitta_Fine
        End If
        'Controllo Versione
        Call MXDB.dbClearUltimoErrore
        strLineErr = "Lettura TabVersioni"
        strLog = MXNU.GetTempFile()
        Call MXNU.ImpostaErroriSuLog(strLog, True)
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT Max(Codice) MaxCodice FROM TABVERSIONIM98")
        If MXDB.dbUltimoErrore(strDesErr) = 0 Then
            If Not MXDB.dbFineTab(hSS) Then
                lngVersione = MXDB.dbGetCampo(hSS, NO_REPOSITION, "MaxCodice", 0)
            End If
            intq = MXDB.dbChiudiSS(hSS)
            Call MXNU.ChiudiErroriSuLog
            If Val(Replace(MXNU.VersioneMetodo, ".", "")) <> lngVersione Then
                If Not MBolSaltaMessaggiConnessione Then  'Per MetodoEvolus, nel caso si faccia annulla della selezione ditte, viene rifatta la connessione ad db precedente
                    'Call MXNU.MsgBoxEX(1186, vbCritical, 1007, Array("", Val(Replace(MXNU.VersioneMetodo, ".", "")), lngVersione))
                End If
            End If
        Else
            intq = MXDB.dbChiudiSS(hSS)
            Call MXNU.ChiudiErroriSuLog
            Call MXNU.MsgBoxEX(1185, vbCritical, 1007, Array("", strDesErr))
        End If
        strLineErr = "Copia File mwpers.ini"
        '>>> FILE INI PERSONALE
        If Dir$(MXNU.File_ini_personale, vbNormal) = "" Then
            FileCopy MXNU.PercorsoPreferenze & "\mwpers.ini", MXNU.File_ini_personale
        End If
        strLineErr = "Copia File mwpersvis.ini"
        If Dir$(MXNU.File_ini_personaleVisioni, vbNormal) = "" Then
            FileCopy MXNU.PercorsoPreferenze & "\mwpersvis.ini", MXNU.File_ini_personaleVisioni
            Call SpostaSezioneVisioni
        End If
        strLineErr = "Salva Impostazioni Utente"
        Call MXNU.SalvaImpostazioniUtente(MXNU.UtenteSistema)

        MXNU.Dsc_Breve_Ditta = LeggiDscBreve(MXNU.DittaAttiva)
        #If TOOLS <> 1 And ISNUCLEO = 0 Then
        'Sviluppo nr. 773
        strLineErr = "Lettura Record ComandiBatch"
'        If EsistonoComandiBatch("DATAORA<{ fn Now()}-1") Then
'            If Not MBolSaltaMessaggiConnessione Then  'Per MetodoEvolus, nel caso si faccia annulla della selezione ditte, viene rifatta la connessione ad db precedente
'                Call MXNU.MsgBoxEX(2585, vbCritical, 1007)
'            End If
'        End If
        #End If
    #Else
        If MXNU.UtenteDB = "" Then
            ApriDitta = False
            GoTo ApriDitta_Fine
        End If
        ApriDitta = bolConnesso
        If bolConnesso Then
            Dim hssDesDitta As MXKit.CRecordSet
            Dim strsql As String
            Dim strDes As String

            strsql = "SELECT DataCostituzione,DesBreve FROM  TabDitte"
            Set hssDesDitta = MXDB.dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
            If Not MXDB.dbFineTab(hssDesDitta, TIPO_SNAPSHOT) Then
                strDes = MXDB.dbGetCampo(hssDesDitta, TIPO_SNAPSHOT, "DesBreve", "")
                MXNU.DataCostituzione = MXDB.dbGetCampo(hssDesDitta, TIPO_SNAPSHOT, "DataCostituzione", MXNU.Default_Data)
            Else
                strDes = ""
            End If
            Call MXDB.dbChiudiSS(hssDesDitta)
            MXNU.Dsc_Breve_Ditta = strDes
        Else
            GoTo ApriDitta_Fine
        End If
    #End If

    MXNU.DSNDittaAttiva = MXDB.UltimoDSNAperto

    strLineErr = "Lettura licenze software"
    '[15/06/2011] Rimozione Chiave Hardware - Controllo spostato all'interno di ApriDitta
    If (MXNU.ChiaveSoftwarePresente(MXNU.DittaAttiva, hndDBArchivi.ConnessioneR) <> EnmStatoChiaveSoftware.ChiavePresente) Then
        Call MXNU.MsgBoxEX(9000, vbOKOnly + vbCritical, 1007)
        ApriDitta = False
        GoTo ApriDitta_Fine
    End If

    'ricarico tutti i dynaset temporanei dei file .DAT
    strLineErr = "Caricamento Validazioni"
    Call MXVA.ApriDyTRValidazione
    strLineErr = "Caricamento Anagrafiche"
    Call MXVA.ApriDyTRAnagraf
    strLineErr = "Caricamento Tabelle"
    Call MXCT.ApriDyTRTabelle
    strLineErr = "Caricamento Visioni"
    Call MXVI.ApriDyTRVisioni
    strLineErr = "Caricamento Situazioni"
    Call MXVI.ApriDyTRSituazioni

ApriDitta_Fine:
    On Local Error GoTo 0
    Exit Function

ApriDitta_Err:
    If blnMessaggio Then Call MXNU.MsgBoxEX(9010, vbCritical, 1007, Array("ApriDitta", Err.Number, Err.Description & " [" & strLineErr & "]"))
    On Local Error GoTo 0
    ApriDitta = False
    Call ChiudiDitta
    #If ISM98SERVER <> 1 Then
'        If MXNU.FrmMetodo Is Nothing Then
'            Call ChiudiMetodo
'        Else
'            Unload MXNU.FrmMetodo
'        End If
    #End If
    Resume ApriDitta_Fine
End Function


Function EsisteAnno() As Boolean
    Dim hSS As CRecordSet, intq As Integer, strsql As String

    EsisteAnno = False
    If MXNU.AnnoAttivo > 0 Then
        strsql = "SELECT * FROM TabEsercizi WHERE Codice=" & MXNU.AnnoAttivo
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, strsql, TIPO_SNAPSHOT)
        EsisteAnno = Not MXDB.dbFineTab(hSS)
        intq = MXDB.dbChiudiSS(hSS)
    End If

End Function




Sub LeggiVincoli()
Dim q As Integer
Dim hndtn As CRecordSet
Dim inti As Integer
Dim strsql As String
Dim hTabEse As MXKit.CRecordSet     'Tabella esercizi


    Set hndtn = MXDB.dbCreaSS(hndDBArchivi, "SELECT * FROM TabVincoliGIC WHERE Esercizio=" & MXNU.AnnoAttivo, TIPO_TABELLA)

    For inti = 1 To 5
        MXNU.VincoliIva(IVA_VEN, inti) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "SCIVADeb" & CStr(inti), "")
        MXNU.VincoliIva(IVA_ACQ, inti) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "SCIVACred" & CStr(inti), "")
        MXNU.VincoliIva(IVA_SOS, inti) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "SCIVASosp" & CStr(inti), "")
        MXNU.VincoliIva(IVA_VENINTRA, inti) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "SCIVAVendIntra" & CStr(inti), "")
        MXNU.VincoliIva(IVA_ACQINTRA, inti) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "SCIVAAcqIntra" & CStr(inti), "")
        MXNU.VincoliIva(IVA_AUTOFATTURE, inti) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "SCIVAAutoFatt" & CStr(inti), "")   'Sviluppo 1368
        MXNU.VincoliIva(IVA_SOS_CRED, inti) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "SCIvaSospCred" & CStr(inti), "")
    Next inti

    MXNU.Vincoli(SC_CLI_CORRISP) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "SCCliCorrisp", "")
    MXNU.Vincoli(REG_IVA_INTRA) = CStr(MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "RegVendIntra", ""))
    MXNU.Vincoli(CAUS_INSOLUTO) = CStr(MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "CausContInsoluto", ""))
    MXNU.Vincoli(CAUS_APERTURA) = CStr(MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "CausContAp", ""))
    MXNU.Vincoli(CAUS_CHIUSURA) = CStr(MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "CausContCh", ""))
    MXNU.Vincoli(CONTO_PATR_APERTURA) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "ContoPatrAP", "")
    MXNU.Vincoli(CONTO_PATR_CHIUSURA) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "ContoPatrCH", "")
    MXNU.Vincoli(CONTO_ECO_CHIUSURA) = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "ContoEcoCH", "")
    MXNU.Vincoli(REG_IVA_AUTOFATT) = CStr(MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "RegVendAutoFatt", ""))   'Sviluppo 1368

    'Rif. Sviluppo 589
    MXNU.Vincoli(CAUS_UTILEPERDITA) = CStr(MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "CausContRilUtPerd", ""))   'Sviluppo 1368
    MXNU.Vincoli(CONTO_UTILE_ESERCIZIO) = CStr(MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "ContoUtileEserc", ""))   'Sviluppo 1368
    MXNU.Vincoli(CONTO_PERDITA_ESERCIZIO) = CStr(MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "ContoPerdEserc", ""))   'Sviluppo 1368

    MXNU.CodCambioLire = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "DivisaLire", 0)
    MXNU.CodCambioEuro = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "DivisaEuro", 0)
    MXNU.DecimaliQuantita = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "nDecimaliQuantita", 0)
    MXNU.DecimaliPesiVolumi = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "nDecimaliPesiVol", 0)
    MXNU.DecimaliLireTotale = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "nDecimaliTotaleLire", 0)
    MXNU.DecimaliLireUnitario = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "nDecimaliUnitarioLire", 0)
    MXNU.FORMATO_QUANTITA = Formato("####,###,##0", MXNU.DecimaliQuantita)
    MXNU.FORMATO_PESIVOLUMI = Formato("####,###,##0", MXNU.DecimaliPesiVolumi)
    MXNU.FORMATO_LIRE_UNITARIO = Formato("####,###,###,##0", MXNU.DecimaliLireUnitario)
    MXNU.FORMATO_LIRE_TOTALE = Formato("####,###,###,##0", MXNU.DecimaliLireTotale)
    'Sviluppo nr. 1566
    MXNU.ImportiSpRipMag = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "IncludiSpRip", 0)

    'Imposto nella proprietà del nucleo l'ultimo anno creato.
    Set hTabEse = MXDB.dbCreaSS(hndDBArchivi, "SELECT MAX(CODICE) AS ULTESE FROM TABESERCIZI")
    MXNU.UltimoEsercizioCreato = MXDB.dbGetCampo(hTabEse, TIPO_SNAPSHOT, "ULTESE", MXNU.AnnoAttivo)
    Call MXDB.dbChiudiSS(hTabEse)

    q = MXDB.dbChiudiSS(hndtn)

    If MXNU.CodCambioLire = MXNU.CodCambioEuro Then
        Call MXNU.MsgBoxEX(1399, vbCritical, 1007)
    End If
    Call GetFormatiEuro

    'lettura vincoli produzione
    strsql = "select NDECIMALICICLO from TABVINCOLIPRODUZIONE order by PROGRESSIVO desc"
    Set hndtn = MXDB.dbCreaSS(hndDBArchivi, strsql, TIPO_TABELLA)
    'RIF.A#6402 - memorizzo il numero di decimali dei centesimi
    MXNU.DecimaliCentesimi = MXDB.dbGetCampo(hndtn, TIPO_SNAPSHOT, "NDECIMALICICLO", 0)
    MXNU.Formato_Centesimi = Formato("####,###,##0", MXNU.DecimaliCentesimi)
    q = MXDB.dbChiudiSS(hndtn)
End Sub


Sub GetFormatiEuro()
    Dim hSS As CRecordSet, intDec As Integer, intq As Integer

    MXNU.FORMATO_EURO_UNITARIO = "###,###,##0.00"
    MXNU.FORMATO_EURO_TOTALE = "###,###,##0.00"
    MXNU.DecimaliEuroTotale = 2
    MXNU.DecimaliEuroUnitario = 2
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT nDecimaliUnitario,nDecimaliTotale FROM TabCambi WHERE Codice=(SELECT DivisaEuro FROM TabVincoliGIC WHERE Esercizio=" & MXNU.AnnoAttivo & ")", TIPO_TABELLA)
    If Not MXDB.dbFineTab(hSS, TIPO_DYNASET) Then
        intDec = MXDB.dbGetCampo(hSS, NO_REPOSITION, "nDecimaliUnitario", 0)
        MXNU.DecimaliEuroUnitario = intDec
        MXNU.FORMATO_EURO_UNITARIO = Formato("####,###,###,##0", intDec)

        intDec = MXDB.dbGetCampo(hSS, NO_REPOSITION, "nDecimaliTotale", 0)
        MXNU.FORMATO_EURO_TOTALE = Formato("####,###,###,##0", intDec)
        MXNU.DecimaliEuroTotale = intDec
    End If
    intq = MXDB.dbChiudiSS(hSS)

    Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT CambioEuro FROM TabCambi WHERE Codice=0")
    If Not MXDB.dbFineTab(hSS, TIPO_SNAPSHOT) Then
        MXNU.CambioLireEuro = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "CambioEuro", 1)
    Else
        MXNU.CambioLireEuro = 1
    End If
    intq = MXDB.dbChiudiSS(hSS)


End Sub

Function Formato(strDes As String, intDec As Integer) As String
    Dim strD As String

    If intDec > 0 Then
        strD = Right(strDes & Left(".000000", intDec + 1), Len(strDes))
    Else
        strD = strDes
    End If
    If Left(strD, 1) = "," Then
        Mid(strD, 1, 1) = "#"
    End If
    Formato = strD

End Function


Private Sub SpostaSezioneVisioni()
    Dim strRiga As String
    Dim vetStr() As String
    Dim lngNum As Long
    Dim i As Integer
    Dim intq As Integer

    strRiga = MXNU.LeggiProfilo(MXNU.File_ini_personale, "VISIONI", 0&, "")
    If strRiga <> "" Then
        vetStr = Split(strRiga, vbNullChar, , vbTextCompare)
        lngNum = UBound(vetStr)
        For i = 0 To lngNum
            strRiga = MXNU.LeggiProfilo(MXNU.File_ini_personale, "VISIONI", vetStr(i), "")
            Call MXNU.ScriviProfilo(MXNU.File_ini_personaleVisioni, "VISIONI", vetStr(i), strRiga)
        Next i
        Call MXNU.ScriviProfilo(MXNU.File_ini_personale, "VISIONI", 0&, "")
    End If

End Sub
Public Sub ControllaOpSchedulate()

#If BATCH <> 1 And ISM98SERVER <> 1 And TOOLS <> 1 And ISNUCLEO = 0 Then
    Dim oSched As MxScheduler.clsScheduler
    Set oSched = New MxScheduler.clsScheduler
    If oSched.Inizializza(MXNU, Command()) Then
        If oSched.SegnalaLog(MXNU) Then
            Call oSched.MostraLogNonLetti(MXNU)
        End If
    End If

#End If
End Sub

Public Sub ApriSchedulatore()

#If BATCH <> 1 And ISM98SERVER <> 1 And TOOLS <> 1 And ISNUCLEO = 0 Then
    Dim oSched As MxScheduler.clsScheduler
    Set oSched = New MxScheduler.clsScheduler
    If oSched.Inizializza(MXNU, Command()) Then
        Call oSched.GestisciOperazioni(False)
    End If

#End If
End Sub


'Public Sub OldControllaOpSchedulate()
'#If BATCH <> 1 And ISM98SERVER <> 1 And TOOLS <> 1 And ISNUCLEO = 0 Then
'    Dim strFile As String
'    Dim q As Integer
'    Dim strTitolo As String
'
'    On Local Error GoTo OP_Err
'
'    strFile = Dir$(MXNU.PercorsoPreferenze & "\SCHEDULA\" & CStr(MXNU.NTerminale) & "_*.log")
'    While strFile <> ""
'        q = InStr(strFile, CStr(MXNU.NTerminale) & "_")
'        strTitolo = Mid$(strFile, q + 2)
'        q = InStr(strTitolo, ".")
'        strTitolo = Left$(strTitolo, q - 1)
'        strTitolo = MXNU.CaricaCaptionInLingua(MXNU.LeggiProfilo(MXNU.PercorsoPgm & "\MWSchedula.ini", "DESCRIZIONE", strTitolo, strTitolo))
'        #If IsMetodo2005 = 1 Then
'            frmLog.MstrTitolo = strTitolo
'            Call frmLog.MostraFileLog(MXNU.PercorsoPreferenze & "\SCHEDULA\" & strFile, , , True)
'        #Else
'            Dim frmLogOpSc As frmLog
'            Set frmLogOpSc = New frmLog
'            frmLogOpSc.MstrTitolo = strTitolo
'            Call frmLogOpSc.MostraFileLog(MXNU.PercorsoPreferenze & "\SCHEDULA\" & strFile, , , True)
'            Set frmLogOpSc = Nothing
'        #End If
'        If MXNU.MsgBoxEX(2580, vbInformation + vbYesNo, 1007, Array(strTitolo)) = vbYes Then
'            Kill MXNU.PercorsoPreferenze & "\SCHEDULA\" & strFile
'        End If
'        strFile = Dir
'    Wend
'OP_Fine:
'    On Local Error GoTo 0
'    Exit Sub
'
'OP_Err:
'    Call MXNU.MsgBoxEX(1009, vbCritical, 1007, Array("ControllaOpSchedulate", Err.Number, Err.Description))
'    Resume OP_Fine
'#End If
'End Sub









'RIF.A#10175 - faccio l'analisi copertura solo se filtro modificato. La modifica della query visione la faccio in ogni caso
Public Function CreaTempPerAnalisiDisp(strQueryVis As String, ByVal strWHERE As String, ByVal strOrderBy As String, ByVal vntDataElab As Variant, Optional bolFiltroModificato As Boolean = True) As Boolean
    Dim intPos As Integer
    Dim objAnalisiDisp As MXBusiness.CAnalisiProd

    MXNU.MostraMsgInfo 70035
    If (bolFiltroModificato) Then
        'creo oggetto per analisi
        Set objAnalisiDisp = MXPROD.CreaCAnalisiProd()
        CreaTempPerAnalisiDisp = Not (objAnalisiDisp Is Nothing)
        If CreaTempPerAnalisiDisp Then
            'effettuo analisi copertura
            CreaTempPerAnalisiDisp = objAnalisiDisp.AnalisiCopertura(strWHERE, vntDataElab)
        End If
    End If
    'modifico query visione
    intPos = InStrRev(strQueryVis, "where", , vbTextCompare)
    strQueryVis = Left$(strQueryVis, intPos - 2) & " WHERE IDSESSIONE=" & MXNU.IDSessione _
        & " ORDER BY " & strOrderBy

END_CreaTempPerAnalisiDisp:
    Set objAnalisiDisp = Nothing
    MXNU.MostraMsgInfo ""
End Function


Private Function AnnoMeseInizioCont() As String
    AnnoMeseInizioCont = CStr(Year(MXNU.DataIniCont)) + "-" + Mid("01-GEN02-FEB03-MAR04-APR05-MAG06-GIU07-LUG08-AGO09-SET10-OTT11-NOV12-DIC", (Month(MXNU.DataIniCont) - 1) * 6 + 1, 6)
End Function

' validazioni personalizzate dei filtri
Public Sub ValidPersFiltri(ByVal strNomeValid As String, ByVal strNomeCmpValid As String, bolEseguiValStd As Boolean, vntNewValore As Variant)
    Select Case strNomeValid
        Case "VALID_ARTICOLO", "VALID_ARTVARIANTI"
            Dim xCodArt As MXBusiness.CVArt
            Set xCodArt = MXART.CreaCVArt()
            xCodArt.Codice = vntNewValore
            If xCodArt.Valida(CHIEDIVAR_TUTTE, False, , 0, False) Then
                vntNewValore = xCodArt.Codice
            End If
            'bolEseguiValStd = True
            Call xCodArt.Termina
            Set xCodArt = Nothing
        Case "VALID_ARTICOLOTIP"  ' rif.sch. A4562
            Dim xCodArtTip As MXBusiness.CVArt
            Set xCodArtTip = MXART.CreaCVArt()
            xCodArtTip.Codice = vntNewValore
            If xCodArtTip.Valida(CHIEDIVAR_TUTTE, False, , 0, False) Then
                vntNewValore = xCodArtTip.Codice
            End If
            'bolEseguiValStd = True
            Call xCodArtTip.Termina
            Set xCodArtTip = Nothing
        Case "VALID_ARTCOMPOSTO"
            Dim xCodDba As MXBusiness.CComposto
            Set xCodDba = MXDBA.CreaCComposto()
            If xCodDba.Valida(vntNewValore, False) Then
                vntNewValore = xCodDba.pCodice
            End If
            Set xCodDba = Nothing
        Case "VALID_CICLOPROD"
            Dim xCodClv As MXBusiness.CCicloLav
            Set xCodClv = MXCICLI.CreaCCiclo()
            If xCodClv.Valida(vntNewValore, False) Then
                vntNewValore = xCodClv.pCodice
            End If
            Set xCodClv = Nothing
    End Select
End Sub

Public Sub Totali_AddRecordIniziali(cTraccia As MXKit.cTraccia, _
                        HrsTot As MXKit.CRecordSet, _
                        bolRichiediParziali As Boolean, _
                        CmbTotali_ListIndex As Integer, _
                        IdRigaFiltro_Esecizio As Long, _
                        IdRigaFiltro_Data As Long, _
                        ssFiltroDati As Object, _
                        bolSituazione As Boolean, _
                        Optional vntCodConto As Variant, _
                        Optional IdRigaFiltro_CodConto As Variant, _
                        Optional IdRigaFiltro_Provisorio As Variant)

    Dim objGruppo As MXKit.CGruppo
    Dim strQuery As String
    Dim strSelect As String
    Dim strWHERE As String
    Dim hSS As MXKit.CRecordSet
    Dim bolFinito As Boolean
    Dim vntValore As Variant
    Dim strNomeCampo As String
    Dim bolSaldoContabile As Boolean
    Dim vntEs As Variant
    Dim vntOldOpData As Variant

    'costruisco la query per i valori totali
    For Each objGruppo In cTraccia.pTotale(CmbTotali_ListIndex).colGruppi
        strNomeCampo = objGruppo.CColGruppo.strDataField
        If StrComp(cTraccia.pNomeTraccia, "VIS_MOVCON", vbTextCompare) = 0 Then
            If StrComp(strNomeCampo, "Mese", vbTextCompare) = 0 Then
                strNomeCampo = "0 as Mese"
                bolSaldoContabile = True
            End If
        End If
        strSelect = ConcatenaEspressione(strSelect, ",", strNomeCampo)
    Next objGruppo
    strQuery = "SELECT DISTINCT " & strSelect
    strQuery = strQuery & " FROM " & cTraccia.pLivelloCorrente.SQLDammiFROM
    vntOldOpData = ssComboListIndex(ssFiltroDati, COLOPERATORE, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Data))
    Call ssComboListIndex(ssFiltroDati, COLOPERATORE, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Data), OPMINUGUALE)
    If StrComp(cTraccia.pNomeTraccia, "VIS_MOVCON", vbTextCompare) = 0 Then
        Call ssFiltroDati.GetText(COLVALOREDA, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Esecizio), vntEs)
        Call ssFiltroDati.SetText(COLVALOREDA, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Esecizio), vntEs - 1)
    Else
        Call ssComboListIndex(ssFiltroDati, COLOPERATORE, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Esecizio), OPMINUGUALE)
    End If
    cTraccia.pLivelloCorrente.strSQLWhr = cTraccia.CFiltroDati.SQLFiltro
    'Anomalia nr. 6310
    strWHERE = cTraccia.pLivelloCorrente.SQLDammiWHERE(False)
    Call ssComboListIndex(ssFiltroDati, COLOPERATORE, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Esecizio), OPUGUALE)
    If StrComp(cTraccia.pNomeTraccia, "VIS_MOVMAG", vbTextCompare) = 0 Or StrComp(cTraccia.pNomeTraccia, "VIS_MOVMAG_BASE", vbTextCompare) = 0 Then
        Call ssCellLock(ssFiltroDati, COLOPERATORE, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Esecizio))
    End If
    If StrComp(cTraccia.pNomeTraccia, "VIS_MOVCON", vbTextCompare) = 0 Then
        Call ssFiltroDati.SetText(COLVALOREDA, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Esecizio), vntEs)
    End If
    Call ssComboListIndex(ssFiltroDati, COLOPERATORE, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Data), vntOldOpData)
    cTraccia.pLivelloCorrente.strSQLWhr = cTraccia.CFiltroDati.SQLFiltro
    strQuery = strQuery & " WHERE " & strWHERE
    'inserisco i dati nel recordset totali
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, strQuery)
    bolFinito = MXDB.dbFineTab(hSS)
    If bolFinito And bolSaldoContabile And Not IsMissing(vntCodConto) Then
        'Situazione
        Call MXDB.dbInserisci(HrsTot)
        For Each objGruppo In cTraccia.pTotale(CmbTotali_ListIndex).colGruppi
            strNomeCampo = objGruppo.CColGruppo.strDataField
            If StrComp(strNomeCampo, "Esercizio", vbTextCompare) <> 0 Then
                Select Case UCase$(strNomeCampo)
                    Case "CONTO"
                        vntValore = vntCodConto
                    Case "MESE"
                        vntValore = AnnoMeseInizioCont() '"01-GEN"  'Modificato x Anomalia 8942
                End Select
            Else
                Call ssFiltroDati.GetText(COLVALOREDA, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Esecizio), vntValore)
            End If
            Call MXDB.dbSetCampo(HrsTot, HrsTot.Tipo, strNomeCampo, vntValore)
        Next objGruppo
        Call MXDB.dbRegistra(HrsTot)
    Else
        Do While Not bolFinito
            Call MXDB.dbInserisci(HrsTot)
            For Each objGruppo In cTraccia.pTotale(CmbTotali_ListIndex).colGruppi
                strNomeCampo = objGruppo.CColGruppo.strDataField
                If StrComp(strNomeCampo, "Esercizio", vbTextCompare) <> 0 Then
                    vntValore = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, strNomeCampo, "")
                    If StrComp(cTraccia.pNomeTraccia, "VIS_MOVCON", vbTextCompare) = 0 Then
                        If StrComp(strNomeCampo, "MESE", vbTextCompare) = 0 Then
                            vntValore = AnnoMeseInizioCont()   '"01-GEN"   'Modificato x Anomalia 8942
                        End If
                    End If
                Else
                    Call ssFiltroDati.GetText(COLVALOREDA, cTraccia.CFiltroDati.IdFiltro2Riga(IdRigaFiltro_Esecizio), vntValore)
                End If
                Call MXDB.dbSetCampo(HrsTot, HrsTot.Tipo, strNomeCampo, vntValore)
            Next objGruppo
            Call MXDB.dbRegistra(HrsTot)
            bolFinito = Not MXDB.dbSuccessivo(hSS)
        Loop
    End If
    Call MXDB.dbChiudiSS(hSS)
    If bolSaldoContabile And IsMissing(vntCodConto) And Not IsMissing(IdRigaFiltro_CodConto) Then
        'Visione
        Dim strWHEIniz As String
        strWHEIniz = swapp(cTraccia.CFiltroDati.SQLFiltro(cTraccia.CFiltroDati.IdFiltro2Riga(Val(IdRigaFiltro_CodConto))), "VistaRigheContabilita.", "")
        If strWHEIniz <> "" Then strWHEIniz = strWHEIniz & " AND "
        strWHEIniz = strWHEIniz & cTraccia.CFiltroDati.SQLFiltro(cTraccia.CFiltroDati.IdFiltro2Riga(Val(IdRigaFiltro_Provisorio)))
        strQuery = "SELECT Conto FROM VistaSaldiInizialiPN WHERE Conto NOT IN (SELECT DISTINCT Conto FROM VistaRigheContabilita WHERE " & strWHERE & " ) AND " & strWHEIniz
        Set hSS = MXDB.dbCreaSS(hndDBArchivi, strQuery)
        bolFinito = MXDB.dbFineTab(hSS)
        Do While Not bolFinito
            Call MXDB.dbInserisci(HrsTot)

            Call MXDB.dbSetCampo(HrsTot, HrsTot.Tipo, "Esercizio", vntEs)
            Call MXDB.dbSetCampo(HrsTot, HrsTot.Tipo, "MESE", AnnoMeseInizioCont())   '"01-GEN"   'Modificato x Anomalia 8942

            vntValore = MXDB.dbGetCampo(hSS, TIPO_SNAPSHOT, "Conto", "")
            Call MXDB.dbSetCampo(HrsTot, HrsTot.Tipo, "Conto", vntValore)

            Call MXDB.dbRegistra(HrsTot)

            bolFinito = Not MXDB.dbSuccessivo(hSS)
        Loop
        Call MXDB.dbChiudiSS(hSS)
    End If
    bolRichiediParziali = False

End Sub
