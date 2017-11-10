Attribute VB_Name = "MObjKitBus"

Option Explicit

Global hndDBArchivi As MXKit.CConnessione

Global MXNU As MXNucleo.XNucleo
Global MXDB As MXKit.XODBC

Global MXCREP As MXKit.CAmbCRW
Global MXAA As MXKit.CAmbAgenti
Global MXCT As MXKit.CAmbTab
Global MXVI As MXKit.CAmbVisioni
Global MXVA As MXKit.CAmbValid
Global MXFT As MXKit.CAmbFiltri

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

'*** modifica ExtensionLoader ***
Private mColAmb As Collection


Public Function InitObjMetodo(ByRef frmForm As Form, _
                              ByVal strCommandLineMet As String, _
                              ByVal strDitta As String, _
                              ByVal strUtente As String, _
                              ByVal strPWD As String) As Boolean
    Dim blnObjInit As Boolean
    
    'inizializzazione del nucleo
    If InitNucleo(strCommandLineMet) Then

        ' Controllo chiave
       ' frmMain.TxtConnessione.Text = "Controllo Chiave Hardware"
        DoEvents
        
'        If InitChiaveHW() Then
        If 1 = 1 Then
            'creazione degli oggetti
            'frmMain.txtConnessione.Text = "Creazione degli oggetti Kit e Business"
            DoEvents
        
            If CreateObjKitBus(frmMain.CTLXKit1, frmMain.CTLXBus1) Then
                'connessione al database
            '    frmMain.txtConnessione.Text = "Connessione al Database"
                DoEvents
                If InitDatabase(strDitta, strUtente, strPWD) Then
                    'inizializzazione degli oggetti
            '        frmMain.txtConnessione.Text = "Inizializzazione degli oggetti Kit e Business"
                    DoEvents
                    
                    If InitObjKitBus(hndDBArchivi) Then
            '            frmMain.txtConnessione.Text = "Caricamento dei vincoli"
                        DoEvents
                        
                        Call LeggiVincoli
                        
                        blnObjInit = True
                    Else
                        MsgBox "Inizializzazione degli oggetti non riuscita", vbCritical
                        blnObjInit = False
                    End If
                Else
                    MsgBox "Inizializzazione del Database non riuscita", vbCritical
                    blnObjInit = False
                End If
            Else
                MsgBox "Creazione degli oggetti Kit e Business non riuscita", vbCritical
                blnObjInit = False
            End If
        Else
            MsgBox "Inizializzazione Chiave Hardware non riuscita", vbCritical
            blnObjInit = False
        End If
    Else
        MsgBox "Inizializzazione del Nucleo non riuscita", vbCritical
        blnObjInit = False
    End If
    
'    frmMain.txtConnessione.Text = ""

    InitObjMetodo = blnObjInit
End Function

Public Function InitNucleo(strCommandLineMet As String) As Boolean
    On Local Error GoTo InitNucleo_Err
    Dim blnRes As Boolean
        
    Set MXNU = New MXNucleo.XNucleo
    InitNucleo = MXNU.Inizializza(strCommandLineMet, "")
    MXNU.VersioneMetodo = App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "00")

InitNucleo_Fine:
    On Local Error GoTo 0
    Exit Function
    
InitNucleo_Err:
    Call MessaggioErrore("InitNucleo", Err.Number, Err.Description)
    InitNucleo = False
    Resume InitNucleo_Fine
End Function

Public Function InitDatabase(ByVal strDitta As String, _
                             ByVal strUtente As String, _
                             ByVal strPWD As String) As Boolean
    
    On Local Error GoTo InitDatabase_Err
    
    InitDatabase = MXDB.dbInizializza(MXNU)
    If InitDatabase Then
        Set hndDBArchivi = MXDB.dbApriDB(strDitta, strUtente, strPWD)
    End If
    InitDatabase = (Not hndDBArchivi Is Nothing)
    
    'carico i file .dat
    If InitDatabase Then
        MXNU.DittaAttiva = strDitta
    
        Call MXVA.ApriDyTRValidazione
        Call MXVA.ApriDyTRAnagraf
        Call MXCT.ApriDyTRTabelle
        Call MXVI.ApriDyTRVisioni
        Call MXVI.ApriDyTRSituazioni
    End If
    
InitDatabase_Fine:
    On Local Error GoTo 0
    Exit Function
    
InitDatabase_Err:
    Call MessaggioErrore("InitDatabase", Err.Number, Err.Description)
    InitDatabase = False
    On Local Error GoTo 0
    Resume InitDatabase_Fine

End Function

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
        #If IsMetodo2005 = 1 Then
            Call CTLXKit.ImpostaClasseLog(frmLog)
        #End If

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
    
        On Local Error Resume Next
    'REMIND: modifiche per MXConsole
    'Set MXALL = New MXConsole.CAmbConsole
    If ((MXNU.ControlloModulichiave(modAllInOneRuntime) = 0) _
        Or MXNU.ControlloModulichiave(modMetodoXPEvolution) = 0) Then
        
        Set MXALL = CreateObject("MXConsole.CAmbConsole")
    End If
    
    'REMIND: modifiche per Quality
    If (MXNU.ControlloModulichiave(modQualityMenagement) = 0) Or (MXNU.ControlloModulichiave(modOfficeUser) = 0) Then
        Set MXQM = CreateObject("M98quality.cAmbQuality")
    End If
    
    'Modifiche per Wizard
    If (MXNU.ControlloModulichiave(modMetodoXPEvolution) = 0) Then
        Set MXWIZARD = CreateObject("MXWizard.cWizard")
    End If
    On Local Error GoTo CreateObjKitBus_Err

    
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
    
        Set mColAmb = Nothing
    
    If (Not MXWIZARD Is Nothing) Then
        Call MXWIZARD.Termina
        Set MXWIZARD = Nothing
    End If

    
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

    
    Set MXDB = Nothing
    Set MXNU = Nothing
    
    DropObjKitBus = bolRes
End Function

Public Function InitObjKitBus(hndDBArchivi As MXKit.CConnessione) As Boolean

    InitObjKitBus = True
    On Local Error GoTo InitObjKitBus_Err
    
    '>>> INZIZIALIZZAZIONE INTERFACCIA CRYSTAL REPORTS
    If Not (MXCREP Is Nothing) Then
        If Not MXCREP.Inizializza(MXNU) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Crystal Reports"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
   
   '        '>>> INIZIALIZZAZIONE INTERFACCIA GESTIONE IMPOSTAZIONI
'        If Not MXGI.Inizializza(MXDB, MXNU, hndDbArch) Then
'            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, "Gestione Impostazioni")
'            Call ChiudiDitta
'            Call ChiudiMetodo
'        End If

   
    '>>> INIZIALIZZAZIONE INTERFACCIA FILTRI DI STAMPA
    If Not (MXFT Is Nothing) Then
        If Not MXFT.Inizializza(MXNU, MXVI, MXDB, hndDBArchivi) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Filtri di Stampa"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    '>>> INIZIALIZZAZIONE INTERFACCIA VISIONI
    If Not (MXVI Is Nothing) Then
        If Not MXVI.Inizializza(MXNU, MXDB, MXFT, MXCREP, hndDBArchivi) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Visioni"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    
    
    
'    '>>> INIZIALIZZAZIONE INTERFACCIA AGENTI
'    If MXNU.ModuloRegole Then
'        MXNU.ModuloRegole = MXAA.Inizializza(MXNU, MXDB, MXVI, MXCREP, hndDBArchivi)
''        'MXAA.CaricaMenuAgenti
'    End If
    
    If MXNU.ModuloRegole Then
        'Anomalia interna (inutile esposizione della proprietà ModuloRegole del nucleo in modifica/scrittura)
        ' La proprietà viene inizializzata in ChiavePresente() del nucleo e solo lì....
        'MXNU.ModuloRegole = MXAA.Inizializza(MXNU, MXDB, MXVI, MXCREP, hndDbArch) '<-- vecchia riga
        Call MXAA.Inizializza(MXNU, MXDB, MXVI, MXCREP, hndDBArchivi)
    End If

    
    
    '>>> INIZIALIZZAZIONE INTERFACCIA VALIDAZIONI
    If Not (MXVA Is Nothing) Then
        If Not MXVA.Inizializza(MXNU, MXDB, MXVI, hndDBArchivi) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Validazioni"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    
    '>>> INIZIALIZZAZIONE INTERFACCIA SCADENZE
    If Not (MXSC Is Nothing) Then
        If Not MXSC.Inizializza(MXNU, MXDB, hndDBArchivi) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Scadenze"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    
    '>>> INIZIALIZZAZIONE INTERFACCIA TABELLE
    If Not (MXCT Is Nothing) Then
        If Not MXCT.Inizializza(MXNU, MXDB, MXVI, MXAA, hndDBArchivi) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Tabelle"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    '>>> INIZIALIZZAZIONE INTERFACCIA VALIDAZIONE ARTICOLI
    If Not (MXART Is Nothing) Then
        If Not MXART.Inizializza(MXNU, MXDB, MXAA, hndDBArchivi) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Validazione Articoli"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    
    '>>> INIZIALIZZAZIONE INTERFACCIA MOVIMENTAZIONE STORICO
    If Not (MXSM Is Nothing) Then
        If Not MXSM.Inizializza(MXNU, MXDB, MXAA, MXART, hndDBArchivi) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Movimentazione Magazzino"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    
    '>>> INIZIALIZZAZIONE INTERFACCIA Prima Nota
    If Not (MXPN Is Nothing) Then
        If Not MXPN.Inizializza(MXNU, MXDB, MXAA, MXCT, MXSC, MXVI, MXVA, hndDBArchivi) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Prima Nota"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    
    '>>> INIZIALIZZAZIONE INTERFACCIA Documenti
    If Not (MXGD Is Nothing) Then
        If Not MXGD.Inizializza(MXNU, MXDB, MXAA, MXART, MXSM, MXCT, MXSC, MXVI, MXPN, MXFT, MXCREP, hndDBArchivi) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Gestione Documenti"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    
    '>>> INIZIALIZZAZIONE INTERFACCIA DISTINTA BASE
    If Not (MXDBA Is Nothing) Then
        If Not MXDBA.Inizializza(MXNU, MXDB, MXART, hndDBArchivi) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Distinta Base"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
        
    '>>> INIZIALIZZAZIONE INTERFACCIA PIANIFICAZIONE
    If Not (MXPIAN Is Nothing) Then
        If Not MXPIAN.Inizializza(MXNU, MXDB, MXART, MXDBA, hndDBArchivi) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Pianificazione"))
            InitObjKitBus = False
        End If
    End If
    
    '>>> INIZIALIZZAZIONE INTERFACCIA ORDINI DI PRODUZIONE
'    If Not (MXPROD Is Nothing) Then
'        If Not MXPROD.Inizializza(MXNU, MXDB, MXAA, MXART, MXSM, MXCT, MXVI, MXDBA, MXPIAN, hndDBArchivi) Then
'            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Produzione"))
'            InitObjKitBus = False
'        End If
'    End If


    If Not (MXPROD Is Nothing) Then
        'RIF.A.ISV.#9 - aggiunto ambiente MXVA
        If Not MXPROD.Inizializza(MXNU, MXDB, MXAA, MXART, MXSM, MXCT, MXVI, MXDBA, MXPIAN, MXVA, hndDBArchivi) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Produzione"))
            InitObjKitBus = False
        End If
    End If

    
    '>>> INIZIALIZZAZIONE INTERFACCIA CICLI DI LAVORAZIONE
    If Not (MXCICLI Is Nothing) Then
        If Not MXCICLI.Inizializza(MXNU, MXDB, MXART, hndDBArchivi) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Cicli Lavorazione"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    '>>> INIZIALIZZAZIONE INTERFACCIA COMMESSE CLIENTI
    If Not (MXCC Is Nothing) Then
        If Not MXCC.Inizializza(MXNU, MXDB, MXAA, MXART, MXVI, MXDBA, hndDBArchivi) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Commesse Clienti"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    '>>> INIZIALIZZAZIONE INTERFACCIA GESTIONE RISORSE
    If Not (MXRIS Is Nothing) Then
        If Not MXRIS.Inizializza(MXNU, MXDB, MXAA, MXPROD, hndDBArchivi) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Gestione Risorse"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If

    '>>> INIZIALIZZAZIONE INTERFACCIA SCHEDULAZIONE
    If Not (MXSCH Is Nothing) Then
        If Not MXSCH.Inizializza(MXNU, MXDB, MXAA, MXART, MXCT, MXVI, MXPROD, MXCICLI, MXRIS, hndDBArchivi) Then
            Call MXNU.MsgBoxEX(9004, vbOKOnly + vbCritical, 1007, Array("Schedulazione"))
            InitObjKitBus = False
            GoTo InitObjKitBus_Fine
        End If
    End If
    
    
        '>>> INIZIALIZZAZIONE AMBIENTE ALLINONE
    'REMIND: modifiche per MXConsole
    #If ISM98SERVER = 0 Then
        If Not (MXALL Is Nothing) Then
            Dim colObjs As Collection
            Dim colAmbs As Collection
                    
            Set colAmbs = Ambienti2Collection(True)
            Set colObjs = New Collection
            colObjs.Add hndDBArchivi
            Call MXALL.Initialize(MXNU.PercorsoPgm & "\AllInOne", colAmbs, colObjs)
        End If
    #End If

    '>>> INIZIALIZZAZIONE AMBIENTE QUALITY
    #If ISM98SERVER = 0 Then
        If Not (MXQM Is Nothing) Then
            Call MXQM.Inizializza(MXNU)
        End If
    #End If
    
    'Wizard
    #If ISM98SERVER = 0 Then
        If Not (MXWIZARD Is Nothing) Then
            Call MXWIZARD.Inizializza(MXNU, MXDB, MXVI, MXVA, MXFT, MXCT, hndDBArchivi)
        End If
    #End If


InitObjKitBus_Fine:
    On Local Error GoTo 0
    Exit Function
    
InitObjKitBus_Err:
    Call MXNU.MsgBoxEX(9010, vbCritical, 1007, Array("InitObjKitBus", Err.Number, Err.Description))

    InitObjKitBus = False
    On Local Error GoTo 0
    Resume InitObjKitBus_Fine

End Function

Private Sub MessaggioErrore(ByVal strSource As String, ByVal lngErrNumber As Long, ByVal strErrDescription As String)
    MsgBox "Si è verificato un errore nella funzione " & strSource & " di tipo [" & lngErrNumber & "] " & strErrDescription, vbCritical, "ERRORE"
End Sub

Public Sub ChiudiDitta()
    On Local Error Resume Next
    Call MXVA.ChiudiDyTRAnagraf
    Call MXVA.ChiudiDyTRValidazione
    Call MXCT.ChiudiDyTRTabelle
    Call MXVI.ChiudiDyTRVisioni
    Call MXVI.ChiudiDyTRSituazioni
    'Call MXNU.SalvaImpostazioniUtente(MXNU.UtenteSistema)
    On Local Error GoTo 0
End Sub

Public Function InitChiaveHW() As Boolean
    ' oltre a dire se c'è la chiave, istanzia anche la proprietà MXNU.IDSessione ( e forse altre...)
    Dim blnRes As Boolean
    blnRes = MXNU.ChiavePresente()
    If Not blnRes Then
        Call MXNU.MsgBoxEX(9000, vbOKOnly + vbCritical, 1007)
    End If
    
    ' Testa la presenza dei moduli ISV necessari al programma
    If blnRes Then
        blnRes = ModuloPresente(150, "Run-time ISV")
    End If
    If blnRes Then
        blnRes = ModuloPresente(161, "Business Prima Nota ISV")
    End If
    If blnRes Then
        blnRes = ModuloPresente(162, "Business Scadenze ISV")
    End If
    If blnRes Then
        blnRes = ModuloPresente(163, "Business Storico Mag. ISV")
    End If
    If blnRes Then
        blnRes = ModuloPresente(164, "Business Documenti ISV")
    End If
    If blnRes Then
        blnRes = ModuloPresente(166, "Business Validazione Articolo ISV")
    End If
    
    InitChiaveHW = blnRes
End Function

Private Function ModuloPresente(ByVal intNrModulo As Integer, _
                                ByVal strDscModulo As String) As Boolean
    Dim blnRes As Boolean
    blnRes = (MXNU.ControlloModulichiave(intNrModulo) = 0)
    If Not blnRes Then
        Call MXNU.MsgBoxEX("Modulo nr. [" & intNrModulo & " - " & strDscModulo & "] non presente sulla chiave. Impossibile continuare!", vbOKOnly + vbCritical, 1007)
    End If
    ModuloPresente = blnRes
End Function

Sub LeggiVincoli()
Dim q As Integer
Dim hndtn As CRecordSet
Dim inti As Integer
Dim hTabEse As MXKit.CRecordSet     'Tabella esercizi

    Set hndtn = MXDB.dbCreaSS(hndDBArchivi, "SELECT * FROM TabVincoliGIC WHERE Esercizio=" & MXNU.AnnoAttivo)
    
    For inti = 1 To 5
        MXNU.VincoliIva(IVA_VEN, inti) = MXDB.dbGetCampo(hndtn, hndtn.Tipo, "SCIVADeb" & CStr(inti), "")
        MXNU.VincoliIva(IVA_ACQ, inti) = MXDB.dbGetCampo(hndtn, hndtn.Tipo, "SCIVACred" & CStr(inti), "")
        MXNU.VincoliIva(IVA_SOS, inti) = MXDB.dbGetCampo(hndtn, hndtn.Tipo, "SCIVASosp" & CStr(inti), "")
        MXNU.VincoliIva(IVA_VENINTRA, inti) = MXDB.dbGetCampo(hndtn, hndtn.Tipo, "SCIVAVendIntra" & CStr(inti), "")
        MXNU.VincoliIva(IVA_ACQINTRA, inti) = MXDB.dbGetCampo(hndtn, hndtn.Tipo, "SCIVAAcqIntra" & CStr(inti), "")
    Next inti
    
    MXNU.Vincoli(SC_CLI_CORRISP) = MXDB.dbGetCampo(hndtn, hndtn.Tipo, "SCCliCorrisp", "")
    MXNU.Vincoli(REG_IVA_INTRA) = CStr(MXDB.dbGetCampo(hndtn, hndtn.Tipo, "RegVendIntra", ""))
    MXNU.Vincoli(CAUS_INSOLUTO) = CStr(MXDB.dbGetCampo(hndtn, hndtn.Tipo, "CausContInsoluto", ""))
    MXNU.Vincoli(CAUS_APERTURA) = CStr(MXDB.dbGetCampo(hndtn, hndtn.Tipo, "CausContAp", ""))
    MXNU.Vincoli(CAUS_CHIUSURA) = CStr(MXDB.dbGetCampo(hndtn, hndtn.Tipo, "CausContCh", ""))
    MXNU.Vincoli(CONTO_PATR_APERTURA) = MXDB.dbGetCampo(hndtn, hndtn.Tipo, "ContoPatrAP", "")
    MXNU.Vincoli(CONTO_PATR_CHIUSURA) = MXDB.dbGetCampo(hndtn, hndtn.Tipo, "ContoPatrCH", "")
    MXNU.Vincoli(CONTO_ECO_CHIUSURA) = MXDB.dbGetCampo(hndtn, hndtn.Tipo, "ContoEcoCH", "")
    
    MXNU.CodCambioLire = MXDB.dbGetCampo(hndtn, hndtn.Tipo, "DivisaLire", 0)
    MXNU.CodCambioEuro = MXDB.dbGetCampo(hndtn, hndtn.Tipo, "DivisaEuro", 0)
    MXNU.DecimaliQuantita = MXDB.dbGetCampo(hndtn, hndtn.Tipo, "nDecimaliQuantita", 0)
    MXNU.DecimaliPesiVolumi = MXDB.dbGetCampo(hndtn, hndtn.Tipo, "nDecimaliPesiVol", 0)
    MXNU.DecimaliLireTotale = MXDB.dbGetCampo(hndtn, hndtn.Tipo, "nDecimaliTotaleLire", 0)
    MXNU.DecimaliLireUnitario = MXDB.dbGetCampo(hndtn, hndtn.Tipo, "nDecimaliUnitarioLire", 0)
    MXNU.FORMATO_QUANTITA = Formato("####,###,##0", MXNU.DecimaliQuantita)
    MXNU.FORMATO_PESIVOLUMI = Formato("####,###,##0", MXNU.DecimaliPesiVolumi)
    MXNU.FORMATO_LIRE_UNITARIO = Formato("####,###,###,##0", MXNU.DecimaliLireUnitario)
    MXNU.FORMATO_LIRE_TOTALE = Formato("####,###,###,##0", MXNU.DecimaliLireTotale)
    
    'Imposto nella proprietà del nucleo l'ultimo anno creato.
    Set hTabEse = MXDB.dbCreaSS(hndDBArchivi, "SELECT MAX(CODICE) AS ULTESE FROM TABESERCIZI")
        MXNU.UltimoEsercizioCreato = MXDB.dbGetCampo(hTabEse, hndtn.Tipo, "ULTESE", MXNU.AnnoAttivo)
    Call MXDB.dbChiudiSS(hTabEse)
    
    q = MXDB.dbChiudiSS(hndtn)
    
    If MXNU.CodCambioLire = MXNU.CodCambioEuro Then
        Call MXNU.MsgBoxEX(1399, vbCritical, 1007)
    End If
    Call GetFormatiEuro
    
    'lettura vincoli produzione
    strSql = "select NDECIMALICICLO from TABVINCOLIPRODUZIONE order by PROGRESSIVO desc"
    Set hndtn = MXDB.dbCreaSS(hndDBArchivi, strSql)
    q = MXDB.dbGetCampo(hndtn, hndtn.Tipo, "NDECIMALICICLO", 0)
    MXNU.Formato_Centesimi = Formato("####,###,##0", q)
    q = MXDB.dbChiudiSS(hndtn)
End Sub

Sub GetFormatiEuro()
    Dim hSS As CRecordSet, intDec As Integer, intq As Integer
    
    MXNU.FORMATO_EURO_UNITARIO = "###,###,##0.00"
    MXNU.FORMATO_EURO_TOTALE = "###,###,##0.00"
    MXNU.DecimaliEuroTotale = 2
    MXNU.DecimaliEuroUnitario = 2
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT nDecimaliUnitario,nDecimaliTotale FROM TabCambi WHERE Codice=(SELECT DivisaEuro FROM TabVincoliGIC WHERE Esercizio=" & MXNU.AnnoAttivo & ")")
    If Not MXDB.dbFineTab(hSS) Then
        intDec = MXDB.dbGetCampo(hSS, hSS.Tipo, "nDecimaliUnitario", 0)
        MXNU.DecimaliEuroUnitario = intDec
        MXNU.FORMATO_EURO_UNITARIO = Formato("####,###,###,##0", intDec)
        
        intDec = MXDB.dbGetCampo(hSS, NO_REPOSITION, "nDecimaliTotale", 0)
        MXNU.FORMATO_EURO_TOTALE = Formato("####,###,###,##0", intDec)
        MXNU.DecimaliEuroTotale = intDec
    End If
    intq = MXDB.dbChiudiSS(hSS)
    
    Set hSS = MXDB.dbCreaSS(hndDBArchivi, "SELECT CambioEuro FROM TabCambi WHERE Codice=0")
    If Not MXDB.dbFineTab(hSS) Then
        MXNU.CambioLireEuro = MXDB.dbGetCampo(hSS, hSS.Tipo, "CambioEuro", 1)
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

Public Function Ambienti2Collection(Optional ByVal bolSkipKey As Boolean = False) As Collection
Dim colAmb As Collection

    If (mColAmb Is Nothing) Then
        'creo la collezione degli ambienti
        Set colAmb = New Collection
        With colAmb
            .Add MXNU, "MXNU"
            If bolSkipKey Or (MXNU.ControlloModulichiave(MD32_KIT) = 0) Then
                .Add MXDB, "MXDB"
                .Add MXCREP, "MXCREP"
                .Add MXCT, "MXCT"
                .Add MXVI, "MXVI"
                .Add MXVA, "MXVA"
                .Add MXFT, "MXFT"
                If MXNU.ControlloModulichiave(modAgentiRunTime) = 0 Then .Add MXAA, "MXAA"
                .Add MXALL, "MXALL"
                .Add MXQM, "MXQM"
            End If
            If bolSkipKey Or (MXNU.ControlloModulichiave(MD32_BUSINESS_SCADENZE) = 0) Then .Add MXSC, "MXSC"
            If bolSkipKey Or (MXNU.ControlloModulichiave(MD32_BUSINESS_CTRLCODARTICOLO) = 0) Then .Add MXART, "MXART"
            If bolSkipKey Or (MXNU.ControlloModulichiave(MD32_BUSINESS_STORICO) = 0) Then .Add MXSM, "MXSM"
            If bolSkipKey Or (MXNU.ControlloModulichiave(MD32_BUSINESS_DBA) = 0) Then .Add MXDBA, "MXDBA"
            If bolSkipKey Or (MXNU.ControlloModulichiave(MD32_BUSINESS_DOCUMENTI) = 0) Then .Add MXGD, "MXGD"
            If bolSkipKey Or (MXNU.ControlloModulichiave(MD32_BUSINESS_PIANIFICAZIONE) = 0) Then .Add MXPIAN, "MXPIAN"
            If bolSkipKey Or (MXNU.ControlloModulichiave(MD32_BUSINESS_PRIMANOTA) = 0) Then .Add MXPN, "MXPN"
            If bolSkipKey Or (MXNU.ControlloModulichiave(MD32_BUSINESS_PRODUZIONE) = 0) Then .Add MXPROD, "MXPROD"
            If bolSkipKey Or (MXNU.ControlloModulichiave(MD32_BUSINESS_CICLILAVORAZIONE) = 0) Then .Add MXCICLI, "MXCICLI"
            If bolSkipKey Or (MXNU.ControlloModulichiave(MD32_BUSINESS_COMMESSECLIENTI) = 0) Then .Add MXCC, "MXCC"
            If bolSkipKey Or (MXNU.ControlloModulichiave(MD32_BUSINESS_GESTIONERISORSE) = 0) Then .Add MXRIS, "MXRIS"
        End With
        'e la bufferizzo
        Set mColAmb = colAmb
    Else
        Set colAmb = mColAmb
    End If
    
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

