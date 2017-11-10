Attribute VB_Name = "MetodoFunction"
Option Explicit

Public FM98 As CFMetodo98
Public blnStampaAnteprima As Boolean
Public strBarCode As String

Public Type Validvalori
    Campo As String
    Valore As String
End Type

Public Type Cliente
    Tipo As String
    CodConto As String
    DscConto1 As String
    DscConto2 As String
    CodMastro As Long
    Indirizzo As String
    Cap As String
    Localita As String
    Provincia As String
    Telefono As Integer
    Fax As String
    Mail As String
    CodFiscale As String
    PartitaIva As String
    Note As String
    codnazione As Integer
    codiceiso As String
    codlingua As Integer
    
    'Riservati
    NrListino As Long
    CodPagamento As Long
    CodSpedizione As Long
    CodAgente1 As String
    Sconto  As Integer
    ClienteFatturazione As String
End Type


Public DocumentiMetodo() As Testa_DOC

Public Type Riga_DOC
    blnAnalizza As Boolean
    idtesta As Long
    idRiga As Long
    TipoRiga As String
    DataConsegna As String
    CodArt As String
    Descrizioneart As String
    UMGest As String
    nRRifPartita As String
    QtaGest As Double
    Qta1Mag As Double
    PrezzoUnitLordo As Double
    Totale As Double
    ScontoEsteso As String
    Annotazioni As String
    CodDeposito As String
    Confronta As Boolean
    versioneDiBA As String
End Type

Public Type Testa_DOC
    blnAnalizza As Boolean
    TipoDoc As String
    TipoDocPartenza As String
    Bis As String
    Esercizio As Long
    progressivo As Long
    NumeroDoc As Long
    CodClifor As String
    CodCFFatt As String
    CodListino As Integer
    CodAgente1 As String
    CodPagamento As Integer
    Bloccato As String
    ModoTrasp As String
    ScontoFinale As Integer
    DestinazioneDiversa As Integer
    PrcScontoIncond As Integer
    DataDoc As String
    Annotazioni As String
    DataInizioTrasp As String
    NumRifDoc As String
    DataRifDoc As String
    Righe() As Riga_DOC
End Type

Public Function CreaCliente_STD(cli As Cliente) As String
Dim strsql As String
Dim lngContaClienti As Long
Dim blnFine As Boolean
On Error GoTo ERR_CreaCliente_STD
            CreaCliente_STD = ""
            lngContaClienti = 1
            blnFine = False
            While Not blnFine
                cli.CodConto = "C" & Right(String(6, " ") & lngContaClienti, 6)
                If Not TrovaCliente(cli.CodConto) Then
                    If Len(cli.ClienteFatturazione) = 0 Then cli.ClienteFatturazione = cli.CodConto
                    
                    strsql = " insert into anagraficacf (tipoconto,codconto,dscconto1,dscconto2,codmastro,indirizzo,localita,Cap,provincia,telefono,Fax"
                    strsql = strsql & " ,Telex,partitaiva,CodFiscale,Note"
                    strsql = strsql & " ,codnazione,codiceiso,codlingua,utentemodifica,datamodifica)"
                    strsql = strsql & " select " & hndDBArchivi.FormatoSQL(cli.Tipo, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.CodConto, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.DscConto1, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.DscConto2, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.CodMastro, DB_INTEGER) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.Indirizzo, DB_TEXT) & " ,"
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.Localita, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.Cap, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.Provincia, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.Telefono, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.Fax, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.Mail, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.PartitaIva, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.CodFiscale, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.Note, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.codnazione, DB_INTEGER) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.codiceiso, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.codlingua, DB_INTEGER) & ","
                    strsql = strsql & "'trm1',getdate()"
                    
                    MXDB.dbEseguiSQL hndDBArchivi, strsql
                    
                    strsql = " insert into anagraficaRiservaticf (esercizio,codconto,codpag,listino,codsped,codcontofatt,codAgente1,rivivaomaggi,usaprzprvpart,utentemodifica,datamodifica)"
                    strsql = strsql & " Select "
                    strsql = strsql & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.CodConto, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.CodPagamento, DB_INTEGER) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.NrListino, DB_INTEGER) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.CodSpedizione, DB_INTEGER) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.ClienteFatturazione, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.CodAgente1, DB_TEXT) & ","
                    strsql = strsql & "1,1,'trm1',getdate()"
                    
                    MXDB.dbEseguiSQL hndDBArchivi, strsql
                    
                    
                    strsql = " insert into ANAGRAFICACFGESTCONT (codconto,utentemodifica,datamodifica)"
                    strsql = strsql & " Select " & hndDBArchivi.FormatoSQL(cli.CodConto, DB_TEXT) & ",'trm1',getdate()"
                    
                    MXDB.dbEseguiSQL hndDBArchivi, strsql
                    
                    strsql = " insert into extraclienti(codconto,utentemodifica,datamodifica)"
                    strsql = strsql & " Select " & hndDBArchivi.FormatoSQL(cli.CodConto, DB_TEXT) & ",'trm1',getdate()"
                    
                    MXDB.dbEseguiSQL hndDBArchivi, strsql
                    blnFine = True
               End If
                lngContaClienti = lngContaClienti + 1
            Wend
            CreaCliente_STD = cli.CodConto
ERR_CreaCliente_STD:
    If Err <> 0 Then
        Resume Next
    End If
End Function

Public Function CreaCliente(cli As Cliente) As String
Dim strsql As String
On Error GoTo Err_CreaCliente
                If Not TrovaCliente(cli.CodConto) Then
                    strsql = " insert into anagraficacf (tipoconto,codconto,dscconto1,dscconto2,codmastro,indirizzo,localita,Cap,provincia,telefono,Fax"
                    strsql = strsql & " ,Telex,partitaiva,CodFiscale,Note"
                    strsql = strsql & " ,codnazione,codiceiso,codlingua,utentemodifica,datamodifica)"
                    strsql = strsql & " select " & hndDBArchivi.FormatoSQL(cli.Tipo, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.CodConto, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.DscConto1, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.DscConto2, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.CodMastro, DB_INTEGER) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.Indirizzo, DB_TEXT) & " ,"
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.Localita, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.Cap, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.Provincia, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.Telefono, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.Fax, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.Mail, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.PartitaIva, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.CodFiscale, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.Note, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.codnazione, DB_INTEGER) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.codiceiso, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.codlingua, DB_INTEGER) & ","
                    strsql = strsql & "'trm1',getdate()"
                    
                    MXDB.dbEseguiSQL hndDBArchivi, strsql
                    
                    strsql = " insert into anagraficaRiservaticf (esercizio,codconto,codpag,listino,codsped,codcontofatt,rivivaomaggi,usaprzprvpart,utentemodifica,datamodifica)"
                    strsql = strsql & " Select "
                    strsql = strsql & hndDBArchivi.FormatoSQL(MXNU.AnnoAttivo, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.CodConto, DB_TEXT) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.CodPagamento, DB_INTEGER) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.NrListino, DB_INTEGER) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.CodSpedizione, DB_INTEGER) & ","
                    strsql = strsql & hndDBArchivi.FormatoSQL(cli.ClienteFatturazione, DB_TEXT) & ","
                    strsql = strsql & "1,1,'trm1',getdate()"
                    
                    MXDB.dbEseguiSQL hndDBArchivi, strsql
                    
                    
                    strsql = " insert into ANAGRAFICACFGESTCONT (codconto,utentemodifica,datamodifica)"
                    strsql = strsql & " Select " & hndDBArchivi.FormatoSQL(cli.CodConto, DB_TEXT) & ",'trm1',getdate()"
                    
                    MXDB.dbEseguiSQL hndDBArchivi, strsql
               End If
Err_CreaCliente:
    If Err <> 0 Then
        Resume Next
    End If
End Function

Public Function TrovaCliente(strCodice As String) As Boolean
    Dim strsql As String
    Dim xTmp As MXKit.CRecordSet

    strsql = "select CodConto from anagraficacf where codconto=" & hndDBArchivi.FormatoSQL(strCodice, DB_TEXT)
    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strsql)
    If Len(MXDB.dbGetCampo(xTmp, xTmp.Tipo, "CodConto", "")) = 0 Then
        TrovaCliente = False
    Else
        TrovaCliente = True
    End If

    Call MXDB.dbChiudiSS(xTmp)
    Set xTmp = Nothing
End Function

Public Function CalcolaFormula(ByVal Formula As String) As Double
On Error GoTo Err_CalcolaFormula
   Dim cF As MXBusiness.CFormula
    Formula = Replace(Formula, "=", "")

   Set cF = New MXBusiness.CFormula
   If cF.ControllaErrori(Formula) Then
      CalcolaFormula = cF.Calcola(Formula)
   Else
        CalcolaFormula = 0
        MsgBox "Formula Errata [" & Formula & "]"
   End If
   Set cF = Nothing
Err_CalcolaFormula:
    If Err <> 0 Then
            
    End If
End Function




Public Function NuovaTesta(ByRef docs() As Testa_DOC) As Long
Dim lngNumero As Long
    lngNumero = UBound(docs) + 1
    ReDim Preserve docs(lngNumero)
    lngNumero = UBound(docs)
    InizializzaTestaDoc docs(lngNumero)
    Erase docs(lngNumero).Righe
    ReDim docs(lngNumero).Righe(0)
    InizializzaRigaDoc docs(lngNumero).Righe(0)
    NuovaTesta = UBound(docs)
End Function

Public Function NuovaRiga(ByRef doc As Testa_DOC) As Long
Dim lngNumero As Long
    lngNumero = UBound(doc.Righe) + 1
    ReDim Preserve doc.Righe(lngNumero)
    lngNumero = UBound(doc.Righe)
    InizializzaRigaDoc doc.Righe(lngNumero)
    NuovaRiga = UBound(doc.Righe)
End Function

Private Sub InizializzaTestaDoc(ByRef doc As Testa_DOC)
    doc.Annotazioni = ""
    doc.blnAnalizza = False
    doc.Bloccato = 0
    doc.CodAgente1 = ""
    doc.CodCFFatt = ""
    doc.CodClifor = ""
    doc.CodListino = -1
    doc.CodPagamento = 0
    doc.DataDoc = "01-01-1900"
    doc.DataInizioTrasp = "01-01-1900"
    doc.ModoTrasp = 0
    doc.PrcScontoIncond = 0
    doc.progressivo = 0
    doc.ScontoFinale = 0
    Erase doc.Righe
    ReDim doc.Righe(0)
End Sub


Private Sub InizializzaRigaDoc(ByRef riga As Riga_DOC)
    riga.Annotazioni = ""
    riga.blnAnalizza = False
    riga.CodArt = ""
    riga.DataConsegna = "01-01-1900"
    riga.idRiga = 0
    riga.idtesta = 0
    riga.PrezzoUnitLordo = 0
    riga.QtaGest = 0
    riga.ScontoEsteso = ""
    riga.TipoRiga = "N"
    riga.UMGest = ""
    riga.Confronta = False
End Sub

Public Sub AnnullaDocumentoSTD(lngIdProgressivo As Long)
Dim xObj As MXBusiness.CGestDoc
If lngIdProgressivo > 0 Then
Set xObj = MXGD.CreaCGestDoc(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
With xObj
        
        'Carico il documento
       Call .xTDoc.AssegnaCampo("PROGRESSIVO", lngIdProgressivo)
      ' Mi pongo in stato di modifica
       .Stato = GD_MODIFICA
        'Me.MousePointer = vbHourglass
        'vntVal = MobjTesta.GrInput("PROGRESSIVO").ValoreCorrente
        If xObj.Annulla Then
            'cancella immagini correlate
            Dim cImmagine As CImage
            On Local Error Resume Next
            Set cImmagine = New CImage
            cImmagine.EliminaImmagine "TESTEDOCUMENTI", lngIdProgressivo
            Set cImmagine = Nothing
            On Local Error GoTo 0
        Else
            'Me.MousePointer = vbNormal
            Exit Sub
        End If
        If Not xObj Is Nothing Then
            Call xObj.Termina
            Set xObj.RisultatoPrelAuto = Nothing
            Set xObj = Nothing
        End If
End With
End If
End Sub

Private Sub Errore(strTesto As String)
    AggiornaLog strTesto
End Sub



Public Function CreaDocumentoSTD(doc As Testa_DOC, Optional blnAnnullaRighe As Boolean = True, Optional blnConfrontaTotaliRiga As Boolean = False) As String
On Error Resume Next
'On Error GoTo err_CreaDocumentoSTD
Dim intNewEsercizio  As Integer
Dim lngNewNrDoc  As Long
Dim strNewBis  As String
Dim mCGestDoc As MXBusiness.CGestDoc
Dim RigaCorrente As Long
Dim NomeFileLog As String
Dim intFileLog As Integer
Dim intConta As Integer
Dim strErrore As String
'Dim Errore As Integer
NomeFileLog = MXNU.GetTempFile()
intFileLog = MXNU.ImpostaErroriSuLog(NomeFileLog, True)


'Errore 0
'Errore 1
        RigaCorrente = 1
'Errore 2
        Set mCGestDoc = MXGD.CreaCGestDoc(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
'Errore 3
            With mCGestDoc
'Errore 4
                If doc.progressivo > 0 Then
'Errore 5
                    'AggiornaLog "modifica"
                    
                    .Stato = GD_MODIFICA
'Errore 6
                    'costruzione testa documento
                    Call .xTDoc.AssegnaCampo("PROGRESSIVO", doc.progressivo)
'Errore 7
                    Call .MostraRighe
'Errore 8
                    If blnAnnullaRighe Then
'Errore 9
                            For intConta = 1 To .NumeroRighe
'Errore 10
                                DoEvents

'Errore 11
                                .RigaAttiva.RigaCorr = intConta
'Errore 12
                                .RigaAttiva.AnnullaRiga
'Errore 13
                            Next
'Errore 14
                    End If
'Errore 15
        Else
'Errore 16
                    'AggiornaLog "inserimento"
                    blnAnnullaRighe = True
'Errore 17
                    .Stato = GD_INSERIMENTO
'Errore 18
                    'costruzione testa documento
                    If Len(doc.TipoDoc) > 0 Then Call .xTDoc.AssegnaCampo("TIPODOC", doc.TipoDoc)
'Errore 19
                    'AggiornaLog doc.TipoDoc
                    If doc.Esercizio > 0 Then Call .xTDoc.AssegnaCampo("ESERCIZIO", doc.Esercizio)
'Errore 20
                    'AggiornaLog CStr(doc.Esercizio)
                    If doc.DataDoc <> "01-01-1900" And IsDate(doc.DataDoc) Then Call .xTDoc.AssegnaCampo("DATADOC", Format(doc.DataDoc, "DD-MM-YYYY"))
'Errore 22
                    
                    If Len(doc.NumRifDoc) > 0 Then Call .xTDoc.AssegnaCampo("NUMRIFDOC", doc.NumRifDoc)
                    'AggiornaLog doc.NumRifDoc
'Errore 21
                    'AggiornaLog doc.DataDoc
                    If doc.DataRifDoc <> "01-01-1900" And IsDate(doc.DataRifDoc) Then Call .xTDoc.AssegnaCampo("DATARIFDOC", Format(doc.DataRifDoc, "DD-MM-YYYY"))
                    'AggiornaLog doc.DataRifDoc
'Errore 23
                    If Len(doc.CodClifor) > 0 Then Call .xTDoc.AssegnaCampo("CODCLIFOR", doc.CodClifor)
                    'AggiornaLog doc.CodClifor
'Errore 24
                    If Len(doc.CodAgente1) > 0 Then Call .xTDoc.AssegnaCampo("CODAGENTE1", doc.CodAgente1)
                    'AggiornaLog doc.CodAgente1
'Errore 25
                    If Len(doc.Annotazioni) > 0 Then Call .xTDoc.AssegnaCampo("ANNOTAZIONI", Mid(doc.Annotazioni, 1, 250))
'Errore 26
                    If Len(doc.Bloccato) > 0 Then Call .xTDoc.AssegnaCampo("BLOCCATO", doc.Bloccato)
'Errore 27
                    If Len(doc.CodCFFatt) > 0 Then Call .xTDoc.AssegnaCampo("CODCFFATT", doc.CodCFFatt)
'Errore 28
                    If Len(doc.CodListino) > 0 Then Call .xTDoc.AssegnaCampo("CODLISTINO", doc.CodListino)
'Errore 29
'Errore "paga " & doc.CodPagamento
                    If Len(doc.CodPagamento) > 0 Then Call .xTDoc.AssegnaCampo("CODPAGAMENTO", doc.CodPagamento)
'Errore 30
                    If doc.DataInizioTrasp <> "01-01-1900" And IsDate(doc.DataInizioTrasp) Then Call .xTDoc.AssegnaCampo("DATAINIZIOTRASP", Format(doc.DataInizioTrasp, "DD-MM-YYYY"))
'Errore 31
                    If Len(doc.ModoTrasp) > 0 Then Call .xTDoc.AssegnaCampo("MODOTRASP", doc.ModoTrasp)
'Errore 32
                    If doc.PrcScontoIncond > 0 Then Call .xTDoc.AssegnaCampo("PRCSCONTOINCOND", doc.PrcScontoIncond)
'Errore 33
                    If doc.ScontoFinale > 0 Then Call .xTDoc.AssegnaCampo("SCONTOFINALE", doc.ScontoFinale)
'Errore 34
                End If
                
                'Attendi 1
'Errore 35
                RigaCorrente = .NumeroRighe + 1
'                'aggiungo intesatzione
'                .RigaAttiva.ValoreCampo(R_TIPORIGA, RigaCorrente, True) = "D"
'                .RigaAttiva.ValoreCampo(R_DESCRIZIONEART, RigaCorrente, True) = "Importazione ordine:" & doc.NumRifDoc
'                RigaCorrente = .NumeroRighe + 1
'Errore 36
                For intConta = LBound(doc.Righe) To UBound(doc.Righe)
'Errore 37
                    DoEvents
'Errore 38
                    If doc.Righe(intConta).blnAnalizza Then
'Errore 39
                             'Attendi 5
                             .RigaAttiva.RigaCorr = RigaCorrente
'Errore 40
'Errore "Riga corrente " & .RigaAttiva.RigaCorr
                             If Len(doc.Righe(intConta).TipoRiga) > 0 Then .RigaAttiva.ValoreCampo(R_TIPORIGA, RigaCorrente, True) = doc.Righe(intConta).TipoRiga
                             'AggiornaLog doc.Righe(intConta).TipoRiga
'Errore "Riga corrente " & .RigaAttiva.RigaCorr
'Errore 41
                            If Len(doc.Righe(intConta).CodArt) > 0 Then .RigaAttiva.ValoreCampo(R_CODART, RigaCorrente, True) = doc.Righe(intConta).CodArt
'Errore "Riga corrente " & .RigaAttiva.RigaCorr
'Errore 42
                             If Len(doc.Righe(intConta).Descrizioneart) > 0 Then .RigaAttiva.ValoreCampo(R_DESCRIZIONEART, RigaCorrente, True) = doc.Righe(intConta).Descrizioneart
'Errore "Riga corrente " & .RigaAttiva.RigaCorr
'Errore 43
                             'AggiornaLog doc.Righe(intConta).Codart
                             If Len(doc.Righe(intConta).UMGest) > 0 Then .RigaAttiva.ValoreCampo(R_UMGEST, RigaCorrente, True) = doc.Righe(intConta).UMGest
'Errore "Riga corrente " & .RigaAttiva.RigaCorr
'Errore 44
                             'AggiornaLog doc.Righe(intConta).UMGest
                             If doc.Righe(intConta).QtaGest >= 0 Then .RigaAttiva.ValoreCampo(R_QTAGEST, RigaCorrente, True) = doc.Righe(intConta).QtaGest
'Errore "Riga corrente " & .RigaAttiva.RigaCorr
'Errore 45
                             'AggiornaLog CStr(doc.Righe(intConta).QtaGest)
                             If doc.Righe(intConta).PrezzoUnitLordo >= 0 Then .RigaAttiva.ValoreCampo(R_PREZZOUNITLORDO, RigaCorrente, True) = doc.Righe(intConta).PrezzoUnitLordo
'Errore "Riga corrente " & .RigaAttiva.RigaCorr
'Errore 46
                             
                             If doc.Righe(intConta).Confronta And doc.Righe(intConta).TipoRiga <> "O" Then
'Errore "Riga corrente " & .RigaAttiva.RigaCorr
'Errore 47
                                If doc.Righe(intConta).ScontoEsteso <> .RigaAttiva.ValoreCampo(R_SCONTIESTESI, RigaCorrente, True) Then
'Errore "Riga corrente " & .RigaAttiva.RigaCorr
'Errore 48
                                    If Len(doc.Righe(intConta).ScontoEsteso) = 0 And .RigaAttiva.ValoreCampo(R_SCONTIESTESI, RigaCorrente, True) = "0" Then
'Errore 49
                                    Else
'Errore 50
                                        AggiornaLog "ATTENZIONE:Discrepanza tra gli sconti importati per l'articolo " & doc.Righe(intConta).CodArt & ", Metodo:(" & .RigaAttiva.ValoreCampo(R_SCONTIESTESI, RigaCorrente, True) & ") Terminalino:(" & doc.Righe(intConta).ScontoEsteso & ")"
'Errore 51
                                    End If
'Errore 52
                                End If
'Errore 53
                                
                                If Fix(doc.Righe(intConta).Totale) <> Fix(.RigaAttiva.ValoreCampo(R_TOTNETTORIGAEURO, RigaCorrente, True)) Then
'Errore 54
                                   AggiornaLog "ATTENZIONE:Discrepanza tra i prezzi importati per l'articolo " & doc.Righe(intConta).CodArt & ", Metodo:" & Fix(.RigaAttiva.ValoreCampo(R_TOTNETTORIGAEURO, RigaCorrente, True)) & " Terminalino:" & Fix(doc.Righe(intConta).Totale)
'Errore 55
                                End If
'Errore 56
                                
                             End If
'Errore 57
                             If Len(doc.Righe(intConta).ScontoEsteso) > 0 Then .RigaAttiva.ValoreCampo(R_SCONTIESTESI, RigaCorrente, True) = doc.Righe(intConta).ScontoEsteso
'Errore 58
                             If Len(doc.Righe(intConta).Annotazioni) > 0 Then .RigaAttiva.ValoreCampo(R_ANNOTAZIONI, RigaCorrente, True) = doc.Righe(intConta).Annotazioni
'Errore 59
                             If IsDate(doc.Righe(intConta).DataConsegna) And doc.Righe(intConta).DataConsegna <> "01-01-1900" Then
'Errore 60
                                 .RigaAttiva.ValoreCampo(R_DATACONSEGNA, RigaCorrente, True) = Format(doc.Righe(intConta).DataConsegna, "DD-MM-YYYY")
'Errore 61
                             End If
'Errore 62
                             If blnConfrontaTotaliRiga Then
'Errore 63
                                'verificare chi utilizza questo confronto... e perchè dovrebbe confrontare il campo
'Errore 64
                                'R_TOTLORDOPREL
                                If Round(doc.Righe(intConta).Totale, 2) <> .RigaAttiva.ValoreCampo(R_TOTNETTORIGAEURO, RigaCorrente, True) Then
'Errore 65
                                   AggiornaLog "ATTENZIONE:Discrepanza tra i prezzi importati, Metodo:" & .RigaAttiva.ValoreCampo(R_TOTNETTORIGAEURO, RigaCorrente, True) & " Terminalino:" & Round(doc.Righe(intConta).Totale, 2)
'Errore 66
                                End If
'Errore 67
                             End If
'Errore 68
                             RigaCorrente = RigaCorrente + 1
'Errore 69
                              
                    End If
'Errore 70
                Next
'Errore 71
                Call .Calcolo_Totali
'Errore 72
                'registrazione documento
                intNewEsercizio = .xTDoc.grinput("ESERCIZIO").ValoreCorrente
'Errore 73
                lngNewNrDoc = .xTDoc.grinput("NUMERODOC").ValoreCorrente
'Errore 74
                strNewBis = .xTDoc.grinput("BIS").ValoreCorrente
'Errore 75
                If .Salva(intNewEsercizio, lngNewNrDoc, strNewBis) Then
'Errore 76
                    doc.progressivo = .xTDoc.grinput("Progressivo").ValoreCorrente
                    lngNewNrDoc = GetNumeroDoc(.xTDoc.grinput("Progressivo").ValoreCorrente)
                    CreaDocumentoSTD = "GENERATO:" & doc.TipoDoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
                    AggiornaLog CreaDocumentoSTD
                    'AggiornaLog "Dettaglio" & LeggiLog(NomeFileLog)
                Else
                    CreaDocumentoSTD = "ERRORE:" & doc.TipoDoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
                    AggiornaLog CreaDocumentoSTD
                    AggiornaLog "Vedi Errore " & NomeFileLog
                    doc.progressivo = 0
                    'AggiornaLog "Dettaglio" & LeggiLog(NomeFileLog)
                End If
            End With
            If Not mCGestDoc Is Nothing Then
                Call mCGestDoc.Termina
                Set mCGestDoc = Nothing
            End If
    
    AggiornaLog NomeFileLog
   Call MXNU.ChiudiErroriSuLog


'err_CreaDocumentoSTD:
'
'
'
'If Err <> 0 Then
'    strErrore = Err.Description
'    Errore strErrore
'    'MsgBox Err.Description, , "xxxx"
'    If Not mCGestDoc Is Nothing Then
'        Call mCGestDoc.Termina
'        Set mCGestDoc = Nothing
'    End If
'    CreaDocumentoSTD = "ERRORE:" & doc.TipoDoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
'    Resume Next
'End If


End Function



'Public Function InitCreaDocumentoSTD(doc As Testa_DOC, Optional blnAnnullaRighe As Boolean = True, Optional blnConfrontaTotaliRiga As Boolean = False) As String
'On Error Resume Next
''On Error GoTo err_CreaDocumentoSTD
'Dim intNewEsercizio  As Integer
'Dim lngNewNrDoc  As Long
'Dim strNewBis  As String
'Dim mCGestDoc As MXBusiness.CGestDoc
'Dim RigaCorrente As Long
'Dim NomeFileLog As String
'Dim intFileLog As Integer
'Dim intConta As Integer
'Dim strErrore As String
''Dim Errore As Integer
'NomeFileLog = MXNU.GetTempFile()
'intFileLog = MXNU.ImpostaErroriSuLog(NomeFileLog, True)
'
'
''Errore 0
''Errore 1
'        RigaCorrente = 1
''Errore 2
'        Set mCGestDoc = MXGD.CreaCGestDoc(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
''Errore 3
'            With mCGestDoc
''Errore 4
'                If doc.Progressivo > 0 Then
''Errore 5
'                    'AggiornaLog "modifica"
'
'                    .Stato = GD_MODIFICA
''Errore 6
'                    'costruzione testa documento
'                    Call .xTDoc.AssegnaCampo("PROGRESSIVO", doc.Progressivo)
''Errore 7
'                    Call .MostraRighe
''Errore 8
'                    If blnAnnullaRighe Then
''Errore 9
'                            For intConta = 1 To .NumeroRighe
''Errore 10
'                                DoEvents
'
''Errore 11
'                                .RigaAttiva.RigaCorr = intConta
''Errore 12
'                                .RigaAttiva.AnnullaRiga
''Errore 13
'                            Next
''Errore 14
'                    End If
''Errore 15
'        Else
''Errore 16
'                    'AggiornaLog "inserimento"
'                    blnAnnullaRighe = True
''Errore 17
'                    .Stato = GD_INSERIMENTO
''Errore 18
'                    'costruzione testa documento
'                    If Len(doc.TipoDoc) > 0 Then Call .xTDoc.AssegnaCampo("TIPODOC", doc.TipoDoc)
''Errore 19
'                    'AggiornaLog doc.TipoDoc
'                    If doc.Esercizio > 0 Then Call .xTDoc.AssegnaCampo("ESERCIZIO", doc.Esercizio)
''Errore 20
'                    'AggiornaLog CStr(doc.Esercizio)
'                    If doc.DataDoc <> "01-01-1900" And IsDate(doc.DataDoc) Then Call .xTDoc.AssegnaCampo("DATADOC", Format(doc.DataDoc, "DD-MM-YYYY"))
''Errore 22
'
'                    If Len(doc.NumRifDoc) > 0 Then Call .xTDoc.AssegnaCampo("NUMRIFDOC", doc.NumRifDoc)
'                    'AggiornaLog doc.NumRifDoc
''Errore 21
'                    'AggiornaLog doc.DataDoc
'                    If doc.DataRifDoc <> "01-01-1900" And IsDate(doc.DataRifDoc) Then Call .xTDoc.AssegnaCampo("DATARIFDOC", Format(doc.DataRifDoc, "DD-MM-YYYY"))
'                    'AggiornaLog doc.DataRifDoc
''Errore 23
'                    If Len(doc.CodClifor) > 0 Then Call .xTDoc.AssegnaCampo("CODCLIFOR", doc.CodClifor)
'                    'AggiornaLog doc.CodClifor
''Errore 24
'                    If Len(doc.CodAgente1) > 0 Then Call .xTDoc.AssegnaCampo("CODAGENTE1", doc.CodAgente1)
'                    'AggiornaLog doc.CodAgente1
''Errore 25
'                    If Len(doc.Annotazioni) > 0 Then Call .xTDoc.AssegnaCampo("ANNOTAZIONI", Mid(doc.Annotazioni, 1, 250))
''Errore 26
'                    If Len(doc.Bloccato) > 0 Then Call .xTDoc.AssegnaCampo("BLOCCATO", doc.Bloccato)
''Errore 27
'                    If Len(doc.CodCFFatt) > 0 Then Call .xTDoc.AssegnaCampo("CODCFFATT", doc.CodCFFatt)
''Errore 28
'                    If Len(doc.CodListino) > 0 Then Call .xTDoc.AssegnaCampo("CODLISTINO", doc.CodListino)
''Errore 29
''Errore "paga " & doc.CodPagamento
'                    If Len(doc.CodPagamento) > 0 Then Call .xTDoc.AssegnaCampo("CODPAGAMENTO", doc.CodPagamento)
''Errore 30
'                    If doc.DataInizioTrasp <> "01-01-1900" And IsDate(doc.DataInizioTrasp) Then Call .xTDoc.AssegnaCampo("DATAINIZIOTRASP", Format(doc.DataInizioTrasp, "DD-MM-YYYY"))
''Errore 31
'                    If Len(doc.ModoTrasp) > 0 Then Call .xTDoc.AssegnaCampo("MODOTRASP", doc.ModoTrasp)
''Errore 32
'                    If doc.PrcScontoIncond > 0 Then Call .xTDoc.AssegnaCampo("PRCSCONTOINCOND", doc.PrcScontoIncond)
''Errore 33
'                    If doc.ScontoFinale > 0 Then Call .xTDoc.AssegnaCampo("SCONTOFINALE", doc.ScontoFinale)
''Errore 34
'                End If
'
'                'Attendi 1
''Errore 35
'                RigaCorrente = .NumeroRighe + 1
''                'aggiungo intesatzione
''                .RigaAttiva.ValoreCampo(R_TIPORIGA, RigaCorrente, True) = "D"
''                .RigaAttiva.ValoreCampo(R_DESCRIZIONEART, RigaCorrente, True) = "Importazione ordine:" & doc.NumRifDoc
''                RigaCorrente = .NumeroRighe + 1
''Errore 36
'                For intConta = LBound(doc.Righe) To UBound(doc.Righe)
''Errore 37
'                    DoEvents
''Errore 38
'                    If doc.Righe(intConta).blnAnalizza Then
''Errore 39
'                             'Attendi 5
'                             .RigaAttiva.RigaCorr = RigaCorrente
''Errore 40
''Errore "Riga corrente " & .RigaAttiva.RigaCorr
'                             If Len(doc.Righe(intConta).TipoRiga) > 0 Then .RigaAttiva.ValoreCampo(R_TIPORIGA, RigaCorrente, True) = doc.Righe(intConta).TipoRiga
'                             'AggiornaLog doc.Righe(intConta).TipoRiga
''Errore "Riga corrente " & .RigaAttiva.RigaCorr
''Errore 41
'                            If Len(doc.Righe(intConta).CodArt) > 0 Then .RigaAttiva.ValoreCampo(R_CODART, RigaCorrente, True) = doc.Righe(intConta).CodArt
''Errore "Riga corrente " & .RigaAttiva.RigaCorr
''Errore 42
'                             If Len(doc.Righe(intConta).Descrizioneart) > 0 Then .RigaAttiva.ValoreCampo(R_DESCRIZIONEART, RigaCorrente, True) = doc.Righe(intConta).Descrizioneart
''Errore "Riga corrente " & .RigaAttiva.RigaCorr
''Errore 43
'                             'AggiornaLog doc.Righe(intConta).Codart
'                             If Len(doc.Righe(intConta).UMGest) > 0 Then .RigaAttiva.ValoreCampo(R_UMGEST, RigaCorrente, True) = doc.Righe(intConta).UMGest
''Errore "Riga corrente " & .RigaAttiva.RigaCorr
''Errore 44
'                             'AggiornaLog doc.Righe(intConta).UMGest
'                             If doc.Righe(intConta).QtaGest >= 0 Then .RigaAttiva.ValoreCampo(R_QTAGEST, RigaCorrente, True) = doc.Righe(intConta).QtaGest
''Errore "Riga corrente " & .RigaAttiva.RigaCorr
''Errore 45
'                             'AggiornaLog CStr(doc.Righe(intConta).QtaGest)
'                             If doc.Righe(intConta).PrezzoUnitLordo >= 0 Then .RigaAttiva.ValoreCampo(R_PREZZOUNITLORDO, RigaCorrente, True) = doc.Righe(intConta).PrezzoUnitLordo
''Errore "Riga corrente " & .RigaAttiva.RigaCorr
''Errore 46
'
'                             If doc.Righe(intConta).Confronta And doc.Righe(intConta).TipoRiga <> "O" Then
''Errore "Riga corrente " & .RigaAttiva.RigaCorr
''Errore 47
'                                If doc.Righe(intConta).ScontoEsteso <> .RigaAttiva.ValoreCampo(R_SCONTIESTESI, RigaCorrente, True) Then
''Errore "Riga corrente " & .RigaAttiva.RigaCorr
''Errore 48
'                                    If Len(doc.Righe(intConta).ScontoEsteso) = 0 And .RigaAttiva.ValoreCampo(R_SCONTIESTESI, RigaCorrente, True) = "0" Then
''Errore 49
'                                    Else
''Errore 50
'                                        AggiornaLog "ATTENZIONE:Discrepanza tra gli sconti importati per l'articolo " & doc.Righe(intConta).CodArt & ", Metodo:(" & .RigaAttiva.ValoreCampo(R_SCONTIESTESI, RigaCorrente, True) & ") Terminalino:(" & doc.Righe(intConta).ScontoEsteso & ")"
''Errore 51
'                                    End If
''Errore 52
'                                End If
''Errore 53
'
'                                If Fix(doc.Righe(intConta).Totale) <> Fix(.RigaAttiva.ValoreCampo(R_TOTNETTORIGAEURO, RigaCorrente, True)) Then
''Errore 54
'                                   AggiornaLog "ATTENZIONE:Discrepanza tra i prezzi importati per l'articolo " & doc.Righe(intConta).CodArt & ", Metodo:" & Fix(.RigaAttiva.ValoreCampo(R_TOTNETTORIGAEURO, RigaCorrente, True)) & " Terminalino:" & Fix(doc.Righe(intConta).Totale)
''Errore 55
'                                End If
''Errore 56
'
'                             End If
''Errore 57
'                             If Len(doc.Righe(intConta).ScontoEsteso) > 0 Then .RigaAttiva.ValoreCampo(R_SCONTIESTESI, RigaCorrente, True) = doc.Righe(intConta).ScontoEsteso
''Errore 58
'                             If Len(doc.Righe(intConta).Annotazioni) > 0 Then .RigaAttiva.ValoreCampo(R_ANNOTAZIONI, RigaCorrente, True) = doc.Righe(intConta).Annotazioni
''Errore 59
'                             If IsDate(doc.Righe(intConta).DataConsegna) And doc.Righe(intConta).DataConsegna <> "01-01-1900" Then
''Errore 60
'                                 .RigaAttiva.ValoreCampo(R_DATACONSEGNA, RigaCorrente, True) = Format(doc.Righe(intConta).DataConsegna, "DD-MM-YYYY")
''Errore 61
'                             End If
''Errore 62
'                             If blnConfrontaTotaliRiga Then
''Errore 63
'                                'verificare chi utilizza questo confronto... e perchè dovrebbe confrontare il campo
''Errore 64
'                                'R_TOTLORDOPREL
'                                If Round(doc.Righe(intConta).Totale, 2) <> .RigaAttiva.ValoreCampo(R_TOTNETTORIGAEURO, RigaCorrente, True) Then
''Errore 65
'                                   AggiornaLog "ATTENZIONE:Discrepanza tra i prezzi importati, Metodo:" & .RigaAttiva.ValoreCampo(R_TOTNETTORIGAEURO, RigaCorrente, True) & " Terminalino:" & Round(doc.Righe(intConta).Totale, 2)
''Errore 66
'                                End If
''Errore 67
'                             End If
''Errore 68
'                             RigaCorrente = RigaCorrente + 1
''Errore 69
'
'                    End If
''Errore 70
'                Next
''Errore 71
'                Call .Calcolo_Totali
''Errore 72
'                'registrazione documento
'                intNewEsercizio = .xTDoc.GrInput("ESERCIZIO").ValoreCorrente
''Errore 73
'                lngNewNrDoc = .xTDoc.GrInput("NUMERODOC").ValoreCorrente
''Errore 74
'                strNewBis = .xTDoc.GrInput("BIS").ValoreCorrente
''Errore 75
'                If .Salva(intNewEsercizio, lngNewNrDoc, strNewBis) Then
''Errore 76
'                    lngNewNrDoc = GetNumeroDoc(.xTDoc.GrInput("Progressivo").ValoreCorrente)
'                    CreaDocumentoSTD = "GENERATO:" & doc.TipoDoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
'                    AggiornaLog CreaDocumentoSTD
'                    'AggiornaLog "Dettaglio" & LeggiLog(NomeFileLog)
'                Else
'                    CreaDocumentoSTD = "ERRORE:" & doc.TipoDoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
'                    AggiornaLog CreaDocumentoSTD
'                    AggiornaLog "Vedi Errore " & NomeFileLog
'                    'AggiornaLog "Dettaglio" & LeggiLog(NomeFileLog)
'                End If
'            End With
'            If Not mCGestDoc Is Nothing Then
'                Call mCGestDoc.Termina
'                Set mCGestDoc = Nothing
'            End If
'
'    AggiornaLog NomeFileLog
'   Call MXNU.ChiudiErroriSuLog
'
'
''err_CreaDocumentoSTD:
''
''
''
''If Err <> 0 Then
''    strErrore = Err.Description
''    Errore strErrore
''    'MsgBox Err.Description, , "xxxx"
''    If Not mCGestDoc Is Nothing Then
''        Call mCGestDoc.Termina
''        Set mCGestDoc = Nothing
''    End If
''    CreaDocumentoSTD = "ERRORE:" & doc.TipoDoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
''    Resume Next
''End If
'
'
'End Function




Public Function CreaDocumentoSTD_test(doc As Testa_DOC, Optional blnAnnullaRighe As Boolean = True, Optional blnConfrontaTotaliRiga As Boolean = False) As String
On Error Resume Next
On Error GoTo err_CreaDocumentoSTD_test
Dim intNewEsercizio  As Integer
Dim lngNewNrDoc  As Long
Dim strNewBis  As String
Dim mCGestDoc As MXBusiness.CGestDoc
Dim RigaCorrente As Long
Dim NomeFileLog As String
Dim intFileLog As Integer
Dim intConta As Integer
Dim strErrore As String
'Dim Errore As Integer
'NomeFileLog = MXNU.GetTempFile()
'intFileLog = MXNU.ImpostaErroriSuLog(NomeFileLog, True)


Errore 0
Errore 1
        RigaCorrente = 1
Errore 2
        Set mCGestDoc = MXGD.CreaCGestDoc(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
Errore 3
            With mCGestDoc
Errore 4
                If doc.progressivo > 0 Then
Errore 5
                    'AggiornaLog "modifica"
                    
                    .Stato = GD_MODIFICA
Errore 6
                    'costruzione testa documento
                    Call .xTDoc.AssegnaCampo("PROGRESSIVO", doc.progressivo)
Errore 7
                    Call .MostraRighe
Errore 8
                    If blnAnnullaRighe Then
Errore 9
                            For intConta = 1 To .NumeroRighe
Errore 10
                                DoEvents

Errore 11
                                .RigaAttiva.RigaCorr = intConta
Errore 12
                                .RigaAttiva.AnnullaRiga
Errore 13
                            Next
Errore 14
                    End If
Errore 15
        Else
Errore 16
                    'AggiornaLog "inserimento"
                    blnAnnullaRighe = True
Errore 17
                    .Stato = GD_INSERIMENTO
Errore 18
                    'costruzione testa documento
                    If Len(doc.TipoDoc) > 0 Then Call .xTDoc.AssegnaCampo("TIPODOC", doc.TipoDoc)
Errore 19
                    'AggiornaLog doc.TipoDoc
                    If doc.Esercizio > 0 Then Call .xTDoc.AssegnaCampo("ESERCIZIO", doc.Esercizio)
Errore 20
                    'AggiornaLog CStr(doc.Esercizio)
                    If doc.DataDoc <> "01-01-1900" And IsDate(doc.DataDoc) Then Call .xTDoc.AssegnaCampo("DATADOC", Format(doc.DataDoc, "DD-MM-YYYY"))
Errore 22
                    
                    If Len(doc.NumRifDoc) > 0 Then Call .xTDoc.AssegnaCampo("NUMRIFDOC", doc.NumRifDoc)
                    'AggiornaLog doc.NumRifDoc
Errore 21
                    'AggiornaLog doc.DataDoc
                    If doc.DataRifDoc <> "01-01-1900" And IsDate(doc.DataRifDoc) Then Call .xTDoc.AssegnaCampo("DATARIFDOC", Format(doc.DataRifDoc, "DD-MM-YYYY"))
                    'AggiornaLog doc.DataRifDoc
Errore 23
                    If Len(doc.CodClifor) > 0 Then Call .xTDoc.AssegnaCampo("CODCLIFOR", doc.CodClifor)
                    'AggiornaLog doc.CodClifor
Errore 24
                    If Len(doc.CodAgente1) > 0 Then Call .xTDoc.AssegnaCampo("CODAGENTE1", doc.CodAgente1)
                    'AggiornaLog doc.CodAgente1
Errore 25
                    If Len(doc.Annotazioni) > 0 Then Call .xTDoc.AssegnaCampo("ANNOTAZIONI", Mid(doc.Annotazioni, 1, 250))
Errore 26
                    If Len(doc.Bloccato) > 0 Then Call .xTDoc.AssegnaCampo("BLOCCATO", doc.Bloccato)
Errore 27
                    If Len(doc.CodCFFatt) > 0 Then Call .xTDoc.AssegnaCampo("CODCFFATT", doc.CodCFFatt)
Errore 28
                    If Len(doc.CodListino) > 0 Then Call .xTDoc.AssegnaCampo("CODLISTINO", doc.CodListino)
Errore 29
Errore "paga " & doc.CodPagamento
                    If Len(doc.CodPagamento) > 0 Then Call .xTDoc.AssegnaCampo("CODPAGAMENTO", doc.CodPagamento)
Errore 30
                    If doc.DataInizioTrasp <> "01-01-1900" And IsDate(doc.DataInizioTrasp) Then Call .xTDoc.AssegnaCampo("DATAINIZIOTRASP", Format(doc.DataInizioTrasp, "DD-MM-YYYY"))
Errore 31
                    If Len(doc.ModoTrasp) > 0 Then Call .xTDoc.AssegnaCampo("MODOTRASP", doc.ModoTrasp)
Errore 32
                    If doc.PrcScontoIncond > 0 Then Call .xTDoc.AssegnaCampo("PRCSCONTOINCOND", doc.PrcScontoIncond)
Errore 33
                    If doc.ScontoFinale > 0 Then Call .xTDoc.AssegnaCampo("SCONTOFINALE", doc.ScontoFinale)
Errore 34
                End If
                
                'Attendi 1
Errore 35
                RigaCorrente = .NumeroRighe + 1
                'aggiungo intesatzione
Errore 36
                For intConta = LBound(doc.Righe) To UBound(doc.Righe)
Errore 37
                    DoEvents
Errore 38
                    If doc.blnAnalizza Then
Errore 39
                             'Attendi 5
                             .RigaAttiva.RigaCorr = RigaCorrente
Errore 40
Errore "Riga corrente " & .RigaAttiva.RigaCorr
                             If Len(doc.Righe(intConta).TipoRiga) > 0 Then .RigaAttiva.ValoreCampo(R_TIPORIGA, RigaCorrente, True) = doc.Righe(intConta).TipoRiga
                             'AggiornaLog doc.Righe(intConta).TipoRiga
Errore "Riga corrente " & .RigaAttiva.RigaCorr
Errore 41
                            If Len(doc.Righe(intConta).CodArt) > 0 Then .RigaAttiva.ValoreCampo(R_CODART, RigaCorrente, True) = doc.Righe(intConta).CodArt
Errore "Riga corrente " & .RigaAttiva.RigaCorr
Errore 42
                             If Len(doc.Righe(intConta).Descrizioneart) > 0 Then .RigaAttiva.ValoreCampo(R_DESCRIZIONEART, RigaCorrente, True) = doc.Righe(intConta).Descrizioneart
Errore "Riga corrente " & .RigaAttiva.RigaCorr
Errore 43
                             'AggiornaLog doc.Righe(intConta).Codart
                             If Len(doc.Righe(intConta).UMGest) > 0 Then .RigaAttiva.ValoreCampo(R_UMGEST, RigaCorrente, True) = doc.Righe(intConta).UMGest
Errore "Riga corrente " & .RigaAttiva.RigaCorr
Errore 44
                             'AggiornaLog doc.Righe(intConta).UMGest
                             If doc.Righe(intConta).QtaGest >= 0 Then .RigaAttiva.ValoreCampo(R_QTAGEST, RigaCorrente, True) = doc.Righe(intConta).QtaGest
Errore "Riga corrente " & .RigaAttiva.RigaCorr
Errore 45
                             'AggiornaLog CStr(doc.Righe(intConta).QtaGest)
                             If doc.Righe(intConta).PrezzoUnitLordo >= 0 Then .RigaAttiva.ValoreCampo(R_PREZZOUNITLORDO, RigaCorrente, True) = doc.Righe(intConta).PrezzoUnitLordo
Errore "Riga corrente " & .RigaAttiva.RigaCorr
Errore 46
                             
                             If doc.Righe(intConta).Confronta And doc.Righe(intConta).TipoRiga <> "O" Then
Errore "Riga corrente " & .RigaAttiva.RigaCorr
Errore 47
                                If doc.Righe(intConta).ScontoEsteso <> .RigaAttiva.ValoreCampo(R_SCONTIESTESI, RigaCorrente, True) Then
Errore "Riga corrente " & .RigaAttiva.RigaCorr
Errore 48
                                    If Len(doc.Righe(intConta).ScontoEsteso) = 0 And .RigaAttiva.ValoreCampo(R_SCONTIESTESI, RigaCorrente, True) = "0" Then
Errore 49
                                    Else
Errore 50
                                        AggiornaLog "ATTENZIONE:Discrepanza tra gli sconti importati per l'articolo " & doc.Righe(intConta).CodArt & ", Metodo:(" & .RigaAttiva.ValoreCampo(R_SCONTIESTESI, RigaCorrente, True) & ") Terminalino:(" & doc.Righe(intConta).ScontoEsteso & ")"
Errore 51
                                    End If
Errore 52
                                End If
Errore 53
                                
                                If Fix(doc.Righe(intConta).Totale) <> Fix(.RigaAttiva.ValoreCampo(R_TOTNETTORIGAEURO, RigaCorrente, True)) Then
Errore 54
                                   AggiornaLog "ATTENZIONE:Discrepanza tra i prezzi importati per l'articolo " & doc.Righe(intConta).CodArt & ", Metodo:" & Fix(.RigaAttiva.ValoreCampo(R_TOTNETTORIGAEURO, RigaCorrente, True)) & " Terminalino:" & Fix(doc.Righe(intConta).Totale)
Errore 55
                                End If
Errore 56
                                
                             End If
Errore 57
                             If Len(doc.Righe(intConta).ScontoEsteso) > 0 Then .RigaAttiva.ValoreCampo(R_SCONTIESTESI, RigaCorrente, True) = doc.Righe(intConta).ScontoEsteso
Errore 58
                             If Len(doc.Righe(intConta).Annotazioni) > 0 Then .RigaAttiva.ValoreCampo(R_ANNOTAZIONI, RigaCorrente, True) = doc.Righe(intConta).Annotazioni
Errore 59
                             If IsDate(doc.Righe(intConta).DataConsegna) And doc.Righe(intConta).DataConsegna <> "01-01-1900" Then
Errore 60
                                 .RigaAttiva.ValoreCampo(R_DATACONSEGNA, RigaCorrente, True) = Format(doc.Righe(intConta).DataConsegna, "DD-MM-YYYY")
Errore 61
                             End If
Errore 62
                             If blnConfrontaTotaliRiga Then
'Errore 63
                                'verificare chi utilizza questo confronto... e perchè dovrebbe confrontare il campo
'Errore 64
                                'R_TOTLORDOPREL
                                If Round(doc.Righe(intConta).Totale, 2) <> .RigaAttiva.ValoreCampo(R_TOTNETTORIGAEURO, RigaCorrente, True) Then
'Errore 65
                                   AggiornaLog "ATTENZIONE:Discrepanza tra i prezzi importati, Metodo:" & .RigaAttiva.ValoreCampo(R_TOTNETTORIGAEURO, RigaCorrente, True) & " Terminalino:" & Round(doc.Righe(intConta).Totale, 2)
'Errore 66
                                End If
'Errore 67
                             End If
'Errore 68
                             RigaCorrente = RigaCorrente + 1
'Errore 69
                              
                    End If
'Errore 70
                Next
'Errore 71
                Call .Calcolo_Totali
'Errore 72
                'registrazione documento
                intNewEsercizio = .xTDoc.grinput("ESERCIZIO").ValoreCorrente
'Errore 73
                lngNewNrDoc = .xTDoc.grinput("NUMERODOC").ValoreCorrente
'Errore 74
                strNewBis = .xTDoc.grinput("BIS").ValoreCorrente
'Errore 75
                If .Salva(intNewEsercizio, lngNewNrDoc, strNewBis) Then
'Errore 76
                    lngNewNrDoc = GetNumeroDoc(.xTDoc.grinput("Progressivo").ValoreCorrente)
                    CreaDocumentoSTD_test = "GENERATO:" & doc.TipoDoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
                    AggiornaLog CreaDocumentoSTD_test
                Else
                    CreaDocumentoSTD_test = "ERRORE:" & doc.TipoDoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
                    AggiornaLog CreaDocumentoSTD_test
                    AggiornaLog "Vedi Errore " & NomeFileLog
                End If
            End With
            If Not mCGestDoc Is Nothing Then
                Call mCGestDoc.Termina
                Set mCGestDoc = Nothing
            End If
    
'    AggiornaLog NomeFileLog
'   Call MXNU.ChiudiErroriSuLog


err_CreaDocumentoSTD_test:



If Err <> 0 Then
    strErrore = Err.Description
    Errore strErrore
    'MsgBox Err.Description, , "xxxx"
'    If Not mCGestDoc Is Nothing Then
'        Call mCGestDoc.Termina
'        Set mCGestDoc = Nothing
'    End If

    CreaDocumentoSTD_test = "ERRORE:" & doc.TipoDoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
    Resume Next
End If


End Function


Public Function GetNomeDocumento(lngProgressivo As Long) As String
    Dim strsql As String
    Dim xTmp As MXKit.CRecordSet
    Dim strCondizione As String
    'da aggiungere condizione se esportabile per terminalino
    strsql = "select cast(tipodoc as varchar) + '/'+ cast(numerodoc as varchar) +'/'+ cast(esercizio as varchar) as nome from testedocumenti where progressivo=" & lngProgressivo

    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strsql)

    GetNomeDocumento = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "nome", "")
    Call MXDB.dbChiudiSS(xTmp)
    Set xTmp = Nothing
End Function


Public Function GetNumeroDoc(lngProgressivo As Long) As Long
    Dim strsql As String
    Dim xTmp As MXKit.CRecordSet
    Dim strCondizione As String
    'da aggiungere condizione se esportabile per terminalino
    strsql = "select numerodoc from testedocumenti where progressivo=" & lngProgressivo

    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strsql)

    GetNumeroDoc = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "numerodoc", 0)
    Call MXDB.dbChiudiSS(xTmp)
    Set xTmp = Nothing
End Function

Public Function getTipoDoc(Codice As String) As String
    Dim strsql As String
    Dim xTmp As MXKit.CRecordSet
    Dim strCondizione As String
    'da aggiungere condizione se esportabile per terminalino
    strsql = "select Tipo from parametridoc where codice=" & hndDBArchivi.FormatoSQL(Codice, DB_TEXT)

    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strsql)

    getTipoDoc = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "Tipo", "")
    Call MXDB.dbChiudiSS(xTmp)
    Set xTmp = Nothing
End Function


Public Function GetCampoDocumento(progressivo As Long, strCampo As String) As String
    Dim strsql As String
    Dim xTmp As MXKit.CRecordSet
    Dim strCondizione As String
    'da aggiungere condizione se esportabile per terminalino
    strsql = "select " & strCampo & " from testedocumenti where progressivo=" & progressivo

    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strsql)

    GetCampoDocumento = MXDB.dbGetCampo(xTmp, xTmp.Tipo, strCampo, "")
    Call MXDB.dbChiudiSS(xTmp)
    Set xTmp = Nothing
End Function

Public Function getCausaleContabile(Codice As String) As Long
    Dim strsql As String
    Dim xTmp As MXKit.CRecordSet
    Dim strCondizione As String
    'da aggiungere condizione se esportabile per terminalino
    strsql = "select causalecontab from parametridoc where codice=" & hndDBArchivi.FormatoSQL(Codice, DB_TEXT)

    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strsql)

    getCausaleContabile = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "causalecontab", 0)
    Call MXDB.dbChiudiSS(xTmp)
    Set xTmp = Nothing
End Function

Public Function GetTipoDocumento(Codice As String) As String
    Dim strsql As String
    Dim xTmp As MXKit.CRecordSet
    Dim strCondizione As String
    'da aggiungere condizione se esportabile per terminalino
    strsql = "select clifor from parametridoc where codice=" & hndDBArchivi.FormatoSQL(Codice, DB_TEXT)

    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strsql)

    GetTipoDocumento = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "clifor", "")
    Call MXDB.dbChiudiSS(xTmp)
    Set xTmp = Nothing
End Function


Public Function GetTipoDoc2(lngProgressivo As Long) As String
    Dim strsql As String
    Dim xTmp As MXKit.CRecordSet
    Dim strCondizione As String
    'da aggiungere condizione se esportabile per terminalino
    strsql = "select tipodoc from testedocmenti where progrssivo=" & lngProgressivo

    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strsql)

    GetTipoDoc2 = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "tipodoc", "")
    Call MXDB.dbChiudiSS(xTmp)
    Set xTmp = Nothing
End Function

Private Function GetDocPrelievo(strCodice As String) As String
    Dim strsql As String
    Dim xTmp As MXKit.CRecordSet
    Dim strCondizione As String
    'da aggiungere condizione se esportabile per terminalino
    strsql = "select codice from DOCDAPRELEVARE where docprelievo=" & hndDBArchivi.FormatoSQL(strCodice, DB_TEXT)

    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strsql)

    GetDocPrelievo = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "codice", "")
    Call MXDB.dbChiudiSS(xTmp)
    Set xTmp = Nothing
End Function


Public Function GetEsercizioDocumento(lngProgressivo As Long) As Long
On Error GoTo Err_GetEsercizioDocumento
Dim idRiga As Long
Dim idtesta As Long
Dim arr() As String
Dim xTmp As MXKit.CRecordSet
Dim strsql  As String

        
        strsql = "select esercizio from testedocumenti "
        strsql = strsql & " where  progressivo=" & hndDBArchivi.FormatoSQL(lngProgressivo, DB_DECIMAL)
       
        Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strsql)
                
        GetEsercizioDocumento = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "esercizio", MXNU.AnnoAttivo)
        
        Call MXDB.dbChiudiSS(xTmp)
        Set xTmp = Nothing
    
Err_GetEsercizioDocumento:
    If Err <> 0 Then
        MsgBox Err.Description, , "Err_GetEsercizioDocumento"
    End If
End Function



Public Function GetProgressivoDocumento(strTipoDoc As String, lngNumeroDoc As Long, lngEsercizio As Long, strBis As String) As Long
On Error GoTo Err_GetProgressivoDocumento
Dim idRiga As Long
Dim idtesta As Long
Dim arr() As String
Dim xTmp As MXKit.CRecordSet
Dim strsql  As String

        
        strsql = "select progressivo from testedocumenti "
        strsql = strsql & " where  Tipodoc=" & hndDBArchivi.FormatoSQL(strTipoDoc, DB_TEXT)
        strsql = strsql & " and numerodoc=" & hndDBArchivi.FormatoSQL(lngNumeroDoc, DB_DECIMAL)
        strsql = strsql & " and esercizio=" & hndDBArchivi.FormatoSQL(lngEsercizio, DB_DECIMAL)
        strsql = strsql & " and bis=" & hndDBArchivi.FormatoSQL(strBis, DB_TEXT)
       
        Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strsql)
                
        GetProgressivoDocumento = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "progressivo", 0)
        
        Call MXDB.dbChiudiSS(xTmp)
        Set xTmp = Nothing
    
Err_GetProgressivoDocumento:
    If Err <> 0 Then
        MsgBox Err.Description, , "Err_GetProgressivoDocumento"
    End If
End Function




Public Function ModificaRigaDocumento(idtesta As Long, idRiga As Long, qta As Double, Causale As Long) As Boolean
On Error Resume Next

Dim intNewEsercizio  As Integer
Dim lngNewNrDoc  As Long
Dim strNewBis  As String
Dim mCGestDoc As MXBusiness.CGestDoc
Dim RigaCorrente As Long
Dim NomeFileLog As String
Dim intFileLog As Integer

NomeFileLog = MXNU.GetTempFile()
intFileLog = MXNU.ImpostaErroriSuLog(NomeFileLog, True)

        RigaCorrente = 1
        Set mCGestDoc = MXGD.CreaCGestDoc(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
        With mCGestDoc
        .Stato = GD_MODIFICA
        'costruzione testa documento
        Call .xTDoc.AssegnaCampo("PROGRESSIVO", idtesta)
        Call .MostraRighe
               '
        If mCGestDoc.RigaAttiva.TrovaIDRiga(idRiga, True) Then
            'mCGestDoc.RigaAttiva.RigaCorr = IdRiga
            RigaCorrente = mCGestDoc.RigaAttiva.RigaCorr
            .RigaAttiva.ValoreCampo(R_QTAGEST, RigaCorrente, True) = .RigaAttiva.ValoreCampo(R_QTAGEST, RigaCorrente, True) - qta
            .RigaAttiva.ValoreCampo(R_CAUSMAG, RigaCorrente, True) = Causale
        End If
        
        
        Call .Calcolo_Totali
                
        intNewEsercizio = .xTDoc.grinput("ESERCIZIO").ValoreCorrente
        lngNewNrDoc = .xTDoc.grinput("NUMERODOC").ValoreCorrente
        strNewBis = .xTDoc.grinput("BIS").ValoreCorrente
        If .Salva(intNewEsercizio, lngNewNrDoc, strNewBis, GD_MOVIMENTA) Then
            ModificaRigaDocumento = True
        Else
            ModificaRigaDocumento = False
        End If
        End With
        
        If Not mCGestDoc Is Nothing Then
            Call mCGestDoc.Termina
            Set mCGestDoc = Nothing
        End If
    
    AggiornaLog NomeFileLog
    Call MXNU.ChiudiErroriSuLog


End Function


Public Function AnnullaRigaDocumento(idtesta As Long, idRiga As Long) As Boolean
On Error Resume Next

Dim intNewEsercizio  As Integer
Dim lngNewNrDoc  As Long
Dim strNewBis  As String
Dim mCGestDoc As MXBusiness.CGestDoc
Dim RigaCorrente As Long
Dim NomeFileLog As String
Dim intFileLog As Integer

NomeFileLog = MXNU.GetTempFile()
intFileLog = MXNU.ImpostaErroriSuLog(NomeFileLog, True)

        RigaCorrente = 1
        Set mCGestDoc = MXGD.CreaCGestDoc(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
        With mCGestDoc
        .Stato = GD_MODIFICA
        'costruzione testa documento
        Call .xTDoc.AssegnaCampo("PROGRESSIVO", idtesta)
        Call .MostraRighe
               '
        If mCGestDoc.RigaAttiva.TrovaIDRiga(idRiga) Then
            .RigaAttiva.AnnullaRiga
        End If
        Call .Calcolo_Totali
                
        intNewEsercizio = .xTDoc.grinput("ESERCIZIO").ValoreCorrente
        lngNewNrDoc = .xTDoc.grinput("NUMERODOC").ValoreCorrente
        strNewBis = .xTDoc.grinput("BIS").ValoreCorrente
        If .Salva(intNewEsercizio, lngNewNrDoc, strNewBis, GD_MOVIMENTA) Then
            AnnullaRigaDocumento = True
        Else
            AnnullaRigaDocumento = False
        End If
        End With
        
        If Not mCGestDoc Is Nothing Then
            Call mCGestDoc.Termina
            Set mCGestDoc = Nothing
        End If
    
    AggiornaLog NomeFileLog
    Call MXNU.ChiudiErroriSuLog


End Function



Public Function RisalvaDocumento(idtesta As Long) As Boolean
On Error Resume Next

Dim intNewEsercizio  As Integer
Dim lngNewNrDoc  As Long
Dim strNewBis  As String
Dim mCGestDoc As MXBusiness.CGestDoc
Dim RigaCorrente As Long
Dim NomeFileLog As String
Dim intFileLog As Integer


'Attendi 4
NomeFileLog = MXNU.GetTempFile()
intFileLog = MXNU.ImpostaErroriSuLog(NomeFileLog, True)

        RigaCorrente = 1
        Set mCGestDoc = MXGD.CreaCGestDoc(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
        With mCGestDoc
        .Stato = GD_MODIFICA
        'costruzione testa documento
        Call .xTDoc.AssegnaCampo("PROGRESSIVO", idtesta)
        Call .MostraRighe
        
        Call .Calcolo_Totali
                
        intNewEsercizio = .xTDoc.grinput("ESERCIZIO").ValoreCorrente
        lngNewNrDoc = .xTDoc.grinput("NUMERODOC").ValoreCorrente
        strNewBis = .xTDoc.grinput("BIS").ValoreCorrente
        If .Salva(intNewEsercizio, lngNewNrDoc, strNewBis) Then
            RisalvaDocumento = True
        Else
            RisalvaDocumento = False
        End If
        End With
        
        If Not mCGestDoc Is Nothing Then
            Call mCGestDoc.Termina
            Set mCGestDoc = Nothing
        End If
    
    Call MXNU.ChiudiErroriSuLog


End Function



'Valida un codice articolo come le righe documenti di metodo
Public Function Valid_Articolo(strArticolo As String, strCliente As String) As String
On Error GoTo Err_Valid_Articolo
        Dim xArt As MXBusiness.CVArt
        
        Set xArt = MXART.CreaCVArt()
        With xArt
            .Codice = strArticolo
                If .Valida(CHIEDIVAR_TUTTE, False, , , , , strCliente) Then
                    Valid_Articolo = xArt.Codice
                Else
                    Valid_Articolo = ""
                End If
                On Local Error Resume Next
        End With
        Set xArt = Nothing
Err_Valid_Articolo:
    If Err <> 0 Then
        Valid_Articolo = ""
    End If
End Function


Public Sub ValidateCampo(strValidazione As String, strValore As String, arr() As Validvalori, blnSelezione As Boolean, Optional Ordina As String = False)
On Error GoTo Err_ValidateCampo
Dim objValid As MXKit.ControlliCampo
Dim intConta As Integer
Dim strCampi As String
    strCampi = ""
    For intConta = LBound(arr) To UBound(arr)
        If Len(arr(intConta).Campo) > 0 Then
            If Len(strCampi) = 0 Then
                strCampi = arr(intConta).Campo
            Else
                strCampi = strCampi & "," & arr(intConta).Campo
            End If
        End If
    Next
    Set objValid = New MXKit.ControlliCampo
    objValid.strWHEAgg = ""
    Call objValid.Inizializza(strValidazione)
    objValid.ListaCampiRit = strCampi
    If blnSelezione Then
        If objValid.Selezione Then
            For intConta = LBound(arr) To UBound(arr)
                    If Len(arr(intConta).Campo) > 0 Then
                        arr(intConta).Valore = objValid.ValoriCampiRit(arr(intConta).Campo)
                    End If
            Next
        End If
    Else
        
        If objValid.Validazione(strValore) Then
            For intConta = LBound(arr) To UBound(arr)
                    If Len(arr(intConta).Campo) > 0 Then
                        arr(intConta).Valore = objValid.ValoriCampiRit(arr(intConta).Campo)
                    End If
            Next
        Else
            For intConta = LBound(arr) To UBound(arr)
                    If Len(arr(intConta).Campo) > 0 Then
                        arr(intConta).Valore = ""
                    End If
            Next
        End If
    End If
    
    Set objValid = Nothing
Err_ValidateCampo:
    If Err <> 0 Then
    End If
End Sub


'Apre la form documenti di metodo
Public Sub ApriDocumentoMetodo(lngIdTesta As Long, frm As Form)
On Error Resume Next
Dim colValoriChiave As Collection

    Call MXNU.FrmMetodo.EseguiAzione("GestioneDocItem", 0, 4000)
    Set colValoriChiave = New Collection
       
    colValoriChiave.Add lngIdTesta
    'frm.WindowState = vbMinimized
    DoEvents
    Call MXNU.FrmMetodo.FormAttiva.AzioniMetodo(11, colValoriChiave)
End Sub


'SET ANSI_NULLS ON
'GO
'SET QUOTED_IDENTIFIER ON
'GO
'CREATE VIEW [dbo].[VISTA_SOL_GIAC] AS
'SELECT
'    CODART,
'    SUM(Giacenza * QTA1UM) As Giacenza
'From
'    STORICOMAG
'Group By
'    Codart
'
'GO
'grant all on VISTA_SOL_GIAC to metodo98

Public Sub GetGIArticolo(strArticolo As String, dblGiacenza As Double, dblImpegnato)
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "SELECT Giacenza,Impegnato FROM VISTA_SOL_GIAC" _
      & " WHERE CODART=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    
    dblGiacenza = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "Giacenza", 0)
    dblImpegnato = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "Impegnato", 0)
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Sub

Public Function GetUltimoEsercizio() As Long
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "SELECT max(codice) as codice FROM tabesercizi"
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    
    GetUltimoEsercizio = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "codice", 0)
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function

Public Function GetDescrizioneUbicazione(strDeposito As String, strUbicazione As String) As String
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "SELECT DESCRIZIONE FROM tabubicazionidepositi" _
      & " WHERE CODDEPOSITO=" & hndDBArchivi.FormatoSQL(strDeposito, DB_TEXT) _
      & " AND CODUBICAZIONE=" & hndDBArchivi.FormatoSQL(strUbicazione, DB_TEXT) _
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    GetDescrizioneUbicazione = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "DESCRIZIONE", "")
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function

Public Function EsisteUbicazione(strDeposito As String, strUbicazione As String) As Boolean
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "SELECT count(*) as numero FROM tabubicazionidepositi" _
      & " WHERE CODDEPOSITO=" & hndDBArchivi.FormatoSQL(strDeposito, DB_TEXT) _
      & " AND CODUBICAZIONE=" & hndDBArchivi.FormatoSQL(strUbicazione, DB_TEXT) _
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    If MXDB.dbGetCampo(Xrs, Xrs.Tipo, "numero", 0) > 0 Then
        EsisteUbicazione = True
    Else
        EsisteUbicazione = False
    End If
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function


Public Function GetGiacenzaArticolo(strArticolo As String) As Double
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "SELECT Giacenza FROM VISTA_SOL_GIAC" _
      & " WHERE CODART=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
    
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    GetGiacenzaArticolo = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "Giacenza", 0)
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function

Public Function GetUbicazioneArticolo(strArticolo As String) As String
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "select CODUBICAZIONE from UBICAZIONIARTICOLI ua where ua.CODDEPOSITO=1 and ua.codiceart=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
        
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    GetUbicazioneArticolo = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "CODUBICAZIONE", "")
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function

Public Function GetIndirizzoCliente(strCliente As String, Tipo As Integer) As String
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "select Indirizzo,Cap,localita,provincia from anagraficacf where codconto=" & hndDBArchivi.FormatoSQL(strCliente, DB_TEXT)
        
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    Select Case Tipo
    Case Is = 0
        GetIndirizzoCliente = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "Indirizzo", "") & " " & MXDB.dbGetCampo(Xrs, Xrs.Tipo, "Cap", "") & " " & MXDB.dbGetCampo(Xrs, Xrs.Tipo, "Localita", "") & " " & MXDB.dbGetCampo(Xrs, Xrs.Tipo, "provincia", "")
    Case Is = 1
        GetIndirizzoCliente = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "Indirizzo", "")
    Case Is = 2
        GetIndirizzoCliente = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "Cap", "") & " " & MXDB.dbGetCampo(Xrs, Xrs.Tipo, "Localita", "") & " " & MXDB.dbGetCampo(Xrs, Xrs.Tipo, "provincia", "")
    End Select
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function

Public Function GetRagioneSociale(strCliente As String, Tipo As Integer) As String
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "select dscconto1 + ' ' + dscconto2 as ragsoc from anagraficacf where codconto=" & hndDBArchivi.FormatoSQL(strCliente, DB_TEXT)
        
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
        
    GetRagioneSociale = Trim(MXDB.dbGetCampo(Xrs, Xrs.Tipo, "ragsoc", ""))
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function


Public Function GetNumeroDestinazioni(strCliente As String) As Long
Dim xTmp As MXKit.CRecordSet
Dim strsql As String
        strsql = "select count(*) as totale from DESTINAZIONiDIVERSe"
        strsql = strsql & " Where CODCONTO = " & hndDBArchivi.FormatoSQL(strCliente, DB_TEXT)
        Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strsql)

        GetNumeroDestinazioni = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "totale", 0)
        
        Call MXDB.dbChiudiSS(xTmp)
        Set xTmp = Nothing
End Function


Public Function GetUnicaDestinazioni(strCliente As String) As Long
Dim xTmp As MXKit.CRecordSet
Dim strsql As String
        strsql = "select codice  from DESTINAZIONiDIVERSe"
        strsql = strsql & " Where CODCONTO = " & hndDBArchivi.FormatoSQL(strCliente, DB_TEXT)
        strsql = strsql & " and (select count(*) from DESTINAZIONiDIVERSe"
        strsql = strsql & " Where CODCONTO = " & hndDBArchivi.FormatoSQL(strCliente, DB_TEXT)
        strsql = strsql & " )=1"
        
        
        Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strsql)

        GetUnicaDestinazioni = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "codice", 0)
        
        Call MXDB.dbChiudiSS(xTmp)
        Set xTmp = Nothing
End Function


Public Function GetDestinazione(strCliente As String, lngDestinazione As Long, Tipo As Long) As String
Dim xTmp As MXKit.CRecordSet
Dim strsql As String
        strsql = "select * from DESTINAZIONiDIVERSe"
        strsql = strsql & " Where CODCONTO = " & hndDBArchivi.FormatoSQL(strCliente, DB_TEXT)
        strsql = strsql & " AND CODICE = " & hndDBArchivi.FormatoSQL(lngDestinazione, DB_DECIMAL)
        Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strsql)

        Select Case Tipo
        Case Is = 0
            GetDestinazione = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "Ragionesociale", "")
        Case Is = 1
            GetDestinazione = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "Indirizzo", "")
        Case Is = 2
            GetDestinazione = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "Cap", "") & " " & MXDB.dbGetCampo(xTmp, xTmp.Tipo, "Localita", "") & " " & MXDB.dbGetCampo(xTmp, xTmp.Tipo, "provincia", "")
        End Select
        Call MXDB.dbChiudiSS(xTmp)
        Set xTmp = Nothing
End Function



Public Function GetGiacenzaDepositoArticolo(strArticolo As String, strDeposito As String) As Double
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    
    strsql = "SELECT CODART , CODDEPOSITO ,"
    strsql = strsql & "              (CASE WHEN NRIFPARTITA IS NULL THEN '' ELSE NRIFPARTITA END )  AS NRIFPARTITA ,"
    strsql = strsql & "              (SUM(Carico)+SUM(ResoDaScarico)-SUM(Scarico)-SUM(ResoDaCarico)) AS GIAC FROM VISTAGIACENZEINIZDEPOSITI"
    strsql = strsql & " where codart=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
    strsql = strsql & " and  CODDEPOSITO=" & hndDBArchivi.FormatoSQL(strDeposito, DB_TEXT)
    strsql = strsql & "GROUP BY  CODART,CODDEPOSITO,NRIFPARTITA"
    strsql = strsql & " order by NRIFPARTITA"
    strsql = strsql & " having (SUM(Carico)+SUM(ResoDaScarico)-SUM(Scarico)-SUM(ResoDaCarico))>0"
    
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    GetGiacenzaDepositoArticolo = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "Giac", 0)
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function



'/****** Oggetto:  View [dbo].[VISTA_SOL_IMP]    Data script: 04/14/2011 10:05:03 ******/
'SET ANSI_NULLS ON
'GO
'SET QUOTED_IDENTIFIER ON
'GO
'CREATE VIEW [dbo].[VISTA_SOL_IMP] AS
'SELECT
'    CODART,
'    SUM(Impegnato * QTA1UM) As Impegnato
'From
'    STORICOMAG
'Where
'    Impegnato = 1 Or Impegnato = -1
'Group By
'    Codart
'grant all on VISTA_SOL_IMP to metodo98

Public Function GetImpegnatoArticolo(strArticolo As String) As Double
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "SELECT Impegnato FROM VISTA_SOL_IMP" _
      & " WHERE CODART=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    GetImpegnatoArticolo = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "Impegnato", 0)
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function

'CREATE VIEW [dbo].[VISTA_SOL_ORD] AS
'SELECT
'    CODART,
'    SUM(Ordinato * QTA1UM) As Ordinato
'From
'    STORICOMAG
'Where
'    Ordinato = 1 Or Ordinato = -1
'Group By
'    Codart
'grant all on [VISTA_SOL_ORD] to metodo98


Public Function GetOrdinatoArticolo(strArticolo As String) As Double
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "SELECT Ordinato FROM VISTA_SOL_ORD" _
      & " WHERE CODART=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    GetOrdinatoArticolo = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "Ordinato", 0)
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function


Public Function GetTassativo(lngIdTesta As Long, lngIdRiga As Long) As Boolean
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
'    strSql = "SELECT tassativo FROM Solinfo_Vista_RigheTerminalino" _
'      & " WHERE idtesta=" & hndDbArchivi.FormatoSQL(lngIdtesta, DB_DECIMAL) _
'      & " AND idriga=" & hndDbArchivi.FormatoSQL(lngIdRiga, DB_DECIMAL)
'
'    Set Xrs = MXDB.dbCreaSS(hndDbArchivi, strSql)
'    If MXDB.dbGetCampo(Xrs, Xrs.Tipo, "Tassativo", 0) = 1 Then
'        GetTassativo = True
'    Else
'    End If
'    Call MXDB.dbChiudiSS(Xrs)
'    Set Xrs = Nothing
        GetTassativo = False
End Function


Public Function GetUM(strArticolo As String, TipoUM As Integer) As String
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "SELECT UM FROM ARTICOLIUMPREFERITE" _
      & " WHERE CODART=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT) _
      & " AND TIPOUM=" & hndDBArchivi.FormatoSQL(TipoUM, DB_DECIMAL)
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    GetUM = Trim(MXDB.dbGetCampo(Xrs, Xrs.Tipo, "UM", ""))
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function

Public Function MyRound(dblNumero As Double, intDecimali As Integer) As Double
On Error GoTo Err_MYRound

Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "SELECT round("
    strsql = strsql & hndDBArchivi.FormatoSQL(dblNumero, DB_DECIMAL)
    strsql = strsql & "," & intDecimali
    strsql = strsql & ") as valore"
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    MyRound = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "valore", 0)
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
Err_MYRound:
    If Err <> 0 Then
        MsgBox "Errore arrotonamento " & Err.Description, , "MyRound"
    End If

End Function


Public Function GetBlackList(strCliente As String) As Boolean
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "select isnull(codstatoestero,0) as BlackList  from anagraficacf "
    strsql = strsql & " where codconto=" & hndDBArchivi.FormatoSQL(strCliente, DB_TEXT)

    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    If MXDB.dbGetCampo(Xrs, Xrs.Tipo, "BlackList", 0) = 0 Then
        GetBlackList = False
    Else
        GetBlackList = True
    End If
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function





Public Function GetDescrizioneOrdine(lngIdTesta As Long, lngIdRiga As Long) As String
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "SELECT descrizioneart FROM righedocumenti" _
      & " WHERE idtesta=" & lngIdTesta _
      & " AND idriga=" & lngIdRiga
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    GetDescrizioneOrdine = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "descrizioneart", "")
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function


Public Function GetTipoUMParametriDoc(strTipoDoc As String) As Long
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "select tipoum from parametridoc " _
      & " WHERE codice=" & hndDBArchivi.FormatoSQL(strTipoDoc, DB_TEXT)
      
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    GetTipoUMParametriDoc = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "tipoum", 0)
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function




Public Function UsaLotto(strArticolo As String) As Boolean
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "SELECT aa.MovimentaPartite FROM anagraficaarticoli as aa"
    strsql = strsql & " WHERE aa.CODICE=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
            
    UsaLotto = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "MovimentaPartite", False)
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function



Public Function GetPesoNetto(strArticolo As String) As Double
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "SELECT pesonetto FROM anagraficaarticoli as aa"
    strsql = strsql & " WHERE aa.CODICE=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    If GetUM(strArticolo, 1) = "KG" Then
        GetPesoNetto = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "Pesonetto", 0)
    Else
        GetPesoNetto = GetFattore(strArticolo, GetUM(strArticolo, 1), "KG")
    End If
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function

Public Function GetListinoClienti(strCoodice As String) As Long
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "SELECT listino FROM anagraficariservaticf "
    strsql = strsql & " WHERE codconto=" & hndDBArchivi.FormatoSQL(strCoodice, DB_TEXT)
    strsql = strsql & " and esercizio=" & MXNU.AnnoAttivo
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    
    GetListinoClienti = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "listino", 0)
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function


Public Function GetPrezzoListino(strArticolo As String, intListino As Integer) As Double
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "SELECT prezzoeuro FROM listiniarticoli "
    strsql = strsql & " WHERE codart=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
    strsql = strsql & " and nrlistino=" & intListino
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    
    GetPrezzoListino = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "prezzoeuro", 0)
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function


Public Function GetUMOrdine(idtesta As Long, idRiga As Long) As String
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "SELECT UMgest FROM righedocumenti" _
      & " WHERE idtesta=" & hndDBArchivi.FormatoSQL(idtesta, DB_DECIMAL) _
      & " and idriga=" & hndDBArchivi.FormatoSQL(idRiga, DB_DECIMAL)
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    GetUMOrdine = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "UMgest", "")
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function
Public Function GetDataCampoOrdine(Campo As String, idtesta As Long, idRiga As Long) As String
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "SELECT isnull(" & Campo & ",'') as Campo  FROM righedocumenti" _
      & " WHERE idtesta=" & hndDBArchivi.FormatoSQL(idtesta, DB_DECIMAL) _
      & " and idriga=" & hndDBArchivi.FormatoSQL(idRiga, DB_DECIMAL)
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    GetDataCampoOrdine = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "Campo", "")
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function

Public Function GetTipoRiferimento(strTipoDoc As String) As String
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    strsql = "select rifdocinrighe from parametridoc where codice =" & hndDBArchivi.FormatoSQL(strTipoDoc, DB_TEXT)
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    GetTipoRiferimento = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "rifdocinrighe", "N")
    
    'V=vostro
    'S=Nostro
    'N=Nessuno
    'E=Entrambi
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function





Public Function GetFattore(strArticolo As String, UM1 As String, UM2 As String) As Double
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
    ' il fatto è da UM1 a um2
    ' quindi moltiplicando qta um1 * fattore si ottiene UM2

    strsql = "select isnull(fattore,1) as fattore from ARTICOLIFATTORICONVERSIONE" _
      & " WHERE CODART=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT) _
      & " AND UM1=" & hndDBArchivi.FormatoSQL(UM1, DB_TEXT) _
      & " AND UM2=" & hndDBArchivi.FormatoSQL(UM2, DB_TEXT)
    
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    GetFattore = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "Fattore", 1)
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function

Public Sub StampaDocumento_RPT(ByVal lngIdTesta As Long, Optional NumeroCopie As Long = 0, Optional strWHERE As String = "")
    Dim strFileRpt As String
    Dim strStampante As String
    Dim lngNumCopie As Long
    Dim strTipoDoc As String
    Dim strTempo As String
    Dim strsql As String
    strsql = _
        "SELECT PD.MODULOSTAMPA, PD.NUMCOPIE, PD.OPZIONISTAMPA, PD.DEVICESTAMPA,PD.CODICE" _
      & " FROM TESTEDOCUMENTI TD" _
      & " INNER JOIN PARAMETRIDOC PD ON PD.CODICE=TD.TIPODOC" _
      & " WHERE TD.PROGRESSIVO = " & hndDBArchivi.FormatoSQL(lngIdTesta, DB_LONG)
    Dim Xrs As MXKit.CRecordSet
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)

    If Not MXDB.dbFineTab(Xrs, Xrs.Tipo) Then
        strFileRpt = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "MODULOSTAMPA", "")
        If NumeroCopie = 0 Then
            lngNumCopie = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "NUMCOPIE", 0)
        Else
            lngNumCopie = NumeroCopie
        End If
        strTipoDoc = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "CODICE", "")
        strStampante = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "DEVICESTAMPA", "")

        strFileRpt = Replace(strFileRpt, "%PATHPERSDITTA%", MXNU.PercorsoPers & "\" & MXNU.DittaAttiva, 1, -1, vbTextCompare)
        strFileRpt = Replace(strFileRpt, "%PATHPERS%", MXNU.PercorsoPers, 1, -1, vbTextCompare)
        strFileRpt = Replace(strFileRpt, "%PATHPGM%", MXNU.PercorsoPgm & "\Stampe", 1, -1, vbTextCompare)

        strsql = " TESTEDOCUMENTI.PROGRESSIVO=" & hndDBArchivi.FormatoSQL(lngIdTesta, DB_LONG) & strWHERE

        If Len(strFileRpt) > 0 Then
'            If strWHERE <> "" Then
'                frmInit.SB.Panels(1).Text = "Msgbox 2"
'                Call GestioneStampa(strFileRpt, strStampante, lngNumCopie, strSql, False)
'            Else
                Call GestioneStampa_RPT(strFileRpt, strStampante, lngNumCopie, strsql)
'            End If
        End If
    End If


    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Sub


Public Sub StampaDocumento_ETC(ByVal lngIdTesta As Long, Optional NumeroCopie As Long = 0)
On Error GoTo Err_StampaDocumento_ETC
    Dim strFileRpt As String
    Dim strStampante As String
    Dim lngNumCopie As Long
    Dim strTipoDoc As String
    Dim strTempo As String
    Dim strsql As String
    strsql = _
        "SELECT PD.MODULOSTAMPAETIC, PD.NUMCOPIE, PD.OPZIONISTAMPA, PD.DEVICESTAMPAEtIC,PD.CODICE" _
      & " FROM TESTEDOCUMENTI TD" _
      & " INNER JOIN PARAMETRIDOC PD ON PD.CODICE=TD.TIPODOC" _
      & " WHERE TD.PROGRESSIVO = " & hndDBArchivi.FormatoSQL(lngIdTesta, DB_LONG)
    Dim Xrs As MXKit.CRecordSet
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)

    If Not MXDB.dbFineTab(Xrs, Xrs.Tipo) Then
        AggiornaLog "Stampa:"
        strFileRpt = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "MODULOSTAMPAETIC", "")
        If NumeroCopie = 0 Then
            lngNumCopie = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "NUMCOPIE", 0)
        Else
            lngNumCopie = NumeroCopie
        End If
        strTipoDoc = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "CODICE", "")
        strStampante = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "DEVICESTAMPAEtIC", "")

        strFileRpt = Replace(strFileRpt, "%PATHPERSDITTA%", MXNU.PercorsoPers & "\" & MXNU.DittaAttiva, 1, -1, vbTextCompare)
        strFileRpt = Replace(strFileRpt, "%PATHPERS%", MXNU.PercorsoPers, 1, -1, vbTextCompare)
        strFileRpt = Replace(strFileRpt, "%PATHPGM%", MXNU.PercorsoPgm & "\Stampe", 1, -1, vbTextCompare)

        strsql = " TESTEDOCUMENTI.PROGRESSIVO=" & hndDBArchivi.FormatoSQL(lngIdTesta, DB_LONG)

        AggiornaLog "Tipo:" & strTipoDoc
        AggiornaLog "Stampante:" & strStampante
        AggiornaLog "Report:" & strFileRpt
        AggiornaLog "where:" & strsql


        If Len(strFileRpt) > 0 Then
            'Call GestioneStampa(strFileRpt, strStampante, lngNumCopie, strSql)
            Call GestioneStampa_RPT(strFileRpt, strStampante, lngNumCopie, strsql)
        End If
    End If


    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
    
Err_StampaDocumento_ETC:
If Err <> 0 Then
    MsgBox Err.Description, , "Err_StampaDocumento_ETC"
End If
End Sub


Public Function CalcolaPrezzoNetto(ByVal strArticolo As String, ByVal strCliFor As String, lngListino As Long, strUM As String, ByVal dblQuantita As Double) As Double
On Error Resume Next
Dim cPrezzi As MXBusiness.CPrzPrv
Dim decPrezzoEuro As Double
Dim decSconto As Double
Dim strsql As String
Dim xTmp As MXKit.CRecordSet
'Dim strUM3 As String
Dim dblFattore As Double
'Dim lngGruppoCF As Long
'Dim lngGruppoART As Long
Dim strSconto As String

    'dblFattore = GetFattorediConversione(strUM1, strUM2, strCliFor, strArticolo)
    Set cPrezzi = New MXBusiness.CPrzPrv
    cPrezzi.LeggiCFGPrezziProvv 0, 0
    If dblQuantita = 0 Then dblQuantita = 1
    
    Call cPrezzi.CalcolaPrezzi(1, strCliFor, strArticolo, "04-09-2012", lngListino, strUM)
    decSconto = cPrezzi.CalcolaSconto(1, True, strCliFor, strArticolo, "04-09-2012", lngListino, strUM)
    

    'Call cPrezzi.CalcolaPrezzi(1, strCliFor, strArticolo, , 5, "PZ")
    
    'lngGruppoCF = 0 'GetGruppo(strCliFor, True)
    'lngGruppoART = 0 'GetGruppo(strArticolo, False)
    
    decPrezzoEuro = cPrezzi.PrezzoEuro - decSconto
    'decPrezzoEuro = decPrezzoEuro - decSconto
    
    Set cPrezzi = Nothing
    CalcolaPrezzoNetto = decPrezzoEuro
    
End Function




Public Function GetSconto(ByVal strArticolo As String, ByVal strCliFor As String, lngListino As Long, strUM As String, strData As String) As String
On Error Resume Next
Dim cPrezzi As MXBusiness.CPrzPrv
Dim decPrezzoEuro As Double
Dim decSconto As Double
Dim strsql As String
Dim xTmp As MXKit.CRecordSet
'Dim strUM3 As String
Dim dblFattore As Double
'Dim lngGruppoCF As Long
'Dim lngGruppoART As Long
Dim strSconto As String

    'dblFattore = GetFattorediConversione(strUM1, strUM2, strCliFor, strArticolo)
    Set cPrezzi = New MXBusiness.CPrzPrv
    cPrezzi.LeggiCFGPrezziProvv 0, 0
   ' If dblQuantita = 0 Then dblQuantita = 1
    
    Call cPrezzi.CalcolaPrezzi(1, strCliFor, strArticolo, strData, lngListino, strUM)
    GetSconto = cPrezzi.CalcolaSconto(1, True, strCliFor, strArticolo, strData, lngListino, strUM)
    
   
    Set cPrezzi = Nothing
    
End Function

Public Function GestioneStampa_RPT(ByVal strFileRpt As String, ByVal strStampante As String, ByVal intNumCopie As Integer, ByVal strSQLWhere As String) As Boolean

    Dim xFiltro As MXKit.CFiltro
    Dim xCRW As MXKit.CCrw
    Dim strErrore As String
    
    'Dim frmAnt As Object

On Local Error GoTo ERR_Stampa
    '-------------------------------------------------
    'Inizializzazione oggetti

    GestioneStampa_RPT = True
    'MsgBox "1"
    Set xFiltro = MXFT.CreaCFiltro()
    'MsgBox "2"
    'Set frmAnt = FM98.GetOggettoMetodo("FrmAnteprima")
    If xFiltro.InizializzaFiltro() Then
        'MsgBox "3"
        If Len(strSQLWhere) <> 0 Then
        'MsgBox "4"
            Call xFiltro.SettaSQLFiltro(strSQLWhere)
        Else
        'MsgBox "5"
            xFiltro.SettaSQLFiltro ("")
        End If
        'MsgBox "6"
        
        If Len(strFileRpt) <> 0 And intNumCopie > 0 Then
            ' Init oggetto Crystal Reports
        'MsgBox "7"
            Set xCRW = MXCREP.CreaCCrw()
        'MsgBox "8"
        '    MsgBox "strFileRpt"
    
            With xCRW
                .ClearOpzioniStp
                Call .Stampante.LeggiVBPrinter(strStampante)
                .Stampante.nCopie = intNumCopie
                .Stampante.SettaVBPrinter
                '07/2017
'                If blnStampaLotto Then
'                    .Stampante.nFromPage = 1
'                    .Stampante.nToPage = 1
'                End If
                '.Titolo = "Stampa ordini"
                '.DSNDitta = MXNU.DittaAttiva
                .Filerpt = strFileRpt
                '.OpzioniForm = STP_TUTTE
                .Periferica = "Stampante"
                AggiornaLog "in Stampa"
                '.MostraFrmStampa
                Call .Stampa(xFiltro, Nothing, False)
                AggiornaLog "Fine Stampa"
            End With
        End If
    Else
        AggiornaLog "Errore init filtro"
    End If

    ' scarico tutto...
    'Set frmAnt = Nothing
    Set xCRW = Nothing
    Set xFiltro = Nothing
    '-------------------------------------------------
ERR_Stampa:
    If Err <> 0 Then
        strErrore = Err.Description
        AggiornaLog "errore " & strErrore
        GestioneStampa_RPT = False
        'Set frmAnt = Nothing
        Set xCRW = Nothing
        Set xFiltro = Nothing
    End If
End Function


Public Function GestioneStampa_ANT(ByVal strFileRpt As String, ByVal strStampante As String, ByVal intNumCopie As Integer, ByVal strSQLWhere As String, Optional blnAnteprima As Boolean = False) As Boolean

    Dim xFiltro As MXKit.CFiltro
    Dim xCRW As MXKit.CCrw
    Dim frmAnt As Object

On Local Error GoTo ERR_Stampa
    '-------------------------------------------------
    'Inizializzazione oggetti
    blnStampaAnteprima = True
    GestioneStampa_ANT = True
    Set xFiltro = MXFT.CreaCFiltro()
    Set FM98 = New CFMetodo98
    Set frmAnt = FM98.GetOggettoMetodo("FrmAnteprima")
    Call xFiltro.InizializzaFiltro
    If Len(strSQLWhere) <> 0 Then
        xFiltro.ParAgg.Add "BARCODE", "BARCODE", strBarCode, "", True, "BARCODE"
        xFiltro.SettaSQLFiltro (strSQLWhere)
    Else
        xFiltro.SettaSQLFiltro ("")
    End If
    If Len(strFileRpt) <> 0 And intNumCopie > 0 Then
        ' Init oggetto Crystal Reports
        Set xCRW = MXCREP.CreaCCrw()
        With xCRW
            
            
            .ClearOpzioniStp
            .Stampante.LeggiVBPrinter (strStampante)
            .Stampante.nCopie = intNumCopie
            .Stampante.SettaVBPrinter
            '.Titolo = "Stampa ordini"
            '.DSNDitta = MXNU.DittaAttiva
            .Filerpt = strFileRpt
            '.OpzioniForm = STP_TUTTE
            .Periferica = "Stampante"
            '.MostraFrmStampa
            Call .Stampa(xFiltro, frmAnt, blnAnteprima)
        End With
    End If


    ' scarico tutto...
    Set FM98 = Nothing
    Set frmAnt = Nothing
    Set xCRW = Nothing
    Set xFiltro = Nothing
    '-------------------------------------------------
ERR_Stampa:
    If Err <> 0 Then
        'MsgBox Err.Description
        GestioneStampa_ANT = False
        Set FM98 = Nothing
        Set frmAnt = Nothing
        Set xCRW = Nothing
        Set xFiltro = Nothing
    End If
    blnStampaAnteprima = False
End Function

Public Sub InserisciNuovoLotto(strArticolo As String, strLotto As String)
On Error GoTo Err_InserisciNuovoLotto
Dim strsql As String
    
    strsql = "insert into ANAGRAFICALOTTI (codarticolo,codlotto,bloccato,utentemodifica,datamodifica)"
    strsql = strsql & "values("
    strsql = strsql & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
    strsql = strsql & "," & hndDBArchivi.FormatoSQL(strLotto, DB_TEXT)
    strsql = strsql & ",0"
    strsql = strsql & "," & hndDBArchivi.FormatoSQL(MXNU.UtenteAttivo, DB_TEXT)
    strsql = strsql & "," & hndDBArchivi.FormatoSQL(Now, DB_DATE)
    strsql = strsql & ")"
    MXDB.dbEseguiSQL hndDBArchivi, strsql
    
    
Err_InserisciNuovoLotto:
    If Err <> 0 Then
        Err.Clear
        Resume Next
    End If
End Sub


'Private Sub Importa()
'On Error Resume Next
'Dim f As File
'Dim fld As Folder
'Dim strNewFile As String
'
'AggiornaLog "Importazione del:" & Format(Now, "DD-MM-YYYY HH:mm")
'
'Set fld = fs.GetFolder(Percorso)
'For Each f In fld.Files
'
'    If Left(UCase(fs.GetExtensionName(f.Path)), 3) = "OUT" Then
'        If Not blnAutomatico Then sbStato.Panels(1).Text = "Analisi file:" & fs.GetFileName(f.Path)
'        AggiornaLog "Analisi file:" & f.Path
'        If ImportaFile(f.Path) Then
'            strNewFile = fs.BuildPath(fs.GetParentFolderName(f.Path), Replace(fs.GetFileName(f.Path), "OUT", "BCK"))
'            MXUTil.flRinominaFile f.Path, strNewFile
'        End If
'    Else
'        If Not blnAutomatico Then sbStato.Panels(1).Text = "file non valido:" & fs.GetFileName(f.Path)
'    End If
'Next
'If Not blnAutomatico Then VisualizzaLog
'InviaLog
'End Sub
'
'Public Function ImportaFile(strNomeFile As String) As Boolean
'On Error Resume Next
'Dim txt As TextStream
'Dim strRiga As String
'Dim arr() As String
'Dim lngIndice As Long
'Dim lngIndiceRiga As Long
'Dim strTempo As String
'Dim strAgente As String
'Dim blnOrdine As Boolean
'
'Erase OrdiniOCG
'ReDim OrdiniOCG(0)
'Erase OrdiniOCI
'ReDim OrdiniOCI(0)
'
'    If Not blnAutomatico Then cmdImporta.Enabled = False
'
'    Set txt = fs.OpenTextFile(strNomeFile, ForReading, True)
'    blnOrdine = False
'    If txt Is Nothing Then ImportaFile = False: Exit Function
'
'
'    While Not txt.AtEndOfStream
'        DoEvents
'        strRiga = txt.ReadLine
'        arr = Split(strRiga, "|", , vbTextCompare)
'        If UBound(arr) > 0 Then
'            Select Case UCase(arr(0))
'                Case Is = UCase("mggiarea")
'                    If GetGruppoArticolo(arr(4) & arr(5) & arr(6)) < 20 Then
'                            lngIndice = UBound(OrdiniOCI) + 1
'                            ReDim Preserve OrdiniOCI(lngIndice)
'                            lngIndice = lngIndice - 1
'                            strAgente = Right(Left(fs.GetBaseName(strNomeFile), 6), 3)
'                            OrdiniOCI(lngIndice).Agente = GetCodiceAgente(strAgente)
'                            OrdiniOCI(lngIndice).Cliente = GetValueINI("ClienteOrdine", OrdiniOCI(lngIndice).Agente, "", True)
'                            OrdiniOCI(lngIndice).Articolo1 = arr(4)
'                            OrdiniOCI(lngIndice).Articolo2 = arr(5)
'                            OrdiniOCI(lngIndice).Articolo3 = arr(6)
'                            OrdiniOCI(lngIndice).Quantita = arr(7)
'                            OrdiniOCI(lngIndice).strDataDocumento = Format(arr(9), "DD/MM/YYYY")
'                            If DatePart("w", Now) = vbFriday And Format(Now, "DD/MM/YYYY") <> "17/12/2010" Then
'                                OrdiniOCI(lngIndice).strDataConsegna = Format(DateAdd("d", 3, Now), "DD/MM/YYYY")
'                            Else
'                                OrdiniOCI(lngIndice).strDataConsegna = Format(DateAdd("d", 1, Now), "DD/MM/YYYY")
'                            End If
'                            OrdiniOCI(lngIndice).UM = GetUMArticolo(OrdiniOCI(lngIndice))
'                            OrdiniOCI(lngIndice).strPartita = GetLotto(OrdiniOCI(lngIndice).Articolo1 & OrdiniOCI(lngIndice).Articolo2 & OrdiniOCI(lngIndice).Articolo3, OrdiniOCI(lngIndice).Cliente, OrdiniOCI(lngIndice).strDataConsegna)
'                            AggiornaLotto OrdiniOCI(lngIndice).Articolo1 & OrdiniOCI(lngIndice).Articolo2 & OrdiniOCI(lngIndice).Articolo3, OrdiniOCI(lngIndice).strPartita
'                            OrdiniOCI(lngIndice).strScadenza = GetDataScadenza(OrdiniOCI(lngIndice).Articolo1 & OrdiniOCI(lngIndice).Articolo2 & OrdiniOCI(lngIndice).Articolo3, OrdiniOCI(lngIndice).Cliente, OrdiniOCI(lngIndice).strDataConsegna)
'                            blnOrdine = True
'                    Else
'                            lngIndice = UBound(OrdiniOCG) + 1
'                            ReDim Preserve OrdiniOCG(lngIndice)
'                            lngIndice = lngIndice - 1
'                            strAgente = Right(Left(fs.GetBaseName(strNomeFile), 6), 3)
'                            OrdiniOCG(lngIndice).Agente = GetCodiceAgente(strAgente)
'                            OrdiniOCG(lngIndice).Cliente = GetValueINI("ClienteOrdine", OrdiniOCG(lngIndice).Agente, "", True)
'                            OrdiniOCG(lngIndice).Articolo1 = arr(4)
'                            OrdiniOCG(lngIndice).Articolo2 = arr(5)
'                            OrdiniOCG(lngIndice).Articolo3 = arr(6)
'                            OrdiniOCG(lngIndice).Quantita = arr(7)
'                            OrdiniOCG(lngIndice).strDataDocumento = Format(arr(9), "DD/MM/YYYY")
'                            If DatePart("w", Now) = vbFriday And Format(Now, "DD/MM/YYYY") <> "17/12/2010" Then
'                                OrdiniOCG(lngIndice).strDataConsegna = Format(DateAdd("d", 3, Now), "DD/MM/YYYY")
'                            Else
'                                OrdiniOCG(lngIndice).strDataConsegna = Format(DateAdd("d", 1, Now), "DD/MM/YYYY")
'                            End If
'                            OrdiniOCG(lngIndice).UM = GetUMArticolo(OrdiniOCG(lngIndice))
'                            OrdiniOCG(lngIndice).strPartita = GetLotto(OrdiniOCG(lngIndice).Articolo1 & OrdiniOCG(lngIndice).Articolo2 & OrdiniOCG(lngIndice).Articolo3, OrdiniOCG(lngIndice).Cliente, OrdiniOCG(lngIndice).strDataConsegna)
'                            OrdiniOCG(lngIndice).lngDescrTV = Getdescriz_tv(OrdiniOCG(lngIndice).Articolo1 & OrdiniOCG(lngIndice).Articolo2 & OrdiniOCG(lngIndice).Articolo3)
'                            AggiornaLotto OrdiniOCI(lngIndice).Articolo1 & OrdiniOCI(lngIndice).Articolo2 & OrdiniOCI(lngIndice).Articolo3, OrdiniOCI(lngIndice).strPartita
'                            OrdiniOCG(lngIndice).strScadenza = GetDataScadenza(OrdiniOCG(lngIndice).Articolo1 & OrdiniOCG(lngIndice).Articolo2 & OrdiniOCG(lngIndice).Articolo3, OrdiniOCG(lngIndice).Cliente, OrdiniOCG(lngIndice).strDataConsegna)
'                            blnOrdine = True
'                    End If
'            End Select
'        End If
'    Wend
'    txt.Close
'    strTempo = CreaOrdine("OCG", OrdiniOCG, Format(Now, "DD-MM-YYYY"))
'    AggiornaLog strTempo
'
'    'If Len(OrdiniOCG(0).Cliente) = 0 Then
'    '    ImportaFile = False
'    'Else
'        If InStr(1, strTempo, "ERRORE") = 0 Then ImportaFile = True
'    'End If
'
'    strTempo = CreaOrdine("OCI", OrdiniOCI, Format(Now, "DD-MM-YYYY"))
'    AggiornaLog strTempo
'
'    'If Len(OrdiniOCI(0).Cliente) = 0 Then
'    '    ImportaFile = False
'    'Else
'        If InStr(1, strTempo, "ERRORE") = 0 Then ImportaFile = True
'    'End If
'
'
'
'
'
'    If Not blnAutomatico Then cmdImporta.Enabled = True
'End Function
'
'
'Private Function AggiornaLotto(strArticolo As String, strLotto As String) As Boolean
'On Error GoTo Err_TrovaLotto
'Dim strSQL As String
'
'    strSQL = "insert into ANAGRAFICALOTTI (codarticolo,codlotto,bloccato,utentemodifica,datamodifica)"
'    strSQL = strSQL & "values("
'    strSQL = strSQL & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
'    strSQL = strSQL & "," & hndDBArchivi.FormatoSQL(strLotto, DB_TEXT)
'    strSQL = strSQL & ",0"
'    strSQL = strSQL & "," & hndDBArchivi.FormatoSQL(MXNU.UtenteAttivo, DB_TEXT)
'    strSQL = strSQL & "," & hndDBArchivi.FormatoSQL(Now, DB_DATE)
'    strSQL = strSQL & " ) "
'    MXDB.dbEseguiSQL hndDBArchivi, strSQL
'
'
'Err_TrovaLotto:
'    If Err <> 0 Then
'        Err.Clear
'        Resume Next
'    End If
'End Function
'
'
'
'Private Function CreaOrdine(strTipoDoc As String, o() As Ordine, strData As String) As String
'On Error Resume Next
'
'Dim intNewEsercizio  As Integer
'Dim lngNewNrDoc  As Long
'Dim strNewBis  As String
'Dim mCGestDoc As MXBusiness.CGestDoc
'Dim RigaCorrente As Long
'Dim NomeFileLog As String
'Dim intFileLog As Integer
'
'Dim intConta As Integer
'Dim intListino As Integer
'Dim dblTotale As Double
'Dim strPrezziArticolo As String
'Dim dblDaPagare As Double
'Dim strCondizione  As String
'Dim blnSalva As Boolean
'
'If Len(o(0).Cliente) = 0 Then
'    CreaOrdine = ""
'    Exit Function
'End If
'NomeFileLog = MXNU.GetTempFile()
'intFileLog = MXNU.ImpostaErroriSuLog(NomeFileLog, True)
'
'        RigaCorrente = 1
'        Set mCGestDoc = MXGD.CreaCGestDoc(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
'            With mCGestDoc
'                .Stato = GD_INSERIMENTO
'                'costruzione testa documento
'                Call .xTDoc.AssegnaCampo("TIPODOC", strTipoDoc)
'                Call .xTDoc.AssegnaCampo("ESERCIZIO", MXNU.AnnoAttivo)
'                Call .xTDoc.AssegnaCampo("DATADOC", strData)
'                Call .xTDoc.AssegnaCampo("DATARIFDOC", o(0).strDataDocumento)
'                Call .xTDoc.AssegnaCampo("CODCLIFOR", o(0).Cliente)
'                Call .xTDoc.AssegnaCampo("CODAGENTE1", "A 88" & o(0).Agente)
'                RigaCorrente = .NumeroRighe + 1
'                'aggiungo intesatzione
'                strCondizione = GetValueINI("CONDIZIONEOCG", "WHERE", "", True)
'                blnSalva = False
'                For intConta = LBound(o) To UBound(o)
'                    DoEvents
'                    If Len(o(intConta).Articolo1 + o(intConta).Articolo2 + o(intConta).Articolo3) > 0 Then
'                        If InStr(1, strCondizione, ";" & o(intConta).lngDescrTV & ";", vbTextCompare) = 0 Or strTipoDoc = "OCI" Then
'                             blnSalva = True
'                             .RigaAttiva.RigaCorr = RigaCorrente
'
'                             .RigaAttiva.ValoreCampo(R_CODART, RigaCorrente, True) = o(intConta).Articolo1 + o(intConta).Articolo2 + o(intConta).Articolo3
'                             .RigaAttiva.ValoreCampo(R_UMGEST, RigaCorrente, True) = o(intConta).UM
'                             .RigaAttiva.ValoreCampo(R_QTAGEST, RigaCorrente, True) = o(intConta).Quantita
'                             If IsDate(o(intConta).strDataConsegna) Then
'                                 .RigaAttiva.ValoreCampo(R_DATACONSEGNA, RigaCorrente, True) = Format(o(intConta).strDataConsegna, "DD-MM-YYYY")
'                             End If
'                             .RigaAttiva.ValoreCampo(R_NRIFPARTITA, RigaCorrente, True) = o(intConta).strPartita
'                             .hrrRigheExtra.RecSet("scadenza_prodotto").Value = Format(o(intConta).strScadenza, "DD/MM/YYYY")
'
'                             RigaCorrente = RigaCorrente + 1
'                        End If
'                    End If
'                Next
'                Call .Calcolo_Totali
'                'registrazione documento
'                intNewEsercizio = .xTDoc.GrInput("ESERCIZIO").ValoreCorrente
'                lngNewNrDoc = .xTDoc.GrInput("NUMERODOC").ValoreCorrente
'                strNewBis = .xTDoc.GrInput("BIS").ValoreCorrente
'                'If .Salva(intNewEsercizio, lngNewNrDoc, strNewBis, GD_MOVIMENTA, GD_CREA_TRANSITORIO) Then
'                If strTipoDoc <> "OCG" Then
'                    blnSalva = True
'                End If
'                If blnSalva Then
'                    If .Salva(intNewEsercizio, lngNewNrDoc, strNewBis) Then
'                        lngNewNrDoc = GetNumeroDoc(.xTDoc.GrInput("Progressivo").ValoreCorrente)
'                        CreaOrdine = "Generato doc :" & strTipoDoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
'                        If strTipoDoc = "OCG" Then
'                            'AggiornaOrdini .xTDoc.GrInput("PROGRESSIVO").ValoreCorrente
'                            AggiornaOrdiniDaLista o
'                        End If
'                        StampaBolla .xTDoc.GrInput("Progressivo").ValoreCorrente
'                    Else
'                        AggiornaLog "Vedi Errore " & NomeFileLog
'                        CreaOrdine = "ERRORE:" & strTipoDoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
'                    End If
'                Else
'                    CreaOrdine = "NESSUNA RIGA:" & strTipoDoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
'                    If strTipoDoc = "OCG" Then
'                        'AggiornaOrdini .xTDoc.GrInput("PROGRESSIVO").ValoreCorrente
'                        AggiornaOrdiniDaLista o
'                    End If
'                End If
'            End With
'            If Not mCGestDoc Is Nothing Then
'                Call mCGestDoc.Termina
'                Set mCGestDoc = Nothing
'            End If
'    Call MXNU.ChiudiErroriSuLog
'End Function
'
'
'
'
'
'Public Sub AggiornaLog(strTesto As String)
'On Error Resume Next
'Dim fs As New FileSystemObject
'Dim txt As TextStream
'Dim strPercorso As String
'    strPercorso = fs.BuildPath(App.Path, "LOG")
'    If Not fs.FolderExists(strPercorso) Then fs.CreateFolder strPercorso
'    Set txt = fs.OpenTextFile(fs.BuildPath(strPercorso, Format(Now, "YYYY-MM-DD") & MXNU.UtenteAttivo & ".txt"), ForAppending, True)
'    txt.WriteLine strTesto
'    txt.Close
'Set fs = Nothing
'End Sub
'
'Public Sub VisualizzaLog()
'Dim fs As New FileSystemObject
'Dim txt As TextStream
'Dim strPercorso As String
'    strPercorso = fs.BuildPath(App.Path, "LOG")
'    strPercorso = fs.BuildPath(strPercorso, Format(Now, "YYYY-MM-DD") & MXNU.UtenteAttivo & ".txt")
'    Shell "notepad.exe " & strPercorso, vbNormalFocus
'    Set fs = Nothing
'End Sub
'
'
'
'
'Public Function GetCodiceAgente(strCodice As String) As String
'    Dim strSQL As String
'    Dim xTmp As MXKit.CRecordSet
'
'    strSQL = "select right(codagente,3) as codagente from extraagenti where isnull(NUM_Ter,0)=" & hndDBArchivi.FormatoSQL(strCodice, DB_INTEGER)
'    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strSQL)
'    GetCodiceAgente = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "codagente", "000")
'
'    Call MXDB.dbChiudiSS(xTmp)
'    Set xTmp = Nothing
'End Function
'
'
'
'
'Public Sub AggiornaOrdini(lngIdTesta As Long)
'On Error Resume Next
'    Dim strSQL As String
'    Dim xTmp As MXKit.CRecordSet
'
'    strSQL = " select rd.idriga,rd.idtesta,rd.codart,rd.qtagest,td.codclifor,'" & GetValueINI("IMPORT", "UTENTE", "IMPORT", True) & "',getdate(),'" & GetValueINI("IMPORT", "UTENTE", "IMPORT", True) & "',2,0,erd.scadenza_prodotto,rd.nrrifpartita,rd.dataconsegna,NULL,NULL"
'    strSQL = strSQL & " from testedocumenti td"
'    strSQL = strSQL & " left outer join righedocumenti rd on td.progressivo=rd.idtesta"
'    strSQL = strSQL & " left outer join extrarighedoc erd on rd.idtesta=erd.idtesta   and rd.idriga=erd.idriga"
'    strSQL = strSQL & " where td.progressivo=" & lngIdTesta
'    strSQL = strSQL & " and rd.qtagest>0"
'
'    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strSQL)
'
'    While Not MXDB.dbFineTab(xTmp)
'        If TrovaRigaOrdine(MXDB.dbGetCampo(xTmp, xTmp.Tipo, "codclifor", ""), MXDB.dbGetCampo(xTmp, xTmp.Tipo, "codart", "")) Then
'            'update
'            strSQL = "update Solinfo_Tab_GeneraOrdini set qta =isnull(qta,0)+ " & MXDB.dbGetCampo(xTmp, xTmp.Tipo, "qtagest", 0)
'            strSQL = strSQL & " where isnull(Cliente,'')=" & hndDBArchivi.FormatoSQL(MXDB.dbGetCampo(xTmp, xTmp.Tipo, "codclifor", ""), DB_TEXT)
'            strSQL = strSQL & " and isnull(articolo,'')=" & hndDBArchivi.FormatoSQL(MXDB.dbGetCampo(xTmp, xTmp.Tipo, "codart", ""), DB_TEXT)
'            strSQL = strSQL & " and isnull(utente,'')='" & GetValueINI("IMPORT", "UTENTE", "IMPORT", True) & "'"
'            MXDB.dbEseguiSQL hndDBArchivi, strSQL
'        Else
'            'insert
'            strSQL = "insert into Solinfo_Tab_GeneraOrdini"
'            strSQL = strSQL & " select rd.codart,rd.qtagest,td.codclifor,'" & GetValueINI("IMPORT", "UTENTE", "IMPORT", True) & "',getdate(),'" & GetValueINI("IMPORT", "UTENTE", "IMPORT", True) & "',2,0,erd.scadenza_prodotto,rd.nrrifpartita,rd.dataconsegna,NULL,NULL"
'            strSQL = strSQL & " from testedocumenti td"
'            strSQL = strSQL & " left outer join righedocumenti rd on td.progressivo=rd.idtesta"
'            strSQL = strSQL & " left outer join extrarighedoc erd on rd.idtesta=erd.idtesta   and rd.idriga=erd.idriga"
'            strSQL = strSQL & " where rd.idtesta=" & MXDB.dbGetCampo(xTmp, xTmp.Tipo, "idtesta", 0)
'            strSQL = strSQL & " and rd.idriga=" & MXDB.dbGetCampo(xTmp, xTmp.Tipo, "idriga", 0)
'
'            MXDB.dbEseguiSQL hndDBArchivi, strSQL
'        End If
'        Call MXDB.dbSuccessivo(xTmp)
'    Wend
'
'
'    Call MXDB.dbChiudiSS(xTmp)
'    Set xTmp = Nothing
'End Sub
'
'
'Private Sub AggiornaOrdiniDaLista(o() As Ordine)
'On Error Resume Next
'    Dim strSQL As String
'    Dim strArticolo As String
'    Dim xTmp As MXKit.CRecordSet
'    Dim strCondizione As String
'    Dim intConta As Integer
'
'        strCondizione = GetValueINI("CONDIZIONEOCG", "WHERE", "", True)
'        For intConta = LBound(o) To UBound(o)
'            DoEvents
'            strArticolo = o(intConta).Articolo1 + o(intConta).Articolo2 + o(intConta).Articolo3
'
'
'            If Len(strArticolo) > 0 And InStr(1, strCondizione, ";" & o(intConta).lngDescrTV & ";", vbTextCompare) > 0 Then
'
'                If TrovaRigaOrdine(o(intConta).Cliente, strArticolo) Then
'                    'update
'                    strSQL = "update Solinfo_Tab_GeneraOrdini set qta =isnull(qta,0)+ " & o(intConta).Quantita
'                    strSQL = strSQL & " where isnull(Cliente,'')=" & hndDBArchivi.FormatoSQL(o(intConta).Cliente, DB_TEXT)
'                    strSQL = strSQL & " and isnull(articolo,'')=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
'                    strSQL = strSQL & " and isnull(utente,'')='" & GetValueINI("IMPORT", "UTENTE", "IMPORT", True) & "'"
'                    MXDB.dbEseguiSQL hndDBArchivi, strSQL
'                Else
'                    'insert
'                    strSQL = "insert into Solinfo_Tab_GeneraOrdini (Articolo,qta,cliente,Utentemodifica,datamodifica,utente,Generaoc,generaof,datascadenza,lotto,dataconsegna,qtaplus,flagfornitore)"
'                    strSQL = strSQL & " Values ("
'                    strSQL = strSQL & " " & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
'                    strSQL = strSQL & " ," & hndDBArchivi.FormatoSQL(o(intConta).Quantita, DB_DECIMAL)
'                    strSQL = strSQL & " ," & hndDBArchivi.FormatoSQL(o(intConta).Cliente, DB_TEXT)
'                    strSQL = strSQL & " ," & hndDBArchivi.FormatoSQL(GetValueINI("IMPORT", "UTENTE", "IMPORT", True), DB_TEXT)
'                    strSQL = strSQL & " ,getdate()"
'                    strSQL = strSQL & " ," & hndDBArchivi.FormatoSQL(GetValueINI("IMPORT", "UTENTE", "IMPORT", True), DB_TEXT)
'                    strSQL = strSQL & " ,0"
'                    strSQL = strSQL & " ,0"
'                    strSQL = strSQL & " ,NULL"
'                    strSQL = strSQL & " ,NULL"
'                    strSQL = strSQL & " ,NULL"
'                    strSQL = strSQL & " ,NULL"
'                    strSQL = strSQL & " ,NULL)"
'                    MXDB.dbEseguiSQL hndDBArchivi, strSQL
'                End If
'            End If
'        Next
'End Sub
'
'
'Public Function TrovaRigaOrdine(strCliente As String, strArticolo As String) As Boolean
'    Dim strSQL As String
'    Dim xTmp As MXKit.CRecordSet
'
'    strSQL = "select count(*) as tot from Solinfo_Tab_GeneraOrdini "
'    strSQL = strSQL & " where isnull(Cliente,'')=" & hndDBArchivi.FormatoSQL(strCliente, DB_TEXT)
'    strSQL = strSQL & " and isnull(articolo,'')=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
'    strSQL = strSQL & " and isnull(utente,'')='IMPORT'"
'
'    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strSQL)
'
'    If MXDB.dbGetCampo(xTmp, xTmp.Tipo, "tot", 0) > 0 Then
'        TrovaRigaOrdine = True
'    Else
'        TrovaRigaOrdine = False
'    End If
'
'    Call MXDB.dbChiudiSS(xTmp)
'    Set xTmp = Nothing
'End Function
'
'
'
'Public Function GetGruppoArticolo(strCodice As String) As Integer
'    Dim strSQL As String
'    Dim xTmp As MXKit.CRecordSet
'
'    strSQL = "select gruppo from anagraficaarticoli where isnull(codice,'')=" & hndDBArchivi.FormatoSQL(strCodice, DB_TEXT)
'    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strSQL)
'    GetGruppoArticolo = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "Gruppo", "0")
'
'    Call MXDB.dbChiudiSS(xTmp)
'    Set xTmp = Nothing
'End Function
'
'Private Function GetUMArticolo(o As Ordine) As String
'    Dim strSQL As String
'    Dim xTmp As MXKit.CRecordSet
'    Dim strCondizione As String
'    'da aggiungere condizione se esportabile per terminalino
'    strSQL = "select isnull(QTA,'') as UM from Solinfo_Vista_MGGIAREP where "
'    strSQL = strSQL & "isnull(articolo1,'')=" & hndDBArchivi.FormatoSQL(o.Articolo1, DB_TEXT)
'    strSQL = strSQL & "and isnull(articolo2,'')=" & hndDBArchivi.FormatoSQL(o.Articolo2, DB_TEXT)
'    strSQL = strSQL & "and isnull(articolo3,'')=" & hndDBArchivi.FormatoSQL(o.Articolo3, DB_TEXT)
'
'    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strSQL)
'
'    GetUMArticolo = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "UM", "")
'    Call MXDB.dbChiudiSS(xTmp)
'    Set xTmp = Nothing
'End Function
'
'
'
'
'
'
'Private Function GetLotto(strArticolo As String, strCliente As String, strDataConsegna As String) As String
'    Dim strSQL As String
'    Dim xTmp As MXKit.CRecordSet
'    Dim strCondizione As String
'    'da aggiungere condizione se esportabile per terminalino
'    strSQL = "select  dbo.SOLINFO_GET_DATASCADENZAORDINI('" & Trim(strArticolo) & "','" & strCliente & "','" & strDataConsegna & "',0) + dbo.SOLINFO_GET_DATASCADENZAORDINI('" & Trim(strArticolo) & "','" & strCliente & "','" & strDataConsegna & "',1  ) AS NRRIFPARTITA "
'
'    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strSQL)
'
'    GetLotto = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "NRRIFPARTITA", "")
'    Call MXDB.dbChiudiSS(xTmp)
'    Set xTmp = Nothing
'End Function
'
'Private Function GetDataScadenza(strArticolo As String, strCliente As String, strDataConsegna As String) As String
'    Dim strSQL As String
'    Dim xTmp As MXKit.CRecordSet
'    Dim strTempo As String
'    'da aggiungere condizione se esportabile per terminalino
'    strSQL = "select  DBO.SOLINFO_GET_DATASCADENZAORDINI('" & strArticolo & "','" & strCliente & "','" & strDataConsegna & "',0) AS DS"
'
'    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strSQL)
'
'    strTempo = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "DS", "")
'    If Len(strTempo) = 6 Then
'        GetDataScadenza = Mid(strTempo, 1, 2) & "-" & Mid(strTempo, 3, 2) & "-" & Mid(strTempo, 5, 2)
'        Call MXDB.dbChiudiSS(xTmp)
'        Set xTmp = Nothing
'    End If
'End Function
'
'Private Function Getdescriz_tv(strArticolo As String) As Long
'    Dim strSQL As String
'    Dim xTmp As MXKit.CRecordSet
'    Dim strTempo As String
'    'da aggiungere condizione se esportabile per terminalino
'    strSQL = "select  descriz_tv from extramag where codart=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
'
'    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strSQL)
'
'    Getdescriz_tv = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "descriz_tv", 0)
'
'    Call MXDB.dbChiudiSS(xTmp)
'    Set xTmp = Nothing
'
'End Function
'
'
'Public Sub InvioMail(strFile As String)
'    Dim lCount      As Long
'    Dim lCtr        As Long
'    Dim t!
'
'    If Not CBool(GetValueINI("Parametro", "AttivaMail", "True", True)) Then Exit Sub
'    Screen.MousePointer = vbHourglass
'
'    With poSendMail
'
'        ' **************************************************************************
'        ' Set the basic properties common to all messages to be sent
'        ' **************************************************************************
'        .SMTPHostValidation = VALIDATE_NONE
'        .EmailAddressValidation = VALIDATE_NONE
'        .Delimiter = ";"
'        '.SMTPHost = GetValueINI("Parametro", "SMTP", "serverdb.solinfo.lan", True)
'
'        .SMTPHost = GetValueINI("Parametro", "SMTP", "server01.fontaneto.lan", True)
'                  ' Required the fist time, optional thereafter
'        .From = GetValueINI("Parametro", "MailDA", "Paolo.sacco@fontaneto.com", True) ' Required the fist time, optional thereafter
'        .FromDisplayName = ""         ' Optional, saved after first use
'        .Message = "Log importazione ordini del " & Now
'        .Recipient = GetValueINI("Parametro", "MailA", "Paolo.sacco@fontaneto.com;Lara.Spanu@fontaneto.com;Monica.zanetta@fontaneto.com;gabriele.sacco@fontaneto.com", True)
'        '.Recipient = GetValueINI("Parametro", "MailA", "marco@websolinfo.it", True)
'        .Subject = "Importazione automatica ordini"
'        .Attachment = strFile
'        .Send
'    End With
'    ' display the results
'    Screen.MousePointer = vbDefault
'End Sub
'
'' *****************************************************************************
'' The following four Subs capture the Events fired by the vbSendMail component
'' *****************************************************************************
'
'Private Sub poSendMail_SendFailed(Explanation As String)
'    ' vbSendMail 'SendFailed Event'
'    If Not blnAutomatico Then MsgBox ("ATTENZIONE problemi nell'invio della mail motivo: " & vbCrLf & Explanation)
'End Sub
'
'Private Sub poSendMail_SendSuccesful()
'    If Not blnAutomatico Then MsgBox "Invio ok"
'End Sub
'
'
'Public Sub InviaLog()
'Dim fs As New FileSystemObject
'Dim txt As TextStream
'Dim strPercorso As String
'    strPercorso = fs.BuildPath(App.Path, "LOG")
'    strPercorso = fs.BuildPath(strPercorso, Format(Now, "YYYY-MM-DD") & MXNU.UtenteAttivo & ".txt")
'    InvioMail strPercorso
'    Set fs = Nothing
'End Sub
'
'
'Public Sub StampaBolla(ByVal lngIDTestaDoc As Long)
'    Dim strFileRpt As String
'    Dim strStampante As String
'    Dim lngNumCopie As Long
'    Dim strTipoDoc As String
'    Dim strTempo As String
'    Dim strSQL As String
'    strSQL = _
'        "SELECT PD.MODULOSTAMPA, PD.NUMCOPIE, PD.OPZIONISTAMPA, PD.DEVICESTAMPA,PD.CODICE" _
'      & " FROM TESTEDOCUMENTI TD" _
'      & " INNER JOIN PARAMETRIDOC PD ON PD.CODICE=TD.TIPODOC" _
'      & " WHERE TD.PROGRESSIVO = " & hndDBArchivi.FormatoSQL(lngIDTestaDoc, DB_LONG)
'    Dim Xrs As MXKit.CRecordSet
'    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strSQL)
'
'    If Not MXDB.dbFineTab(Xrs, Xrs.Tipo) Then
'        strFileRpt = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "MODULOSTAMPA", "")
'        lngNumCopie = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "NUMCOPIE", 0)
'        strTipoDoc = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "CODICE", "")
'        strStampante = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "DEVICESTAMPA", "")
'
'        strFileRpt = Replace(strFileRpt, "%PATHPERSDITTA%", MXNU.PercorsoPers & "\" & MXNU.DittaAttiva, 1, -1, vbTextCompare)
'        strFileRpt = Replace(strFileRpt, "%PATHPERS%", MXNU.PercorsoPers, 1, -1, vbTextCompare)
'        strFileRpt = Replace(strFileRpt, "%PATHPGM%", MXNU.PercorsoPgm & "\Stampe", 1, -1, vbTextCompare)
'
'        strSQL = " TESTEDOCUMENTI.PROGRESSIVO=" & hndDBArchivi.FormatoSQL(lngIDTestaDoc, DB_LONG)
'
''Modifica 18-01-2007 Bollani Marco
''La stampa deve andare sempre sulla stampante predefinita a prescindere dal terminalino utilizzato
'
''        If UCase(MXNU.UtenteAttivo) = "POCKETPC1" Then
''            strStampante = "POCKETPC1"
''            Call MsgBox("L'utente attivo è " & MXNU.UtenteAttivo & " e la stampante è " & strStampante)
''        End If
'    'TESTO LA STAMPANTE
'        'MsgBox strTipoDoc
'        strTempo = GetValueINI("STAMPA", "Stampante", "", True)
'        If Len(strTempo) > 0 Then
'            strStampante = strTempo
'        End If
'
'
''        'MsgBox "valore=" & strTempo
''        If Len(strTempo) = 0 Then
''            strTempo = GetValueINI(MXNU.UtenteAttivo, "DEFAULT", "", True)
''            If Len(strTempo) <> 0 Then
''                strStampante = strTempo
''            End If
''        Else
''            strStampante = strTempo
''        End If
'        'MsgBox "stampante" & strStampante
'        'MsgBox "report" & strFileRpt
'        'Exit Sub
'        If strStampante <> "" Then
'            Call GestioneStampa(strFileRpt, strStampante, lngNumCopie, strSQL)
'        Else
'            Call MsgBox("Attenzione! Nessuna stampante definita per la stampa del documento. Stampa annullata.", vbExclamation)
'        End If
'    End If
'
'
'    Call MXDB.dbChiudiSS(Xrs)
'    Set Xrs = Nothing
'End Sub
'
'


'
'Public Function GetQUANTITARIF(strCodart As String) As Double
'On Error GoTo Err_GetQUANTITARIF
'Dim IdRiga As Long
'Dim IdTesta As Long
'Dim arr() As String
'Dim xTmp As MXKit.CRecordSet
'Dim strSQL  As String
'
'
'        strSQL = "select isnull(QUANTITARIF,0) as QUANTITARIF from distintaartcomposti "
'        strSQL = strSQL & " where  VersioneDba='STN'"
'        strSQL = strSQL & " and  ARTCOMPOSTO=" & hndDBArchivi.FormatoSQL(strCodart, DB_TEXT)
'
'        Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strSQL)
'
'        GetQUANTITARIF = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "QUANTITARIF", 0)
'
'        Call MXDB.dbChiudiSS(xTmp)
'        Set xTmp = Nothing
'
'Err_GetQUANTITARIF:
'    If Err <> 0 Then
'        MsgBox Err.Description, , "Err_GetQUANTITARIF"
'    End If
'End Function
'
'
'
'
'Private Function AggiornaLotto(strArticolo As String, strLotto As String, strData As String) As Boolean
'On Error GoTo Err_AggiornaLotto
'Dim strSQL As String
'
'    strSQL = "insert into ANAGRAFICALOTTI (codarticolo,codlotto,bloccato,utentemodifica,datamodifica)"
'    strSQL = strSQL & "values("
'    strSQL = strSQL & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
'    strSQL = strSQL & "," & hndDBArchivi.FormatoSQL(strLotto, DB_TEXT)
'    strSQL = strSQL & ",0"
'    strSQL = strSQL & "," & hndDBArchivi.FormatoSQL(MXNU.UtenteAttivo, DB_TEXT)
'    strSQL = strSQL & "," & hndDBArchivi.FormatoSQL(Now, DB_DATE)
'    strSQL = strSQL & " ) "
'    MXDB.dbEseguiSQL hndDBArchivi, strSQL
'
'    strSQL = "insert into anagrcarlotti (codarticolo,codlotto,NRRIGA,valore,utentemodifica,datamodifica)"
'    strSQL = strSQL & "values("
'    strSQL = strSQL & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
'    strSQL = strSQL & "," & hndDBArchivi.FormatoSQL(strLotto, DB_TEXT)
'    strSQL = strSQL & ",0"
'    strSQL = strSQL & "," & hndDBArchivi.FormatoSQL(strData, DB_TEXT)
'    strSQL = strSQL & "," & hndDBArchivi.FormatoSQL(MXNU.UtenteAttivo, DB_TEXT)
'    strSQL = strSQL & "," & hndDBArchivi.FormatoSQL(Now, DB_DATE)
'    strSQL = strSQL & " ) "
'    MXDB.dbEseguiSQL hndDBArchivi, strSQL
'
'Err_AggiornaLotto:
'    If Err <> 0 Then
'        Err.Clear
'        Resume Next
'    End If
'End Function
'
'
'
'Private Function TrovaArticolo(arrTotale() As Ripieno, strCodart As String) As Long
'    TrovaArticolo = -1
'    For intConta = LBound(arrTotale) To UBound(arrTotale)
'        If Len(arrTotale(intConta).Codart) > 0 Then
'            If arrTotale(intConta).Codart = strCodart Then
'                TrovaArticolo = intConta
'            End If
'        End If
'    Next intConta
'End Function
'
'Private Function TrovaDocumento(lngProgressivo As Long, strTipodoc As String) As Long
'Dim strSQL As String
'Dim xTmp As MXKit.CRecordSet
'    'da modificare
'    strSQL = " SELECT idtesta FROM righedocumenti where tipodoc= '" & strTipodoc & "' and "
'    strSQL = strSQL & " IdTestaRP =" & lngProgressivo
'    strSQL = strSQL & " group by idtesta"
'
'    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strSQL)
'
'    TrovaDocumento = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "idtesta", 0)
'
'    Call MXDB.dbChiudiSS(xTmp)
'    Set xTmp = Nothing
'
'End Function
'
'
'Private Function GetIdRiga(lngProgressivo As Long, strCodart As String) As Long
'Dim strSQL As String
'Dim xTmp As MXKit.CRecordSet
'    'da modificare
'    strSQL = " SELECT idRiga FROM righedocumenti where codart= '" & strCodart & "' and "
'    strSQL = strSQL & " IdTesta =" & lngProgressivo
'
'    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strSQL)
'
'    GetIdRiga = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "idRiga", 0)
'
'    Call MXDB.dbChiudiSS(xTmp)
'    Set xTmp = Nothing
'
'End Function
'
'Private Function GetNumeroDoc(lngProgressivo As Long) As Long
'Dim strSQL As String
'Dim xTmp As MXKit.CRecordSet
'    'da modificare
'    strSQL = " SELECT numerodoc FROM testedocumenti where "
'    strSQL = strSQL & " progressivo =" & lngProgressivo
'
'    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strSQL)
'
'    GetNumeroDoc = MXDB.dbGetCampo(xTmp, xTmp.Tipo, "numerodoc", 0)
'
'    Call MXDB.dbChiudiSS(xTmp)
'    Set xTmp = Nothing
'
'End Function
'
'
'Private Function GeneraDocumento(arr() As Ripieno, strTipodoc As String, lngNumeroDoc As Long, lngEsercizio As Long) As String
'On Error Resume Next
'Dim intNewEsercizio  As Integer
'Dim lngNewNrDoc  As Long
'Dim strNewBis  As String
'Dim mCGestDoc As MXBusiness.CGestDoc
'Dim RigaCorrente As Long
'Dim NomeFileLog As String
'Dim intFileLog As Integer
'Dim intContaRipieno As Integer
'Dim intConta As Integer
'Dim lngProgressivo As Long
'Dim strLotto As String
'
'    NomeFileLog = MXNU.GetTempFile()
'    intFileLog = MXNU.ImpostaErroriSuLog(NomeFileLog, True)
'    lngProgressivo = TrovaDocumento(lngDocumento, strTipodoc)
'
'        RigaCorrente = 1
'        'MsgBox "1"
'        Set mCGestDoc = MXGD.CreaCGestDoc(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
'        'MsgBox "2"
'            With mCGestDoc
'        'MsgBox ""
'                If lngProgressivo = 0 Then
'        'MsgBox "3"
'                    .Stato = GD_INSERIMENTO
'        'MsgBox "4"
'                    Call .xTDoc.AssegnaCampo("TIPODOC", strTipodoc)
'        'MsgBox "5"
'                    Call .xTDoc.AssegnaCampo("ESERCIZIO", lngEsercizio)
'        'MsgBox "6"
'                    Call .xTDoc.AssegnaCampo("CODCLIFOR", "F    59")
'        'MsgBox "7"
'                    Call .xTDoc.AssegnaCampo("DATADOC", Format(Now, "DD/MM/YYYY"))
'        'MsgBox "8"
'                Else
'        'MsgBox "9"
'                    Call .xTDoc.AssegnaCampo("PROGRESSIVO", lngProgressivo)
'        'MsgBox "10"
'                    .Stato = GD_MODIFICA
'        'MsgBox "11"
'                    Call .xTDoc.AssegnaCampo("ESERCIZIO", lngEsercizio)
'        'MsgBox "12"
'                    .MostraRighe
'        'MsgBox "13"
'                End If
'        'MsgBox "14"
'                Dim blnTrovaRiga As Boolean
'        'MsgBox "15"
'
'        'MsgBox "16"
'                For intContaRipieno = LBound(arr) To UBound(arr)
'        'MsgBox "17"
'                    If Len(arr(intContaRipieno).Codart) > 0 Then
'        'MsgBox "18"
'                        blnTrovaRiga = False
'        'MsgBox "19"
'                        For intContaRighe = 1 To .NumeroRighe
'        'MsgBox "20"
'                            DoEvents
'        'MsgBox "21"
'                            .RigaAttiva.RigaCorr = intContaRighe
'        'MsgBox "22"
'                            If .RigaAttiva.ValoreCampo(R_CODART) = arr(intContaRipieno).Codart Then
'        'MsgBox "23"
'                                blnTrovaRiga = True
'        'MsgBox "24"
'                                Exit For
'        'MsgBox "25"
'                            End If
'        'MsgBox "26"
'                        Next intContaRighe
'        'MsgBox "27"
'                        ' se non trovo articolo mi posizioni sull'ultima riga
'        'MsgBox "28"
'                        If Not blnTrovaRiga Then
'        'MsgBox "29"
'                            intContaRighe = .NumeroRighe + 1
'        'MsgBox "30"
'                            .RigaAttiva.RigaCorr = intContaRighe
'        'MsgBox "31"
'                        End If
'        'MsgBox arr(intContaRipieno).CodArt
'                        .RigaAttiva.ValoreCampo(R_CODART, intContaRighe, True) = arr(intContaRipieno).Codart
'        'MsgBox arr(intContaRipieno).Qta
'                        .RigaAttiva.ValoreCampo(R_QTAGEST, intContaRighe, True) = arr(intContaRipieno).Qta
'                        strlottoGenerato = "L" & Format(lngNumeroDoc, "0000") & Right(.xTDoc.GrInput("ESERCIZIO").ValoreCorrente, 2)
'                        AggiornaLotto arr(intContaRipieno).Codart, strlottoGenerato, .xTDoc.GrInput("DATADOC").ValoreCorrente
'                        .RigaAttiva.ValoreCampo(R_NRIFPARTITA) = strlottoGenerato
'
'
'        'MsgBox "34"
'                        '.RigaAttiva.ValoreCampo(R_VERDIBA, intContaRighe) = "STN"
'        'MsgBox "35"
'                        .RigaAttiva.ValoreCampo(R_IDTESTARP, intContaRighe, True) = arr(intContaRipieno).IdTesta
'        'MsgBox "36"
'                        .RigaAttiva.ValoreCampo(R_IDRIGARP, intContaRighe, True) = 0
'        'MsgBox "37"
'                    End If
'        'MsgBox "38"
'                Next intContaRipieno
'        'MsgBox "39"
'
'
'                For intContaRighe = 1 To .NumeroRighe
'                    .RigaAttiva.RigaCorr = intContaRighe
'                    blnTrovaRiga = False
'                    For intContaRipieno = LBound(arr) To UBound(arr)
'                        If Len(arr(intContaRipieno).Codart) > 0 Then
'                            If .RigaAttiva.ValoreCampo(R_CODART, intContaRighe, True) = arr(intContaRipieno).Codart Then
'                                blnTrovaRiga = True
'                                Exit For
'                            End If
'                        End If
'                    Next intContaRipieno
'                    If Not blnTrovaRiga Then
'                        .RigaAttiva.AnnullaRiga
'                    End If
'                Next intContaRighe
'
'
'                Call .Calcolo_Totali
'                'registrazione documento
'                intNewEsercizio = .xTDoc.GrInput("ESERCIZIO").ValoreCorrente
'                lngNewNrDoc = .xTDoc.GrInput("NUMERODOC").ValoreCorrente
'                strNewBis = .xTDoc.GrInput("BIS").ValoreCorrente
'                If .Salva(intNewEsercizio, lngNewNrDoc, strNewBis) Then
'                    lngDocGenerato = .xTDoc.GrInput("Progressivo").ValoreCorrente
'                    'strlottoGenerato = "L" & Format(lngNewNrDoc, "0000") & Right(intNewEsercizio, 2)
'                    Stato "Generato doc :" & strTipodoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
'                    GeneraDocumento = "Generato doc :" & strTipodoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
'                Else
'                    Stato "ERRORE GENERAZIONE " & strTipodoc
'                    lngDocGenerato = -1
'                    strlottoGenerato = ""
'                    GeneraDocumento = ""
'                    AggiornaLog "####################"
'                    AggiornaLog "ERRORE GENERAZIONE " & strTipodoc
'                    AggiornaLog "per maggiori informazioni aprire file:"
'                    AggiornaLog NomeFileLog
'                    AggiornaLog "####################"
'                End If
'            End With
'            Attendi 1
'
'            'termino l'oggetto gestione documenti
'            If Not mCGestDoc Is Nothing Then
'                Call mCGestDoc.Termina
'                Set mCGestDoc = Nothing
'            End If
'    Call MXNU.ChiudiErroriSuLog
'    'Attendi 1
'End Function
'
'
'
'
'Private Function GeneraDocumentoPRO(arr() As Ripieno, strTipodoc As String, lngDocDaPrelevare As Long, strLotto As String, lngEsercizio As Long) As String
'On Error Resume Next
'Dim intNewEsercizio  As Integer
'Dim lngNewNrDoc  As Long
'Dim strNewBis  As String
'Dim mCGestDoc As MXBusiness.CGestDoc
'Dim RigaCorrente As Long
'Dim NomeFileLog As String
'Dim intFileLog As Integer
'Dim intContaRipieno As Integer
'Dim intConta As Integer
'Dim lngProgressivo As Long
'
'    NomeFileLog = MXNU.GetTempFile()
'    intFileLog = MXNU.ImpostaErroriSuLog(NomeFileLog, True)
'    lngProgressivo = TrovaDocumento(lngDocDaPrelevare, strTipodoc)
'
'        RigaCorrente = 1
'        Set mCGestDoc = MXGD.CreaCGestDoc(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
'            With mCGestDoc
'                If lngProgressivo = 0 Then
'                    .Stato = GD_INSERIMENTO
'                    Call .xTDoc.AssegnaCampo("TIPODOC", strTipodoc)
'                    Call .xTDoc.AssegnaCampo("ESERCIZIO", lngEsercizio)
'                    Call .xTDoc.AssegnaCampo("CODCLIFOR", "F    59")
'                    Call .xTDoc.AssegnaCampo("DATADOC", Format(Now, "DD/MM/YYYY"))
'                Else
'                    Call .xTDoc.AssegnaCampo("PROGRESSIVO", lngProgressivo)
'                    .Stato = GD_MODIFICA
'                    Call .xTDoc.AssegnaCampo("ESERCIZIO", MXNU.AnnoAttivo)
'                    .MostraRighe
'                End If
'                Dim blnTrovaRiga As Boolean
'
'                For intContaRipieno = LBound(arr) To UBound(arr)
'                    If Len(arr(intContaRipieno).Codart) > 0 Then
'                        blnTrovaRiga = False
'                        For intContaRighe = 1 To .NumeroRighe
'                            DoEvents
'                            .RigaAttiva.RigaCorr = intContaRighe
'                            If .RigaAttiva.ValoreCampo(R_CODART) = arr(intContaRipieno).Codart Then
'                                blnTrovaRiga = True
'                                Exit For
'                            End If
'                        Next intContaRighe
'                        ' se non trovo articolo mi posizioni sull'ultima riga
'                        If Not blnTrovaRiga Then
'                            intContaRighe = .NumeroRighe + 1
'                            .RigaAttiva.RigaCorr = intContaRighe
'                        End If
'                        .RigaAttiva.ValoreCampo(R_CODART) = arr(intContaRipieno).Codart
'                        .RigaAttiva.ValoreCampo(R_QTAGEST) = arr(intContaRipieno).Qta
'                        '.RigaAttiva.ValoreCampo(R_VERDIBA) = "STN"
'                        AggiornaLotto arr(intContaRipieno).Codart, strLotto, .xTDoc.GrInput("DATADOC").ValoreCorrente
'                        .RigaAttiva.ValoreCampo(R_NRIFPARTITA) = strLotto
'                        .RigaAttiva.ValoreCampo(R_IDTESTARP) = lngDocDaPrelevare
'                        .RigaAttiva.ValoreCampo(R_IDRIGARP) = 0 'GetIdRiga(lngDocDaPrelevare, arr(intContaRipieno).Codart) * -1
'                        '.RigaAttiva.ValoreCampo(R_FLAG_PRELEVA) = 1
'                    End If
'                Next intContaRipieno
'
'
'                For intContaRighe = 1 To .NumeroRighe
'                    .RigaAttiva.RigaCorr = intContaRighe
'                    blnTrovaRiga = False
'                    For intContaRipieno = LBound(arr) To UBound(arr)
'                        If Len(arr(intContaRipieno).Codart) > 0 Then
'                            If .RigaAttiva.ValoreCampo(R_CODART, intContaRighe, True) = arr(intContaRipieno).Codart Then
'                                blnTrovaRiga = True
'                                Exit For
'                            End If
'                        End If
'                    Next intContaRipieno
'                    If Not blnTrovaRiga Then
'                        .RigaAttiva.AnnullaRiga
'                    End If
'                Next intContaRighe
'
'
'                Call .Calcolo_Totali
'                'registrazione documento
'                intNewEsercizio = .xTDoc.GrInput("ESERCIZIO").ValoreCorrente
'                lngNewNrDoc = .xTDoc.GrInput("NUMERODOC").ValoreCorrente
'                strNewBis = .xTDoc.GrInput("BIS").ValoreCorrente
'                If .Salva(intNewEsercizio, lngNewNrDoc, strNewBis) Then
'                    Stato "Generato doc :" & strTipodoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
'                    GeneraDocumentoPRO = "Generato doc :" & strTipodoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
'                Else
'                    Stato "ERRORE GENERAZIONE " & strTipodoc
'                    GeneraDocumentoPRO = ""
'                    AggiornaLog "####################"
'                    AggiornaLog "ERRORE GENERAZIONE " & strTipodoc
'                    AggiornaLog "per maggiori informazioni aprire file:"
'                    AggiornaLog NomeFileLog
'                    AggiornaLog "####################"
'                End If
'            End With
'            Attendi 1
'
'            'termino l'oggetto gestione documenti
'            If Not mCGestDoc Is Nothing Then
'                Call mCGestDoc.Termina
'                Set mCGestDoc = Nothing
'            End If
'    Call MXNU.ChiudiErroriSuLog
'    'Attendi 1
'End Function
'
'
'



Public Function GetRiferimentoDocumentoNs(lngProgressivo As Long) As String
Dim strsql As String
Dim xTmp As MXKit.CRecordSet
    strsql = "select  numerodoc,tipodoc,datadoc,clifor,descrizione from testedocumenti td"
    strsql = strsql & " left outer join parametridoc pd on pd.codice=td.tipodoc"
    
    strsql = strsql & " Where progressivo= " & lngProgressivo
    
    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strsql)
        
    GetRiferimentoDocumentoNs = "Ns. " & MXDB.dbGetCampo(xTmp, xTmp.Tipo, "descrizione", "") & " Nr. " & MXDB.dbGetCampo(xTmp, xTmp.Tipo, "Numerodoc", 0) & " del " & MXDB.dbGetCampo(xTmp, xTmp.Tipo, "Datadoc", Now)
    
    Call MXDB.dbChiudiSS(xTmp)
    Set xTmp = Nothing
End Function

Public Function GetRiferimentoDocumentoVs(lngProgressivo As Long) As String
Dim strsql As String
Dim xTmp As MXKit.CRecordSet
    strsql = "select  numrifdoc,datarifdoc, numerodoc,tipodoc,datadoc,clifor,descrizione from testedocumenti td"
    strsql = strsql & " left outer join parametridoc pd on pd.codice=td.tipodoc"
    
    strsql = strsql & " Where progressivo= " & lngProgressivo
    
    Set xTmp = MXDB.dbCreaSS(hndDBArchivi, strsql)
        
    GetRiferimentoDocumentoVs = "Vs. " & MXDB.dbGetCampo(xTmp, xTmp.Tipo, "descrizione", "") & " Nr. " & MXDB.dbGetCampo(xTmp, xTmp.Tipo, "numrifdoc", 0) & " del " & MXDB.dbGetCampo(xTmp, xTmp.Tipo, "Datarifdoc", Now)
    
    Call MXDB.dbChiudiSS(xTmp)
    Set xTmp = Nothing
End Function

Public Function GetConnectionString() As String
    GetConnectionString = MXNU.GetstrConnection(MXNU.DittaAttiva) & ";UID=" & MXNU.UtenteDB & ";PWD=" & MXNU.PasswordDB
End Function

Public Function CreaNuovoLotto(strArticolo As String, strLotto As String, strData As String) As Boolean
On Error GoTo Err_CreaNuovoLotto
Dim strsql As String
     
    strsql = "insert into ANAGRaficaLOTTI (codarticolo,codlotto,bloccato,utentemodifica,datamodifica)"
    strsql = strsql & "values("
    strsql = strsql & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
    strsql = strsql & "," & hndDBArchivi.FormatoSQL(strLotto, DB_TEXT)
    strsql = strsql & ",0"
    strsql = strsql & "," & hndDBArchivi.FormatoSQL(MXNU.UtenteAttivo, DB_TEXT)
    strsql = strsql & "," & hndDBArchivi.FormatoSQL(Now, DB_DATE)
    strsql = strsql & " ) "
    MXDB.dbEseguiSQL hndDBArchivi, strsql
    
    strsql = "insert into ANAGRCARLOTTI (codarticolo,codlotto,nrriga,valore,utentemodifica,datamodifica,codlotto)"
    strsql = strsql & "values("
    strsql = strsql & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
    strsql = strsql & "," & hndDBArchivi.FormatoSQL(strLotto, DB_TEXT)
    strsql = strsql & ",1"
    strsql = strsql & "," & hndDBArchivi.FormatoSQL(strData, DB_TEXT)
    strsql = strsql & "," & hndDBArchivi.FormatoSQL(MXNU.UtenteAttivo, DB_TEXT)
    strsql = strsql & "," & hndDBArchivi.FormatoSQL(Now, DB_DATE)
    '03/05/2016 inserimento codclifor nei lotti ?
    strsql = strsql & "," & hndDBArchivi.FormatoSQL("", DB_TEXT)
    strsql = strsql & " ) "
    MXDB.dbEseguiSQL hndDBArchivi, strsql
    
    strsql = "update ANAGRCARLOTTI "
    strsql = strsql & " set valore=" & hndDBArchivi.FormatoSQL(strData, DB_TEXT)
    
    strsql = strsql & " where codarticolo=" & hndDBArchivi.FormatoSQL(strArticolo, DB_TEXT)
    strsql = strsql & " and codlotto=" & hndDBArchivi.FormatoSQL(strLotto, DB_TEXT)
    
    MXDB.dbEseguiSQL hndDBArchivi, strsql
    
    
Err_CreaNuovoLotto:
    If Err <> 0 Then
        Err.Clear
        Resume Next
    End If
End Function

Public Sub CreaPrezzinuovo()
Dim strsql As String
Dim Xrs As MXKit.CRecordSet
Dim lngProgressivo As Long
    strsql = "select * from listcli$ lc "
    strsql = strsql & " left outer join extraclienti ec on ec.ex_codice=LC.F1"
    strsql = strsql & " where isnull(f3,'') = 'F'  and len(isnull(codconto,''))>0 "

    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
lngProgressivo = 1
    While Not MXDB.dbFineTab(Xrs)
    
        strsql = "insert into gestioneprezzi"
        strsql = strsql & " (progressivo,codgruppoprezzicf,codclifor,codart,codgruppoprezzimag,iniziovalidita,finevalidita,usanrlistino,tipoarrot,arrotalire,arrotaeuro,codartric,utentemodifica,datamodifica,progressivoctr) values "
        strsql = strsql & "(" & lngProgressivo & ",0," & hndDBArchivi.FormatoSQL(MXDB.dbGetCampo(Xrs, Xrs.Tipo, "codconto", ""), DB_TEXT) & "," & hndDBArchivi.FormatoSQL(MXDB.dbGetCampo(Xrs, Xrs.Tipo, "F2", ""), DB_TEXT) & ",0,'01-01-1900','31-12-2090',1,'N',0,0," & hndDBArchivi.FormatoSQL(MXDB.dbGetCampo(Xrs, Xrs.Tipo, "F2", ""), DB_TEXT) & ",'trm1',getdate(),0)"
        MXDB.dbEseguiSQL hndDBArchivi, strsql
    
    
        If MXDB.dbGetCampo(Xrs, Xrs.Tipo, "F4", 0) < 0 Then
            strsql = "insert into [GESTIONEPREZZIRIGHE]"
            strsql = strsql & " ([IDRIGA],[RIFPROGRESSIVO],[NRLISTINO],[UM],[QTAMINIMA],[PREZZO_MAGG],[PREZZO_MAGGEURO],[SCONTO_UNICO],[SCONTO_AGGIUNTIVO]"
            strsql = strsql & " ,[TIPO],[UTENTEMODIFICA],[DATAMODIFICA],[TP_QTASCONTO],[TP_QTACOEFF]) values"
            strsql = strsql & " (" & lngProgressivo & "," & lngProgressivo & ",10,'',0," & hndDBArchivi.FormatoSQL(Abs(MXDB.dbGetCampo(Xrs, Xrs.Tipo, "F4", 0)), DB_DOUBLE) & "," & hndDBArchivi.FormatoSQL(Abs(MXDB.dbGetCampo(Xrs, Xrs.Tipo, "F4", 0)), DB_DOUBLE) & ",'','',-2,'trm1',getdate(),NULL,NULL )"
        Else
            strsql = "insert into [GESTIONEPREZZIRIGHE]"
            strsql = strsql & " ([IDRIGA],[RIFPROGRESSIVO],[NRLISTINO],[UM],[QTAMINIMA],[PREZZO_MAGG],[PREZZO_MAGGEURO],[SCONTO_UNICO],[SCONTO_AGGIUNTIVO]"
            strsql = strsql & " ,[TIPO],[UTENTEMODIFICA],[DATAMODIFICA],[TP_QTASCONTO],[TP_QTACOEFF]) values"
            strsql = strsql & " (" & lngProgressivo & "," & lngProgressivo & ",10,'',0," & hndDBArchivi.FormatoSQL(MXDB.dbGetCampo(Xrs, Xrs.Tipo, "F4", 0), DB_DOUBLE) & "," & hndDBArchivi.FormatoSQL(MXDB.dbGetCampo(Xrs, Xrs.Tipo, "F4", 0), DB_DOUBLE) & ",'','',-1,'trm1',getdate(),NULL,NULL )"
        End If
        MXDB.dbEseguiSQL hndDBArchivi, strsql
    
    
        lngProgressivo = lngProgressivo + 1
        Call MXDB.dbSuccessivo(Xrs)
    Wend
    
    
    
    
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Sub



