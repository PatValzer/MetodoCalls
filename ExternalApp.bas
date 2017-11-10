Attribute VB_Name = "ExternalApp"
'-- Here we want to open the database
Dim sConnectionString As String
'DB WORK
Dim db As New ADODB.Connection



Private Sub openConnection()
    sConnectionString = "PROVIDER = MSDASQL;driver={SQL Server};database=PastificioMoroApp;server=10.0.0.204;uid=sa;pwd=Psa01#;"
    db.ConnectionString = sConnectionString
    db.Open 'open connection
End Sub


Private Function getRecordset(strsql As String) As Recordset
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    With cmd
      .ActiveConnection = db
      .CommandText = strsql
      .CommandType = adCmdText
    End With
    
    With rs
      .CursorType = adOpenStatic
      .CursorLocation = adUseClient
      .LockType = adLockOptimistic
      .Open cmd
    End With
    
    Set getRecordset = rs
End Function


Private Sub closeRecordset(rs As Recordset)
    rs.Close
    Set rs = Nothing
End Sub
 
Private Sub closeConnection()
    db.Close
    Set db = Nothing
    Set cmd = Nothing
End Sub
 
 
 Public Sub creaBolla(idBolla As String)
    
    Dim mxMyDoc As MXBusiness.CGestDoc
    Dim lngProgressivoOrdineDaPrelevare As Long
    Dim strIdBolla As String
    Dim rsBolle As New ADODB.Recordset
    Dim rsBolleRighe As ADODB.Recordset
    Dim blnNuovaRiga, blnChiudiRiga, blnChiudiOrdine As Boolean
    Dim intNewEsercizio As Integer
    Dim lngNewNrDoc As Long
    Dim strNewBis As String
    
    DoEvents
    Dim strLogFile As String
    strLogFile = "C:\Temp\ImpDoc_" & Format(Now, "yyyymmdd_HhNn") & ".log"
    Call MXNU.ImpostaErroriSuLog(strLogFile, True)
    
    openConnection
    strsql = "SELECT * FROM BOLLA WHERE IDBOLLA='" & idBolla & "'"
    Set rsBolle = getRecordset(strsql)
    blnChiudiOrdine = rsBolle.Fields.Item("ORDINECHIUSO").Value
    strIdBolla = rsBolle.Fields.Item("IDBOLLA").Value
    
    Set mxMyDoc = MXGD.CreaCGestDoc(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
    Do While Not rsBolle.EOF
        With mxMyDoc
            .Stato = GD_INSERIMENTO
            Call .xTDoc.AssegnaCampo("TIPODOC", "DAF") 'tipoDocumento
            Call .xTDoc.AssegnaCampo("NUMRIFDOC", rsBolle.Fields.Item("NUMEROBOLLA").Value)
            Call .xTDoc.AssegnaCampo("DATARIFDOC", rsBolle.Fields.Item("DATABOLLA").Value)
            Call .xTDoc.AssegnaCampo("DATADOC", rsBolle.Fields.Item("DATABOLLA").Value)
            Call .xTDoc.AssegnaCampo("CODCLIFOR", rsBolle.Fields.Item("CODICEFORNITORE").Value)
            Call .xTDoc.AssegnaCampo("DATAINIZIOTRASP", rsBolle.Fields.Item("DATABOLLA").Value)
            Call .xTDoc.AssegnaCampo("NRBANCALI", rsBolle.Fields.Item("NUMEROBANCALIENTRATA").Value)
            
            
            Dim i As Integer
            i = 1
            
            lngProgressivoOrdineDaPrelevare = rsBolle.Fields.Item("PROGRESSIVO").Value
            If lngProgressivoOrdineDaPrelevare > 0 Then
                .RigaAttiva.ValoreCampo(R_TIPORIGA, i, True) = "R"
                .RigaAttiva.ValoreCampo(R_DESCRIZIONEART, i, True) = GetRiferimentoDocumentoVs(lngProgressivoOrdineDaPrelevare)
                .RigaAttiva.ValoreCampo(R_PRELEVA, i, True) = 1
                i = mxMyDoc.NumeroRighe + 1
            End If
            Dim strsqlRighe As String
            strsqlRighe = "SELECT * FROM BollaRiga WHERE IDBOLLA='" & strIdBolla & "'"
            Set rsBolleRighe = getRecordset(strsqlRighe)

            Do While Not rsBolleRighe.EOF
                blnNuovaRiga = rsBolleRighe.Fields.Item("RIGAFUORIORDINE").Value
                Dim lngIdRiga As Long
                lngIdRiga = rsBolleRighe.Fields.Item("IDRIGA").Value
                strTipoRiga = GetTipoRiga(lngProgressivoOrdineDaPrelevare, lngIdRiga)
                If blnChiudiOrdine = True Then
                    blnChiudiRiga = True
                Else
                    blnChiudiRiga = rsBolleRighe.Fields.Item("RIGACHIUSA").Value
                End If
                mxMyDoc.RigaAttiva.RigaCorr = mxMyDoc.RigaAttiva.RigaCorr + 1
                If blnNuovaRiga = False Then
                    .PrelevaRiga lngProgressivoOrdineDaPrelevare, lngIdRiga, , , , , rsBolleRighe.Fields.Item("QUANTITAORDINATA").Value, , rsBolleRighe.Fields.Item("LOTTO").Value
                    .RigaAttiva.ValoreCampo(R_TIPORIGA, idRiga, True) = strTipoRiga
                    .RigaAttiva.ValoreCampo(51, idRiga, True) = blnChiudiRiga
                Else
                    .RigaAttiva.RigaCorr = i
                    .RigaAttiva.ValoreCampo(R_CODART, i, True) = rsBolleRighe.Fields.Item("CODICEARTICOLO").Value
                    .RigaAttiva.ValoreCampo(R_QTAGEST, i, True) = rsBolleRighe.Fields.Item("QUANTITAORDINATA").Value
                    .RigaAttiva.ValoreCampo(R_NRRIFPARTITA, i, True) = rsBolleRighe.Fields.Item("LOTTO").Value
                End If
                i = i + 1
                rsBolleRighe.MoveNext
            Loop
            
            closeRecordset rsBolleRighe
            
            Call .Calcolo_Totali
            intNewEsercizio = .xTDoc.grinput("ESERCIZIO").ValoreCorrente
            lngNewNrDoc = .xTDoc.grinput("NUMERODOC").ValoreCorrente
            strNewBis = .xTDoc.grinput("BIS").ValoreCorrente
           
            If .Salva(intNewEsercizio, lngNewNrDoc, strNewBis) Then
                db.Execute ("UPDATE BOLLA SET idBollaInserita =" & lngNewNrDoc & ", bollaInserita = 'true' WHERE IDBOLLA='" & idBolla & "'")
            Else
                Call AggiornaLog("Errore nella generazione del documento - Attenzione!")
            End If
            
        End With
        
        Attendi 1
        rsBolle.MoveNext
    Loop
    
    Call MXNU.ChiudiErroriSuLog
    
    If Not mxMyDoc Is Nothing Then
        Call mxMyDoc.Termina
        Set mxMyDoc = Nothing
    End If
    
    closeRecordset rsBolle
    closeConnection
 
'err_CreaBolla:
'    If Err <> 0 Then
'        strErrore = Err.Description
'        Errore strErrore
'        'MsgBox Err.Description, , "xxxx"
'    '    If Not mCGestDoc Is Nothing Then
'    '        Call mCGestDoc.Termina
'    '        Set mCGestDoc = Nothing
'    '    End If
'
'        CreaDocumentoSTD_test = "ERRORE:" & doc.TipoDoc & "\" & lngNewNrDoc & "\" & intNewEsercizio
'        Resume Next
'    End If
 End Sub
 
 
 
Public Sub GeneraDocumento()
    creaBolla (3613)
End Sub


Public Function GetTipoRiga(ByRef lngProgressivo As Long, ByRef lngIdRiga As Long) As String
Dim blnRes As Boolean

    strsql = "SELECT TIPORIGA FROM RIGHEDOCUMENTI WHERE IDTESTA=" & lngProgressivo & " AND IDRIGA='" & lngIdRiga & "'"
    Debug.Print strsql
Dim Xrs As MXKit.CRecordSet
    Set Xrs = MXDB.dbCreaSS(hndDBArchivi, strsql)
    
    If MXDB.dbFineTab(Xrs) Then
        GetTipoRiga = "N"
    Else
        Do While Not MXDB.dbFineTab(Xrs)
                GetTipoRiga = MXDB.dbGetCampo(Xrs, Xrs.Tipo, "TIPORIGA", 0)
            Call MXDB.dbSuccessivo(Xrs)
        Loop
    End If
    
    Call MXDB.dbChiudiSS(Xrs)
    Set Xrs = Nothing
End Function

