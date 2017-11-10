Attribute VB_Name = "modUtil"
Option Explicit

'==========================================================================
'Funzioni create per il KMA
'==========================================================================
Public Sub LogPrint(ByVal Livello As Integer, ByVal Rif, ByVal Msg)
End Sub

Public Function fSalvaDoc(ByRef xObj As MXBusiness.CGestDoc, _
                          ByVal strTipoDoc As String, _
                          ByRef strIDDoc As String, _
                          ByRef lngIDTestaDoc As Long) As Boolean
                          
On Local Error GoTo err_fSalvaDoc
    '---------------------------------------------------
    Dim intNewEsercizio As Integer
    Dim lngNewNumeroDoc As Long
    Dim strNewBis As String
    
    Call xObj.Calcolo_Totali
    
    intNewEsercizio = xObj.xTDoc.GrInput("ESERCIZIO").ValoreCorrente
    lngNewNumeroDoc = xObj.xTDoc.GrInput("NUMERODOC").ValoreCorrente
    strNewBis = xObj.xTDoc.GrInput("BIS").ValoreCorrente
    
    lngIDTestaDoc = 0
    fSalvaDoc = xObj.Salva(intNewEsercizio, lngNewNumeroDoc, strNewBis, GD_MOVIMENTA_BATCH, GD_CREA_TRANSITORIO_BATCH)
    If fSalvaDoc Then
        ' rileggo i valori perchè potrebbero essere cambiati (se un altro utente ha salvato prima)
        strIDDoc = _
            xObj.xTDoc.GrInput("TIPODOC").ValoreCorrente & "\" _
          & xObj.xTDoc.GrInput("ESERCIZIO").ValoreCorrente & "\" _
          & xObj.xTDoc.GrInput("NUMERODOC").ValoreCorrente
        
        lngIDTestaDoc = xObj.xTDoc.GrInput("PROGRESSIVO").ValoreCorrente
    Else
        Call LogPrint(2, "fSalva", "Errore nella funzione fSalvaDoc. TipoDoc='" & strTipoDoc & "'  esercizio='" & intNewEsercizio & "'  NumeroDoc='" & lngNewNumeroDoc & "'")
    End If
    ' --------------------------------------------------

Fine_fSalvaDoc:
    On Local Error GoTo 0
    Exit Function

err_fSalvaDoc:
    Call LogPrint(0, "fSalvaDoc", "Errore nella funzione fSalvaDoc")
    fSalvaDoc = False
    Resume Fine_fSalvaDoc

End Function
'Public Function fInizializzaDoc(ByRef xObj As Object, ByVal Stato As setStatoGestioneDocumenti) As Boolean
Public Function fInizializzaDoc(ByRef xObj As MXBusiness.CGestDoc, ByVal Stato As setStatoGestioneDocumenti) As Boolean
                                
    On Local Error GoTo err_fInizializzaDoc
    '---------------------------------------------------
    Set xObj = MXGD.CreaCGestDoc(Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
    xObj.Stato = Stato
    fInizializzaDoc = True
    ' --------------------------------------------------

fine_fInizializzaDoc:
    On Local Error GoTo 0
    Exit Function

err_fInizializzaDoc:
    Call LogPrint(0, "fInizializzaDoc", "Errore nella funzione MXGD.CreaCGestDoc")
    fInizializzaDoc = False
    Resume fine_fInizializzaDoc
End Function
Public Function fTerminaDoc(ByRef xObj As Object) As Boolean
    On Local Error GoTo err_fTerminaDoc
    '---------------------------------------------------
    Call xObj.Termina
    Set xObj = Nothing
    fTerminaDoc = True
    ' --------------------------------------------------

fine_fTerminaDoc:
    On Local Error GoTo 0
    Exit Function

err_fTerminaDoc:
    Call LogPrint(0, "fTerminaDoc", "Errore nella funzione xObj.Termina")
    fTerminaDoc = False
    Resume fine_fTerminaDoc
End Function

'Public Sub Attendi(ByVal Sec As Double)
'Dim Inizio As Single
'
'Inizio = VBA.Timer
'Do While (VBA.Timer - Inizio < Sec)
'    DoEvents
'Loop
'End Sub

Public Sub MostraMsgStato(ByVal strMsg As String, _
                         ByRef objlbl As Object, _
                         Optional ByVal cColor As ColorConstants)
    
    objlbl.Caption = strMsg
    
    If Not IsMissing(cColor) Then
        objlbl.BackColor = cColor
    End If
    
    Call Attendi(0.25)
    
    If Not IsMissing(cColor) Then
        objlbl.BackColor = frmMain.BackColor
    End If
End Sub

Public Sub MostraScheda(ByRef frmForm As Form, _
                        ByVal lngCtrlIdx As Long)
    With frmForm
        .Scheda(lngCtrlIdx).Visible = True
        .Scheda(lngCtrlIdx).ZOrder 0
        
        If lngCtrlIdx < .Ling.Count Then
        ' Abbasso tutte le linguette
        Dim lngCurrLing As Long
        For lngCurrLing = 0 To frmForm.Ling.Count - 1
            .Ling(lngCurrLing).OnTop = False
        Next
        ' Alzo la linguetta corrente
        .Ling(lngCtrlIdx).OnTop = True
        End If
    End With
End Sub

Public Sub NascondiScheda(ByRef frmForm As Form, _
                          ByVal lngCtrlIdx As Long)
    With frmForm
        .Scheda(lngCtrlIdx).Visible = False
    End With
End Sub

Public Sub CancellaMsgStato(ByRef objlbl As Object)
    objlbl.Caption = ""
End Sub

Public Sub GetCommandLine(ByRef ArgArray() As String)

    Const MAXARGS = 10
    
    Dim C, CmdLine, CmdLnLen, InArg, i, NumArgs 'Declare variables.
        
    ReDim ArgArray(MAXARGS) 'Make array of the correct size.
    NumArgs = 0: InArg = False
        
    CmdLine = Command() 'Get command line arguments.
    CmdLnLen = Len(CmdLine)
    
    'Go thru command line one character at a time.
    For i = 1 To CmdLnLen
        C = Mid(CmdLine, i, 1)

        'Test for space or tab.
        If (C <> " " And C <> vbTab) Then
            'Neither space nor tab.
            If Not InArg Then   'Test if already in argument.
            'New argument begins.
                If NumArgs = MAXARGS Then Exit For  'Test for too many arguments.
                NumArgs = NumArgs + 1
                InArg = True
            End If
            ArgArray(NumArgs) = ArgArray(NumArgs) & C   'Concatenate character to current argument.
        Else
            'Found a space or tab
            InArg = False   'Set InArg flag to False.
        End If
    Next i
    
    ReDim Preserve ArgArray(NumArgs) 'Resize array just enough to hold arguments.
End Sub

