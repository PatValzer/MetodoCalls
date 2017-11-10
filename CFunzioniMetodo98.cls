VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFMetodo98"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

'Gestione Accessi
'Implements MXInterfacce.IFunzioniM98

Private Function IFunzioniM98_AccessoAnnulla(ByVal intAccesso As Integer) As Boolean
    'IFunzioniM98_AccessoAnnulla = AccessoAnnulla(intAccesso)
End Function

Private Function IFunzioniM98_AccessoInserimento(ByVal intAccesso As Integer) As Boolean
    'IFunzioniM98_AccessoInserimento = AccessoInserimento(ByVal intAccesso)
End Function

Private Function IFunzioniM98_AccessoLettura(ByVal intAccesso As Integer) As Boolean
    'IFunzioniM98_AccessoLettura = AccessoLettura(intAccesso)
End Function

Private Function IFunzioniM98_AccessoModifica(ByVal intAccesso As Integer) As Boolean
    'IFunzioniM98_AccessoModifica = AccessoModifica(intAccesso)
End Function


Private Function IFunzioniM98_DefExtra(pAnagraficaExtra As Object, Foglio As Object, Sezione As String) As Integer
'    Dim AnagraficaExtra As MXKit.Anagrafica
'    Dim SS As FPSpreadADO.fpSpread
'    If StrComp(TypeName(pAnagraficaExtra), "mxkit.Anagrafica", vbTextCompare) = 0 Then
'        Set AnagraficaExtra = pAnagraficaExtra
'    Else
'        Err.Raise 13, "DefExtra", "Richiesto oggetto mxkit.Anagrafica"
'    End If
'    If StrComp(TypeName(Foglio), "FPSpreadADO.fpSpread", vbTextCompare) = 0 Then
'        Set SS = Foglio
'    Else
'        Err.Raise 13, "DefExtra", "Richiesto oggetto FPSpreadADO.fpSpread"
'    End If
'    IFunzioniM98_DefExtra = DefExtra(AnagraficaExtra, SS, Sezione)
End Function

Private Function IFunzioniM98_FormImpostaAccessi(ByVal frmDef As Object, lngButtonMask As Long) As Integer
    'IFunzioniM98_FormImpostaAccessi = FormImpostaAccessi(frmDef, lngButtonMask)
End Function

Private Function IFunzioniM98_GetOggettoMetodo(ByVal strOggetto As String, Optional ByVal bolNew As Boolean = True) As Object

    Dim objRes As Object
    Set objRes = Nothing
'#If BATCH <> 1 And TOOLS <> 1 And ISNUCLEO = 0 Then
    Select Case UCase(strOggetto)
        Case "COLL_MENU"
            'Set objRes = frmModuli.MMenus
        Case "FRMGESTIONEDOC"
'            Set objRes = New frmGestioneDoc
        Case "FRMANTEPRIMA"
            'Set objRes = New FrmAnteprima
        Case "FRMFILTRO"
            'Set objRes = New FrmFiltro
        Case "FRMLOG"
            'Set objRes = New frmLog
        Case "FRMSITUAZIONE"
            'Set objRes = New frmSituazione
        Case "FRMTABELLE"
            'Set objRes = New frmTabelle
        Case "FRMTROVAGEN"
            'Set objRes = New FrmTrovaGen
        Case "FRMVISIONI"
            'Set objRes = New frmVisioni
        Case "FRMNAVIGATORE"
'            Set objRes = New frmNavigatore
        Case "FRMVISIONICONSELEZ"
'            #If TOOLS = 0 Then
'                Set objRes = New frmVisioniConSelez
'            #End If
        Case "FRMFILTROTABELLA"
            'Set objRes = New frmFiltroTabella
        Case "FRMCONTROLLI"
            'Set objRes = New FrmNomiControlli
        Case "FRMSITUAZIONE"
            'Set objRes = New frmSituazione
        Case "FRMANALISIATP"
'            Set objRes = New frmAnalisiATP
        Case "FRMIMMAGINE"
'            Set objRes = New frmImage
        Case "FRMFUNZIONI"
'            Set objRes = New frmFunzioni
        Case "FRMINFOSTP"
            'Set objRes = New frmInfoStp
        Case "FRMVISUTMOD"
            'Set objRes = New FrmVisUtMod
        Case "FRMANACLI"
'            Set objRes = New frmAnaCli
        Case "FRMDISTINTABASE"
'            Set objRes = New frmDistintaBase
        Case "FRMANAGRAFICACOMMESSE"
'            Set objRes = New frmAnagraficaCommesse
        Case "FRMASSEGNAZIONICOMMCLI"
'            Set objRes = New frmAssegnazioniCommCli
        Case "FRMEXTCHILD"
            'Set objRes = New frmExtChild
        Case "FRMSTORICOAVANZ"
'#If TOOLS <> 1 Then
''            Set objRes = frmStoricoAvanz
'            Call FormLoader(objRes, 6625)
'#End If
        'Case "FRMSCHEDULA"
        '    Set objRes = New FrmSchedula
    End Select
    
'#End If
    If (UCase(strOggetto) = "FRMLOG") Then
      'Set objRes = New frmLog
    End If

    Set IFunzioniM98_GetOggettoMetodo = objRes
    Set objRes = Nothing

End Function


Private Sub IFunzioniM98_SSExtraButtonClicked(pAgenteAuto As Object, pAnagraficaExtra As Object, SSExtra As Object, ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
'    Dim MWAgt1 As MXKit.CAgenteAuto
'    Dim AnagraficaExtra As MXKit.Anagrafica
'    Dim SS As FPSpreadADO.fpSpread
'
'    If StrComp(TypeName(pAgenteAuto), "mxkit.CAgenteAuto", vbTextCompare) = 0 Then
'        Set MWAgt1 = pAgenteAuto
'    Else
'        Err.Raise 13, "SSExtraButtonClicked", "Richiesto oggetto mxkit.CAgenteAuto"
'    End If
'    If StrComp(TypeName(pAnagraficaExtra), "mxkit.Anagrafica", vbTextCompare) = 0 Then
'        Set AnagraficaExtra = pAnagraficaExtra
'    Else
'        Err.Raise 13, "SSExtraButtonClicked", "Richiesto oggetto mxkit.Anagrafica"
'    End If
'    If StrComp(TypeName(SSExtra), "FPSpreadADO.fpSpread", vbTextCompare) = 0 Then
'        Set SS = SSExtra
'    Else
'        Err.Raise 13, "SSExtraButtonClicked", "Richiesto oggetto FPSpreadADO.fpSpread"
'    End If
'
'    Call SSExtraButtonClicked(MWAgt1, AnagraficaExtra, SS, Col, Row, ButtonDown)
End Sub


Private Sub IFunzioniM98_SSExtraEditMode(pAgenteAuto As Object, pAnagraficaExtra As Object, SSExtra As Object, ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
'    Dim MWAgt1 As MXKit.CAgenteAuto
'    Dim AnagraficaExtra As MXKit.Anagrafica
'    Dim SS As FPSpreadADO.fpSpread
'
'    If StrComp(TypeName(pAgenteAuto), "mxkit.CAgenteAuto", vbTextCompare) = 0 Then
'        Set MWAgt1 = pAgenteAuto
'    Else
'        Err.Raise 13, "SSExtraButtonClicked", "Richiesto oggetto mxkit.CAgenteAuto"
'    End If
'    If StrComp(TypeName(pAnagraficaExtra), "mxkit.Anagrafica", vbTextCompare) = 0 Then
'        Set AnagraficaExtra = pAnagraficaExtra
'    Else
'        Err.Raise 13, "SSExtraButtonClicked", "Richiesto oggetto mxkit.Anagrafica"
'    End If
'    If StrComp(TypeName(SSExtra), "FPSpreadADO.fpSpread", vbTextCompare) = 0 Then
'        Set SS = SSExtra
'    Else
'        Err.Raise 13, "SSExtraButtonClicked", "Richiesto oggetto FPSpreadADO.fpSpread"
'    End If
'
'    Call SSExtraEditMode(MWAgt1, AnagraficaExtra, SS, Col, Row, Mode, ChangeMade)
End Sub

Private Sub IFunzioniM98_VisDatiUtenteModifica(strNomeTabella As String, strWHE As String, strDesRecord As String)
'#If BATCH <> 1 Then
'    Call MXCT.VisDatiUtenteModifica(FrmVisUtMod, strNomeTabella, strWHE, strDesRecord)
'#End If
End Sub

Public Function GetInterface() As MXInterfacce.IFunzioniM98
    Set GetInterface = Me
End Function

Public Function GetOggettoMetodo(ByVal strOggetto As String, Optional ByVal bolNew As Boolean = True) As Object
    Set GetOggettoMetodo = IFunzioniM98_GetOggettoMetodo(strOggetto, bolNew)
End Function

