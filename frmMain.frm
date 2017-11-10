VERSION 5.00
Object = "{F01BB9F2-F3EA-469F-A8DB-FF6D60F1C3ED}#1.0#0"; "MXKit.ocx"
Object = "{AD2ECC11-B7B8-4117-864B-CAEBB056E412}#1.0#0"; "MXBusiness.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3024
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3024
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MXKit.CTLXKit CTLXKit1 
      Left            =   1260
      Top             =   0
      _ExtentX        =   2244
      _ExtentY        =   1291
      _ExtentID       =   "DD9E93B1"
   End
   Begin MXBusiness.CTLXBus CTLXBus1 
      Left            =   0
      Top             =   0
      _ExtentX        =   2138
      _ExtentY        =   1291
      _ExtentID       =   "DD9E93B1"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

 
Private Sub Form_Load()



    Dim Lista_arg() As String
    Dim operazione As String
    
    Call GetCommandLine(Lista_arg)

    Dim userArg As String
    Dim dittaArg As String
    Dim preferenze As String
    Dim pwd As String

    preferenze = Lista_arg(1)
    userArg = Lista_arg(2)
    pwd = Replace(Lista_arg(3), "-p", "")
    dittaArg = Lista_arg(4)
    operazione = Lista_arg(5)
    
    Select Case operazione
        Case "inserimentoBolla":
            Dim idBolla As String
            idBolla = Lista_arg(6)
            If InizializzaMetodo(userArg, pwd, dittaArg, preferenze, CTLXKit1, CTLXBus1, False) Then
                creaBolla (idBolla)
            End If
    End Select
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TerminaMetodo
    'FreeConsole
End Sub

