VERSION 5.00
Begin VB.Form fPrincipal 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "GO !"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "fPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim res As Long
    Dim szFileOut As String
    
    '//Iniciar libreria
    Dim hHeap As Long
    hHeap = GetProcessHeap
    Call AVIFileInit
    
    szFileOut = App.Path & "\temp.avi"
    res = GrabarStream(szFileOut) '// Devuelve el compressedStream para guardar en el
    If res <> AVIERR_OK Then GoTo errorout
    
    res = AVIERR_OK
errorout:
    CerrarTodo
    
    MsgBox "fin"
End Sub

