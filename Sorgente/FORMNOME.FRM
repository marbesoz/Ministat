VERSION 5.00
Begin VB.Form FormNomeVariabile 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambia il nome della variabile"
   ClientHeight    =   1395
   ClientLeft      =   2940
   ClientTop       =   3120
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "FORMNOME.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1395
   ScaleWidth      =   3615
   Begin VB.CommandButton CommandAnnulla 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Annulla"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox TextNomeVariabile 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1200
      MaxLength       =   12
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      Top             =   360
      Width           =   1212
   End
   Begin VB.CommandButton CommandOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Nuovo nome:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   1212
   End
End
Attribute VB_Name = "FormNomeVariabile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandAnnulla_Click()
    
    Unload FormNomeVariabile
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
    DeselezionaColonna

End Sub

Private Sub CommandOK_Click()

    FormPrincipale!Grid1.Row = 0
    FormPrincipale!Grid1.Text = TextNomeVariabile.Text
    
    Unload FormNomeVariabile
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
    DeselezionaColonna

End Sub

Private Sub Form_Activate()
    CommandAnnulla.SetFocus
End Sub

Private Sub Form_Load()

    Screen.MousePointer = 0
' centra il form
    FormNomeVariabile.Top = Screen.Height / 2 - Height / 2
    FormNomeVariabile.Left = Screen.Width / 2 - Width / 2
' legge nome variabile preesistente
' e lo utilizza come default
    FormPrincipale!Grid1.Row = 0
    TextNomeVariabile.Text = FormPrincipale!Grid1.Text
End Sub

