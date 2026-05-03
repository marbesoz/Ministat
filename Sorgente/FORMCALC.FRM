VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormCalcoli 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calcoli"
   ClientHeight    =   1395
   ClientLeft      =   3045
   ClientTop       =   3060
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
   Icon            =   "FORMCALC.frx":0000
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
      TabIndex        =   1
      Top             =   840
      Width           =   1335
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
   Begin MSMask.MaskEdBox MaskedEdit1 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MousePointer    =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "0"
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Numero:"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   1212
   End
End
Attribute VB_Name = "FormCalcoli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandAnnulla_Click()

    Unload FormCalcoli
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
    DeselezionaColonna

End Sub

Private Sub CommandOK_Click()
    Dim Dummy
    On Error GoTo Calcoli

    MaskedEdit1.Text = FormattaStringa(MaskedEdit1.Text)
    FormCalcoli.Refresh
    NumeroPerCalcoli = MaskedEdit1.Text
    Dummy = CDbl(NumeroPerCalcoli)
    
    Unload FormCalcoli
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
'    DeselezionaColonna

Calcoli:
CommandAnnulla_Click

End Sub

Private Sub Form_Activate()
    CommandAnnulla.SetFocus
End Sub

Private Sub Form_Load()

    Screen.MousePointer = 0
' centra il form
    FormCalcoli.Top = Screen.Height / 2 - Height / 2
    FormCalcoli.Left = Screen.Width / 2 - Width / 2
    
    NumeroPerCalcoli = ""

End Sub

