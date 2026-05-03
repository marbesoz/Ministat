VERSION 5.00
Begin VB.Form FormInfo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informazioni su Ministat"
   ClientHeight    =   4770
   ClientLeft      =   4305
   ClientTop       =   2700
   ClientWidth     =   6480
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
   Icon            =   "FORMINFO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4770
   ScaleWidth      =   6480
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   372
      Left            =   2520
      TabIndex        =   4
      Top             =   4200
      Width           =   1572
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   240
         Picture         =   "FORMINFO.frx":030A
         ScaleHeight     =   855
         ScaleWidth      =   855
         TabIndex        =   5
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Milano - Italia"
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
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Istituto di Ricovero e Cura a Carattere Scientifico"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   2280
         Width           =   4335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "ISTITUTO AUXOLOGICO ITALIANO"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   2040
         Width           =   4095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Prodotto nei Laboratori di Analisi dell'"
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
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "e-mail: marco.besozzi@auxologico.it"
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
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "(C) 2000 - Versione 1.1.3216 (32 bit) by M. Besozzi"
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
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   1200
         Width           =   5175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Statistica per applicazioni biomediche"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   960
         Width           =   5055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Ministat 1.1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1080
         TabIndex        =   1
         Top             =   300
         Width           =   4095
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"FORMINFO.frx":0B8C
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
      Height          =   680
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   6015
   End
End
Attribute VB_Name = "FormInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    Unload FormInfo
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
    
End Sub

Private Sub Form_Activate()
    
    FormInfo.Command1.SetFocus

End Sub

Private Sub Form_Load()
    
    If FormPrincipale!Grid1.SelEndCol = FormPrincipale!Grid1.SelStartCol And FormPrincipale!Grid1.SelEndRow = FormPrincipale!Grid1.SelStartRow Then
        FormPrincipale!Grid1.HighLight = False
    End If

' centra il form
    FormInfo.Top = Screen.Height / 2 - Height / 2
    FormInfo.Left = Screen.Width / 2 - Width / 2

End Sub

Private Sub Picture1_Click()

' centra il form
    FormCaricamentoProgramma.Top = Screen.Height / 2 - Height / 2
    FormCaricamentoProgramma.Left = Screen.Width / 2 - Width / 2


End Sub

