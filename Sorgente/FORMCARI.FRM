VERSION 5.00
Begin VB.Form FormCaricamentoProgramma 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendere prego..."
   ClientHeight    =   4260
   ClientLeft      =   390
   ClientTop       =   960
   ClientWidth     =   6465
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
   Icon            =   "FORMCARI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4260
   ScaleWidth      =   6465
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
         Left            =   120
         Picture         =   "FORMCARI.frx":030A
         ScaleHeight     =   855
         ScaleWidth      =   855
         TabIndex        =   4
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
         TabIndex        =   5
         Top             =   2520
         Width           =   2175
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
         Left            =   1200
         TabIndex        =   9
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Istituto di Ricovero e Cura a Carattere Scientifico"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   8
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
         Left            =   960
         TabIndex        =   7
         Top             =   2040
         Width           =   4095
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
         TabIndex        =   6
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
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   5415
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
      Caption         =   $"FORMCARI.frx":0B8C
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
      TabIndex        =   10
      Top             =   3360
      Width           =   6015
   End
End
Attribute VB_Name = "FormCaricamentoProgramma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

' centra il form
    FormCaricamentoProgramma.Top = Screen.Height / 2 - Height / 2
    FormCaricamentoProgramma.Left = Screen.Width / 2 - Width / 2

End Sub

Private Sub Picture1_Click()

' centra il form
    FormCaricamentoProgramma.Top = Screen.Height / 2 - Height / 2
    FormCaricamentoProgramma.Left = Screen.Width / 2 - Width / 2


End Sub

