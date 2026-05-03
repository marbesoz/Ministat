VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormIstogramma 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Statistiche elementari parametriche"
   ClientHeight    =   6375
   ClientLeft      =   1095
   ClientTop       =   1755
   ClientWidth     =   7365
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
   Icon            =   "FORMISTO.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6375
   ScaleWidth      =   7365
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   4320
      ScaleHeight     =   345
      ScaleWidth      =   945
      TabIndex        =   36
      Top             =   5400
      Width           =   972
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   612
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   5880
      ScaleHeight     =   345
      ScaleWidth      =   945
      TabIndex        =   26
      Top             =   120
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   4092
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   7092
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ripristina valori originari"
         Height          =   372
         Left            =   2280
         TabIndex        =   18
         Top             =   3360
         Width           =   2532
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Annulla"
         Height          =   372
         Left            =   5280
         TabIndex        =   22
         Top             =   3360
         Width           =   1212
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "OK"
         Height          =   372
         Left            =   600
         TabIndex        =   16
         Top             =   3360
         Width           =   1212
      End
      Begin VB.TextBox Text3 
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
         Height          =   285
         Left            =   3720
         MaxLength       =   40
         TabIndex        =   7
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox Text2 
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
         Height          =   285
         Left            =   240
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox Text1 
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
         Height          =   285
         Left            =   240
         MaxLength       =   50
         TabIndex        =   3
         Top             =   480
         Width           =   5775
      End
      Begin MSMask.MaskEdBox MaskedEdit5 
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Top             =   2640
         Width           =   375
         _ExtentX        =   661
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
      Begin MSMask.MaskEdBox MaskedEdit4 
         Height          =   285
         Left            =   5400
         TabIndex        =   15
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSMask.MaskEdBox MaskedEdit3 
         Height          =   285
         Left            =   3720
         TabIndex        =   14
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSMask.MaskEdBox MaskedEdit2 
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSMask.MaskEdBox MaskedEdit1 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Numero di classi:"
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
         Height          =   252
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Width           =   1572
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Y superiore:"
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
         Left            =   5400
         TabIndex        =   11
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Y inferiore:"
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
         Left            =   3720
         TabIndex        =   10
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "X superiore:"
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
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "X inferiore:"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "In ordinate:"
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
         Left            =   3720
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "In ascisse:"
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
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Titolo:"
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
         Height          =   252
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   225
      Left            =   6900
      TabIndex        =   20
      Top             =   240
      Width           =   75
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   21
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Left            =   120
      TabIndex        =   35
      Top             =   2220
      Width           =   972
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Left            =   120
      TabIndex        =   34
      Top             =   3440
      Width           =   972
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Left            =   120
      TabIndex        =   33
      Top             =   4610
      Width           =   972
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Left            =   120
      TabIndex        =   32
      Top             =   5740
      Width           =   972
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Left            =   5760
      TabIndex        =   31
      Top             =   6000
      Width           =   972
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Left            =   4920
      TabIndex        =   30
      Top             =   6000
      Width           =   972
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Left            =   3480
      TabIndex        =   29
      Top             =   6000
      Width           =   972
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Left            =   2160
      TabIndex        =   28
      Top             =   6000
      Width           =   972
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Left            =   1200
      TabIndex        =   27
      Top             =   6000
      Width           =   972
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Left            =   1200
      TabIndex        =   25
      Top             =   840
      Width           =   3972
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   1680
      TabIndex        =   24
      Top             =   6240
      Width           =   5055
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   720
      TabIndex        =   23
      Top             =   360
      Width           =   6492
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   4452
      Left            =   1200
      Top             =   720
      Width           =   7332
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&File"
      Begin VB.Menu FileStampa 
         Caption         =   "&Stampa"
      End
      Begin VB.Menu FileSeparatoreUno 
         Caption         =   "-"
      End
      Begin VB.Menu FileEsci 
         Caption         =   "&Esci"
      End
   End
   Begin VB.Menu MenuModifica 
      Caption         =   "&Modifica"
      Begin VB.Menu ModificaCopia 
         Caption         =   "&Copia"
      End
      Begin VB.Menu ModificaTaglia 
         Caption         =   "&Taglia"
      End
   End
   Begin VB.Menu MenuGrafica 
      Caption         =   "&Grafica"
      Begin VB.Menu IstogrammaDellaDistribuzione 
         Caption         =   "&Istogramma della distribuzione"
         Begin VB.Menu IstoCon 
            Caption         =   "&Con curva gaussiana"
         End
         Begin VB.Menu IstoSenza 
            Caption         =   "&Senza curva gaussiana"
         End
      End
      Begin VB.Menu GraficaSeparatoreUno 
         Caption         =   "-"
      End
      Begin VB.Menu GraficaConfigurazioneColori 
         Caption         =   "&Configurazione colori per grafica"
         Begin VB.Menu GraficaSfondoGrigio 
            Caption         =   "&Sfondo grigio"
            Checked         =   -1  'True
         End
         Begin VB.Menu GraficaDatiAColori 
            Caption         =   "&Istogramma a colori"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu GraficaScaleLegende 
         Caption         =   "C&onfigurazione scale assi e legende"
      End
   End
   Begin VB.Menu MenuStatistica 
      Caption         =   "&Statistica"
      Begin VB.Menu StatisticaStatisticheElementari 
         Caption         =   "&Statistiche elementari"
      End
      Begin VB.Menu StatisticaTabellaDeiPercentili 
         Caption         =   "&Tabella dei percentili"
      End
      Begin VB.Menu StatisticaSuddivisioneInClassi 
         Caption         =   "S&uddivisione in classi"
      End
   End
   Begin VB.Menu MenuAiuto 
      Caption         =   "&?"
      Begin VB.Menu AiutoGuida 
         Caption         =   "&Guida"
      End
      Begin VB.Menu AiutoSeparatoreUno 
         Caption         =   "-"
      End
      Begin VB.Menu AiutoCalcolatrice 
         Caption         =   "&Calcolatrice"
      End
   End
End
Attribute VB_Name = "FormIstogramma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim X() As Double
    Dim Y(), CC(), DD()
    Dim Riga, ColonnaX, ColonnaY, FineDati, Z1
    Dim XMI, XMA, YMI, YMA
    Dim XMI_Default, XMA_Default
    Dim YMI_Default, YMA_Default
    Dim LW, HI, CL, AM, MX
    Dim Titolo
    Dim InAscisse, InOrdinate
    Dim FlagGrafico, FlagStatistica
    Dim FlagSfondoGrigio
    Dim FlagDatiAColori
    Dim FlagCon, FlagSenza
    Dim SomX, SomXquad, SomXCubo, SomXquar, NumDati
    Dim Media, DevStnd, CV, ErrStnd
    Dim EmmeDue, EmmeTre, EmmeQua
    Dim GiUno, GiDue, EsseUno, EsseDue

Private Sub AiutoCalcolatrice_Click()
    Dim Calcolatrice
    On Error GoTo ErroreCalcolatriceQuattro:
    
    Calcolatrice = Shell("Calc.exe", 1)
        
Exit Sub

ErroreCalcolatriceQuattro:
    MsgBox "Calcolatrice non disponibile.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcolatrice"
Exit Sub

End Sub

Private Sub AiutoGuida_Click()

    FormPrincipale!CMDialog1.HelpFile = App.Path & "\ministat.hlp"
    FormPrincipale!CMDialog1.HelpCommand = cdlHelpContents
    FormPrincipale!CMDialog1.ShowHelp

End Sub

Private Sub Command1_Click()
    Dim FlagErroreScale, N, R, I, NewMX
    On Error GoTo Istogramma
        
    FlagErroreScale = False
    MaskedEdit1.Text = FormattaStringa(MaskedEdit1.Text)
    MaskedEdit2.Text = FormattaStringa(MaskedEdit2.Text)
    MaskedEdit3.Text = Int(CDbl(FormattaStringa(MaskedEdit3.Text)))
    MaskedEdit4.Text = Int(CDbl(FormattaStringa(MaskedEdit4.Text)))
    MaskedEdit5.Text = Int(CDbl(FormattaStringa(MaskedEdit5.Text)))
    If CDbl(MaskedEdit1.Text) > XMI_Default Then
        MaskedEdit1.Text = Format(XMI_Default, "+#####0.0####;-#####0.0####")
        FlagErroreScale = True
    End If
    If CDbl(MaskedEdit2.Text) < XMA_Default Then
        MaskedEdit2.Text = Format(XMA_Default, "+#####0.0####;-#####0.0####")
        FlagErroreScale = True
    End If
    If CDbl(MaskedEdit3.Text) <> YMI_Default Then
        MaskedEdit3.Text = Format(YMI_Default, "+#####0.0####;-#####0.0####")
        FlagErroreScale = True
    End If
    If CDbl(MaskedEdit4.Text) < YMA_Default Then
        MaskedEdit4.Text = Format(YMA_Default, "+#####0.0####;-#####0.0####")
        FlagErroreScale = True
    End If
    If CDbl(MaskedEdit5.Text) < 1 Then
        MaskedEdit5.Text = Format(CL, "##")
        FlagErroreScale = True
    End If
    If CDbl(MaskedEdit5.Text) > 40 Then
        MaskedEdit5.Text = Format(CL, "##")
        FlagErroreScale = True
    End If
        
        If FlagErroreScale = False Then
            
            Screen.MousePointer = 11

            Frame1.Visible = False
            Shape1.Visible = False
            FormIstogramma.Refresh
            Command5.Enabled = True
            
            Titolo = Text1.Text
            InAscisse = Text2.Text
            InOrdinate = Text3.Text
            XMI = MaskedEdit1.Text
            XMA = MaskedEdit2.Text
            YMI = MaskedEdit3.Text
            YMA = MaskedEdit4.Text
            MX = YMA
            CL = MaskedEdit5.Text
            
            Label13.Caption = Titolo
            Label14.Caption = InAscisse
            Label15.Caption = InOrdinate

            LW = XMI
            HI = XMA
            AM = (HI - LW) / CL
            
            For I = 1 To CL
                CC(I) = 0
            Next I

            For N = 1 To Z1
                R = CL
                If X(N) = HI Then GoTo a13300
                R = Int((X(N) - LW) / AM) + 1
a13300:
                CC(R) = CC(R) + 1
            Next N
            
            NewMX = CC(1)
                For I = 2 To CL
                    If CC(I) > NewMX Then NewMX = CC(I)
                Next I
            If CDbl(Format(NewMX, "+#####0.0####;-#####0.0####")) > YMA Then
                YMA = NewMX
                MX = NewMX
            End If
            
'        If FlagGrafico = True Then

        If FlagCon = True Then
            IstoCon_Click
        End If
        If FlagSenza = True Then
            IstoSenza_Click
        End If

'        End If
        
        Screen.MousePointer = 1

        End If

Istogramma:
        MaskedEdit1.Text = Format(XMI_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit2.Text = Format(XMA_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit3.Text = Format(YMI_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit4.Text = Format(YMA_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit5.Text = Format(CL, "##")

End Sub

Private Sub Command4_Click()
    
    Screen.MousePointer = 11
    
'    If FlagGrafico = True Then

        If FlagCon = True Then
            IstoCon_Click
        End If
        If FlagSenza = True Then
            IstoSenza_Click
        End If

'    End If
    
    Frame1.Visible = False
    Shape1.Visible = False
    Command5.Enabled = True
    Screen.MousePointer = 1

End Sub

Private Sub Command5_Click()
    
    Screen.MousePointer = 11
    
    MaskedEdit1.Text = Format(XMI_Default, "+#####0.0####;-#####0.0####")
    MaskedEdit2.Text = Format(XMA_Default, "+#####0.0####;-#####0.0####")
    MaskedEdit3.Text = Format(YMI_Default, "+#####0.0####;-#####0.0####")
    MaskedEdit4.Text = Format(YMA_Default, "+#####0.0####;-#####0.0####")
    MaskedEdit5.Text = Format(Int(Sqr(Z1) + 1), "##")

    Screen.MousePointer = 1

End Sub

Private Sub FileEsci_Click()
    ReDim X(0)
    ReDim Y(0)
    ReDim CC(0)
    ReDim DD(0)
'    ReDim AA(0), SUP(0), INF(0)
    ReDim Centile(0)
    ReDim Zeta(0)
    FormPrincipale.Show
    Unload FormIstogramma
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
End Sub

Private Sub FileStampa_Click()
    On Error GoTo ErroreFileStampaIsto

    If FlagStatistica = True Then
        Printer.Print Text4.Text
        Printer.EndDoc
    End If

    If FlagGrafico = True Then
        Picture2.Left = 100
        Picture2.Top = 80
        Picture2.Height = 340
        Picture2.Width = 460
        Picture2.ScaleHeight = 340
        Picture2.ScaleWidth = 460
        Picture2.Picture = Picture1.Image
        Picture2.Visible = True
        PrintForm
        Picture2.Visible = False
    End If

    Exit Sub

ErroreFileStampaIsto:
    Exit Sub

End Sub

Private Sub Form_Load()
    On Error GoTo GestioneErroriStatElem
    ReDim X(1 To NUM_RIGHE - 1) As Double
    ReDim Y(1 To NUM_RIGHE - 1)
    ReDim CC(40), DD(40)
    Dim D, I, J, L, M1, M2, ValoreInizialeD
    Dim XX, N, R

    Screen.MousePointer = 11
    AnnullatoCaricamentoForm = False
    FlagGrafico = False
    FlagStatistica = False
    FlagSfondoGrigio = True
    FlagDatiAColori = True
    FlagCon = False
    FlagSenza = False
    Picture2.Visible = False
    
' centra il form
    FormIstogramma.Top = Screen.Height / 2 - Height / 2
    FormIstogramma.Left = Screen.Width / 2 - Width / 2
' sposta in alto a sinistra
    Text4.Move 0, 0
' dimensionamento, ripreso nella sub resize
    Text4.Width = ScaleWidth
    Text4.Height = ScaleHeight
' inizializza
    Text4.Text = ""
    Text4.Enabled = False
' colonne in cui si trovano i dati
    ColonnaX = FormPrincipale!Grid1.SelStartCol
' nome delle variabili
    FormPrincipale!Grid1.Row = 0
    FormPrincipale!Grid1.Col = ColonnaX
    InAscisse = FormPrincipale!Grid1.Text
    Titolo = "Istogramma della distribuzione"
    InOrdinate = "Frequenza assoluta"
' carica i dati nel vettore X()
    FormGauge.Show
    FormGauge.Caption = "Statistiche elementari parametriche"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    FineDati = False
    Z1 = 0
    
        FormGauge!Label6.Caption = "[1/5]"
        For Riga = 1 To NUM_RIGHE - 1

            FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
            FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        FormPrincipale!Grid1.Row = Riga
        
        If FineDati = False Then
            FormPrincipale!Grid1.Col = ColonnaX
            If FormPrincipale!Grid1.Text <> "" Then
                Z1 = Z1 + 1
                X(Riga) = CDbl(-(-FormPrincipale!Grid1.Text))
            Else
                FineDati = True
            End If
        End If
        
        Next Riga

' controlli finali sui dati
    If Z1 > 1 Then

' inizio ordinamento dati
pc330:
    FormGauge!Label6.Caption = "[2/5]"
    D = 1
pc340:
    D = 2 * D: If D <= Z1 Then GoTo pc340
pc350:
    D = Int((D - 1) / 2): If D = 0 Then GoTo pc420
pc360:
    For I = 1 To Z1 - D
pc370:
        J = I
pc380:
        L = J + D
        If X(J) <= X(L) Then GoTo pc400
                    
pc390:
            XX = X(J): X(J) = X(L): X(L) = XX

            If I > ValoreInizialeD And 100 * I / Z1 >= 1 Then
                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If
            J = J - D: If J > 0 Then GoTo pc380
pc400:
    Next I
pc410:
    ValoreInizialeD = Z1 - D
    GoTo pc350
pc420:
' fine ordinamento dati
                
' calcola percentili
    
            FormGauge!Label6.Caption = "[3/5]"
            For I = 1 To Z1
                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                M1 = 0: M2 = 0
pes840:
                    For J = 1 To Z1
pes850:
                       If X(J) > X(I) Then GoTo pes890
pes860:
                       If X(J) = X(I) Then GoTo pes880
pes870:
                       M1 = M1 + 1: GoTo pes890
pes880:
                       M2 = M2 + 1
pes890:
                    Next J
pes900:
                Y(I) = M1 + 0.5 * M2 + 0.5
            Next I

            For I = 1 To Z1
                Y(I) = Format(100 * Y(I) / (Z1 + 1), "##0.0####")
            Next I

            XMI = X(1): XMA = X(1)
            For I = 2 To Z1
                If X(I) < XMI Then XMI = X(I)
                If X(I) > XMA Then XMA = X(I)
            Next
            XMI_Default = XMI
            XMA_Default = XMA

            If XMI = XMA Then GoTo GestioneErroriStatElem ' meglio così

' suddivisione in classi
            LW = XMI
            HI = XMA
            CL = Int(Sqr(Z1) + 1)
            AM = (HI - LW) / CL
            
            FormGauge!Label6.Caption = "[4/5]"
            For N = 1 To Z1
                FormGauge.Label1.Caption = Int(100 * N / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * N / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                R = CL
                If X(N) = HI Then GoTo 13300
                R = Int((X(N) - LW) / AM) + 1
13300:
                CC(R) = CC(R) + 1
            Next N

            MX = CC(1)
                For I = 2 To CL
                    If CC(I) > MX Then MX = CC(I)
                Next I
            
            YMI = 0
            YMA = MX
            YMI_Default = YMI
            YMA_Default = YMA

            SomX = 0: SomXquad = 0: SomXCubo = 0: SomXquar = 0: NumDati = 0
    
            FormGauge!Label6.Caption = "[5/5]"
            For I = 1 To Z1
        
                SomX = SomX + X(I)
                SomXquad = SomXquad + X(I) ^ 2
                SomXCubo = SomXCubo + X(I) ^ 3
                SomXquar = SomXquar + X(I) ^ 4
                NumDati = NumDati + 1
            
                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            
            Next I

            Media = SomX / NumDati
            DevStnd = Sqr((SomXquad - SomX ^ 2 / NumDati) / (NumDati - 1))
            CV = DevStnd * 100 / Media
            ErrStnd = DevStnd / Sqr(NumDati)
    
            EmmeDue = (SomXquad - NumDati * Media ^ 2) / NumDati
            EmmeTre = (SomXCubo - 3 * Media * SomXquad + 2 * NumDati * Media ^ 3) / NumDati
            EmmeQua = (SomXquar - 4 * Media * SomXCubo + 6 * Media ^ 2 * SomXquad - 3 * NumDati * Media ^ 4) / NumDati
            GiUno = EmmeTre / (EmmeDue * Sqr(EmmeDue))
            GiDue = EmmeQua / EmmeDue ^ 2 - 3
            EsseUno = Sqr(6 / NumDati)
            EsseDue = Sqr(24 / NumDati)

            FormGauge.Label1.Caption = "100%"
            FormGauge!Gauge1.Value = 100
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            
            Unload FormGauge
            Text4.Visible = True
            Shape1.Visible = False
            Frame1.Visible = False
            Picture1.Visible = False
            Picture2.Visible = False
            Label26.Visible = False
            FormPrincipale!Grid1.HighLight = True
            RipristinaSchermo
            FormIstogramma.Show
            ' FormPrincipale.Hide
            StatisticaStatisticheElementari_Click
            Screen.MousePointer = 1
            Exit Sub
    Else
        Unload FormGauge
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        MsgBox "Per le statistiche elementari" + Chr(10) + "sono necessari almeno due dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Statistiche elementari non eseguibili"
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
    End If

    Exit Sub

GestioneErroriStatElem:
        Unload FormGauge
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico dei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Statistiche elementari non eseguibili"
        RipristinaSchermo
        Exit Sub

End Sub

Private Sub Form_Resize()
    
    Text4.Width = ScaleWidth
    Text4.Height = ScaleHeight

        If FlagCon = True Then
            IstoCon_Click
        End If
        If FlagSenza = True Then
            IstoSenza_Click
        End If

End Sub

Private Sub GraficaDatiAColori_Click()
    
    GraficaDatiAColori.Checked = Not GraficaDatiAColori.Checked
    FlagDatiAColori = Not FlagDatiAColori

    If FlagGrafico = True Then
        Picture1.Visible = False
        Label13.Visible = False
        Label14.Visible = False
        Label15.Visible = False
        Label16.Visible = False
        Label17.Visible = False
        Label18.Visible = False
        Label19.Visible = False
        Label20.Visible = False
        Label21.Visible = False
        Label22.Visible = False
        Label23.Visible = False
        Label24.Visible = False
        Label25.Visible = False
        Label26.Visible = False

        If FlagCon = True Then
            IstoCon_Click
        End If
        If FlagSenza = True Then
            IstoSenza_Click
        End If

    End If

End Sub

Private Sub GraficaScaleLegende_Click()

    Screen.MousePointer = 0
    FlagGrafico = False
    Cls
    Shape1.Left = ScaleWidth / 2 - Shape1.Width / 2
    Shape1.Top = ScaleHeight / 2 - Shape1.Height / 2 * 1.2
    Frame1.Left = ScaleWidth / 2 - Frame1.Width / 2
    Frame1.Top = ScaleHeight / 2 - Frame1.Height / 2 * 1.2
    Text4.Visible = False
    Shape1.Visible = True
    Frame1.Visible = True
    Picture1.Visible = False
    Label13.Visible = False
    Label14.Visible = False
    Label15.Visible = False
    Label16.Visible = False
    Label17.Visible = False
    Label18.Visible = False
    Label19.Visible = False
    Label20.Visible = False
    Label21.Visible = False
    Label22.Visible = False
    Label23.Visible = False
    Label24.Visible = False
    Label25.Visible = False
    Label26.Visible = False
        
    Text1.Text = Titolo
    Text2.Text = InAscisse
    Text3.Text = InOrdinate
    MaskedEdit1.Text = Format(XMI, "+#####0.0####;-#####0.0####")
    MaskedEdit2.Text = Format(XMA, "+#####0.0####;-#####0.0####")
    MaskedEdit3.Text = Format(YMI, "+#####0.0####;-#####0.0####")
    MaskedEdit4.Text = Format(YMA, "+#####0.0####;-#####0.0####")
    MaskedEdit5.Text = Format(CL, "##")
    Command4.SetFocus

End Sub

Private Sub GraficaSfondoGrigio_Click()
    
    GraficaSfondoGrigio.Checked = Not GraficaSfondoGrigio.Checked
    FlagSfondoGrigio = Not FlagSfondoGrigio

    If FlagGrafico = True Then
        Picture1.Visible = False
        Label13.Visible = False
        Label14.Visible = False
        Label15.Visible = False
        Label16.Visible = False
        Label17.Visible = False
        Label18.Visible = False
        Label19.Visible = False
        Label20.Visible = False
        Label21.Visible = False
        Label22.Visible = False
        Label23.Visible = False
        Label24.Visible = False
        Label25.Visible = False
        Label26.Visible = False

        If FlagCon = True Then
            IstoCon_Click
        End If
        If FlagSenza = True Then
            IstoSenza_Click
        End If
    
    End If

End Sub

Private Sub IstoCon_Click()
    
    Dim I, XX, YIN, YSU, FX, FY, CX, CY
    Dim X1, X2, X3, X4, X5
    Dim Y1, Y2, Y3, Y4, Y5
    Dim FF, LargCol, XF, XI
    Dim SommaAree, ClassiValide, Altmed
    Dim Inizio, Fine, Ampiezza, Deviata, Ordinata

    Screen.MousePointer = 11
    Shape1.Visible = False
    Frame1.Visible = False
    Label13.Visible = False
    Label14.Visible = False
    Label15.Visible = False
    Label16.Visible = False
    Label17.Visible = False
    Label18.Visible = False
    Label19.Visible = False
    Label20.Visible = False
    Label21.Visible = False
    Label22.Visible = False
    Label23.Visible = False
    Label24.Visible = False
    Label25.Visible = False
    Label26.Visible = False
    Picture1.Visible = False
    Picture2.Visible = False
    Text4.Visible = False
    FormIstogramma.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
            Label26.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
            Label26.BackColor = &H80000005
        End If
    Picture1.Cls
    FormIstogramma.ScaleWidth = 800
    FormIstogramma.ScaleHeight = 480
    FormIstogramma.Refresh

    YIN = YMI: YSU = YMA

    Picture1.Left = 100
    Picture1.Top = 80
    Picture1.Height = 340
    Picture1.Width = 460
    Picture1.ScaleHeight = 340
    Picture1.ScaleWidth = 460
    Picture1.Visible = True

    Picture1.Line (0, 0)-(11, 0)
    Picture1.Line (0, 85)-(11, 85)
    Picture1.Line (0, 170)-(11, 170)
    Picture1.Line (0, 255)-(11, 255)
    Picture1.Line (0, 339)-(11, 339)
    
    Picture1.Line (459, 0)-(448, 0)
    Picture1.Line (459, 85)-(448, 85)
    Picture1.Line (459, 170)-(448, 170)
    Picture1.Line (459, 255)-(448, 255)
    Picture1.Line (459, 339)-(448, 339)
    
    Picture1.Line (0, 43)-(6, 43)
    Picture1.Line (0, 128)-(6, 128)
    Picture1.Line (0, 213)-(6, 213)
    Picture1.Line (0, 298)-(6, 298)
    
    Picture1.Line (459, 43)-(453, 43)
    Picture1.Line (459, 128)-(453, 128)
    Picture1.Line (459, 213)-(453, 213)
    Picture1.Line (459, 298)-(453, 298)
    
    Picture1.Line (0, 339)-(0, 329)
    Picture1.Line (115, 339)-(115, 329)
    Picture1.Line (230, 339)-(230, 329)
    Picture1.Line (345, 339)-(345, 329)
    Picture1.Line (459, 339)-(459, 329)
    
    Picture1.Line (0, 0)-(0, 10)
    Picture1.Line (115, 0)-(115, 10)
    Picture1.Line (230, 0)-(230, 10)
    Picture1.Line (345, 0)-(345, 10)
    Picture1.Line (459, 0)-(459, 10)
    
    Picture1.Line (58, 339)-(58, 332)
    Picture1.Line (173, 339)-(173, 332)
    Picture1.Line (288, 339)-(288, 332)
    Picture1.Line (403, 339)-(403, 332)
    
    Picture1.Line (58, 0)-(58, 7)
    Picture1.Line (173, 0)-(173, 7)
    Picture1.Line (288, 0)-(288, 7)
    Picture1.Line (403, 0)-(403, 7)
    
    FF = 340 / MX
    For I = 1 To CL
        DD(I) = Int(CC(I) * FF)
    Next I

    LargCol = 459 / CL
    XF = 0
    I = 1
        If FlagDatiAColori = True Then
            Picture1.FillStyle = 0
            Picture1.FillColor = QBColor(12)
        Else
            Picture1.FillStyle = 5
            Picture1.FillColor = QBColor(0)
        End If
'    For XI = XF To XF + LARGCOL * (CL - 1) Step LARGCOL
    For XI = XF To XF + LargCol * CL Step LargCol
        Picture1.Line (XI, 340)-Step(LargCol, -DD(I)), , B
'        Picture1.Line (XI, 340)-(XI, 340 - DD(I))
'        Picture1.Line (XI, 340 - DD(I))-(XI + LARGCOL, 340 - DD(I))
'        Picture1.Line (XI + LARGCOL, 340 - DD(I))-(XI + LARGCOL, 340)
'        Picture1.Line (XI + LARGCOL, 340)-(XI, 340)
        I = I + 1
    Next XI
    
'traccia gaussiana teorica
    
    Inizio = (((Media - 4 * DevStnd) - LW) / (HI - LW)) * ((XF + (459 / CL) * CL) - XF)
    Fine = (((Media + 4 * DevStnd) - LW) / (HI - LW)) * ((XF + (459 / CL) * CL) - XF)
    
' calcola l'area dell'istogramma e deriva l'altezza della gaussiana
' assunta come un triangolo equilatero con uguale superficie
    SommaAree = 0
    For I = 1 To CL
        SommaAree = SommaAree + (DD(I) * LargCol)
    Next I
    Altmed = SommaAree * 2 / (Fine - Inizio)
    Altmed = Altmed * 1.1 ' fattore di correzione

        I = XF
        Ampiezza = 4 / (Fine - Inizio)
        Deviata = ((I - (XF + Inizio)) * Ampiezza) - 2
        Ordinata = Exp(-Deviata * Deviata)
        Picture1.PSet (I, 340 - (Ordinata * Altmed))
        
        Picture1.PSet (XF + Inizio, 340)
        For I = XF + Inizio To XF + Fine
            Ampiezza = 4 / (Fine - Inizio)
            Deviata = ((I - (XF + Inizio)) * Ampiezza) - 2
            Ordinata = Exp(-Deviata * Deviata)
            Picture1.Line -(I, 340 - (Ordinata * Altmed))
        Next I

' fine tracciamento gaussiana teorica

    Y5 = YSU: Y4 = Int((YSU - (YSU - YIN) / 4) + 0.5): Y3 = Int((YIN + (YSU - YIN) / 2) + 0.5): Y2 = Int((YIN + (YSU - YIN) / 4) + 0.5): Y1 = YIN
    X1 = XMI: X2 = XMI + (XMA - XMI) / 4: X3 = XMI + (XMA - XMI) / 2: X4 = XMA - (XMA - XMI) / 4: X5 = XMA
    
    Label13.Left = 330 - (Label13.Width / 2): Label13.Top = 27
    Label14.Left = 330 - (Label14.Width / 2): Label14.Top = 448
    Label15.Left = 100: Label15.Top = 61
    Label16.Left = 100: Label16.Top = 428
    Label17.Left = 215 - (Label17.Width / 2): Label17.Top = 428
    Label18.Left = 330 - (Label18.Width / 2): Label18.Top = 428
    Label19.Left = 445 - (Label19.Width / 2): Label19.Top = 428
    Label20.Left = 560 - Label20.Width: Label20.Top = 428
    Label21.Left = 90 - Label21.Width: Label21.Top = 409
    Label22.Left = 90 - Label22.Width: Label22.Top = 329
    Label23.Left = 90 - Label23.Width: Label23.Top = 244
    Label24.Left = 90 - Label24.Width: Label24.Top = 160
    Label25.Left = 90 - Label25.Width: Label25.Top = 79

    Label13.Visible = True
    Label14.Visible = True
    Label15.Visible = True
    Label16.Visible = True
    Label17.Visible = True
    Label18.Visible = True
    Label19.Visible = True
    Label20.Visible = True
    Label21.Visible = True
    Label22.Visible = True
    Label23.Visible = True
    Label24.Visible = True
    Label25.Visible = True

    FlagGrafico = True
    FlagStatistica = False
    Label13.Caption = Titolo
    Label14.Caption = InAscisse
    Label15.Caption = InOrdinate
    Label16.Caption = Format(X1, "#####0.0####;-#####0.0####")
    Label17.Caption = Format(X2, "#####0.0####;-#####0.0####")
    Label18.Caption = Format(X3, "#####0.0####;-#####0.0####")
    Label19.Caption = Format(X4, "#####0.0####;-#####0.0####")
    Label20.Caption = Format(X5, "#####0.0####;-#####0.0####")
    Label21.Caption = Format(Y1, "#####0")
    Label22.Caption = Format(Y2, "#####0")
    Label23.Caption = Format(Y3, "#####0")
    Label24.Caption = Format(Y4, "#####0")
    Label25.Caption = Format(Y5, "#####0")
    
    Label26.Visible = True
    Label26.Left = 574
    Label26.Top = Picture1.Top + (Picture1.Height / 3)
    Label26.Caption = "" & Chr(13)
    Label26.Caption = Label26.Caption & "   Dati = " & NumDati & "   " & Chr(13) & Chr(10)
    Label26.Caption = Label26.Caption & "   Media = " & Format(Media, "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   DS = " & Format(DevStnd, "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   CV = " & Format(CV, "#####0.00") & " %   " & Chr(13) & Chr(10)
    Label26.Caption = Label26.Caption & "   Minimo = " & Format(X(1), "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   1" & Chr(176) & " quartile = " & Format(Media - (0.674 * DevStnd), "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   2" & Chr(176) & " quartile = " & Format(Media, "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   3" & Chr(176) & " quartile = " & Format(Media + (0.674 * DevStnd), "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   Massimo = " & Format(X(NumDati), "#####0.0####") & "   " & Chr(13)
    Label26.Top = Picture1.Top + (Picture1.Height / 2) - (Label26.Height / 2)
    FlagCon = True
    FlagSenza = False
    Screen.MousePointer = 1

End Sub

Private Sub IstoSenza_Click()
    
    Dim I, XX, YIN, YSU, FX, FY, CX, CY
    Dim X1, X2, X3, X4, X5
    Dim Y1, Y2, Y3, Y4, Y5
    Dim FF, LargCol, XF, XI
'    Dim SommaAree, ClassiValide, AltMed
'    Dim Inizio, Fine, Ampiezza, Deviata, Ordinata

    Screen.MousePointer = 11
    Shape1.Visible = False
    Frame1.Visible = False
    Label13.Visible = False
    Label14.Visible = False
    Label15.Visible = False
    Label16.Visible = False
    Label17.Visible = False
    Label18.Visible = False
    Label19.Visible = False
    Label20.Visible = False
    Label21.Visible = False
    Label22.Visible = False
    Label23.Visible = False
    Label24.Visible = False
    Label25.Visible = False
    Label26.Visible = False
    Picture1.Visible = False
    Picture2.Visible = False
    Text4.Visible = False
    FormIstogramma.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
            Label26.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
            Label26.BackColor = &H80000005
        End If
    Picture1.Cls
    FormIstogramma.ScaleWidth = 800
    FormIstogramma.ScaleHeight = 480
    FormIstogramma.Refresh

    YIN = YMI: YSU = YMA

    Picture1.Left = 100
    Picture1.Top = 80
    Picture1.Height = 340
    Picture1.Width = 460
    Picture1.ScaleHeight = 340
    Picture1.ScaleWidth = 460
    Picture1.Visible = True

    Picture1.Line (0, 0)-(11, 0)
    Picture1.Line (0, 85)-(11, 85)
    Picture1.Line (0, 170)-(11, 170)
    Picture1.Line (0, 255)-(11, 255)
    Picture1.Line (0, 339)-(11, 339)
    
    Picture1.Line (459, 0)-(448, 0)
    Picture1.Line (459, 85)-(448, 85)
    Picture1.Line (459, 170)-(448, 170)
    Picture1.Line (459, 255)-(448, 255)
    Picture1.Line (459, 339)-(448, 339)
    
    Picture1.Line (0, 43)-(6, 43)
    Picture1.Line (0, 128)-(6, 128)
    Picture1.Line (0, 213)-(6, 213)
    Picture1.Line (0, 298)-(6, 298)
    
    Picture1.Line (459, 43)-(453, 43)
    Picture1.Line (459, 128)-(453, 128)
    Picture1.Line (459, 213)-(453, 213)
    Picture1.Line (459, 298)-(453, 298)
    
    Picture1.Line (0, 339)-(0, 329)
    Picture1.Line (115, 339)-(115, 329)
    Picture1.Line (230, 339)-(230, 329)
    Picture1.Line (345, 339)-(345, 329)
    Picture1.Line (459, 339)-(459, 329)
    
    Picture1.Line (0, 0)-(0, 10)
    Picture1.Line (115, 0)-(115, 10)
    Picture1.Line (230, 0)-(230, 10)
    Picture1.Line (345, 0)-(345, 10)
    Picture1.Line (459, 0)-(459, 10)
    
    Picture1.Line (58, 339)-(58, 332)
    Picture1.Line (173, 339)-(173, 332)
    Picture1.Line (288, 339)-(288, 332)
    Picture1.Line (403, 339)-(403, 332)
    
    Picture1.Line (58, 0)-(58, 7)
    Picture1.Line (173, 0)-(173, 7)
    Picture1.Line (288, 0)-(288, 7)
    Picture1.Line (403, 0)-(403, 7)
    
    FF = 340 / MX
    For I = 1 To CL
        DD(I) = Int(CC(I) * FF)
    Next I

    LargCol = 459 / CL
    XF = 0
    I = 1
        If FlagDatiAColori = True Then
            Picture1.FillStyle = 0
            Picture1.FillColor = QBColor(12)
        Else
            Picture1.FillStyle = 5
            Picture1.FillColor = QBColor(0)
        End If
'    For XI = XF To XF + LARGCOL * (CL - 1) Step LARGCOL
    For XI = XF To XF + LargCol * CL Step LargCol
        Picture1.Line (XI, 340)-Step(LargCol, -DD(I)), , B
'        Picture1.Line (XI, 340)-(XI, 340 - DD(I))
'        Picture1.Line (XI, 340 - DD(I))-(XI + LARGCOL, 340 - DD(I))
'        Picture1.Line (XI + LARGCOL, 340 - DD(I))-(XI + LARGCOL, 340)
'        Picture1.Line (XI + LARGCOL, 340)-(XI, 340)
        I = I + 1
    Next XI
    
'traccia gaussiana teorica
    
'    Inizio = (((Media - 3 * DevStnd) - LW) / (HI - LW)) * ((XF + (459 / CL) * CL) - XF)
'    Fine = (((Media + 3 * DevStnd) - LW) / (HI - LW)) * ((XF + (459 / CL) * CL) - XF)
'
' calcola l'area dell'istogramma e deriva l'altezza della gaussiana
' assunta come un triangolo equilatero con uguale superficie
'    SommaAree = 0
'    For I = 1 To CL
'        SommaAree = SommaAree + (DD(I) * LargCol)
'    Next I
'    AltMed = SommaAree * 2 / (Fine - Inizio)
'    Altmed = Altmed * 1.1 ' fattore di correzione
'
'        I = XF
'        Ampiezza = 4 / (Fine - Inizio)
'        Deviata = ((I - (XF + Inizio)) * Ampiezza) - 2
'        Ordinata = Exp(-Deviata * Deviata)
'        Picture1.PSet (I, 340 - (Ordinata * AltMed))
'
'        For I = XF + Inizio To XF + Fine
'            Ampiezza = 4 / (Fine - Inizio)
'            Deviata = ((I - (XF + Inizio)) * Ampiezza) - 2
'            Ordinata = Exp(-Deviata * Deviata)
'            Picture1.Line -(I, 340 - (Ordinata * AltMed))
'        Next I
'
' fine tracciamento gaussiana teorica

    Y5 = YSU: Y4 = Int((YSU - (YSU - YIN) / 4) + 0.5): Y3 = Int((YIN + (YSU - YIN) / 2) + 0.5): Y2 = Int((YIN + (YSU - YIN) / 4) + 0.5): Y1 = YIN
    X1 = XMI: X2 = XMI + (XMA - XMI) / 4: X3 = XMI + (XMA - XMI) / 2: X4 = XMA - (XMA - XMI) / 4: X5 = XMA
    
    Label13.Left = 330 - (Label13.Width / 2): Label13.Top = 27
    Label14.Left = 330 - (Label14.Width / 2): Label14.Top = 448
    Label15.Left = 100: Label15.Top = 61
    Label16.Left = 100: Label16.Top = 428
    Label17.Left = 215 - (Label17.Width / 2): Label17.Top = 428
    Label18.Left = 330 - (Label18.Width / 2): Label18.Top = 428
    Label19.Left = 445 - (Label19.Width / 2): Label19.Top = 428
    Label20.Left = 560 - Label20.Width: Label20.Top = 428
    Label21.Left = 90 - Label21.Width: Label21.Top = 409
    Label22.Left = 90 - Label22.Width: Label22.Top = 329
    Label23.Left = 90 - Label23.Width: Label23.Top = 244
    Label24.Left = 90 - Label24.Width: Label24.Top = 160
    Label25.Left = 90 - Label25.Width: Label25.Top = 79

    Label13.Visible = True
    Label14.Visible = True
    Label15.Visible = True
    Label16.Visible = True
    Label17.Visible = True
    Label18.Visible = True
    Label19.Visible = True
    Label20.Visible = True
    Label21.Visible = True
    Label22.Visible = True
    Label23.Visible = True
    Label24.Visible = True
    Label25.Visible = True

    FlagGrafico = True
    FlagStatistica = False
    Label13.Caption = Titolo
    Label14.Caption = InAscisse
    Label15.Caption = InOrdinate
    Label16.Caption = Format(X1, "#####0.0####;-#####0.0####")
    Label17.Caption = Format(X2, "#####0.0####;-#####0.0####")
    Label18.Caption = Format(X3, "#####0.0####;-#####0.0####")
    Label19.Caption = Format(X4, "#####0.0####;-#####0.0####")
    Label20.Caption = Format(X5, "#####0.0####;-#####0.0####")
    Label21.Caption = Format(Y1, "#####0")
    Label22.Caption = Format(Y2, "#####0")
    Label23.Caption = Format(Y3, "#####0")
    Label24.Caption = Format(Y4, "#####0")
    Label25.Caption = Format(Y5, "#####0")
    
    Label26.Visible = True
    Label26.Left = 574
    Label26.Top = Picture1.Top + (Picture1.Height / 3)
    Label26.Caption = "" & Chr(13)
    Label26.Caption = Label26.Caption & "   Dati = " & NumDati & "   " & Chr(13) & Chr(10)
    Label26.Caption = Label26.Caption & "   Media = " & Format(Media, "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   DS = " & Format(DevStnd, "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   CV = " & Format(CV, "#####0.00") & " %   " & Chr(13) & Chr(10)
    Label26.Caption = Label26.Caption & "   Minimo = " & Format(X(1), "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   1" & Chr(176) & " quartile = " & Format(Media - (0.674 * DevStnd), "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   2" & Chr(176) & " quartile = " & Format(Media, "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   3" & Chr(176) & " quartile = " & Format(Media + (0.674 * DevStnd), "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   Massimo = " & Format(X(NumDati), "#####0.0####") & "   " & Chr(13)
    Label26.Top = Picture1.Top + (Picture1.Height / 2) - (Label26.Height / 2)
    FlagCon = False
    FlagSenza = True
    Screen.MousePointer = 1

End Sub

Private Sub MenuFile_Click()
    
    FileStampa.Enabled = False
        If FlagGrafico = True Then
            If FormIstogramma.Width < Screen.Width Then
                FileStampa.Enabled = True
            End If
        End If
        If FlagStatistica = True Then
            FileStampa.Enabled = True
        End If

End Sub

Private Sub MenuGrafica_Click()
    
    GraficaConfigurazioneColori = False
    GraficaScaleLegende = False
    
    If FlagGrafico = True Then
        GraficaConfigurazioneColori = True
        GraficaScaleLegende = True
    End If

End Sub

Private Sub MenuModifica_Click()

    If FlagStatistica = True Then
        ModificaCopia.Enabled = True
        ModificaTaglia.Enabled = True
    Else
        ModificaCopia.Enabled = False
        ModificaTaglia.Enabled = False
    End If

End Sub

Private Sub ModificaCopia_Click()

    Clipboard.Clear
    Clipboard.SetText Text4.SelText

End Sub

Private Sub ModificaTaglia_Click()

    Clipboard.Clear
    Clipboard.SetText Text4.SelText
    Text4.SelText = ""

End Sub

Private Sub StatisticaStatisticheElementari_Click()
    Dim N, p, c0, c1, c2, t1, t2, e1
    Dim t5, i1, i2, i3, i4, u1, t
    Dim I

    Screen.MousePointer = 1
    Frame1.Visible = False
    Shape1.Visible = False
    Label13.Visible = False
    Label14.Visible = False
    Label15.Visible = False
    Label16.Visible = False
    Label17.Visible = False
    Label18.Visible = False
    Label19.Visible = False
    Label20.Visible = False
    Label21.Visible = False
    Label22.Visible = False
    Label23.Visible = False
    Label24.Visible = False
    Label25.Visible = False
    Label26.Visible = False
    Picture1.Visible = False
    Text4.Visible = False
    Cls
    FormIstogramma.Refresh
    

        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Statistiche elementari parametriche (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File      : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella   : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & "Variabile : " & InAscisse & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        
        Text4.Text = Text4.Text & "Numero dei dati             = " & NumDati & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Media                       = " & Format(Media, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Deviazione standard         = " & Format(DevStnd, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Coefficiente di variazione  = " & Format(CV, "#####0.00") & " %" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Errore standard della media = " & Format(ErrStnd, "#####0.0####") & Chr(13) & Chr(10)
        
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Coefficiente di asimmetria  = " & Format(GiUno, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Deviazione standard         = " & Format(EsseUno, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Rapporto                    = " & Format(Abs(GiUno / EsseUno), "#####0.00") & Chr(13) & Chr(10)
    
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Coefficiente di curtosi     = " & Format(GiDue, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Deviazione standard         = " & Format(EsseDue, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Rapporto                    = " & Format(Abs(GiDue / EsseDue), "#####0.00") & Chr(13) & Chr(10)
        
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Valore minimo               = " & Format(X(1), "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Primo quartile              = " & Format(Media - (0.674 * DevStnd), "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Secondo quartile            = " & Format(Media, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Terzo quartile              = " & Format(Media + (0.674 * DevStnd), "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Valore massimo              = " & Format(X(NumDati), "#####0.0####") & Chr(13) & Chr(10)

        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Statistiche elementari parametriche (*fine*)" & Chr(13) & Chr(10)
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        FlagGrafico = False
        Screen.MousePointer = 1

Exit Sub

' calcolo di t a partire da p e dai gradi di libertà n
        N = NumDati - 2
        p = 0.05
        c0 = 2.515517
        c1 = 0.802853
        c2 = 0.010328
        t1 = 1.432788
        t2 = 0.189269
        e1 = p / 2
        t5 = Sqr(Abs(2 * Log(e1)))
        i1 = c0 + (c1 * t5) + (c2 * t5 ^ 2)
        i2 = 1 + (t1 * t5) + (t2 * t5 ^ 2) + (0.001308 * t5 ^ 3)
        u1 = t5 - (i1 / i2)
        i1 = (u1 ^ 3 + u1) / (4 * N)
        i2 = ((5 * u1 ^ 5) + (16 * u1 ^ 3) + (3 * u1)) / (96 * N ^ 2)
        i3 = ((3 * u1 ^ 7) + (19 * u1 ^ 5) + (17 * u1 ^ 3) - (15 * u1)) / (384 * N ^ 3)
        i4 = ((79 * u1 ^ 9) + (776 * u1 ^ 7) + (1482 * u1 ^ 5) - (1920 * u1 ^ 3) - (945 * u1)) / (92160 * N ^ 4)
        e1 = ((27 * u1 ^ 11) + (339 * u1 ^ 9) + (930 * u1 ^ 7) - (1782 * u1 ^ 5) - (765 * u1 ^ 3) + (17955 * u1)) / (368640 * N ^ 5)
        t = u1 + i1 + i2 + i3 + i4 + e1

End Sub

Private Sub StatisticaSuddivisioneInClassi_Click()
    Dim I
    Dim Stringa As String

    Screen.MousePointer = 11
    Frame1.Visible = False
    Shape1.Visible = False
    Label13.Visible = False
    Label14.Visible = False
    Label15.Visible = False
    Label16.Visible = False
    Label17.Visible = False
    Label18.Visible = False
    Label19.Visible = False
    Label20.Visible = False
    Label21.Visible = False
    Label22.Visible = False
    Label23.Visible = False
    Label24.Visible = False
    Label25.Visible = False
    Label26.Visible = False
    Picture1.Visible = False
    Text4.Visible = False
    Cls
    FormIstogramma.Refresh
    
        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Suddivisione dei dati in classi (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File      : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella   : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & "Variabile : " & InAscisse & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
                                                                    
        Text4.Text = Text4.Text & "Classe      da...         a...      Dati" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "------  ------------  ------------  ----" & Chr(13) & Chr(10)

        For I = 1 To CL
            Stringa = Format(I, "##")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@") & "  "
            Stringa = Format(LW + ((I - 1) * AM), "#####0.0####")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@") & "  "
            Stringa = Format(LW + (I * AM), "#####0.0####")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@") & "  "
            Stringa = Format(CC(I), "##0"): If Stringa = "" Then Stringa = "0"
            Text4.Text = Text4.Text & Format(Stringa, "@@@@") & "  "
            Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Next I
        
        Text4.Text = Text4.Text & "------  ------------  ------------  ----" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Suddivisione dei dati in classi (*fine*)" & Chr(13) & Chr(10)
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        FlagGrafico = False
        Screen.MousePointer = 1

End Sub

Private Sub StatisticaTabellaDeiPercentili_Click()
    Dim I, K
    ReDim Centile(21)
    ReDim Zeta(21)
    Dim Dummy As Double
    Dim Percentile As Double
    Dim Valore As Double
    Dim LimiteInferiore As Double
    Dim LimiteSuperiore As Double
    Dim Stringa As String

    Screen.MousePointer = 11
    Frame1.Visible = False
    Shape1.Visible = False
    Label13.Visible = False
    Label14.Visible = False
    Label15.Visible = False
    Label16.Visible = False
    Label17.Visible = False
    Label18.Visible = False
    Label19.Visible = False
    Label20.Visible = False
    Label21.Visible = False
    Label22.Visible = False
    Label23.Visible = False
    Label24.Visible = False
    Label25.Visible = False
    Label26.Visible = False
    Picture1.Visible = False
    Text4.Visible = False
    Cls
    FormIstogramma.Refresh
    
    FormGauge.Show
    FormGauge.Caption = "Calcolo dei percentili"
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Tabella dei percentili parametrici (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File      : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella   : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & "Variabile : " & InAscisse & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
                                                                    
        Text4.Text = Text4.Text & "                          Limiti di confidenza al 90%      Valore     " & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                          ---------------------------  non parametrico" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Percentile     Valore       inferiore     superiore      comparativo  " & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "----------  ------------  ------------  -------------  ---------------" & Chr(13) & Chr(10)
        
        Centile(1) = 2.5: Zeta(1) = -1.96
        Centile(2) = 5: Zeta(2) = -1.645
        Centile(3) = 10: Zeta(3) = -1.282
        Centile(4) = 15: Zeta(4) = -1.036
        Centile(5) = 20: Zeta(5) = -0.841
        Centile(6) = 25: Zeta(6) = -0.674
        Centile(7) = 30: Zeta(7) = -0.524
        Centile(8) = 35: Zeta(8) = -0.385
        Centile(9) = 40: Zeta(9) = -0.253
        Centile(10) = 45: Zeta(10) = -0.125
        Centile(11) = 50: Zeta(11) = 0
        Centile(12) = 55: Zeta(12) = 0.125
        Centile(13) = 60: Zeta(13) = 0.253
        Centile(14) = 65: Zeta(14) = 0.385
        Centile(15) = 70: Zeta(15) = 0.524
        Centile(16) = 75: Zeta(16) = 0.674
        Centile(17) = 80: Zeta(17) = 0.841
        Centile(18) = 85: Zeta(18) = 1.036
        Centile(19) = 90: Zeta(19) = 1.282
        Centile(20) = 95: Zeta(20) = 1.645
        Centile(21) = 97.5: Zeta(21) = 1.96

        For I = 1 To 21
            Valore = Media + (Zeta(I) * DevStnd)
            Dummy = Sqr(1 + (((Zeta(I) ^ 2) * NumDati) / (2 * (NumDati - 1))))
            Dummy = DevStnd * Dummy / Sqr(NumDati)
            LimiteInferiore = Valore - (1.645 * Dummy)
            LimiteSuperiore = Valore + (1.645 * Dummy)
            Stringa = Format(Centile(I), "##0.0")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@") & "  "
            Stringa = Format(Valore, "#####0.0####")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@") & "  "
            Stringa = Format(LimiteInferiore, "#####0.0####")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@") & "  "
            Stringa = Format(LimiteSuperiore, "#####0.0####")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@@") & "  "

' valore comparativo non parametrico del percentile
            Valore = 0
            Percentile = Format(Centile(I), "##0.0")
            For K = 1 To Z1 - 1
                If Percentile = Y(K) Then
                    Valore = Format(X(K), "#####0.0####")
                    Exit For
                End If
                If Percentile = Y(K + 1) Then
                    Valore = Format(X(K + 1), "#####0.0####")
                    Exit For
                End If
                If Percentile > Y(K) And Percentile < Y(K + 1) Then
                    Valore = Format(X(K) + ((X(K + 1) - X(K)) * (Percentile - Y(K)) / (Y(K + 1) - Y(K))), "#####0.0####")
                    Exit For
                End If
            Next K
            If Valore <> 0 Then
                Stringa = Format(Valore, "#####0.0####")
                Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@@@@")
            Else
                Text4.Text = Text4.Text & Format("###", "@@@@@@@@@@@@@@@")
            End If
            
            Text4.Text = Text4.Text & Chr(13) & Chr(10)
            
            FormGauge.Label1.Caption = Int(100 * I / 21) & "%"
            FormGauge!Gauge1.Value = Int(100 * I / 21)
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
        Next I
        
        Text4.Text = Text4.Text & "----------  ------------  ------------  -------------  ---------------" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Tabella dei percentili parametrici (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        FlagGrafico = False
        Screen.MousePointer = 1

End Sub

