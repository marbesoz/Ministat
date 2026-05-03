VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormCumulativa 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Statistiche elementari non parametriche"
   ClientHeight    =   6210
   ClientLeft      =   1110
   ClientTop       =   1725
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
   Icon            =   "FORMCUMU.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6210
   ScaleWidth      =   7365
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   4560
      ScaleHeight     =   345
      ScaleWidth      =   945
      TabIndex        =   42
      Top             =   5640
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
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   615
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
      TabIndex        =   29
      Top             =   120
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   7095
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ripristina valori originari"
         Height          =   372
         Left            =   2280
         TabIndex        =   40
         Top             =   2520
         Width           =   2532
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Annulla"
         Height          =   372
         Left            =   5160
         TabIndex        =   25
         Top             =   2520
         Width           =   1212
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Valore X  >>     "
         Height          =   375
         Left            =   2880
         TabIndex        =   24
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Percentile Y  >>"
         Height          =   375
         Left            =   2880
         TabIndex        =   20
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "OK"
         Height          =   372
         Left            =   720
         TabIndex        =   16
         Top             =   2520
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
      Begin MSMask.MaskEdBox MaskedEdit6 
         Height          =   285
         Left            =   1200
         TabIndex        =   23
         Top             =   3765
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
      Begin MSMask.MaskEdBox MaskedEdit5 
         Height          =   285
         Left            =   1200
         TabIndex        =   19
         Top             =   3285
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
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   255
         Left            =   5280
         TabIndex        =   18
         Top             =   3770
         Width           =   1455
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Percentile Y:"
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
         TabIndex        =   22
         Top             =   3780
         Width           =   975
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   255
         Left            =   5280
         TabIndex        =   21
         Top             =   3300
         Width           =   1455
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Valore X:"
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
         TabIndex        =   17
         Top             =   3315
         Width           =   975
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
      TabIndex        =   41
      Top             =   120
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
      TabIndex        =   39
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
      TabIndex        =   38
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
      TabIndex        =   37
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
      TabIndex        =   36
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
      TabIndex        =   35
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
      TabIndex        =   34
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
      TabIndex        =   33
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
      TabIndex        =   32
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
      TabIndex        =   31
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
      TabIndex        =   30
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
      TabIndex        =   28
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
      TabIndex        =   27
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
      TabIndex        =   26
      Top             =   360
      Width           =   6492
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   4692
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
      Begin VB.Menu DistribuzioneCumulativStandard 
         Caption         =   "&Distribuzione cumulativa standard"
      End
      Begin VB.Menu DistribuzioneCumulativaRipiegata 
         Caption         =   "D&istribuzione cumulativa ripiegata"
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
            Caption         =   "&Percentili a colori"
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
Attribute VB_Name = "FormCumulativa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim X() As Double
    Dim Y()
    Dim Riga, ColonnaX, ColonnaY, FineDati, Z1
    Dim XMI, XMA, YMI, YMA
    Dim XMI_Default, XMA_Default
    Dim YMI_Default, YMA_Default
    Dim Titolo1, Titolo2
    Dim InAscisse, InOrdinate1, InOrdinate2
    Dim FlagGrafico, FlagStatistica
    Dim FlagSfondoGrigio
    Dim FlagDatiAColori
    Dim Flag1, Flag2

Private Sub AiutoCalcolatrice_Click()
    Dim Calcolatrice
    On Error GoTo ErroreCalcolatriceTre:
    
    Calcolatrice = Shell("Calc.exe", 1)
        
Exit Sub

ErroreCalcolatriceTre:
    MsgBox "Calcolatrice non disponibile.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcolatrice"
Exit Sub

End Sub

Private Sub AiutoGuida_Click()

    FormPrincipale!CMDialog1.HelpFile = App.Path & "\ministat.hlp"
    FormPrincipale!CMDialog1.HelpCommand = cdlHelpContents
    FormPrincipale!CMDialog1.ShowHelp

End Sub

Private Sub Command1_Click()
    Dim FlagErroreScale
    On Error GoTo Cumulativa
        
    FlagErroreScale = False
    MaskedEdit1.Text = FormattaStringa(MaskedEdit1.Text)
    MaskedEdit2.Text = FormattaStringa(MaskedEdit2.Text)
    MaskedEdit3.Text = FormattaStringa(MaskedEdit3.Text)
    MaskedEdit4.Text = FormattaStringa(MaskedEdit4.Text)
    If CDbl(MaskedEdit1.Text) > XMI_Default Then
        MaskedEdit1.Text = Format(XMI_Default, "+#####0.0####;-#####0.0####")
        FlagErroreScale = True
    End If
    If CDbl(MaskedEdit2.Text) < XMA_Default Then
        MaskedEdit2.Text = Format(XMA_Default, "+#####0.0####;-#####0.0####")
        FlagErroreScale = True
    End If
    If CDbl(MaskedEdit3.Text) > YMI_Default Then
        MaskedEdit3.Text = Format(YMI_Default, "+#####0.0####;-#####0.0####")
        FlagErroreScale = True
    End If
    If CDbl(MaskedEdit4.Text) < YMA_Default Then
        MaskedEdit4.Text = Format(YMA_Default, "+#####0.0####;-#####0.0####")
        FlagErroreScale = True
    End If
        
        If FlagErroreScale = False Then
            
            If Flag1 = True Then
                Titolo1 = Text1.Text
                InOrdinate1 = Text3.Text
            Else
                Titolo2 = Text1.Text
                InOrdinate2 = Text3.Text
            End If
            InAscisse = Text2.Text
            XMI = MaskedEdit1.Text
            XMA = MaskedEdit2.Text
            YMI = MaskedEdit3.Text
            YMA = MaskedEdit4.Text
                If Flag1 = True Then
                    Label13.Caption = Titolo1
                    Label15.Caption = InOrdinate1
                Else
                    Label13.Caption = Titolo2
                    Label15.Caption = InOrdinate2
                End If
            Label14.Caption = InAscisse

'        If FlagGrafico = True Then

        If Flag1 = True Then
            DistribuzioneCumulativStandard_Click
        End If
        If Flag2 = True Then
            DistribuzioneCumulativaRipiegata_Click
        End If

'        End If
        
        Frame1.Visible = False
        Shape1.Visible = False
        Command5.Enabled = True

        End If

Cumulativa:
        MaskedEdit1.Text = Format(XMI_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit2.Text = Format(XMA_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit3.Text = Format(YMI_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit4.Text = Format(YMA_Default, "+#####0.0####;-#####0.0####")

End Sub

Private Sub Command2_Click()
    On Error GoTo GestioneErroriCommand2
    Dim YY, I

        Screen.MousePointer = 11
        Label10.Caption = ""

    MaskedEdit5.Text = FormattaStringa(MaskedEdit5.Text)
    If CDbl(MaskedEdit5.Text) >= X(1) And CDbl(MaskedEdit5.Text) <= X(Z1) Then
    
    For I = 1 To Z1 - 1
        If CDbl(MaskedEdit5.Text) = X(I) Then
            YY = Y(I)
            Label10.Caption = Format(YY, "#####0.0####")
            Exit For
        End If
        If CDbl(MaskedEdit5.Text) = X(I + 1) Then
            YY = Y(I + 1)
            Label10.Caption = Format(YY, "#####0.0####")
            Exit For
        End If
        If CDbl(MaskedEdit5.Text) > X(I) And CDbl(MaskedEdit5.Text) < X(I + 1) Then
            YY = Y(I) + ((Y(I + 1) - Y(I)) * (MaskedEdit5.Text - X(I)) / (X(I + 1) - X(I)))
            Label10.Caption = Format(YY, "#####0.0####")
            Exit For
        End If
    Next I
        Screen.MousePointer = 1
        
    Else
        MsgBox "Il percentile corrispondente al valore" + Chr(10) + "richiesto non è computabile.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "Calcolo non eseguito"
        Screen.MousePointer = 1
    
    End If


Exit Sub

GestioneErroriCommand2:
    Label10.Caption = ""
    Screen.MousePointer = 1
    Exit Sub

End Sub

Private Sub Command3_Click()
    On Error GoTo GestioneErroriCommand3
    Dim XX, I

        Screen.MousePointer = 11
        Label12.Caption = ""
    
    MaskedEdit6.Text = FormattaStringa(MaskedEdit6.Text)
    If CDbl(MaskedEdit6.Text) >= Y(1) And CDbl(MaskedEdit6.Text) <= Y(Z1) Then

    For I = 1 To Z1 - 1
        If CDbl(MaskedEdit6.Text) = Y(I) Then
            XX = X(I)
            Label12.Caption = Format(XX, "#####0.0####")
            Exit For
        End If
        If CDbl(MaskedEdit6.Text) = Y(I + 1) Then
            XX = X(I + 1)
            Label12.Caption = Format(XX, "#####0.0####")
            Exit For
        End If
        If CDbl(MaskedEdit6.Text) > Y(I) And CDbl(MaskedEdit6.Text) < Y(I + 1) Then
            XX = X(I) + ((X(I + 1) - X(I)) * (MaskedEdit6.Text - Y(I)) / (Y(I + 1) - Y(I)))
            Label12.Caption = Format(XX, "#####0.0####")
            Exit For
        End If
    Next I
        Screen.MousePointer = 1

    Else
        MsgBox "Il valore corrispondente al percentile" + Chr(10) + "richiesto non è computabile.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "Calcolo non eseguito"
        Screen.MousePointer = 1

    End If
        
Exit Sub

GestioneErroriCommand3:
    Label12.Caption = ""
    Screen.MousePointer = 1
    Exit Sub

End Sub

Private Sub Command4_Click()
    
'    If FlagGrafico = True Then
        If Flag1 = True Then
            DistribuzioneCumulativStandard_Click
        End If
        If Flag2 = True Then
            DistribuzioneCumulativaRipiegata_Click
        End If

'    End If
    
    Frame1.Visible = False
    Shape1.Visible = False
    Command5.Enabled = True

End Sub

Private Sub Command5_Click()
    
    MaskedEdit1.Text = Format(XMI_Default, "+#####0.0####;-#####0.0####")
    MaskedEdit2.Text = Format(XMA_Default, "+#####0.0####;-#####0.0####")
    MaskedEdit3.Text = Format(YMI_Default, "+#####0.0####;-#####0.0####")
    MaskedEdit4.Text = Format(YMA_Default, "+#####0.0####;-#####0.0####")

    Screen.MousePointer = 1
    
End Sub

Private Sub DistribuzioneCumulativaRipiegata_Click()
    On Error GoTo GestioneErroriGraficaKrouwer
    Dim I, XX, YIN, YSU, FX, FY, CX, CY
    Dim X1, X2, X3, X4, X5
    Dim Y1, Y2, Y3, Y4, Y5
    Dim Dummy As Double

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
    FormCumulativa.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
            Label26.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
            Label26.BackColor = &H80000005
        End If
    Picture1.Cls
    FormCumulativa.ScaleWidth = 800
    FormCumulativa.ScaleHeight = 480
    FormCumulativa.Refresh

    YIN = YMI: YSU = YMA
    
    Picture1.Left = 100
    Picture1.Top = 80
    Picture1.Height = 340
    Picture1.Width = 460
    Picture1.ScaleHeight = 340
    Picture1.ScaleWidth = 460

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
    
    FX = 459 / (XMA - XMI): FY = 339 / (YSU - YIN)

        CX = (X(1) - XMI) * FX
        CY = 339 - (Y(1) - YIN) * FY
        Picture1.CurrentX = CX
        Picture1.CurrentY = CY
    
    Dummy = Format(50, "##0.0")
    For I = 1 To Z1 - 1
        If Dummy = Y(I) Then
            Exit For
        End If
        If Dummy = Y(I + 1) Then
            Exit For
        End If
        If Dummy > Y(I) And Dummy < Y(I + 1) Then
            XX = X(I) + ((X(I + 1) - X(I)) * (Dummy - Y(I)) / (Y(I + 1) - Y(I)))
            Exit For
        End If
    Next I
    
    For I = 2 To Z1
        If Y(I - 1) < 50 And Y(I) > 50 Then
            CX = (XX - XMI) * FX
            CY = 339 - (50 - YIN) * FY
                If FlagDatiAColori = True Then
                    Picture1.Line -(CX, CY), QBColor(12)
                Else
                    Picture1.Line -(CX, CY), QBColor(0)
                End If
            CX = (X(I) - XMI) * FX
            CY = 339 - ((100 - Y(I)) - YIN) * FY
                If FlagDatiAColori = True Then
                    Picture1.Line -(CX, CY), QBColor(12)
                Else
                    Picture1.Line -(CX, CY), QBColor(0)
                End If
        Else
            CX = (X(I) - XMI) * FX ' : If CX < 0 Or CX > 459 Then GoTo 10565
            CY = 339 - (Y(I) - YIN) * FY ' : If CY < 0 Or CY > 339 Then GoTo 10565
                If Y(I) > 50 Then
                    CY = 339 - ((100 - Y(I)) - YIN) * FY
                End If
                If FlagDatiAColori = True Then
                    Picture1.Line -(CX, CY), QBColor(12)
                Else
                    Picture1.Line -(CX, CY), QBColor(0)
                End If
' 10565 :
        End If
    Next I
    
    Y5 = YSU: Y4 = YSU - (YSU - YIN) / 4: Y3 = YIN + (YSU - YIN) / 2: Y2 = YIN + (YSU - YIN) / 4: Y1 = YIN
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
    Flag1 = False
    Flag2 = True
    Label13.Caption = Titolo2
    Label14.Caption = InAscisse
    Label15.Caption = InOrdinate2
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
    Picture1.Visible = True
    
    Label26.Visible = True
    Label26.Left = 574
    Label26.Top = Picture1.Top + (Picture1.Height / 3)
    Label26.Caption = "" & Chr(13)
    Label26.Caption = Label26.Caption & "   Dati = " & Z1 & "   " & Chr(13) & Chr(10)
    Dummy = Format(50, "##0.0")
    For I = 1 To Z1 - 1
        If Dummy = Y(I) Then
            XX = X(I)
            Label26.Caption = Label26.Caption & "   Mediana = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
        If Dummy = Y(I + 1) Then
            XX = X(I + 1)
            Label26.Caption = Label26.Caption & "   Mediana = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
        If Dummy > Y(I) And Dummy < Y(I + 1) Then
            XX = X(I) + ((X(I + 1) - X(I)) * (Dummy - Y(I)) / (Y(I + 1) - Y(I)))
            Label26.Caption = Label26.Caption & "   Mediana = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
    Next I
    Label26.Caption = Label26.Caption & "   Range = " & Format(Abs(X(Z1) - X(1)), "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   Range % = " & Format(Abs(100 * (X(Z1) - X(1)) / XX), "#####0.0") & " %   " & Chr(13) & Chr(10)
    Label26.Caption = Label26.Caption & "   Minimo = " & Format(X(1), "#####0.0####") & "   " & Chr(13)
    Dummy = Format(25, "##0.0")
    For I = 1 To Z1 - 1
        If Dummy = Y(I) Then
            XX = X(I)
            Label26.Caption = Label26.Caption & "   1" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
        If Dummy = Y(I + 1) Then
            XX = X(I + 1)
            Label26.Caption = Label26.Caption & "   1" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
        If Dummy > Y(I) And Dummy < Y(I + 1) Then
            XX = X(I) + ((X(I + 1) - X(I)) * (Dummy - Y(I)) / (Y(I + 1) - Y(I)))
            Label26.Caption = Label26.Caption & "   1" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
    Next I
    Dummy = Format(50, "##0.0")
    For I = 1 To Z1 - 1
        If Dummy = Y(I) Then
            XX = X(I)
            Label26.Caption = Label26.Caption & "   2" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
        If Dummy = Y(I + 1) Then
            XX = X(I + 1)
            Label26.Caption = Label26.Caption & "   2" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
        If Dummy > Y(I) And Dummy < Y(I + 1) Then
            XX = X(I) + ((X(I + 1) - X(I)) * (Dummy - Y(I)) / (Y(I + 1) - Y(I)))
            Label26.Caption = Label26.Caption & "   2" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
    Next I
    Dummy = Format(75, "##0.0")
    For I = 1 To Z1 - 1
        If Dummy = Y(I) Then
            XX = X(I)
            Label26.Caption = Label26.Caption & "   3" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
        If Dummy = Y(I + 1) Then
            XX = X(I + 1)
            Label26.Caption = Label26.Caption & "   3" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
        If Dummy > Y(I) And Dummy < Y(I + 1) Then
            XX = X(I) + ((X(I + 1) - X(I)) * (Dummy - Y(I)) / (Y(I + 1) - Y(I)))
            Label26.Caption = Label26.Caption & "   3" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
    Next I
    Label26.Caption = Label26.Caption & "   Massimo = " & Format(X(Z1), "#####0.0####") & "   " & Chr(13)
'    Label26.Caption = Label26.Caption & "   Range = " & Abs(X(Z1) - X(1)) & "   " & Chr(13)
'    Label26.Caption = Label26.Caption & Chr(13)
    Label26.Top = Picture1.Top + (Picture1.Height / 2) - (Label26.Height / 2)
    Screen.MousePointer = 1
    Exit Sub

GestioneErroriGraficaKrouwer:
    FlagGrafico = False
    FlagStatistica = False
    Screen.MousePointer = 1
    MsgBox "E' presente un errore logico" + Chr(10) + "nei dati da rappresentare.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
    Exit Sub

End Sub

Private Sub DistribuzioneCumulativStandard_Click()
    On Error GoTo GestioneErroriGraficaStandard
    Dim I, XX, YIN, YSU, FX, FY, CX, CY
    Dim X1, X2, X3, X4, X5
    Dim Y1, Y2, Y3, Y4, Y5
    Dim Dummy As Double

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
    FormCumulativa.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
            Label26.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
            Label26.BackColor = &H80000005
        End If
    Picture1.Cls
    FormCumulativa.ScaleWidth = 800
    FormCumulativa.ScaleHeight = 480
    FormCumulativa.Refresh

    YIN = YMI: YSU = YMA
    
    Picture1.Left = 100
    Picture1.Top = 80
    Picture1.Height = 340
    Picture1.Width = 460
    Picture1.ScaleHeight = 340
    Picture1.ScaleWidth = 460

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
    
    FX = 459 / (XMA - XMI): FY = 339 / (YSU - YIN)

        CX = (X(1) - XMI) * FX
        CY = 339 - (Y(1) - YIN) * FY
        Picture1.CurrentX = CX
        Picture1.CurrentY = CY
    For I = 1 To Z1
        CX = (X(I) - XMI) * FX ' : If CX < 0 Or CX > 459 Then GoTo 10565
        CY = 339 - (Y(I) - YIN) * FY ' : If CY < 0 Or CY > 339 Then GoTo 10565
            If FlagDatiAColori = True Then
                Picture1.Line -(CX, CY), QBColor(12)
            Else
                Picture1.Line -(CX, CY), QBColor(0)
            End If
' 10565 :
    Next I
    
    Y5 = YSU: Y4 = YSU - (YSU - YIN) / 4: Y3 = YIN + (YSU - YIN) / 2: Y2 = YIN + (YSU - YIN) / 4: Y1 = YIN
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
    Flag1 = True
    Flag2 = False
    Label13.Caption = Titolo1
    Label14.Caption = InAscisse
    Label15.Caption = InOrdinate1
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
    Picture1.Visible = True
    
    Label26.Visible = True
    Label26.Left = 574
    Label26.Top = Picture1.Top + (Picture1.Height / 3)
    Label26.Caption = "" & Chr(13)
    Label26.Caption = Label26.Caption & "   Dati = " & Z1 & "   " & Chr(13) & Chr(10)
    Dummy = Format(50, "##0.0")
    For I = 1 To Z1 - 1
        If Dummy = Y(I) Then
            XX = X(I)
            Label26.Caption = Label26.Caption & "   Mediana = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
        If Dummy = Y(I + 1) Then
            XX = X(I + 1)
            Label26.Caption = Label26.Caption & "   Mediana = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
        If Dummy > Y(I) And Dummy < Y(I + 1) Then
            XX = X(I) + ((X(I + 1) - X(I)) * (Dummy - Y(I)) / (Y(I + 1) - Y(I)))
            Label26.Caption = Label26.Caption & "   Mediana = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
    Next I
    Label26.Caption = Label26.Caption & "   Range = " & Format(Abs(X(Z1) - X(1)), "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   Range % = " & Format(Abs(100 * (X(Z1) - X(1)) / XX), "#####0.0") & " %   " & Chr(13) & Chr(10)
    Label26.Caption = Label26.Caption & "   Minimo = " & Format(X(1), "#####0.0####") & "   " & Chr(13)
    Dummy = Format(25, "##0.0")
    For I = 1 To Z1 - 1
        If Dummy = Y(I) Then
            XX = X(I)
            Label26.Caption = Label26.Caption & "   1" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
        If Dummy = Y(I + 1) Then
            XX = X(I + 1)
            Label26.Caption = Label26.Caption & "   1" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
        If Dummy > Y(I) And Dummy < Y(I + 1) Then
            XX = X(I) + ((X(I + 1) - X(I)) * (Dummy - Y(I)) / (Y(I + 1) - Y(I)))
            Label26.Caption = Label26.Caption & "   1" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
    Next I
    Dummy = Format(50, "##0.0")
    For I = 1 To Z1 - 1
        If Dummy = Y(I) Then
            XX = X(I)
            Label26.Caption = Label26.Caption & "   2" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
        If Dummy = Y(I + 1) Then
            XX = X(I + 1)
            Label26.Caption = Label26.Caption & "   2" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
        If Dummy > Y(I) And Dummy < Y(I + 1) Then
            XX = X(I) + ((X(I + 1) - X(I)) * (Dummy - Y(I)) / (Y(I + 1) - Y(I)))
            Label26.Caption = Label26.Caption & "   2" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
    Next I
    Dummy = Format(75, "##0.0")
    For I = 1 To Z1 - 1
        If Dummy = Y(I) Then
            XX = X(I)
            Label26.Caption = Label26.Caption & "   3" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
        If Dummy = Y(I + 1) Then
            XX = X(I + 1)
            Label26.Caption = Label26.Caption & "   3" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
        If Dummy > Y(I) And Dummy < Y(I + 1) Then
            XX = X(I) + ((X(I + 1) - X(I)) * (Dummy - Y(I)) / (Y(I + 1) - Y(I)))
            Label26.Caption = Label26.Caption & "   3" & Chr(176) & " quartile = " & Format(XX, "#####0.0####") & "   " & Chr(13)
            Exit For
        End If
    Next I
    Label26.Caption = Label26.Caption & "   Massimo = " & Format(X(Z1), "#####0.0####") & "   " & Chr(13)
    Label26.Top = Picture1.Top + (Picture1.Height / 2) - (Label26.Height / 2)
    Screen.MousePointer = 1
    Exit Sub

GestioneErroriGraficaStandard:
    FlagGrafico = False
    FlagStatistica = False
    Screen.MousePointer = 1
    MsgBox "E' presente un errore logico" + Chr(10) + "nei dati da rappresentare.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
    Exit Sub

End Sub

Private Sub FileEsci_Click()
    ReDim X(0)
    ReDim Y(0)
    FormPrincipale.Show
    Unload FormCumulativa
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
End Sub

Private Sub FileStampa_Click()
    On Error GoTo ErroreFileStampaCumu

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

ErroreFileStampaCumu:
    Exit Sub

End Sub

Private Sub Form_Load()
    On Error GoTo GestioneErroriStatElemNP
    ReDim X(1 To NUM_RIGHE - 1) As Double
    ReDim Y(1 To NUM_RIGHE - 1)
    Dim D, I, J, L, M1, M2, ValoreInizialeD
    Dim XX

    Screen.MousePointer = 11
    AnnullatoCaricamentoForm = False
    FlagGrafico = False
    FlagStatistica = False
    FlagSfondoGrigio = True
    FlagDatiAColori = True
    Picture2.Visible = False
    
' centra il form
    FormCumulativa.Top = Screen.Height / 2 - Height / 2
    FormCumulativa.Left = Screen.Width / 2 - Width / 2
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
    Titolo1 = "Distribuzione cumulativa standard"
    Titolo2 = "Distribuzione cumulativa ripiegata"
    InOrdinate1 = "Percentili"
    InOrdinate2 = "Percentili ripiegati secondo Krouwer"
' carica i dati nel vettore X()
    FormGauge.Show
    FormGauge.Caption = "Statistiche elementari non parametriche"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    FineDati = False
    Z1 = 0
    
        FormGauge.Label6.Caption = "[1/3]"
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
    
    ValoreInizialeD = 0

' inizio ordinamento dati
pc330:
    FormGauge.Label6.Caption = "[2/3]"
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
                
                FormGauge.Label1.Caption = "100%"
                FormGauge!Gauge1.Value = 100
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            FormGauge.Label6.Caption = "[3/3]"
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

            
            FormGauge.Label1.Caption = "100%"
            FormGauge!Gauge1.Value = 100
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            Unload FormGauge
            
            XMI = X(1): XMA = X(1)
            For I = 2 To Z1
                If X(I) < XMI Then XMI = X(I)
                If X(I) > XMA Then XMA = X(I)
            Next
            YMI = 0
            YMA = 100
            XMI_Default = XMI
            XMA_Default = XMA
            
            If XMI = XMA Then GoTo GestioneErroriStatElemNP ' meglio così

            YMI_Default = YMI
            YMA_Default = YMA
            Text4.Visible = True
            Shape1.Visible = False
            Frame1.Visible = False
            Picture1.Visible = False
            Picture2.Visible = False
            Label26.Visible = False
            FormPrincipale!Grid1.HighLight = True
            RipristinaSchermo
            FormCumulativa.Show
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

GestioneErroriStatElemNP:
        Unload FormGauge
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico dei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Statistiche elementari non eseguibili"
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
        Exit Sub

End Sub

Private Sub Form_Resize()
    
    Text4.Width = ScaleWidth
    Text4.Height = ScaleHeight

    If Flag1 = True Then
        DistribuzioneCumulativStandard_Click
    End If
    If Flag2 = True Then
        DistribuzioneCumulativaRipiegata_Click
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

    If Flag1 = True Then
        DistribuzioneCumulativStandard_Click
    End If
    If Flag2 = True Then
        DistribuzioneCumulativaRipiegata_Click
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
        
    If Flag1 = True Then
        Text1.Text = Titolo1
        Text3.Text = InOrdinate1
    Else
        Text1.Text = Titolo2
        Text3.Text = InOrdinate2
    End If
    Text2.Text = InAscisse
    MaskedEdit1.Text = Format(XMI, "+#####0.0####;-#####0.0####")
    MaskedEdit2.Text = Format(XMA, "+#####0.0####;-#####0.0####")
    MaskedEdit3.Text = Format(YMI, "+#####0.0####;-#####0.0####")
    MaskedEdit4.Text = Format(YMA, "+#####0.0####;-#####0.0####")
    
    MaskedEdit3.Enabled = False
    MaskedEdit4.Enabled = False

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

    If Flag1 = True Then
        DistribuzioneCumulativStandard_Click
    End If
    If Flag2 = True Then
        DistribuzioneCumulativaRipiegata_Click
    End If

    End If

End Sub

Private Sub MenuFile_Click()

    FileStampa.Enabled = False
        If FlagGrafico = True Then
            If FormCumulativa.Width < Screen.Width Then
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
    Dim I, XX
    Dim Dummy As Double

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
    FormCumulativa.Refresh
    

        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Statistiche elementari non parametriche (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File      : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella   : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & "Variabile : " & InAscisse & Chr(13) & Chr(10)
        
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Numero dei dati             = " & Z1 & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Dummy = Format(50, "##0.0")
        For I = 1 To Z1 - 1
            If Dummy = Y(I) Then
                XX = X(I)
                Text4.Text = Text4.Text & "Mediana                     = " & Format(XX, "#####0.0####") & "   " & Chr(13) & Chr(10)
                Exit For
            End If
            If Dummy = Y(I + 1) Then
                XX = X(I + 1)
                Text4.Text = Text4.Text & "Mediana                     = " & Format(XX, "#####0.0####") & "   " & Chr(13) & Chr(10)
                Exit For
            End If
            If Dummy > Y(I) And Dummy < Y(I + 1) Then
                XX = X(I) + ((X(I + 1) - X(I)) * (Dummy - Y(I)) / (Y(I + 1) - Y(I)))
                Text4.Text = Text4.Text & "Mediana                     = " & Format(XX, "#####0.0####") & "   " & Chr(13) & Chr(10)
                Exit For
            End If
        Next I
        Text4.Text = Text4.Text & "Range                       = " & Format(Abs(X(Z1) - X(1)), "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Range in % della mediana    = " & Format(Abs(100 * (X(Z1) - X(1)) / XX), "#####0.0") & " %" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Valore minimo               = " & Format(X(1), "#####0.0####") & Chr(13) & Chr(10)
        Dummy = Format(25, "##0.0")
        For I = 1 To Z1 - 1
            If Dummy = Y(I) Then
                XX = X(I)
                Text4.Text = Text4.Text & "Primo quartile              = " & Format(XX, "#####0.0####") & "   " & Chr(13) & Chr(10)
                Exit For
            End If
            If Dummy = Y(I + 1) Then
                XX = X(I + 1)
                Text4.Text = Text4.Text & "Primo quartile              = " & Format(XX, "#####0.0####") & "   " & Chr(13) & Chr(10)
                Exit For
            End If
            If Dummy > Y(I) And Dummy < Y(I + 1) Then
                XX = X(I) + ((X(I + 1) - X(I)) * (Dummy - Y(I)) / (Y(I + 1) - Y(I)))
                Text4.Text = Text4.Text & "Primo quartile              = " & Format(XX, "#####0.0####") & "   " & Chr(13) & Chr(10)
                Exit For
            End If
        Next I
        Dummy = Format(50, "##0.0")
        For I = 1 To Z1 - 1
            If Dummy = Y(I) Then
                XX = X(I)
                Text4.Text = Text4.Text & "Secondo quartile            = " & Format(XX, "#####0.0####") & "   " & Chr(13) & Chr(10)
                Exit For
            End If
            If Dummy = Y(I + 1) Then
                XX = X(I + 1)
                Text4.Text = Text4.Text & "Secondo quartile            = " & Format(XX, "#####0.0####") & "   " & Chr(13) & Chr(10)
                Exit For
            End If
            If Dummy > Y(I) And Dummy < Y(I + 1) Then
                XX = X(I) + ((X(I + 1) - X(I)) * (Dummy - Y(I)) / (Y(I + 1) - Y(I)))
                Text4.Text = Text4.Text & "Secondo quartile            = " & Format(XX, "#####0.0####") & "   " & Chr(13) & Chr(10)
                Exit For
            End If
        Next I
        Dummy = Format(75, "##0.0")
        For I = 1 To Z1 - 1
            If Dummy = Y(I) Then
                XX = X(I)
                Text4.Text = Text4.Text & "Terzo quartile              = " & Format(XX, "#####0.0####") & "   " & Chr(13) & Chr(10)
                Exit For
            End If
            If Dummy = Y(I + 1) Then
                XX = X(I + 1)
                Text4.Text = Text4.Text & "Terzo quartile              = " & Format(XX, "#####0.0####") & "   " & Chr(13) & Chr(10)
                Exit For
            End If
            If Dummy > Y(I) And Dummy < Y(I + 1) Then
                XX = X(I) + ((X(I + 1) - X(I)) * (Dummy - Y(I)) / (Y(I + 1) - Y(I)))
                Text4.Text = Text4.Text & "Terzo quartile              = " & Format(XX, "#####0.0####") & "   " & Chr(13) & Chr(10)
                Exit For
            End If
        Next I
        Text4.Text = Text4.Text & "Valore massimo              = " & Format(X(Z1), "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Statistiche elementari non parametriche (*fine*)" & Chr(13) & Chr(10)
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        FlagGrafico = False
        Flag1 = False
        Flag2 = False
        Screen.MousePointer = 1

End Sub

Private Sub StatisticaTabellaDeiPercentili_Click()
    Dim I, XX
    Dim K As Double
    Dim Dummy As Double
    Dim Stringa As String
    Dim Centile As Double
    Dim Zeta As Double
    Dim UnoMenoZeta As Double
    Dim Percentile As Double
    Dim Valore As Double
    Dim LimiteInferiore As Double
    Dim LimiteSuperiore As Double

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
    FormCumulativa.Refresh
    
    FormGauge.Show
    FormGauge.Caption = "Calcolo dei percentili"
    FormGauge.Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Tabella dei percentili non parametrici (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File      : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella   : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & "Variabile : " & InAscisse & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
                                                                    
        Text4.Text = Text4.Text & "                          Limiti di confidenza al 90%" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                          ---------------------------" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Percentile     Valore       inferiore     superiore" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "----------  ------------  ------------  -------------" & Chr(13) & Chr(10)
        
        Centile = Format(2.5, "##0.0")
        Zeta = 0.025
        UnoMenoZeta = 1 - Zeta
        GoSub CalcolaPercentili
        Stringa = Format(Centile, "##0.0")
        Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@") & "  "
        If Valore <> 0 Then
            Stringa = Format(Valore, "#####0.0####")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@") & "  "
        Else
            Text4.Text = Text4.Text & Format("###", "@@@@@@@@@@@@") & "  "
        End If
        If LimiteInferiore <> 0 And Valore <> 0 Then
            Stringa = Format(LimiteInferiore, "#####0.0####")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@") & "  "
        Else
            Text4.Text = Text4.Text & Format("###", "@@@@@@@@@@@@") & "  "
        End If
        If LimiteSuperiore <> 0 And Valore <> 0 Then
            Stringa = Format(LimiteSuperiore, "#####0.0####")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@@")
        Else
            Text4.Text = Text4.Text & Format("###", "@@@@@@@@@@@@@") & "  "
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        
        For K = 5# To 95# Step 5#
            Centile = Format(K, "##0.0")
            Zeta = K / 100
            UnoMenoZeta = 1 - Zeta
            Valore = 0
            LimiteInferiore = 0
            LimiteSuperiore = 0
            GoSub CalcolaPercentili
            Stringa = Format(Centile, "##0.0")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@") & "  "
            If Valore <> 0 Then
                Stringa = Format(Valore, "#####0.0####")
                Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@") & "  "
            Else
                Text4.Text = Text4.Text & Format("###", "@@@@@@@@@@@@") & "  "
            End If
            If LimiteInferiore <> 0 And Valore <> 0 Then
                Stringa = Format(LimiteInferiore, "#####0.0####")
                Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@") & "  "
            Else
                Text4.Text = Text4.Text & Format("###", "@@@@@@@@@@@@") & "  "
            End If
            If LimiteSuperiore <> 0 And Valore <> 0 Then
                Stringa = Format(LimiteSuperiore, "#####0.0####")
                Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@@")
            Else
                Text4.Text = Text4.Text & Format("###", "@@@@@@@@@@@@@") & "  "
            End If
            Text4.Text = Text4.Text & Chr(13) & Chr(10)
            FormGauge.Label1.Caption = Int(100 * K / 95) & "%"
            FormGauge!Gauge1.Value = Int(100 * K / 95)
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
        Next K
        
        Centile = Format(97.5, "##0.0")
        Zeta = 0.975
        UnoMenoZeta = 1 - Zeta
        Valore = 0
        LimiteInferiore = 0
        LimiteSuperiore = 0
        GoSub CalcolaPercentili
        Stringa = Format(Centile, "##0.0")
        Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@") & "  "
        If Valore <> 0 Then
            Stringa = Format(Valore, "#####0.0####")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@") & "  "
        Else
            Text4.Text = Text4.Text & Format("###", "@@@@@@@@@@@@") & "  "
        End If
        If LimiteInferiore <> 0 And Valore <> 0 Then
            Stringa = Format(LimiteInferiore, "#####0.0####")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@") & "  "
        Else
            Text4.Text = Text4.Text & Format("###", "@@@@@@@@@@@@") & "  "
        End If
        If LimiteSuperiore <> 0 And Valore <> 0 Then
            Stringa = Format(LimiteSuperiore, "#####0.0####")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@@")
        Else
            Text4.Text = Text4.Text & Format("###", "@@@@@@@@@@@@@") & "  "
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        
        Text4.Text = Text4.Text & "----------  ------------  ------------  -------------" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Tabella dei percentili non parametrici (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        FlagGrafico = False
        Flag1 = False
        Flag2 = False
        Screen.MousePointer = 1

Exit Sub
        
CalcolaPercentili:
        
        Percentile = Centile
        For I = 1 To Z1 - 1
            If Percentile = Y(I) Then
                Valore = Format(X(I), "#####0.0####")
                Exit For
            End If
            If Percentile = Y(I + 1) Then
                Valore = Format(X(I + 1), "#####0.0####")
                Exit For
            End If
            If Percentile > Y(I) And Percentile < Y(I + 1) Then
                Valore = Format(X(I) + ((X(I + 1) - X(I)) * (Percentile - Y(I)) / (Y(I + 1) - Y(I))), "#####0.0####")
                Exit For
            End If
        Next I
        
        Dummy = Centile * (Z1 + 1) / 100
        Dummy = Dummy - (1.645 * Sqr(Z1 * Zeta * UnoMenoZeta))
        Percentile = Format(Dummy * 100 / (Z1 + 1), "##0.0")
        For I = 1 To Z1 - 1
            If Percentile = Y(I) Then
                LimiteInferiore = Format(X(I), "#####0.0####")
                Exit For
            End If
            If Percentile = Y(I + 1) Then
                LimiteInferiore = Format(X(I + 1), "#####0.0####")
                Exit For
            End If
            If Percentile > Y(I) And Percentile < Y(I + 1) Then
                LimiteInferiore = Format(X(I) + ((X(I + 1) - X(I)) * (Percentile - Y(I)) / (Y(I + 1) - Y(I))), "#####0.0####")
                Exit For
            End If
        Next I
        
        Dummy = Centile * (Z1 + 1) / 100
        Dummy = Dummy + (1.645 * Sqr(Z1 * Zeta * UnoMenoZeta))
        Percentile = Format(Dummy * 100 / (Z1 + 1), "##0.0")
        For I = 1 To Z1 - 1
            If Percentile = Y(I) Then
                LimiteSuperiore = Format(X(I), "#####0.0####")
                Exit For
            End If
            If Percentile = Y(I + 1) Then
                LimiteSuperiore = Format(X(I + 1), "#####0.0####")
                Exit For
            End If
            If Percentile > Y(I) And Percentile < Y(I + 1) Then
                LimiteSuperiore = Format(X(I) + ((X(I + 1) - X(I)) * (Percentile - Y(I)) / (Y(I + 1) - Y(I))), "#####0.0####")
                Exit For
            End If
        Next I

Return

End Sub

