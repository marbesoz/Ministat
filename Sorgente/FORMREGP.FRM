VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormPolinomiale 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Regressione polinomiale"
   ClientHeight    =   6360
   ClientLeft      =   1095
   ClientTop       =   1770
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
   Icon            =   "FORMREGP.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6360
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
      Height          =   1092
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   732
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
         TabIndex        =   41
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
         Caption         =   "X corrispondente  >>"
         Height          =   375
         Left            =   2880
         TabIndex        =   24
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Y corrispondente  >>"
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
         Caption         =   "Y osservata:"
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
         Caption         =   "X osservata:"
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
      TabIndex        =   40
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
      Height          =   252
      Left            =   120
      TabIndex        =   39
      Top             =   1100
      Width           =   972
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
      Begin VB.Menu GraficaX 
         Caption         =   "Regressione &polinomiale di primo grado"
         Begin VB.Menu GraficaXcon 
            Caption         =   "&Con i limiti di confidenza al 90%"
         End
         Begin VB.Menu GraficaXsenza 
            Caption         =   "&Senza i limiti di confidenza al 90%"
         End
      End
      Begin VB.Menu GraficaY 
         Caption         =   "Regressione p&olinomiale di secondo grado"
      End
      Begin VB.Menu GraficaCPS 
         Caption         =   "Regressione po&linomiale di terzo grado"
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
            Caption         =   "&Dati a colori"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu GraficaScaleLegende 
         Caption         =   "C&onfigurazione scale assi e legende"
      End
   End
   Begin VB.Menu MenuStatistica 
      Caption         =   "&Statistica"
      Begin VB.Menu StatisticaX 
         Caption         =   "Regressione &polinomiale di primo grado"
      End
      Begin VB.Menu StatisticaY 
         Caption         =   "Regressione p&olinomiale di secondo grado"
      End
      Begin VB.Menu StatisticaCPS 
         Caption         =   "Regressione po&linomiale di terzo grado"
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
Attribute VB_Name = "FormPolinomiale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim X() As Double
    Dim Y() As Double
    Dim AA(), SUP(), INF()
    Dim DEVIATAT, MEDIAX, DEVIANZAX
    Dim Riga, ColonnaX, ColonnaY, FineDati
    Dim Z1, Z2
    Dim SLB, INA, COR, DER, DES, DEI
    Dim APO, BPO, CPO, DPO As Double
    Dim XMI, XMA, YMI, YMA
    Dim XMI_Default, XMA_Default
    Dim YMI_Default, YMA_Default
    Dim Titolo1, Titolo2, Titolo3
    Dim InAscisse, InOrdinate
    Dim Flag1, Flag2, Flag3, Flag1con, Flag1senza
    Dim FlagStatistica, FlagGrafico
    Dim FlagSfondoGrigio
    Dim FlagDatiAColori

Private Sub AiutoCalcolatrice_Click()
    Dim Calcolatrice
    On Error GoTo ErroreCalcolatriceSette:
    
    Calcolatrice = Shell("Calc.exe", 1)
        
Exit Sub

ErroreCalcolatriceSette:
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
    On Error GoTo Polinomiale
            
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
            End If
            If Flag2 = True Then
                Titolo2 = Text1.Text
            End If
            If Flag3 = True Then
                Titolo3 = Text1.Text
            End If
            InAscisse = Text2.Text
            InOrdinate = Text3.Text
            XMI = MaskedEdit1.Text
            XMA = MaskedEdit2.Text
            YMI = MaskedEdit3.Text
            YMA = MaskedEdit4.Text
                If Flag1 = True Then
                    Label13.Caption = Titolo1
                End If
                If Flag2 = True Then
                    Label13.Caption = Titolo2
                End If
                If Flag3 = True Then
                    Label13.Caption = Titolo3
                End If
            Label14.Caption = InAscisse
            Label15.Caption = InOrdinate

'        If FlagGrafico = True Then

        If Flag1 = True Then
            If Flag1con = True Then
                GraficaXcon_Click
            End If
            If Flag1senza = True Then
                GraficaXsenza_Click
            End If
        End If
        
        If Flag2 = True Then
            GraficaY_Click
        End If
        If Flag3 = True Then
            GraficaCPS_Click
        End If

'        End If
        
        Frame1.Visible = False
        Shape1.Visible = False
        Command5.Enabled = True

        End If

Polinomiale:
        MaskedEdit1.Text = Format(XMI_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit2.Text = Format(XMA_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit3.Text = Format(YMI_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit4.Text = Format(YMA_Default, "+#####0.0####;-#####0.0####")

End Sub

Private Sub Command2_Click()
    On Error GoTo GestioneErroriCommand2
    Dim Y

    Screen.MousePointer = 11
    
    MaskedEdit5.Text = FormattaStringa(MaskedEdit5.Text)
    If Flag1 = True Then
        Y = INA - (-SLB * MaskedEdit5.Text)
        Label10.Caption = Format(Y, "#####0.0####")
    End If

    If Flag2 = True Then
        Y = APO - (-BPO * MaskedEdit5.Text) - (-CPO * MaskedEdit5.Text * MaskedEdit5.Text)
        Label10.Caption = Format(Y, "#####0.0####")
    End If
    
    If Flag3 = True Then
        Y = APO - (-BPO * MaskedEdit5.Text) - (-CPO * MaskedEdit5.Text * MaskedEdit5.Text) - (-DPO * MaskedEdit5.Text * MaskedEdit5.Text * MaskedEdit5.Text)
        Label10.Caption = Format(Y, "#####0.0####")
    End If
    
    Screen.MousePointer = 1

Exit Sub

GestioneErroriCommand2:
    Label10.Caption = ""
    Screen.MousePointer = 1
    Exit Sub

End Sub

Private Sub Command3_Click()
    On Error GoTo GestioneErroriCommand3
    Dim X, A1, B1, c1 As Double
    ReDim D(10) As Double

    Screen.MousePointer = 11
    
    MaskedEdit6.Text = FormattaStringa(MaskedEdit6.Text)
    If Flag1 = True Then
        X = (MaskedEdit6.Text - INA) / SLB
        Label12.Caption = Format(X, "#####0.0####")
    End If

    If Flag2 = True Then
        A1 = Sgn(APO + BPO * XMI + CPO * XMI * XMI - MaskedEdit6.Text)
        B1 = Sgn(APO + BPO * XMA + CPO * XMA * XMA - MaskedEdit6.Text)
        If A1 * B1 = 0 Then GoTo s5240
        D(2 + A1) = XMI
        D(2 - A1) = XMA
s5180:
        X = (D(1) + D(3)) / 2
        c1 = Sgn(APO + BPO * X + CPO * X * X - MaskedEdit6.Text)
        If c1 = 0 Then GoTo s5280
        D(2 + c1) = X
        If Abs(D(1) - D(3)) / Abs(D(1) + Abs(D(3))) < 0.000001 Then GoTo s5280
        GoTo s5180
s5240:
        If A1 = 0 Then GoTo s5270
        X = XMA
        GoTo s5280
s5270:
        X = XMI
s5280:
        Label12.Caption = Format(X, "#####0.0####")
    End If
    
    If Flag3 = True Then
        A1 = Sgn(APO + BPO * XMI + CPO * XMI * XMI + DPO * XMI * XMI * XMI - MaskedEdit6.Text)
        B1 = Sgn(APO + BPO * XMA + CPO * XMA * XMA + DPO * XMA * XMA * XMA - MaskedEdit6.Text)
        If A1 * B1 = 0 Then GoTo t5240
        D(2 + A1) = XMI
        D(2 - A1) = XMA
t5180:
        X = (D(1) + D(3)) / 2
        c1 = Sgn(APO + BPO * X + CPO * X * X + DPO * X * X * X - MaskedEdit6.Text)
        If c1 = 0 Then GoTo t5280
        D(2 + c1) = X
        If Abs(D(1) - D(3)) / Abs(D(1) + Abs(D(3))) < 0.000001 Then GoTo t5280
        GoTo t5180
t5240:
        If A1 = 0 Then GoTo t5270
        X = XMA
        GoTo t5280
t5270:
        X = XMI
t5280:
        Label12.Caption = Format(X, "#####0.0####")
    End If

    Screen.MousePointer = 1

Exit Sub

GestioneErroriCommand3:
    Label12.Caption = ""
    Screen.MousePointer = 1
    Exit Sub

End Sub

Private Sub Command4_Click()
    
'    If FlagGrafico = True Then
        If Flag1 = True Then
            If Flag1con = True Then
                GraficaXcon_Click
            End If
            If Flag1senza = True Then
                GraficaXsenza_Click
            End If
        End If
        If Flag2 = True Then
            GraficaY_Click
        End If
        If Flag3 = True Then
            GraficaCPS_Click
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

Private Sub FileEsci_Click()
    ReDim X(0)
    ReDim Y(0)
    ReDim AA(0), SUP(0), INF(0)
    FormPrincipale.Show
    Unload FormPolinomiale
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
End Sub

Private Sub FileStampa_Click()
    On Error GoTo ErroreFileStampaRegp

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

ErroreFileStampaRegp:
    Exit Sub

End Sub

Private Sub Form_Load()
    ReDim X(1 To NUM_RIGHE - 1) As Double
    ReDim Y(1 To NUM_RIGHE - 1) As Double
    Dim I

    Screen.MousePointer = 11
    AnnullatoCaricamentoForm = False
    Flag1 = False
    Flag1con = False
    Flag1senza = False
    Flag2 = False
    Flag3 = False
    FlagStatistica = False
    FlagGrafico = False
    FlagSfondoGrigio = True
    FlagDatiAColori = True
    Picture2.Visible = False
    
' centra il form
    FormPolinomiale.Top = Screen.Height / 2 - Height / 2
    FormPolinomiale.Left = Screen.Width / 2 - Width / 2
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
    ColonnaY = FormPrincipale!Grid1.SelEndCol
' nome delle variabili
    FormPrincipale!Grid1.Row = 0
    FormPrincipale!Grid1.Col = ColonnaX
    InAscisse = FormPrincipale!Grid1.Text
    FormPrincipale!Grid1.Col = ColonnaY
    InOrdinate = FormPrincipale!Grid1.Text
    Titolo1 = "Regressione polinomiale di primo grado"
    Titolo2 = "Regressione polinomiale di secondo grado"
    Titolo3 = "Regressione polinomiale di terzo grado"
' carica i dati nei vettori X() e Y()
    FormGauge.Show
    FormGauge.Caption = "Regressione polinomiale"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    FineDati = False
    Z1 = 0
    Z2 = 0
    
        FormGauge!Label6.Caption = "[1/1]"
        For Riga = 1 To NUM_RIGHE - 1
            If 100 * Riga / (NUM_RIGHE - 1) = Int(100 * Riga / (NUM_RIGHE - 1)) Then
                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If
        FormPrincipale!Grid1.Row = Riga
        
        If FineDati = False Then
            FormPrincipale!Grid1.Col = ColonnaX
            If FormPrincipale!Grid1.Text <> "" Then
                Z1 = Z1 + 1
                X(Riga) = CDbl(-(-FormPrincipale!Grid1.Text))
            Else
                FineDati = True
            End If
        FormPrincipale!Grid1.Col = ColonnaY
            If FormPrincipale!Grid1.Text <> "" Then
                Z2 = Z2 + 1
                Y(Riga) = CDbl(-(-FormPrincipale!Grid1.Text))
            Else
                FineDati = True
            End If
        End If
        Next Riga
' controlli finali sui dati
' per la regressione devono esserci più di due punti. ..
    If Z1 > 3 Then
' ...e i dati devono essere in numero uguale
        If Z1 = Z2 Then
            XMI = X(1): XMA = X(1): YMI = Y(1): YMA = Y(1)
            For I = 2 To Z1
                If X(I) < XMI Then XMI = X(I)
                If X(I) > XMA Then XMA = X(I)
                If Y(I) < YMI Then YMI = Y(I)
                If Y(I) > YMA Then YMA = Y(I)
            Next
            XMI_Default = XMI
            XMA_Default = XMA
            YMI_Default = YMI
            YMA_Default = YMA
            Unload FormGauge
            Text4.Visible = True
            Shape1.Visible = False
            Frame1.Visible = False
            Picture1.Visible = False
            Picture2.Visible = False
            Label26.Visible = False
            FormPrincipale!Grid1.HighLight = True
            RipristinaSchermo
            FormPolinomiale.Show
            ' FormPrincipale.Hide
            Screen.MousePointer = 1
            Exit Sub
        Else
            Unload FormGauge
            AnnullatoCaricamentoForm = True
            Screen.MousePointer = 1
            MsgBox "Per la regressione polinomiale" + Chr(10) + "le due variabili devono contenere" + Chr(10) + "lo stesso numero di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Regressione polinomiale non eseguibile"
            FormPrincipale!Grid1.HighLight = True
            RipristinaSchermo
        End If
    Else
        Unload FormGauge
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        MsgBox "Per la regressione polinomiale" + Chr(10) + "sono necessari almeno quattro dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Regressione polinomiale non eseguibile"
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
    End If

End Sub

Private Sub Form_Resize()
    
    Text4.Width = ScaleWidth
    Text4.Height = ScaleHeight

    If Flag1 = True Then
        If Flag1con = True Then
            GraficaXcon_Click
        End If
        If Flag1senza = True Then
            GraficaXsenza_Click
        End If
    End If
    If Flag2 = True Then
        GraficaY_Click
    End If
    If Flag3 = True Then
        GraficaCPS_Click
    End If

End Sub

Private Sub GraficaCPS_Click()
    On Error GoTo GestioneErroriGraficaCPSpoli
    ReDim AA(460), INF(460), SUP(460)
    Dim I, XX, YIN, YSU, FX, FY, CX, CY
    Dim X1, X2, X3, X4, X5
    Dim Y1, Y2, Y3, Y4, Y5

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
    FormPolinomiale.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
            Label26.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
            Label26.BackColor = &H80000005
        End If
    Picture1.Cls
    FormPolinomiale.ScaleWidth = 800
    FormPolinomiale.ScaleHeight = 480
    FormPolinomiale.Refresh

    For I = 1 To 460
        XX = XMI + (I - 1) * (XMA - XMI) / 460
        AA(I) = APO + BPO * XX + CPO * XX * XX + DPO * XX * XX * XX
    Next I

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
    
    FX = 459 / (XMA - XMI): FY = 339 / (YSU - YIN)
    
    For I = 1 To 460
'        If (339 - (AA(I) - YIN) * FY) >= 0 And (339 - (AA(I) - YIN) * FY) <= 339 Then
        If I >= (XMI_Default - XMI) * FX + 1 And I <= (XMA_Default - XMI) * FX + 1 Then
            Picture1.CurrentX = I
            Picture1.CurrentY = 339 - (AA(I) - YIN) * FY
            Exit For
        End If
    Next I

    For I = 1 To 460
'        If (339 - (AA(I) - YIN) * FY) >= 0 And (339 - (AA(I) - YIN) * FY) <= 339 Then Picture1.Line -(I - 1, 339 - (AA(I) - YIN) * FY)
'        Picture1.Line -(I - 1, 339 - (AA(I) - YIN) * FY)
        If I >= (XMI_Default - XMI) * FX + 1 And I <= (XMA_Default - XMI) * FX + 1 Then Picture1.Line -(I, 339 - (AA(I) - YIN) * FY)
    Next

    For I = 1 To Z1
        CX = (X(I) - XMI) * FX ' : If CX < 0 Or CX > 459 Then GoTo 10565
        CY = 339 - (Y(I) - YIN) * FY ' : If CY < 0 Or CY > 339 Then GoTo 10565
            If FlagDatiAColori = True Then
                Picture1.Line (CX - 4, CY)-(CX + 4, CY), QBColor(12)
                Picture1.Line (CX, CY - 4)-(CX, CY + 4), QBColor(12)
            Else
                Picture1.Line (CX - 4, CY)-(CX + 4, CY), QBColor(0)
                Picture1.Line (CX, CY - 4)-(CX, CY + 4), QBColor(0)
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

    FlagStatistica = False
    FlagGrafico = True
    Label13.Caption = Titolo3
    Label14.Caption = InAscisse
    Label15.Caption = InOrdinate
    Label16.Caption = Format(X1, "#####0.0####;-#####0.0####")
    Label17.Caption = Format(X2, "#####0.0####;-#####0.0####")
    Label18.Caption = Format(X3, "#####0.0####;-#####0.0####")
    Label19.Caption = Format(X4, "#####0.0####;-#####0.0####")
    Label20.Caption = Format(X5, "#####0.0####;-#####0.0####")
    Label21.Caption = Format(Y1, "#####0.0####;-#####0.0####")
    Label22.Caption = Format(Y2, "#####0.0####;-#####0.0####")
    Label23.Caption = Format(Y3, "#####0.0####;-#####0.0####")
    Label24.Caption = Format(Y4, "#####0.0####;-#####0.0####")
    Label25.Caption = Format(Y5, "#####0.0####;-#####0.0####")
    
    Label26.Visible = True
    Label26.Left = 574
    Label26.Top = Picture1.Top + (Picture1.Height / 3)
    Label26.Caption = "" & Chr(13)
    Label26.Caption = Label26.Caption & "   y = a + bx + cx" & Chr(178) & "+ dx" & Chr(179) & "   " & Chr(13) & Chr(10)
    Label26.Caption = Label26.Caption & "   n = " & Z1 & "   " & Chr(13) & Chr(10)
    Label26.Caption = Label26.Caption & "   a = " & Format(APO, "#######0.0#########") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   b = " & Format(BPO, "#######0.0#########") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   c = " & Format(CPO, "#######0.0#########") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   d = " & Format(DPO, "#######0.0#########") & "   "
    Label26.Caption = Label26.Caption & Chr(13)
    Label26.Top = Picture1.Top + (Picture1.Height / 2) - (Label26.Height / 2)
    Screen.MousePointer = 1
    Exit Sub

GestioneErroriGraficaCPSpoli:
    Screen.MousePointer = 1
    MsgBox "E' presente un errore logico" + Chr(10) + "nei dati da rappresentare.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
    Exit Sub

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
        If Flag1con = True Then
            GraficaXcon_Click
        End If
        If Flag1senza = True Then
            GraficaXsenza_Click
        End If
    End If
    If Flag2 = True Then
        GraficaY_Click
    End If
    If Flag3 = True Then
        GraficaCPS_Click
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
    End If
    If Flag2 = True Then
        Text1.Text = Titolo2
    End If
    If Flag3 = True Then
        Text1.Text = Titolo3
    End If
    Text2.Text = InAscisse
    Text3.Text = InOrdinate
    MaskedEdit1.Text = Format(XMI, "+#####0.0####;-#####0.0####")
    MaskedEdit2.Text = Format(XMA, "+#####0.0####;-#####0.0####")
    MaskedEdit3.Text = Format(YMI, "+#####0.0####;-#####0.0####")
    MaskedEdit4.Text = Format(YMA, "+#####0.0####;-#####0.0####")
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
        If Flag1con = True Then
            GraficaXcon_Click
        End If
        If Flag1senza = True Then
            GraficaXsenza_Click
        End If
    End If
    If Flag2 = True Then
        GraficaY_Click
    End If
    If Flag3 = True Then
        GraficaCPS_Click
    End If

    End If

End Sub

Private Sub GraficaXcon_Click()

    On Error GoTo GestioneErroriGraficaXcon
    ReDim AA(460), INF(460), SUP(460)
    Dim I, XX, YIN, YSU, FX, FY, CX, CY
    Dim X1, X2, X3, X4, X5
    Dim Y1, Y2, Y3, Y4, Y5
    Dim SCARTO

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
    FormPolinomiale.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
            Label26.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
            Label26.BackColor = &H80000005
        End If
    Picture1.Cls
    FormPolinomiale.ScaleWidth = 800
    FormPolinomiale.ScaleHeight = 480
    FormPolinomiale.Refresh

    For I = 1 To 460
        XX = XMI + (I - 1) * (XMA - XMI) / 460
        AA(I) = INA + SLB * XX
        SCARTO = DEVIATAT * DER * Sqr(1 + (1 / Z1) + ((AA(I) - MEDIAX) ^ 2 / DEVIANZAX))
        SUP(I) = AA(I) + SCARTO
        INF(I) = AA(I) - SCARTO
    Next I

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
    
    FX = 459 / (XMA - XMI): FY = 339 / (YSU - YIN)
    
    For I = 1 To 460
'        If (339 - (AA(I) - YIN) * FY) >= 0 And (339 - (AA(I) - YIN) * FY) <= 339 Then
        If I >= (XMI_Default - XMI) * FX + 1 And I <= (XMA_Default - XMI) * FX + 1 Then
            Picture1.CurrentX = I
            Picture1.CurrentY = 339 - (AA(I) - YIN) * FY
            Exit For
        End If
    Next I

    For I = 1 To 460
'        If (339 - (AA(I) - YIN) * FY) >= 0 And (339 - (AA(I) - YIN) * FY) <= 339 Then Picture1.Line -(I - 1, 339 - (AA(I) - YIN) * FY)
'        Picture1.Line -(I - 1, 339 - (AA(I) - YIN) * FY)
        If I >= (XMI_Default - XMI) * FX + 1 And I <= (XMA_Default - XMI) * FX + 1 Then Picture1.Line -(I, 339 - (AA(I) - YIN) * FY)
    Next

'    For I = 1 To 460
'        If (339 - (AA(I) - YIN) * FY) >= 0 And (339 - (AA(I) - YIN) * FY) <= 339 Then Picture1.PSet (I - 1, 339 - (AA(I) - YIN) * FY)
'    Next
  
    For I = 4 To 460 Step 4
'        If (339 - (SUP(I) - YIN) * FY) > 0 And (339 - (SUP(I) - YIN) * FY) < 339 Then Picture1.PSet (I - 1, 339 - (SUP(I) - YIN) * FY)
'        If (339 - (INF(I) - YIN) * FY) > 0 And (339 - (INF(I) - YIN) * FY) < 339 Then Picture1.PSet (I - 1, 339 - (INF(I) - YIN) * FY)
        If I >= (XMI_Default - XMI) * FX + 1 And I <= (XMA_Default - XMI) * FX + 1 Then Picture1.PSet (I, 339 - (SUP(I) - YIN) * FY)
        If I >= (XMI_Default - XMI) * FX + 1 And I <= (XMA_Default - XMI) * FX + 1 Then Picture1.PSet (I, 339 - (INF(I) - YIN) * FY)
    Next I

    For I = 1 To Z1
        CX = (X(I) - XMI) * FX ' : If CX < 0 Or CX > 459 Then GoTo 10565
        CY = 339 - (Y(I) - YIN) * FY ' : If CY < 0 Or CY > 339 Then GoTo 10565
            If FlagDatiAColori = True Then
                Picture1.Line (CX - 4, CY)-(CX + 4, CY), QBColor(12)
                Picture1.Line (CX, CY - 4)-(CX, CY + 4), QBColor(12)
            Else
                Picture1.Line (CX - 4, CY)-(CX + 4, CY), QBColor(0)
                Picture1.Line (CX, CY - 4)-(CX, CY + 4), QBColor(0)
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

    FlagStatistica = False
    FlagGrafico = True
    Label13.Caption = Titolo1
    Label14.Caption = InAscisse
    Label15.Caption = InOrdinate
    Label16.Caption = Format(X1, "#####0.0####;-#####0.0####")
    Label17.Caption = Format(X2, "#####0.0####;-#####0.0####")
    Label18.Caption = Format(X3, "#####0.0####;-#####0.0####")
    Label19.Caption = Format(X4, "#####0.0####;-#####0.0####")
    Label20.Caption = Format(X5, "#####0.0####;-#####0.0####")
    Label21.Caption = Format(Y1, "#####0.0####;-#####0.0####")
    Label22.Caption = Format(Y2, "#####0.0####;-#####0.0####")
    Label23.Caption = Format(Y3, "#####0.0####;-#####0.0####")
    Label24.Caption = Format(Y4, "#####0.0####;-#####0.0####")
    Label25.Caption = Format(Y5, "#####0.0####;-#####0.0####")
    
    Label26.Visible = True
    Label26.Left = 574
    Label26.Top = Picture1.Top + (Picture1.Height / 3)
    Label26.Caption = "" & Chr(13)
    If SLB >= 0 Then
        Label26.Caption = Label26.Caption & "   y = " & Format(INA, "#####0.0####") & " + " & Format(Abs(SLB), "#####0.0####") & " x   "
    Else
        Label26.Caption = Label26.Caption & "   y = " & Format(INA, "#####0.0####") & " - " & Format(Abs(SLB), "#####0.0####") & " x   "
    End If
    Label26.Caption = Label26.Caption & Chr(13) & Chr(10) & "   n = " & Z1
    Label26.Caption = Label26.Caption & Chr(13) & "   r = " & Format(COR, "#####0.0####")
    Label26.Caption = Label26.Caption & Chr(13)
    Label26.Top = Picture1.Top + (Picture1.Height / 2) - (Label26.Height / 2)
    Flag1con = True
    Flag1senza = False
    Screen.MousePointer = 1

    Exit Sub

GestioneErroriGraficaXcon:
    Screen.MousePointer = 1
    MsgBox "E' presente un errore logico" + Chr(10) + "nei dati da rappresentare.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
    Exit Sub

End Sub

Private Sub GraficaXsenza_Click()
    
    On Error GoTo GestioneErroriGraficaXsenza
    ReDim AA(460), INF(460), SUP(460)
    Dim I, XX, YIN, YSU, FX, FY, CX, CY
    Dim X1, X2, X3, X4, X5
    Dim Y1, Y2, Y3, Y4, Y5
    Dim SCARTO

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
    FormPolinomiale.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
            Label26.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
            Label26.BackColor = &H80000005
        End If
    Picture1.Cls
    FormPolinomiale.ScaleWidth = 800
    FormPolinomiale.ScaleHeight = 480
    FormPolinomiale.Refresh

    For I = 1 To 460
        XX = XMI + (I - 1) * (XMA - XMI) / 460
        AA(I) = INA + SLB * XX
        SCARTO = DEVIATAT * DER * Sqr(1 + (1 / Z1) + ((AA(I) - MEDIAX) ^ 2 / DEVIANZAX))
        SUP(I) = AA(I) + SCARTO
        INF(I) = AA(I) - SCARTO
    Next I

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
    
    FX = 459 / (XMA - XMI): FY = 339 / (YSU - YIN)
    
    For I = 1 To 460
'        If (339 - (AA(I) - YIN) * FY) >= 0 And (339 - (AA(I) - YIN) * FY) <= 339 Then
        If I >= (XMI_Default - XMI) * FX + 1 And I <= (XMA_Default - XMI) * FX + 1 Then
            Picture1.CurrentX = I
            Picture1.CurrentY = 339 - (AA(I) - YIN) * FY
            Exit For
        End If
    Next I

    For I = 1 To 460
'        If (339 - (AA(I) - YIN) * FY) >= 0 And (339 - (AA(I) - YIN) * FY) <= 339 Then Picture1.Line -(I - 1, 339 - (AA(I) - YIN) * FY)
'        Picture1.Line -(I - 1, 339 - (AA(I) - YIN) * FY)
        If I >= (XMI_Default - XMI) * FX + 1 And I <= (XMA_Default - XMI) * FX + 1 Then Picture1.Line -(I, 339 - (AA(I) - YIN) * FY)
    Next

'    For I = 1 To 460
'        If (339 - (AA(I) - YIN) * FY) >= 0 And (339 - (AA(I) - YIN) * FY) <= 339 Then Picture1.PSet (I - 1, 339 - (AA(I) - YIN) * FY)
'    Next
  
'    For I = 4 To 460 Step 4
'        If (339 - (SUP(I) - YIN) * FY) > 0 And (339 - (SUP(I) - YIN) * FY) < 339 Then Picture1.PSet (I - 1, 339 - (SUP(I) - YIN) * FY)
'        If (339 - (INF(I) - YIN) * FY) > 0 And (339 - (INF(I) - YIN) * FY) < 339 Then Picture1.PSet (I - 1, 339 - (INF(I) - YIN) * FY)
'        If I >= (XMI_Default - XMI) * FX + 1 And I <= (XMA_Default - XMI) * FX + 1 Then Picture1.PSet (I, 339 - (SUP(I) - YIN) * FY)
'        If I >= (XMI_Default - XMI) * FX + 1 And I <= (XMA_Default - XMI) * FX + 1 Then Picture1.PSet (I, 339 - (INF(I) - YIN) * FY)
'    Next I

    For I = 1 To Z1
        CX = (X(I) - XMI) * FX ' : If CX < 0 Or CX > 459 Then GoTo 10565
        CY = 339 - (Y(I) - YIN) * FY ' : If CY < 0 Or CY > 339 Then GoTo 10565
            If FlagDatiAColori = True Then
                Picture1.Line (CX - 4, CY)-(CX + 4, CY), QBColor(12)
                Picture1.Line (CX, CY - 4)-(CX, CY + 4), QBColor(12)
            Else
                Picture1.Line (CX - 4, CY)-(CX + 4, CY), QBColor(0)
                Picture1.Line (CX, CY - 4)-(CX, CY + 4), QBColor(0)
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

    FlagStatistica = False
    FlagGrafico = True
    Label13.Caption = Titolo1
    Label14.Caption = InAscisse
    Label15.Caption = InOrdinate
    Label16.Caption = Format(X1, "#####0.0####;-#####0.0####")
    Label17.Caption = Format(X2, "#####0.0####;-#####0.0####")
    Label18.Caption = Format(X3, "#####0.0####;-#####0.0####")
    Label19.Caption = Format(X4, "#####0.0####;-#####0.0####")
    Label20.Caption = Format(X5, "#####0.0####;-#####0.0####")
    Label21.Caption = Format(Y1, "#####0.0####;-#####0.0####")
    Label22.Caption = Format(Y2, "#####0.0####;-#####0.0####")
    Label23.Caption = Format(Y3, "#####0.0####;-#####0.0####")
    Label24.Caption = Format(Y4, "#####0.0####;-#####0.0####")
    Label25.Caption = Format(Y5, "#####0.0####;-#####0.0####")
    
    Label26.Visible = True
    Label26.Left = 574
    Label26.Top = Picture1.Top + (Picture1.Height / 3)
    Label26.Caption = "" & Chr(13)
    If SLB >= 0 Then
        Label26.Caption = Label26.Caption & "   y = " & Format(INA, "#####0.0####") & " + " & Format(Abs(SLB), "#####0.0####") & " x   "
    Else
        Label26.Caption = Label26.Caption & "   y = " & Format(INA, "#####0.0####") & " - " & Format(Abs(SLB), "#####0.0####") & " x   "
    End If
    Label26.Caption = Label26.Caption & Chr(13) & Chr(10) & "   n = " & Z1
    Label26.Caption = Label26.Caption & Chr(13) & "   r = " & Format(COR, "#####0.0####")
    Label26.Caption = Label26.Caption & Chr(13)
    Label26.Top = Picture1.Top + (Picture1.Height / 2) - (Label26.Height / 2)
    Flag1con = False
    Flag1senza = True
    Screen.MousePointer = 1
    Exit Sub

GestioneErroriGraficaXsenza:
    Screen.MousePointer = 1
    MsgBox "E' presente un errore logico" + Chr(10) + "nei dati da rappresentare.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
    Exit Sub

End Sub

Private Sub GraficaY_Click()
    On Error GoTo GestioneErroriGraficaYpoli
    ReDim AA(460), INF(460), SUP(460)
    Dim I, XX, YIN, YSU, FX, FY, CX, CY
    Dim X1, X2, X3, X4, X5
    Dim Y1, Y2, Y3, Y4, Y5

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
    FormPolinomiale.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
            Label26.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
            Label26.BackColor = &H80000005
        End If
    Picture1.Cls
    FormPolinomiale.ScaleWidth = 800
    FormPolinomiale.ScaleHeight = 480
    FormPolinomiale.Refresh

    For I = 1 To 460
        XX = XMI + (I - 1) * (XMA - XMI) / 460
        AA(I) = APO + BPO * XX + CPO * XX * XX
    Next I

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
    
    FX = 459 / (XMA - XMI): FY = 339 / (YSU - YIN)
    
    For I = 1 To 460
'        If (339 - (AA(I) - YIN) * FY) >= 0 And (339 - (AA(I) - YIN) * FY) <= 339 Then
        If I >= (XMI_Default - XMI) * FX + 1 And I <= (XMA_Default - XMI) * FX + 1 Then
            Picture1.CurrentX = I
            Picture1.CurrentY = 339 - (AA(I) - YIN) * FY
            Exit For
        End If
    Next I

    For I = 1 To 460
'        If (339 - (AA(I) - YIN) * FY) >= 0 And (339 - (AA(I) - YIN) * FY) <= 339 Then Picture1.Line -(I - 1, 339 - (AA(I) - YIN) * FY)
'        Picture1.Line -(I - 1, 339 - (AA(I) - YIN) * FY)
        If I >= (XMI_Default - XMI) * FX + 1 And I <= (XMA_Default - XMI) * FX + 1 Then Picture1.Line -(I, 339 - (AA(I) - YIN) * FY)
    Next

    For I = 1 To Z1
        CX = (X(I) - XMI) * FX ' : If CX < 0 Or CX > 459 Then GoTo 10565
        CY = 339 - (Y(I) - YIN) * FY ' : If CY < 0 Or CY > 339 Then GoTo 10565
            If FlagDatiAColori = True Then
                Picture1.Line (CX - 4, CY)-(CX + 4, CY), QBColor(12)
                Picture1.Line (CX, CY - 4)-(CX, CY + 4), QBColor(12)
            Else
                Picture1.Line (CX - 4, CY)-(CX + 4, CY), QBColor(0)
                Picture1.Line (CX, CY - 4)-(CX, CY + 4), QBColor(0)
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

    FlagStatistica = False
    FlagGrafico = True
    Label13.Caption = Titolo2
    Label14.Caption = InAscisse
    Label15.Caption = InOrdinate
    Label16.Caption = Format(X1, "#####0.0####;-#####0.0####")
    Label17.Caption = Format(X2, "#####0.0####;-#####0.0####")
    Label18.Caption = Format(X3, "#####0.0####;-#####0.0####")
    Label19.Caption = Format(X4, "#####0.0####;-#####0.0####")
    Label20.Caption = Format(X5, "#####0.0####;-#####0.0####")
    Label21.Caption = Format(Y1, "#####0.0####;-#####0.0####")
    Label22.Caption = Format(Y2, "#####0.0####;-#####0.0####")
    Label23.Caption = Format(Y3, "#####0.0####;-#####0.0####")
    Label24.Caption = Format(Y4, "#####0.0####;-#####0.0####")
    Label25.Caption = Format(Y5, "#####0.0####;-#####0.0####")
    
    Label26.Visible = True
    Label26.Left = 574
    Label26.Top = Picture1.Top + (Picture1.Height / 3)
    Label26.Caption = "" & Chr(13)
    Label26.Caption = Label26.Caption & "   y = a + bx + cx" & Chr(178) & "   " & Chr(13) & Chr(10)
    Label26.Caption = Label26.Caption & "   n = " & Z1 & "   " & Chr(13) & Chr(10)
    Label26.Caption = Label26.Caption & "   a = " & Format(APO, "#######0.0#########") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   b = " & Format(BPO, "#######0.0#########") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   c = " & Format(CPO, "#######0.0#########") & "   "
    Label26.Caption = Label26.Caption & Chr(13)
    Label26.Top = Picture1.Top + (Picture1.Height / 2) - (Label26.Height / 2)
    Screen.MousePointer = 1
    Exit Sub

GestioneErroriGraficaYpoli:
    Screen.MousePointer = 1
    MsgBox "E' presente un errore logico" + Chr(10) + "nei dati da rappresentare.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
    Exit Sub

End Sub

Private Sub MenuFile_Click()

    FileStampa.Enabled = False
        If FlagGrafico = True Then
            If FormPolinomiale.Width < Screen.Width Then
                FileStampa.Enabled = True
            End If
        End If
        If FlagStatistica = True Then
            FileStampa.Enabled = True
        End If

End Sub

Private Sub MenuGrafica_Click()
        
    GraficaX.Enabled = False
    GraficaY.Enabled = False
    GraficaCPS.Enabled = False
    GraficaScaleLegende.Enabled = False
    GraficaConfigurazioneColori.Enabled = False

    If Flag1 = True Then
        GraficaX.Enabled = True
    End If
    
    If Flag2 = True Then
        GraficaY.Enabled = True
    End If
    
    If Flag3 = True Then
        GraficaCPS.Enabled = True
    End If
    
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

Private Sub StatisticaCPS_Click()
    On Error GoTo GestioneErroriCPSpoli
    Dim A, B, C, D, e, F, G, H, I, J, K, L, M, N, O As Double
    Dim p, Q, A1, A2, A3, A4, A5, A6, A7, A8, A9, AA, AB As Double
    Dim XX, YY
    
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
    FormPolinomiale.Refresh
    
    FormGauge.Show
    FormGauge.Caption = "Polinomio di terzo grado"
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    A1 = 0: A2 = 0: A3 = 0: A4 = 0: A5 = 0: A6 = 0: A7 = 0: A8 = 0: A9 = 0: AA = 0: AB = 0
        For I = 1 To Z1
            FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
            FormGauge!Gauge1.Value = Int(100 * I / Z1)
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            XX = X(I): YY = Y(I): A1 = A1 + 1: A2 = A2 + XX: A3 = A3 + XX * XX: A4 = A4 + XX * XX * XX
            A5 = A5 + YY: A6 = A6 + XX * XX * XX * XX: A7 = A7 + XX * YY: A8 = A8 + XX * XX * XX * XX * XX
            A9 = A9 + YY * XX * XX: AA = AA + XX * XX * XX * XX * XX * XX: AB = AB + YY * XX * XX * XX
        Next
    B = A2 / A1
    C = A3 / A1
    D = A4 / A1
    e = A5 / A1
    F = (A2 * C - A4) / (A2 * B - A3)
    G = (A2 * D - A6) / (A2 * B - A3)
    H = (A2 * e - A7) / (A2 * B - A3)
    I = B * F - C
    J = B * G - D
    K = B * H - e
    L = (-A8 - A3 * J + A4 * G) / (-A6 - A3 * I + A4 * F)
    M = (-A9 - A3 * K + A4 * H) / (-A6 - A3 * I + A4 * F)
    N = I * L - J
    O = I * M - K
    p = F * L - G
    Q = F * M - H
    DPO = (-AB + A4 * O - A6 * Q + A8 * M) / (-AA + A4 * N - A6 * p + A8 * L)
    APO = O - N * DPO
    BPO = p * DPO - Q
    CPO = M - L * DPO
            
        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Regressione polinomiale di terzo grado (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File        : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella     : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "In ascisse  : " & InAscisse & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "In ordinate : " & InOrdinate & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Polinomio di equazione y = a + bx + cx" & Chr(178) & "+ dx" & Chr(179) & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Numero dei dati  = " & Z1 & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Termine noto a   = " & Format(APO, "#######0.0#########") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Coefficiente b   = " & Format(BPO, "#######0.0#########") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Coefficiente c   = " & Format(CPO, "#######0.0#########") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Coefficiente d   = " & Format(DPO, "#######0.0#########") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Regressione polinomiale di terzo grado (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge

        Text4.Visible = True
        Text4.Enabled = True
        Flag1 = False
        Flag2 = False
        Flag3 = True
        FlagStatistica = True
        FlagGrafico = False
'        Titolo = "Regressione polinomiale di terzo grado"
        Screen.MousePointer = 1

Exit Sub

GestioneErroriCPSpoli:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico nei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        Exit Sub

End Sub

Private Sub StatisticaX_Click()
    On Error GoTo GestioneErroriXpoli
    Dim A, B, C, D, e, F, G, H, I, L, M, O
    Dim N, p, c0, c1, c2, t1, t2, e1
    Dim t5, i1, i2, i3, i4, u1, XX, YY
    Dim M33, Sommat, Residuo, PI, T7, A3, J2, S1
    Dim K, S2, J, P1
    Dim t, PT As Double

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
    FormPolinomiale.Refresh
    
    FormGauge.Show
    FormGauge.Caption = "Polinomio di primo grado"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        A = 0: B = 0: C = 0: D = 0: e = 0: F = 0
        
        FormGauge!Label6.Caption = "[1/3]"
        For I = 1 To Z1
            A = A + 1: B = B + X(I): C = C + Y(I): D = D + X(I) * X(I): e = e + Y(I) * Y(I): F = F + X(I) * Y(I)
            If 100 * I / Z1 = Int(100 * I / Z1) Then
                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If
        Next I

        G = D - B * B / A: H = e - C * C / A: I = F - B * C / A: L = I * I / G: M = (H - L) / (A - 2)
        SLB = I / G
        INA = (C / A) - (I / G) * (B / A)
        COR = I / Sqr(G * H)
        DER = Sqr(M)
        DES = Sqr(M) * Sqr(1 / G)
        DEI = Sqr(M) * Sqr(1 / G) * Sqr(D / A)

' calcolo di t a partire da p e dai gradi di libertà n
        N = A - 2
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

' necessari per il calcolo dei limiti di confidenza
        DEVIATAT = t
        MEDIAX = B / A
        DEVIANZAX = G

' test t per la linearità
        
        A = 0: B = 0: C = 0: D = 0: e = 0: F = 0: G = 0: H = 0
    
        FormGauge!Label6.Caption = "[2/3]"
        For I = 1 To Z1
            FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
            FormGauge!Gauge1.Value = Int(100 * I / Z1)
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            XX = X(I): YY = Y(I): A = A + 1: B = B + XX: C = C + XX * XX: D = D + XX * XX * XX: e = e + XX * XX * XX * XX: F = F + YY: G = G + XX * YY: H = H + YY * XX * XX
        Next
        L = C * (B * D - C * C) + D * (C * B - A * D) + e * (A * C - B * B)
        M = H * (B * D - C * C) + D * (C * G - F * D) + e * (F * C - B * G)
        N = C * (F * D - C * G) + H * (C * B - A * D) + e * (A * G - F * B)
        O = C * (B * G - F * C) + D * (F * B - A * G) + H * (A * C - B * B)
        APO = M / L: BPO = N / L: CPO = O / L
    
        M33 = e - (C * C / A) - (A * A * D * D - 2 * A * C * B * D + C * C * B * B) / (A * A * C - A * B * B)
        M33 = 1 / M33
        Sommat = 0
    
        FormGauge!Label6.Caption = "[3/3]"
        For I = 1 To Z1
            FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
            FormGauge!Gauge1.Value = Int(100 * I / Z1)
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            Sommat = Sommat + (APO + BPO * X(I) + CPO * X(I) * X(I) - Y(I)) ^ 2
        Next I
            Residuo = Sommat / (Z1 - 3)

    t = CPO / Sqr(Residuo * M33): N = Z1 - 3
    
    PI = 3.141592654
    T7 = Atn(Abs(t) / Sqr(N))
    A3 = 0
    J2 = N - 2
    If N <> Int(N / 2) * 2 Then GoTo tt3200
    For K = J2 To 2 Step -2
    S1 = 0: S2 = 0
    For I = K - 1 To 1 Step -2: S1 = S1 + Log(I): Next
    For J = K To 2 Step -2: S2 = S2 + Log(J): Next
    A3 = A3 + Exp(S1 - S2) * Cos(T7) ^ K: Next
    P1 = (A3 + 1) * Sin(T7): GoTo tt3280
tt3200:
    If N <> 1 Then GoTo tt3220
    P1 = 2 * T7 / PI: GoTo tt3280
tt3220:
    For K = J2 To 2 Step -2
    S1 = 0: S2 = 0
    For I = K - 1 To 2 Step -2: S1 = S1 + Log(I): Next
    For J = K To 1 Step -2: S2 = S2 + Log(J): Next
    A3 = A3 + Exp(S1 - S2) * Cos(T7) ^ K: Next
    P1 = ((A3 + Cos(T7)) * Sin(T7) + T7) * 2 / PI
tt3280:
    PT = 1 - P1

tt3300:

        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Regressione lineare parametrica x variabile indipendente (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File        : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella     : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "In ascisse  : " & InAscisse & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "In ordinate : " & InOrdinate & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Retta di equazione y = a + bx" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Numero dei dati                = " & Z1 & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Intercetta a                   = " & Format(INA, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Coefficiente angolare b        = " & Format(SLB, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Coefficiente di correlazione r = " & Format(COR, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Errore standard di a           = " & Format(DEI, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Errore standard di b           = " & Format(DES, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Errore standard della stima    = " & Format(DER, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Test t per la linearità        = " & Format(Abs(t), "#####0.0####") & Chr(13) & Chr(10)
        If PT >= 0.00001 Then
            Text4.Text = Text4.Text & "Probabilità p                  = " & Format(PT, "#####0.0####") & Chr(13) & Chr(10)
        Else
            PT = 0.00001
            Text4.Text = Text4.Text & "Probabilità p                  < " & Format(PT, "#####0.0####") & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Regressione lineare parametrica x variabile indipendente (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge

        Text4.Visible = True
        Text4.Enabled = True
        Flag1 = True
        Flag2 = False
        Flag3 = False
        FlagStatistica = True
        FlagGrafico = False
'        Titolo = "Regressione lineare x variabile indipendente"
        Screen.MousePointer = 1

Exit Sub

GestioneErroriXpoli:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico nei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        Exit Sub

End Sub

Private Sub StatisticaY_Click()
    On Error GoTo GestioneErroriYpoli
    Dim A, B, C, D, e, F, G, H, I, L, M, N, O As Double
    Dim XX, YY

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
    FormPolinomiale.Refresh
    
    FormGauge.Show
    FormGauge.Caption = "Polinomio di secondo grado"
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        A = 0: B = 0: C = 0: D = 0: e = 0: F = 0: G = 0: H = 0
        For I = 1 To Z1
            FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
            FormGauge!Gauge1.Value = Int(100 * I / Z1)
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            XX = X(I): YY = Y(I): A = A + 1: B = B + XX: C = C + XX * XX: D = D + XX * XX * XX: e = e + XX * XX * XX * XX: F = F + YY: G = G + XX * YY: H = H + YY * XX * XX
        Next
        L = C * (B * D - C * C) + D * (C * B - A * D) + e * (A * C - B * B)
        M = H * (B * D - C * C) + D * (C * G - F * D) + e * (F * C - B * G)
        N = C * (F * D - C * G) + H * (C * B - A * D) + e * (A * G - F * B)
        O = C * (B * G - F * C) + D * (F * B - A * G) + H * (A * C - B * B)
        APO = M / L: BPO = N / L: CPO = O / L

        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Regressione polinomiale di secondo grado (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File        : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella     : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "In ascisse  : " & InAscisse & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "In ordinate : " & InOrdinate & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Polinomio di equazione y = a + bx + cx" & Chr(178) & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Numero dei dati  = " & Z1 & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Termine noto a   = " & Format(APO, "#######0.0#########") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Coefficiente b   = " & Format(BPO, "#######0.0#########") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Coefficiente c   = " & Format(CPO, "#######0.0#########") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Regressione polinomiale di secondo grado (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge

        Text4.Visible = True
        Text4.Enabled = True
        Flag1 = False
        Flag2 = True
        Flag3 = False
        FlagStatistica = True
        FlagGrafico = False
'        Titolo = "Regressione polinomiale di secondo grado"
        Screen.MousePointer = 1

Exit Sub

GestioneErroriYpoli:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico nei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        Exit Sub

End Sub

