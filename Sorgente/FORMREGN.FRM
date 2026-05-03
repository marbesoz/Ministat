VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormNonParametrica 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Regressione lineare non parametrica"
   ClientHeight    =   6315
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
   Icon            =   "FORMREGN.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6315
   ScaleWidth      =   7365
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   4080
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
         Caption         =   "&Regressione lineare x variabile indipendente"
      End
      Begin VB.Menu GraficaY 
         Caption         =   "R&egressione lineare y variabile indipendente"
      End
      Begin VB.Menu GraficaCPS 
         Caption         =   "Regre&ssione lineare di Passing e Bablok"
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
         Caption         =   "&Regressione lineare x variabile indipendente"
      End
      Begin VB.Menu StatisticaY 
         Caption         =   "R&egressione lineare y variabile indipendente"
      End
      Begin VB.Menu StatisticaCPS 
         Caption         =   "Regre&ssione lineare di Passing e Bablok"
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
Attribute VB_Name = "FormNonParametrica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim X() As Double
    Dim Y() As Double
    Dim AA(), SUP(), INF()
    Dim Riga, ColonnaX, ColonnaY, FineDati
    Dim Z1, Z2, A, B, C, D, e, F, G, H, I, L, M
    Dim SLB, INA, COR, DER, DES, DEI
    Dim XMI, XMA, YMI, YMA
    Dim XMI_Default, XMA_Default
    Dim YMI_Default, YMA_Default
    Dim DEVIATAT, MEDIAX, DEVIANZAX
    Dim Titolo1, Titolo2, Titolo3
    Dim InAscisse, InOrdinate
    Dim Flag1, Flag2, Flag3
    Dim FlagStatistica, FlagGrafico
    Dim FlagSfondoGrigio
    Dim FlagDatiAColori

Private Sub AiutoCalcolatrice_Click()
    Dim Calcolatrice
    On Error GoTo ErroreCalcolatriceSei:
    
    Calcolatrice = Shell("Calc.exe", 1)
        
Exit Sub

ErroreCalcolatriceSei:
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
    On Error GoTo NonParametrica
        
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
            GraficaX_Click
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

NonParametrica:
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
    Y = INA - (-SLB * MaskedEdit5.Text)
    Label10.Caption = Format(Y, "#####0.0####")

    Screen.MousePointer = 1

Exit Sub

GestioneErroriCommand2:
    Label10.Caption = ""
    Screen.MousePointer = 1
    Exit Sub

End Sub

Private Sub Command3_Click()
    On Error GoTo GestioneErroriCommand3
    Dim X

    Screen.MousePointer = 11
    
    MaskedEdit6.Text = FormattaStringa(MaskedEdit6.Text)
    X = (MaskedEdit6.Text - INA) / SLB
    Label12.Caption = Format(X, "#####0.0####")

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
            GraficaX_Click
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
    ReDim AA(0)
    FormPrincipale.Show
    Unload FormNonParametrica
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
End Sub

Private Sub FileStampa_Click()
    On Error GoTo ErroreFileStampaRegn

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

ErroreFileStampaRegn:
    Exit Sub

End Sub

Private Sub Form_Load()
    ReDim X(1 To NUM_RIGHE - 1) As Double
    ReDim Y(1 To NUM_RIGHE - 1) As Double

    Screen.MousePointer = 11
    AnnullatoCaricamentoForm = False
    Flag1 = False
    Flag2 = False
    Flag3 = False
    FlagStatistica = False
    FlagGrafico = False
    FlagSfondoGrigio = True
    FlagDatiAColori = True
    Picture2.Visible = False
    
' centra il form
    FormNonParametrica.Top = Screen.Height / 2 - Height / 2
    FormNonParametrica.Left = Screen.Width / 2 - Width / 2
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
    Titolo1 = "Regressione lineare x variabile indipendente"
    Titolo2 = "Regressione lineare y variabile indipendente"
    Titolo3 = "Regressione lineare di Passing e Bablok"
' carica i dati nei vettori X() e Y()
    FormGauge.Show
    FormGauge.Caption = "Regressione lineare non parametrica"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    FineDati = False
    Z1 = 0
    Z2 = 0
    
        FormGauge!Label6.Caption = "[1/1]"
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
    If Z1 > 2 Then
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
            FormNonParametrica.Show
            ' FormPrincipale.Hide
            Screen.MousePointer = 1
            Exit Sub
        Else
            Unload FormGauge
            AnnullatoCaricamentoForm = True
            Screen.MousePointer = 1
            MsgBox "Per la regressione lineare" + Chr(10) + "le due variabili devono contenere" + Chr(10) + "lo stesso numero di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Regressione lineare non eseguibile"
            FormPrincipale!Grid1.HighLight = True
            RipristinaSchermo
        End If
    Else
        Unload FormGauge
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        MsgBox "Per la regressione lineare" + Chr(10) + "sono necessari almeno tre dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Regressione lineare non eseguibile"
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
    End If

End Sub

Private Sub Form_Resize()
    
    Text4.Width = ScaleWidth
    Text4.Height = ScaleHeight

    If FlagGrafico = True Then
    
    If Flag1 = True Then
        GraficaX_Click
    End If
    If Flag2 = True Then
        GraficaY_Click
    End If
    If Flag3 = True Then
        GraficaCPS_Click
    End If

    End If

End Sub

Private Sub GraficaCPS_Click()
    On Error GoTo GestioneErroriGraficaCPSnon
    ReDim AA(460)
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
    FormNonParametrica.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
            Label26.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
            Label26.BackColor = &H80000005
        End If
    Picture1.Cls
    FormNonParametrica.ScaleWidth = 800
    FormNonParametrica.ScaleHeight = 480
    FormNonParametrica.Refresh

    For I = 1 To 460
        XX = XMI + (I - 1) * (XMA - XMI) / 460
        AA(I) = INA + SLB * XX
    Next I

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
    
    For I = 1 To 460
        If (339 - (AA(I) - YIN) * FY) >= 0 And (339 - (AA(I) - YIN) * FY) <= 339 Then
            Picture1.CurrentX = I - 1
            Picture1.CurrentY = 339 - (AA(I) - YIN) * FY
            Exit For
        End If
    Next I

    For I = 1 To 460
        If (339 - (AA(I) - YIN) * FY) >= 0 And (339 - (AA(I) - YIN) * FY) <= 339 Then Picture1.Line -(I - 1, 339 - (AA(I) - YIN) * FY)
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
    If SLB >= 0 Then
        Label26.Caption = Label26.Caption & "   y = " & Format(INA, "#####0.0####") & " + " & Format(Abs(SLB), "#####0.0####") & " x   "
    Else
        Label26.Caption = Label26.Caption & "   y = " & Format(INA, "#####0.0####") & " - " & Format(Abs(SLB), "#####0.0####") & " x   "
    End If
    Label26.Caption = Label26.Caption & Chr(13) & Chr(10) & "   n = " & Z1
    Label26.Caption = Label26.Caption & Chr(13)
    Label26.Top = Picture1.Top + (Picture1.Height / 2) - (Label26.Height / 2)
    
    Picture1.Visible = True
    Screen.MousePointer = 1
    Exit Sub

GestioneErroriGraficaCPSnon:
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
        GraficaX_Click
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
        GraficaX_Click
    End If
    If Flag2 = True Then
        GraficaY_Click
    End If
    If Flag3 = True Then
        GraficaCPS_Click
    End If

    End If

End Sub

Private Sub GraficaX_Click()
    On Error GoTo GestioneErroriGraficaXnon
    ReDim AA(460)
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
    FormNonParametrica.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
            Label26.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
            Label26.BackColor = &H80000005
        End If
    Picture1.Cls
    FormNonParametrica.ScaleWidth = 800
    FormNonParametrica.ScaleHeight = 480
    FormNonParametrica.Refresh

    For I = 1 To 460
        XX = XMI + (I - 1) * (XMA - XMI) / 460
        AA(I) = INA + SLB * XX
    Next I

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
    
    For I = 1 To 460
        If (339 - (AA(I) - YIN) * FY) >= 0 And (339 - (AA(I) - YIN) * FY) <= 339 Then
            Picture1.CurrentX = I - 1
            Picture1.CurrentY = 339 - (AA(I) - YIN) * FY
            Exit For
        End If
    Next I

    For I = 1 To 460
        If (339 - (AA(I) - YIN) * FY) >= 0 And (339 - (AA(I) - YIN) * FY) <= 339 Then Picture1.Line -(I - 1, 339 - (AA(I) - YIN) * FY)
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
    Label26.Caption = Label26.Caption & Chr(13)
    Label26.Top = Picture1.Top + (Picture1.Height / 2) - (Label26.Height / 2)
    
    Picture1.Visible = True
    Screen.MousePointer = 1
    Exit Sub

GestioneErroriGraficaXnon:
    Screen.MousePointer = 1
    MsgBox "E' presente un errore logico" + Chr(10) + "nei dati da rappresentare.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
    Exit Sub

End Sub

Private Sub GraficaY_Click()
    On Error GoTo GestioneErroriGraficaYnon
    ReDim AA(460)
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
    FormNonParametrica.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
            Label26.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
            Label26.BackColor = &H80000005
        End If
    Picture1.Cls
    FormNonParametrica.ScaleWidth = 800
    FormNonParametrica.ScaleHeight = 480
    FormNonParametrica.Refresh

    For I = 1 To 460
        XX = XMI + (I - 1) * (XMA - XMI) / 460
        AA(I) = INA + SLB * XX
    Next I

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
    
    For I = 1 To 460
        If (339 - (AA(I) - YIN) * FY) >= 0 And (339 - (AA(I) - YIN) * FY) <= 339 Then
            Picture1.CurrentX = I - 1
            Picture1.CurrentY = 339 - (AA(I) - YIN) * FY
            Exit For
        End If
    Next I

    For I = 1 To 460
        If (339 - (AA(I) - YIN) * FY) >= 0 And (339 - (AA(I) - YIN) * FY) <= 339 Then Picture1.Line -(I - 1, 339 - (AA(I) - YIN) * FY)
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
    If SLB >= 0 Then
        Label26.Caption = Label26.Caption & "   y = " & Format(INA, "#####0.0####") & " + " & Format(Abs(SLB), "#####0.0####") & " x   "
    Else
        Label26.Caption = Label26.Caption & "   y = " & Format(INA, "#####0.0####") & " - " & Format(Abs(SLB), "#####0.0####") & " x   "
    End If
    Label26.Caption = Label26.Caption & Chr(13) & Chr(10) & "   n = " & Z1
    Label26.Caption = Label26.Caption & Chr(13)
    Label26.Top = Picture1.Top + (Picture1.Height / 2) - (Label26.Height / 2)
    
    Picture1.Visible = True
    Screen.MousePointer = 1
    Exit Sub

GestioneErroriGraficaYnon:
    Screen.MousePointer = 1
    MsgBox "E' presente un errore logico" + Chr(10) + "nei dati da rappresentare.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
    Exit Sub

End Sub

Private Sub MenuFile_Click()

    FileStampa.Enabled = False
        If FlagGrafico = True Then
            If FormNonParametrica.Width < Screen.Width Then
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
    On Error GoTo GestioneErroriNonparamPB
    Dim A, B, C, D, e, F, J, L, XX
    Dim MX, MY, ValoreInizialeD
    ReDim AA(Z1 * (Z1 - 1) / 2)

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
    FormNonParametrica.Refresh
    
    FormGauge.Show
    FormGauge.Caption = "Regressione di Passing e Bablok"
    FormGauge!Label6.Caption = "[1/2]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    e = 0
    F = 0
    For A = 1 To Z1 - 1
        For B = A + 1 To Z1
        C = X(A) - X(B)
        D = Y(A) - Y(B)
            If C = -D Then GoTo a1206
            If C = 0 And D = 0 Then GoTo a1206
            If C = 0 And D < 0 Then GoTo a1182
            GoTo a1190
a1182:
        e = e + 1
        AA(e) = -9.99E+37
        F = F + 1
        GoTo a1206
a1190:
            If C = 0 And D > 0 Then GoTo a1194
            GoTo a1200
a1194:
        e = e + 1
        AA(e) = 9.99E+37
        GoTo a1206
a1200:
        e = e + 1
        AA(e) = D / C
        If AA(e) < -1 Then F = F + 1
a1206:
        Next B
    Next A
    GoSub OrdinamentoDatiPB
    SLB = AA((e + 1) / 2 + F)
    If e / 2 = Int(e / 2) Then SLB = 0.5 * (AA(e / 2 + F) + AA(e / 2 + 1 + F))
    
    For I = 1 To Z1
        AA(I) = Y(I) - SLB * X(I)
    Next I
    e = Z1
    FormGauge!Label6.Caption = "[2/2]"
    GoSub OrdinamentoDatiPB
    INA = AA((Z1 + 1) / 2)
    If Z1 / 2 = Int(Z1 / 2) Then INA = 0.5 * (AA(Z1 / 2) + AA(Z1 / 2 + 1))

        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Regressione lineare non parametrica di Passing e Bablok (*inizio*)" & Chr(13) & Chr(10)
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
        Text4.Text = Text4.Text & "Numero dei dati                = " & Format(Z1, "#####0") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Intercetta a                   = " & Format(INA, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Coefficiente angolare b        = " & Format(SLB, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Regressione lineare non parametrica di Passing e Bablok (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge

        Text4.Visible = True
        Text4.Enabled = True
        Flag1 = False
        Flag2 = False
        Flag3 = True
        FlagStatistica = True
        FlagGrafico = False
'        Titolo = "Regressione lineare di Passing e Bablok"
        Screen.MousePointer = 1

Exit Sub

OrdinamentoDatiPB:

    ValoreInizialeD = 0

' inizio ordinamento dati

apb330:
    D = 1
apb340:
    D = 2 * D: If D <= e Then GoTo apb340
apb350:
    D = Int((D - 1) / 2): If D = 0 Then GoTo apb420
apb360:
    For I = 1 To e - D
apb370:
        J = I
apb380:
        L = J + D
        If AA(J) <= AA(L) Then GoTo apb400
                    
apb390:
            XX = AA(J): AA(J) = AA(L): AA(L) = XX

            If I > ValoreInizialeD And 100 * I / e >= 1 Then
                FormGauge.Label1.Caption = Int(100 * I / e) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / e)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If
            J = J - D: If J > 0 Then GoTo apb380
apb400:
    Next I
apb410:
    ValoreInizialeD = e - D
    GoTo apb350
apb420:
' fine ordinamento dati
                
                FormGauge.Label1.Caption = "100%"
                FormGauge!Gauge1.Value = 100
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

Return

GestioneErroriNonparamPB:
        Unload FormGauge
        Flag1 = False: Flag2 = False: Flag3 = False
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico nei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        Exit Sub

End Sub

Private Sub StatisticaX_Click()
    On Error GoTo GestioneErroriNonparamX
    Dim A, B, C, D, e, J, L, XX
    Dim MX, MY, ValoreInizialeD
    ReDim AA(Z1 * (Z1 - 1) / 2)

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
    FormNonParametrica.Refresh

    FormGauge.Show
    FormGauge.Caption = "Regressione x variabile indipendente"
    FormGauge!Label6.Caption = "[1/3]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    e = 0
    For A = 1 To Z1 - 1
        For B = A + 1 To Z1
        C = X(A) - X(B)
        D = Y(A) - Y(B)
            If C = 0 Then GoTo 1260
        e = e + 1
        AA(e) = D / C
1260:
        Next B
    Next A
    GoSub OrdinamentoDati
    SLB = AA((e + 1) / 2)
    If e / 2 = Int(e / 2) Then SLB = 0.5 * (AA(e / 2) + AA(e / 2 + 1))
    
    For I = 1 To Z1
        AA(I) = X(I)
    Next I
    e = Z1
    FormGauge!Label6.Caption = "[2/3]"
    GoSub OrdinamentoDati
    MX = AA((Z1 + 1) / 2)
    If Z1 / 2 = Int(Z1 / 2) Then MX = 0.5 * (AA(Z1 / 2) + AA(Z1 / 2 + 1))
    
    For I = 1 To Z1
        AA(I) = Y(I)
    Next I
    e = Z1
    FormGauge!Label6.Caption = "[3/3]"
    GoSub OrdinamentoDati
    MY = AA((Z1 + 1) / 2)
    If Z1 / 2 = Int(Z1 / 2) Then MY = 0.5 * (AA(Z1 / 2) + AA(Z1 / 2 + 1))
    
    INA = MY - SLB * MX

        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Regressione lineare non parametrica x variabile indipendente (*inizio*)" & Chr(13) & Chr(10)
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
        Text4.Text = Text4.Text & "Numero dei dati                = " & Format(Z1, "#####0") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Intercetta a                   = " & Format(INA, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Coefficiente angolare b        = " & Format(SLB, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Regressione lineare non parametrica x variabile indipendente (*fine*)" & Chr(13) & Chr(10)
        
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

OrdinamentoDati:

    ValoreInizialeD = 0

' inizio ordinamento dati

ac330:
    D = 1
ac340:
    D = 2 * D: If D <= e Then GoTo ac340
ac350:
    D = Int((D - 1) / 2): If D = 0 Then GoTo ac420
ac360:
    For I = 1 To e - D
ac370:
        J = I
ac380:
        L = J + D
        If AA(J) <= AA(L) Then GoTo ac400
                    
ac390:
            XX = AA(J): AA(J) = AA(L): AA(L) = XX

            If I > ValoreInizialeD And 100 * I / e >= 1 Then
                FormGauge.Label1.Caption = Int(100 * I / e) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / e)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If
            J = J - D: If J > 0 Then GoTo ac380
ac400:
    Next I
ac410:
    ValoreInizialeD = e - D
    GoTo ac350
ac420:
' fine ordinamento dati
                
                FormGauge.Label1.Caption = "100%"
                FormGauge!Gauge1.Value = 100
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

Return

GestioneErroriNonparamX:
        Unload FormGauge
        Flag1 = False: Flag2 = False: Flag3 = False
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico nei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        Exit Sub

End Sub

Private Sub StatisticaY_Click()
    On Error GoTo GestioneErroriNonparamY
    Dim A, B, C, D, e, J, L, XX
    Dim MX, MY, ValoreInizialeD
    ReDim AA(Z1 * (Z1 - 1) / 2)

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
    FormNonParametrica.Refresh
    
    FormGauge.Show
    FormGauge.Caption = "Regressione y variabile indipendente"
    FormGauge!Label6.Caption = "[1/3]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    e = 0
    For A = 1 To Z1 - 1
        For B = A + 1 To Z1
        C = Y(A) - Y(B)
        D = X(A) - X(B)
            If C = 0 Then GoTo a1260
        e = e + 1
        AA(e) = D / C
a1260:
        Next B
    Next A
    GoSub OrdinamentoDatiY
    SLB = AA((e + 1) / 2)
    If e / 2 = Int(e / 2) Then SLB = 0.5 * (AA(e / 2) + AA(e / 2 + 1))
    
    For I = 1 To Z1
        AA(I) = Y(I)
    Next I
    e = Z1
    FormGauge!Label6.Caption = "[2/3]"
    GoSub OrdinamentoDatiY
    MX = AA((Z1 + 1) / 2)
    If Z1 / 2 = Int(Z1 / 2) Then MX = 0.5 * (AA(Z1 / 2) + AA(Z1 / 2 + 1))
    
    For I = 1 To Z1
        AA(I) = X(I)
    Next I
    e = Z1
    FormGauge!Label6.Caption = "[3/3]"
    GoSub OrdinamentoDatiY
    MY = AA((Z1 + 1) / 2)
    If Z1 / 2 = Int(Z1 / 2) Then MY = 0.5 * (AA(Z1 / 2) + AA(Z1 / 2 + 1))
    
    INA = MY - SLB * MX

    INA = -INA / SLB
    SLB = 1 / SLB

        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Regressione lineare non parametrica y variabile indipendente (*inizio*)" & Chr(13) & Chr(10)
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
        Text4.Text = Text4.Text & "Numero dei dati                = " & Format(Z1, "#####0") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Intercetta a                   = " & Format(INA, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Coefficiente angolare b        = " & Format(SLB, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Regressione lineare non parametrica y variabile indipendente (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge

        Text4.Visible = True
        Text4.Enabled = True
        Flag1 = False
        Flag2 = True
        Flag3 = False
        FlagStatistica = True
        FlagGrafico = False
'        Titolo = "Regressione lineare y variabile indipendente"
        Screen.MousePointer = 1

Exit Sub

OrdinamentoDatiY:

    ValoreInizialeD = 0

' inizio ordinamento dati

yac330:
    D = 1
yac340:
    D = 2 * D: If D <= e Then GoTo yac340
yac350:
    D = Int((D - 1) / 2): If D = 0 Then GoTo yac420
yac360:
    For I = 1 To e - D
yac370:
        J = I
yac380:
        L = J + D
        If AA(J) <= AA(L) Then GoTo yac400
                    
yac390:
            XX = AA(J): AA(J) = AA(L): AA(L) = XX

            If I > ValoreInizialeD And 100 * I / e >= 1 Then
                FormGauge.Label1.Caption = Int(100 * I / e) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / e)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If
            J = J - D: If J > 0 Then GoTo yac380
yac400:
    Next I
yac410:
    ValoreInizialeD = e - D
    GoTo yac350
yac420:
' fine ordinamento dati
                
                FormGauge.Label1.Caption = "100%"
                FormGauge!Gauge1.Value = 100
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

Return

GestioneErroriNonparamY:
        Unload FormGauge
        Flag1 = False: Flag2 = False: Flag3 = False
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico nei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        Exit Sub

End Sub

