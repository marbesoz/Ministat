VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormBayes 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Statistica bayesiana"
   ClientHeight    =   6540
   ClientLeft      =   1095
   ClientTop       =   1755
   ClientWidth     =   7365
   ForeColor       =   &H80000008&
   Icon            =   "FORMBAYE.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6540
   ScaleWidth      =   7365
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
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
      Height          =   372
      Left            =   3240
      ScaleHeight     =   345
      ScaleWidth      =   945
      TabIndex        =   38
      Top             =   5520
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
      Height          =   4092
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   7092
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ripristina valori originari"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   600
         TabIndex        =   16
         Top             =   3360
         Width           =   1212
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         MaxLength       =   40
         TabIndex        =   7
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         MaxLength       =   50
         TabIndex        =   3
         Top             =   480
         Width           =   5775
      End
      Begin MSMask.MaskEdBox MaskedEdit6 
         Height          =   288
         Left            =   3720
         TabIndex        =   37
         Top             =   2640
         Width           =   852
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   327680
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Prevalenza della malattia:"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   3720
         TabIndex        =   36
         Top             =   2400
         Width           =   2052
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Numero di classi:"
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
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5760
      TabIndex        =   31
      Top             =   6120
      Width           =   972
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   4920
      TabIndex        =   30
      Top             =   6120
      Width           =   972
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   3480
      TabIndex        =   29
      Top             =   6120
      Width           =   972
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2160
      TabIndex        =   28
      Top             =   6120
      Width           =   972
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1200
      TabIndex        =   27
      Top             =   6120
      Width           =   972
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1680
      TabIndex        =   24
      Top             =   6360
      Width           =   5052
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
      Height          =   315
      Left            =   720
      TabIndex        =   23
      Top             =   360
      Width           =   6495
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
         Caption         =   "&Istogrammi delle distribuzioni"
      End
      Begin VB.Menu GraficaSensibilita 
         Caption         =   "&Sensibilità e specificità del test"
      End
      Begin VB.Menu GraficaValorePredittivo 
         Caption         =   "&Valore predittivo del test"
      End
      Begin VB.Menu GraficaEfficienzaDiagnostica 
         Caption         =   "&Efficienza diagnostica del test"
      End
      Begin VB.Menu GraficaRoc 
         Caption         =   "&Curva ROC del test"
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
      Begin VB.Menu StatisticaSuddivisioneInClassi 
         Caption         =   "S&uddivisione in classi"
      End
      Begin VB.Menu StatisticaStatisticheDelleClassi 
         Caption         =   "S&tatistiche delle classi"
      End
      Begin VB.Menu StatisticaValorePredittivo 
         Caption         =   "&Valore predittivo del test"
      End
      Begin VB.Menu StatisticaEfficienzaDiagnostica 
         Caption         =   "&Efficienza diagnostica del test"
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
Attribute VB_Name = "FormBayes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim X() As Double, Y() As Double
    Dim PercX() As Double, PercY() As Double
    Dim CCS(), DDS(), CCM(), DDM()
    Dim VP(), VN(), FP(), FN()
    Dim Limite(), Sens(), Spec(), VPPos(), VPNeg()
    Dim Riga, ColonnaX, ColonnaY, FineDati, Z1, Z2
    Dim XMI, XMA, YMI, YMA
    Dim XMI_Default, XMA_Default
    Dim YMI_Default, YMA_Default, CL_Default
    Dim LW, HI, CL, AM, MX
    Dim LWS, HIS, LWM, HIM
    Dim MedianaS, MedianaM, NumS, NumM
    Dim Prev, Prev_Default
    Dim PrimoQuartS, TerzoQuartS, PrimoQuartM, TerzoQuartM
    Dim Titolo1, Titolo2, Titolo3, Titolo4, Titolo5
    Dim InAscisse, InOrdinate
    Dim InAscisseBis, InOrdinateBis
    Dim FlagGrafico, FlagStatistica
    Dim FlagSfondoGrigio
    Dim FlagDatiAColori
    Dim Flag1, Flag2, Flag3, Flag4, Flag5

Private Sub AiutoCalcolatrice_Click()
    Dim Calcolatrice
    On Error GoTo ErroreCalcolatriceDue:
    
    Calcolatrice = Shell("Calc.exe", 1)
        
Exit Sub

ErroreCalcolatriceDue:
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
    Dim SommaS, SommaM, DummyNumS, DummyNumM
    On Error GoTo Bayes
        
    FlagErroreScale = False
    MaskedEdit1.Text = FormattaStringa(MaskedEdit1.Text)
    MaskedEdit2.Text = FormattaStringa(MaskedEdit2.Text)
    MaskedEdit3.Text = Int(CDbl(FormattaStringa(MaskedEdit3.Text)))
    MaskedEdit4.Text = Int(CDbl(FormattaStringa(MaskedEdit4.Text)))
    MaskedEdit5.Text = Int(CDbl(FormattaStringa(MaskedEdit5.Text)))
    MaskedEdit6.Text = FormattaStringa(MaskedEdit6.Text)
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
    If CDbl(MaskedEdit6.Text) < 0.001 Then
        MaskedEdit6.Text = Format(Prev_Default, "0.0##")
        FlagErroreScale = True
    End If
    If CDbl(MaskedEdit6.Text) > 0.999 Then
        MaskedEdit6.Text = Format(Prev_Default, "0.0##")
        FlagErroreScale = True
    End If
        
        If FlagErroreScale = False Then
            
            Screen.MousePointer = 11
        
            Frame1.Visible = False
            Shape1.Visible = False
            FormBayes.Refresh
            Command5.Enabled = True
            
            FormGauge.Show
            FormGauge.Caption = "Preparazione istogrammi"
            FormGauge!Label6.Caption = "[1/3]"
            FormGauge!Label1.Caption = "0%"
            FormGauge!Gauge1.Value = 0
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            
            If Flag1 = True Then
                Titolo1 = Text1.Text
            End If
            If Flag2 = True Then
                Titolo2 = Text1.Text
            End If
            If Flag3 = True Then
                Titolo3 = Text1.Text
            End If
            If Flag4 = True Then
                Titolo4 = Text1.Text
            End If
            If Flag5 = True Then
                Titolo5 = Text1.Text
            End If
            
            InAscisseBis = Text2.Text
            InOrdinateBis = Text3.Text
            XMI = MaskedEdit1.Text
            XMA = MaskedEdit2.Text
            YMI = MaskedEdit3.Text
            YMA = MaskedEdit4.Text
            MX = YMA
            CL = MaskedEdit5.Text
            Prev = MaskedEdit6.Text

            If Flag1 = True Then
                Label13.Caption = Titolo1
            End If
            If Flag2 = True Then
                Label13.Caption = Titolo2
            End If
            If Flag3 = True Then
                Label13.Caption = Titolo3
            End If
            If Flag4 = True Then
                Label13.Caption = Titolo4
            End If
            If Flag5 = True Then
                Label13.Caption = Titolo5
            End If

            Label14.Caption = InAscisseBis
            Label15.Caption = InOrdinateBis

            For I = 1 To CL
                CCS(I) = 0: CCM(I) = 0
            Next I

' suddivisione in classi dei sani
            LW = XMI
            HI = XMA
            AM = (HI - LW) / CL
            
            For N = 1 To Z1
                FormGauge.Label1.Caption = Int(100 * N / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * N / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                R = CL
                If X(N) = HI Then GoTo 6111
                R = Int((X(N) - LW) / AM) + 1
6111:
                CCS(R) = CCS(R) + 1
            Next N

            DummyNumS = 0
            NewMX = 0
                For I = 1 To CL
                    CCS(I) = CCS(I) * ((1 / Prev) - 1) * (NumM / NumS)
                    DummyNumS = DummyNumS + CCS(I)
                    If CCS(I) > NewMX Then NewMX = CCS(I)
                Next I
            
' suddivisione in classi dei malati
            LW = XMI
            HI = XMA
            AM = (HI - LW) / CL
            
            FormGauge.Label6.Caption = "[2/3]"
            For N = 1 To Z2
                FormGauge.Label1.Caption = Int(100 * N / Z2) & "%"
                FormGauge!Gauge1.Value = Int(100 * N / Z2)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                R = CL
                If Y(N) = HI Then GoTo 61111
                R = Int((Y(N) - LW) / AM) + 1
61111:
                CCM(R) = CCM(R) + 1
            Next N

                DummyNumM = 0
                For I = 1 To CL
                    DummyNumM = DummyNumM + CCM(I)
                    If CCM(I) > NewMX Then NewMX = CCM(I)
                Next I

            If CDbl(Format(NewMX, "+#####0.0####;-#####0.0####")) > YMA Then
                YMA = NewMX
                MX = NewMX
            End If

' calcola VP, VN, FP, FN

            If MedianaM <= MedianaS Then ' cut-off preceduto dal segno <
                SommaS = 0: SommaM = 0
                    Limite(0) = LW
                    VP(0) = 0: FP(0) = 0
                    FN(0) = DummyNumM: VN(0) = DummyNumS
                    Sens(0) = 0: Spec(0) = 1
                    VPPos(0) = 0: VPNeg(0) = 1 - Prev
                For I = 1 To CL
                    FormGauge.Label6.Caption = "[3/3]"
                    FormGauge.Label1.Caption = Int(100 * I / CL) & "%"
                    FormGauge!Gauge1.Value = Int(100 * I / CL)
                    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                    Limite(I) = LW + (I * ((HI - LW) / CL))
                    SommaS = SommaS + CCS(I)
                    SommaM = SommaM + CCM(I)
                    VP(I) = SommaM: FN(I) = DummyNumM - SommaM
                    FP(I) = SommaS: VN(I) = DummyNumS - SommaS
                    Sens(I) = VP(I) / (VP(I) + FN(I))
                    Spec(I) = VN(I) / (VN(I) + FP(I))
                        If VP(I) + FP(I) <> 0 Then
                            VPPos(I) = VP(I) / (VP(I) + FP(I))
                        Else
                            VPPos(I) = 0
                        End If
                        If VN(I) + FN(I) <> 0 Then
                            VPNeg(I) = VN(I) / (VN(I) + FN(I))
                        Else
                            VPNeg(I) = 0
                        End If
                Next I
            Else ' cut-off preceduto dal segno >
                SommaS = 0: SommaM = 0
                    Limite(0) = LW
                    VP(0) = DummyNumM: FP(0) = DummyNumS
                    FN(0) = 0: VN(0) = 0
                    Sens(0) = 1: Spec(0) = 0
                    VPPos(0) = Prev: VPNeg(0) = 0
                For I = 1 To CL
                    FormGauge.Label6.Caption = "[3/3]"
                    FormGauge.Label1.Caption = Int(100 * I / CL) & "%"
                    FormGauge!Gauge1.Value = Int(100 * I / CL)
                    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                    Limite(I) = LW + (I * ((HI - LW) / CL))
                    SommaS = SommaS + CCS(I)
                    SommaM = SommaM + CCM(I)
                    VP(I) = DummyNumM - SommaM: FN(I) = SommaM
                    FP(I) = DummyNumS - SommaS: VN(I) = SommaS
                    Sens(I) = VP(I) / (VP(I) + FN(I))
                    Spec(I) = VN(I) / (VN(I) + FP(I))
                        If VP(I) + FP(I) <> 0 Then
                            VPPos(I) = VP(I) / (VP(I) + FP(I))
                        Else
                            VPPos(I) = 0
                        End If
                        If VN(I) + FN(I) <> 0 Then
                            VPNeg(I) = VN(I) / (VN(I) + FN(I))
                        Else
                            VPNeg(I) = 0
                        End If
                Next I
            End If
            
            Unload FormGauge

'    If FlagGrafico = True Then
        If Flag1 = True Then
            IstogrammaDellaDistribuzione_Click
        End If
        If Flag2 = True Then
            GraficaSensibilita_Click
        End If
        If Flag3 = True Then
            GraficaValorePredittivo_Click
        End If
        If Flag4 = True Then
            GraficaEfficienzaDiagnostica_Click
        End If
        If Flag5 = True Then
            GraficaROC_Click
        End If
'    End If
        
            Screen.MousePointer = 1

        End If

Bayes:
        MaskedEdit1.Text = Format(XMI_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit2.Text = Format(XMA_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit3.Text = Format(YMI_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit4.Text = Format(YMA_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit5.Text = Format(CL, "##")
        MaskedEdit6.Text = Format(Prev_Default, "0.0##")

End Sub

Private Sub Command4_Click()
    
    Screen.MousePointer = 11
    
'    If FlagGrafico = True Then
        If Flag1 = True Then
            IstogrammaDellaDistribuzione_Click
        End If
        If Flag2 = True Then
            GraficaSensibilita_Click
        End If
        If Flag3 = True Then
            GraficaValorePredittivo_Click
        End If
        If Flag4 = True Then
            GraficaEfficienzaDiagnostica_Click
        End If
        If Flag5 = True Then
            GraficaROC_Click
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
    MaskedEdit5.Text = Format(CL_Default, "##")
    MaskedEdit6.Text = Format(Prev_Default, "0.0##")

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
    Unload FormBayes
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
End Sub

Private Sub FileStampa_Click()
    On Error GoTo ErroreFileStampaBaye

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

ErroreFileStampaBaye:
    Exit Sub

End Sub

Private Sub Form_Load()
    On Error GoTo GestioneErroriBayes
    ReDim X(NUM_RIGHE - 1) As Double, Y(NUM_RIGHE - 1) As Double
    ReDim CCS(40), DDS(40), CCM(40), DDM(40)
    ReDim VP(0 To 40), VN(0 To 40), FP(0 To 40), FN(0 To 40)
    ReDim Limite(0 To 40), Sens(0 To 40), Spec(0 To 40)
    ReDim VPPos(0 To 40), VPNeg(0 To 40)
    Dim D, I, J, L, M1, M2, ValoreInizialeD
    Dim XX, YY, N, R, Dummy As Double
    Dim SommaS, SommaM, FineDatiX, FineDatiY

    Screen.MousePointer = 11
    AnnullatoCaricamentoForm = False
    FlagGrafico = False
    FlagStatistica = False
    FlagSfondoGrigio = True
    FlagDatiAColori = True
    Flag1 = False: Flag2 = False: Flag3 = False: Flag4 = False: Flag5 = False
    Picture2.Visible = False
    
' centra il form
    FormBayes.Top = Screen.Height / 2 - Height / 2
    FormBayes.Left = Screen.Width / 2 - Width / 2
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
    Titolo1 = "Istogrammi delle distribuzioni"
    Titolo2 = "Sensibilità e specificità del test"
    Titolo3 = "Valore predittivo del test"
    Titolo4 = "Efficienza diagnostica del test"
    Titolo5 = "Curva ROC del test"
    InAscisseBis = "Risultato del test"
    InOrdinateBis = "Frequenza assoluta"
' carica i dati nei vettori X() e Y()
    FormGauge.Show
    FormGauge.Caption = "Statistica bayesiana"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    FineDatiX = False
    FineDatiY = False
    Z1 = 0
    Z2 = 0
    
        FormGauge!Label6.Caption = "[1/8]"
        For Riga = 1 To NUM_RIGHE - 1
            If 100 * Riga / (NUM_RIGHE - 1) = Int(100 * Riga / (NUM_RIGHE - 1)) Then
                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If
        FormPrincipale!Grid1.Row = Riga
        
        If FineDatiX = False Then
            FormPrincipale!Grid1.Col = ColonnaX
            If FormPrincipale!Grid1.Text <> "" Then
                Z1 = Z1 + 1
                X(Riga) = CDbl(-(-FormPrincipale!Grid1.Text))
            Else
                FineDatiX = True
            End If
        End If
        If FineDatiY = False Then
            FormPrincipale!Grid1.Col = ColonnaY
            If FormPrincipale!Grid1.Text <> "" Then
                Z2 = Z2 + 1
                Y(Riga) = CDbl(-(-FormPrincipale!Grid1.Text))
            Else
                FineDatiY = True
            End If
        End If
        Next Riga
    
    ReDim PercX(Z1) As Double, PercY(Z2) As Double

' controlli finali sui dati
' devono esserci almeno due dati
    If Z1 >= 2 And Z2 >= 2 Then
' inizio ordinamento dati X (sani)
pcx330:
    FormGauge.Label6.Caption = "[2/8]"
    D = 1
pcx340:
    D = 2 * D: If D <= Z1 Then GoTo pcx340
pcx350:
    D = Int((D - 1) / 2): If D = 0 Then GoTo pcx420
pcx360:
    For I = 1 To Z1 - D
pcx370:
        J = I
pcx380:
        L = J + D
        If X(J) <= X(L) Then GoTo pcx400
                    
pcx390:
            XX = X(J): X(J) = X(L): X(L) = XX

            If I > ValoreInizialeD And 100 * I / Z1 >= 1 Then
                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If
            J = J - D: If J > 0 Then GoTo pcx380
pcx400:
    Next I
pcx410:
    ValoreInizialeD = Z1 - D
    GoTo pcx350
pcx420:
' fine ordinamento dati
                
' calcola percentili X (sani)
            FormGauge.Label6.Caption = "[3/8]"
            For I = 1 To Z1
                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                M1 = 0: M2 = 0
pex840:
                    For J = 1 To Z1
pex850:
                       If X(J) > X(I) Then GoTo pex890
pex860:
                       If X(J) = X(I) Then GoTo pex880
pex870:
                       M1 = M1 + 1: GoTo pex890
pex880:
                       M2 = M2 + 1
pex890:
                    Next J
pex900:
                PercX(I) = M1 + 0.5 * M2 + 0.5
            Next I

            For I = 1 To Z1
                PercX(I) = Format(100 * PercX(I) / (Z1 + 1), "##0.0####")
            Next I
' calcola mediana, primo e terzo quartile X (sani)
        Dummy = Format(50, "##0.0")
        For I = 1 To Z1 - 1
            If Dummy = PercX(I) Then
                MedianaS = X(I)
                Exit For
            End If
            If Dummy = PercX(I + 1) Then
                MedianaS = X(I + 1)
                Exit For
            End If
            If Dummy > PercX(I) And Dummy < PercX(I + 1) Then
                MedianaS = X(I) + ((X(I + 1) - X(I)) * (Dummy - PercX(I)) / (PercX(I + 1) - PercX(I)))
                Exit For
            End If
        Next I
        Dummy = Format(25, "##0.0")
        For I = 1 To Z1 - 1
            If Dummy = PercX(I) Then
                PrimoQuartS = X(I)
                Exit For
            End If
            If Dummy = PercX(I + 1) Then
                PrimoQuartS = X(I + 1)
                Exit For
            End If
            If Dummy > PercX(I) And Dummy < PercX(I + 1) Then
                PrimoQuartS = X(I) + ((X(I + 1) - X(I)) * (Dummy - PercX(I)) / (PercX(I + 1) - PercX(I)))
                Exit For
            End If
        Next I
        Dummy = Format(75, "##0.0")
        For I = 1 To Z1 - 1
            If Dummy = PercX(I) Then
                TerzoQuartS = X(I)
                Exit For
            End If
            If Dummy = PercX(I + 1) Then
                TerzoQuartS = X(I + 1)
                Exit For
            End If
            If Dummy > PercX(I) And Dummy < PercX(I + 1) Then
                TerzoQuartS = X(I) + ((X(I + 1) - X(I)) * (Dummy - PercX(I)) / (PercX(I + 1) - PercX(I)))
                Exit For
            End If
        Next I
' inizio ordinamento dati Y (malati)
pcy330:
    FormGauge.Label6.Caption = "[4/8]"
    D = 1
pcy340:
    D = 2 * D: If D <= Z2 Then GoTo pcy340
pcy350:
    D = Int((D - 1) / 2): If D = 0 Then GoTo pcy420
pcy360:
    For I = 1 To Z2 - D
pcy370:
        J = I
pcy380:
        L = J + D
        If Y(J) <= Y(L) Then GoTo pcy400
                    
pcy390:
            YY = Y(J): Y(J) = Y(L): Y(L) = YY

            If I > ValoreInizialeD And 100 * I / Z2 >= 1 Then
                FormGauge.Label1.Caption = Int(100 * I / Z2) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z2)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If
            J = J - D: If J > 0 Then GoTo pcy380
pcy400:
    Next I
pcy410:
    ValoreInizialeD = Z2 - D
    GoTo pcy350
pcy420:
' fine ordinamento dati
                
' calcola percentili Y (malati)
            FormGauge.Label6.Caption = "[5/8]"
            For I = 1 To Z2
                FormGauge.Label1.Caption = Int(100 * I / Z2) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z2)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                M1 = 0: M2 = 0
pey840:
                    For J = 1 To Z2
pey850:
                       If Y(J) > Y(I) Then GoTo pey890
pey860:
                       If Y(J) = Y(I) Then GoTo pey880
pey870:
                       M1 = M1 + 1: GoTo pey890
pey880:
                       M2 = M2 + 1
pey890:
                    Next J
pey900:
                PercY(I) = M1 + 0.5 * M2 + 0.5
            Next I

            For I = 1 To Z2
                PercY(I) = Format(100 * PercY(I) / (Z2 + 1), "##0.0")
            Next I
' calcola mediana, primo e terzo quartile Y (malati)
        Dummy = Format(50, "##0.0")
        For I = 1 To Z2 - 1
            If Dummy = PercY(I) Then
                MedianaM = Y(I)
                Exit For
            End If
            If Dummy = PercY(I + 1) Then
                MedianaM = Y(I + 1)
                Exit For
            End If
            If Dummy > PercY(I) And Dummy < PercY(I + 1) Then
                MedianaM = Y(I) + ((Y(I + 1) - Y(I)) * (Dummy - PercY(I)) / (PercY(I + 1) - PercY(I)))
                Exit For
            End If
        Next I
        Dummy = Format(25, "##0.0")
        For I = 1 To Z2 - 1
            If Dummy = PercY(I) Then
                PrimoQuartM = Y(I)
                Exit For
            End If
            If Dummy = PercY(I + 1) Then
                PrimoQuartM = Y(I + 1)
                Exit For
            End If
            If Dummy > PercY(I) And Dummy < PercY(I + 1) Then
                PrimoQuartM = Y(I) + ((Y(I + 1) - Y(I)) * (Dummy - PercY(I)) / (PercY(I + 1) - PercY(I)))
                Exit For
            End If
        Next I
        Dummy = Format(75, "##0.0")
        For I = 1 To Z2 - 1
            If Dummy = PercY(I) Then
                TerzoQuartM = Y(I)
                Exit For
            End If
            If Dummy = PercY(I + 1) Then
                TerzoQuartM = Y(I + 1)
                Exit For
            End If
            If Dummy > PercY(I) And Dummy < PercY(I + 1) Then
                TerzoQuartM = Y(I) + ((Y(I + 1) - Y(I)) * (Dummy - PercY(I)) / (PercY(I + 1) - PercY(I)))
                Exit For
            End If
        Next I

            XMI = X(1): XMA = X(1)
            For I = 2 To Z1
                If X(I) < XMI Then XMI = X(I)
                If X(I) > XMA Then XMA = X(I)
            Next
            For I = 1 To Z2
                If Y(I) < XMI Then XMI = Y(I)
                If Y(I) > XMA Then XMA = Y(I)
            Next
            XMI_Default = XMI
            XMA_Default = XMA

            Prev = Z2 / (Z1 + Z2)
            Prev_Default = Prev
            NumS = Z1
            NumM = Z2

            CL = Int(Sqr(Z1) + 1)
            If Z1 > Z2 Then CL = Int(Sqr(Z2))
            CL_Default = CL

' suddivisione in classi dei sani
            LW = XMI
            HI = XMA
            AM = (HI - LW) / CL
            
            FormGauge.Label6.Caption = "[6/8]"
            For N = 1 To Z1
                FormGauge.Label1.Caption = Int(100 * N / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * N / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                R = CL
                If X(N) = HI Then GoTo 111
                R = Int((X(N) - LW) / AM) + 1
111:
                CCS(R) = CCS(R) + 1
            Next N

            MX = CCS(1)
                For I = 2 To CL
                    If CCS(I) > MX Then MX = CCS(I)
                Next I
            
' suddivisione in classi dei malati
            LW = XMI
            HI = XMA
            AM = (HI - LW) / CL
            
            FormGauge.Label6.Caption = "[7/8]"
            For N = 1 To Z2
                FormGauge.Label1.Caption = Int(100 * N / Z2) & "%"
                FormGauge!Gauge1.Value = Int(100 * N / Z2)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                R = CL
                If Y(N) = HI Then GoTo 1111
                R = Int((Y(N) - LW) / AM) + 1
1111:
                CCM(R) = CCM(R) + 1
            Next N

                For I = 1 To CL
                    If CCM(I) > MX Then MX = CCM(I)
                Next I

            YMI = 0
            YMA = MX
            YMI_Default = YMI
            YMA_Default = YMA

' calcola VP, VN, FP, FN

            If MedianaM <= MedianaS Then ' cut-off preceduto dal segno <
                SommaS = 0: SommaM = 0
                    Limite(0) = LW
                    VP(0) = 0: FP(0) = 0
                    FN(0) = NumM: VN(0) = NumS
                    Sens(0) = 0: Spec(0) = 1
                    VPPos(0) = 0: VPNeg(0) = 1 - Prev
                For I = 1 To CL
                    FormGauge.Label6.Caption = "[8/8]"
                    FormGauge.Label1.Caption = Int(100 * I / CL) & "%"
                    FormGauge!Gauge1.Value = Int(100 * I / CL)
                    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                    Limite(I) = LW + (I * ((HI - LW) / CL))
                    SommaS = SommaS + CCS(I)
                    SommaM = SommaM + CCM(I)
                    VP(I) = SommaM: FN(I) = NumM - SommaM
                    FP(I) = SommaS: VN(I) = NumS - SommaS
                    Sens(I) = VP(I) / (VP(I) + FN(I))
                    Spec(I) = VN(I) / (VN(I) + FP(I))
                        If VP(I) + FP(I) <> 0 Then
                            VPPos(I) = VP(I) / (VP(I) + FP(I))
                        Else
                            VPPos(I) = 0
                        End If
                        If VN(I) + FN(I) <> 0 Then
                            VPNeg(I) = VN(I) / (VN(I) + FN(I))
                        Else
                            VPNeg(I) = 0
                        End If
                Next I
            Else ' cut-off preceduto dal segno >
                SommaS = 0: SommaM = 0
                    Limite(0) = LW
                    VP(0) = NumM: FP(0) = NumS
                    FN(0) = 0: VN(0) = 0
                    Sens(0) = 1: Spec(0) = 0
                    VPPos(0) = Prev: VPNeg(0) = 0
                For I = 1 To CL
                    FormGauge.Label6.Caption = "[8/8]"
                    FormGauge.Label1.Caption = Int(100 * I / CL) & "%"
                    FormGauge!Gauge1.Value = Int(100 * I / CL)
                    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                    Limite(I) = LW + (I * ((HI - LW) / CL))
                    SommaS = SommaS + CCS(I)
                    SommaM = SommaM + CCM(I)
                    VP(I) = NumM - SommaM: FN(I) = SommaM
                    FP(I) = NumS - SommaS: VN(I) = SommaS
                    Sens(I) = VP(I) / (VP(I) + FN(I))
                    Spec(I) = VN(I) / (VN(I) + FP(I))
                        If VP(I) + FP(I) <> 0 Then
                            VPPos(I) = VP(I) / (VP(I) + FP(I))
                        Else
                            VPPos(I) = 0
                        End If
                        If VN(I) + FN(I) <> 0 Then
                            VPNeg(I) = VN(I) / (VN(I) + FN(I))
                        Else
                            VPNeg(I) = 0
                        End If
                Next I
            End If
            
            Unload FormGauge
            Text4.Visible = True
            Shape1.Visible = False
            Frame1.Visible = False
            Picture1.Visible = False
            Picture2.Visible = False
            Label26.Visible = False
            FormPrincipale!Grid1.HighLight = True
            RipristinaSchermo
            FormBayes.Show
            ' FormPrincipale.Hide
            StatisticaValorePredittivo_Click
            Screen.MousePointer = 1
            Exit Sub
    Else
        Unload FormGauge
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        MsgBox "Per la statistica bayesiana" + Chr(10) + "sono necessari almeno due dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Statistica bayesiana non eseguibile"
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
    End If

    Exit Sub

GestioneErroriBayes:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente: riprovare" + Chr(10) + "con un numero inferiore di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Statistica bayesiana non eseguibile"
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
        Exit Sub

End Sub

Private Sub Form_Resize()
    
    Text4.Width = ScaleWidth
    Text4.Height = ScaleHeight

        If Flag1 = True Then
            IstogrammaDellaDistribuzione_Click
        End If
        If Flag2 = True Then
            GraficaSensibilita_Click
        End If
        If Flag3 = True Then
            GraficaValorePredittivo_Click
        End If
        If Flag4 = True Then
            GraficaEfficienzaDiagnostica_Click
        End If
        If Flag5 = True Then
            GraficaROC_Click
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
            IstogrammaDellaDistribuzione_Click
        End If
        If Flag2 = True Then
            GraficaSensibilita_Click
        End If
        If Flag3 = True Then
            GraficaValorePredittivo_Click
        End If
        If Flag4 = True Then
            GraficaEfficienzaDiagnostica_Click
        End If
        If Flag5 = True Then
            GraficaROC_Click
        End If

    End If

End Sub

Private Sub GraficaEfficienzaDiagnostica_Click()
    Dim I, XX, YIN, YSU, FX, FY, CX, CY
    Dim X1, X2, X3, X4, X5
    Dim Y1, Y2, Y3, Y4, Y5
    Dim FF, LargCol, XF, XI, DummyUno, DummyDue

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
    FormBayes.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
            Label26.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
            Label26.BackColor = &H80000005
        End If
    Picture1.Cls
    FormBayes.ScaleWidth = 800
    FormBayes.ScaleHeight = 480
    FormBayes.Refresh

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
    
    Picture1.Line (0, 170)-(459, 170)
    
    LargCol = 459 / CL
    
    FF = 170 / MX
    For I = 1 To CL
        DDS(I) = Int(CCS(I) * FF)
    Next I
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
        Picture1.Line (XI, 170)-Step(LargCol, -DDS(I)), , B
'        Picture1.Line (XI, 340)-(XI, 340 - DD(I))
'        Picture1.Line (XI, 340 - DD(I))-(XI + LARGCOL, 340 - DD(I))
'        Picture1.Line (XI + LARGCOL, 340 - DD(I))-(XI + LARGCOL, 340)
'        Picture1.Line (XI + LARGCOL, 340)-(XI, 340)
        I = I + 1
    Next XI
    
    FF = 169 / MX
    For I = 1 To CL
        DDM(I) = Int(CCM(I) * FF)
    Next I
    XF = 0
    I = 1
        If FlagDatiAColori = True Then
            Picture1.FillStyle = 0
            Picture1.FillColor = QBColor(11)
        Else
            Picture1.FillStyle = 3
            Picture1.FillColor = QBColor(0)
        End If
'    For XI = XF To XF + LARGCOL * (CL - 1) Step LARGCOL
    For XI = XF To XF + LargCol * CL Step LargCol
        Picture1.Line (XI, 170)-Step(LargCol, DDM(I)), , B
'        Picture1.Line (XI, 340)-(XI, 340 - DD(I))
'        Picture1.Line (XI, 340 - DD(I))-(XI + LARGCOL, 340 - DD(I))
'        Picture1.Line (XI + LARGCOL, 340 - DD(I))-(XI + LARGCOL, 340)
'        Picture1.Line (XI + LARGCOL, 340)-(XI, 340)
        I = I + 1
    Next XI

        XF = 0
        I = 0
        Picture1.DrawStyle = 2
        For XI = XF To XF + LargCol * CL Step LargCol
            DummyUno = Sens(I) - (1 - Spec(I))
            If DummyUno < 0 Then DummyUno = 0
            DummyDue = Sens(I + 1) - (1 - Spec(I + 1))
            If DummyDue < 0 Then DummyDue = 0
            If FlagDatiAColori = True Then
                Picture1.Line (XI, 340 - DummyUno * 340)-(XI + LargCol, 340 - DummyDue * 340), QBColor(13)
            Else
                Picture1.Line (XI, 340 - DummyUno * 340)-(XI + LargCol, 340 - DummyDue * 340), QBColor(0)
            End If
        I = I + 1
        Next XI
        
    Y5 = YSU: Y4 = Int((YIN + (YSU - YIN) / 2) + 0.5): Y3 = YIN: Y2 = Int((YIN + (YSU - YIN) / 2) + 0.5): Y1 = YSU
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
    Label13.Caption = Titolo4
    Label14.Caption = InAscisseBis
    Label15.Caption = InOrdinateBis
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
    
    FormBayes.CurrentX = 460: FormBayes.CurrentY = 60
    FormBayes.Print "Probabilità p"
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 408
    FormBayes.Print "0"
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 330
    FormBayes.Print Format(0.25, "0.00")
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 160
    FormBayes.Print Format(0.75, "0.00")
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 78
    FormBayes.Print Format(1, "0.00")
    FormBayes.CurrentX = 580: FormBayes.CurrentY = 177
    FormBayes.Print "Efficienza del test"
    FormBayes.DrawStyle = 2
        If FlagDatiAColori = True Then
            FormBayes.Line (580, 193)-(680, 193), QBColor(13)
        Else
            FormBayes.Line (580, 193)-(680, 193), QBColor(0)
        End If
    
    Label26.Visible = True
    Label26.Left = 574
    Label26.Top = Picture1.Top + (Picture1.Height / 3)
    Label26.Caption = "" & Chr(13)
    Label26.Caption = Label26.Caption & "   SANI" & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   [" & InAscisse & "]" & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   Prevalenza = " & Format(Prev, "0.000") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   MALATI" & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   [" & InOrdinate & "]" & "   " & Chr(13)
    Label26.Top = Picture1.Top + (Picture1.Height / 2) - (Label26.Height / 2)
    Flag1 = False: Flag2 = False: Flag3 = False: Flag4 = True: Flag5 = False
    Picture1.DrawStyle = 0
    FormBayes.DrawStyle = 0
    Screen.MousePointer = 1

End Sub

Private Sub GraficaROC_Click()
    Dim I, XX, YIN, YSU, FX, FY, CX, CY
    Dim X1, X2, X3, X4, X5
    Dim Y1, Y2, Y3, Y4, Y5
    Dim FF, LargCol, XF, XI, DummyUno, DummyDue

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
    FormBayes.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
            Label26.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
            Label26.BackColor = &H80000005
        End If
    Picture1.Cls
    FormBayes.ScaleWidth = 800
    FormBayes.ScaleHeight = 480
    FormBayes.Refresh

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
    
    Picture1.Line (0, 170)-(459, 170)
    
    LargCol = 459 / CL
    
    FF = 170 / MX
    For I = 1 To CL
        DDS(I) = Int(CCS(I) * FF)
    Next I
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
        Picture1.Line (XI, 170)-Step(LargCol, -DDS(I)), , B
'        Picture1.Line (XI, 340)-(XI, 340 - DD(I))
'        Picture1.Line (XI, 340 - DD(I))-(XI + LARGCOL, 340 - DD(I))
'        Picture1.Line (XI + LARGCOL, 340 - DD(I))-(XI + LARGCOL, 340)
'        Picture1.Line (XI + LARGCOL, 340)-(XI, 340)
        I = I + 1
    Next XI
    
    FF = 169 / MX
    For I = 1 To CL
        DDM(I) = Int(CCM(I) * FF)
    Next I
    XF = 0
    I = 1
        If FlagDatiAColori = True Then
            Picture1.FillStyle = 0
            Picture1.FillColor = QBColor(11)
        Else
            Picture1.FillStyle = 3
            Picture1.FillColor = QBColor(0)
        End If
'    For XI = XF To XF + LARGCOL * (CL - 1) Step LARGCOL
    For XI = XF To XF + LargCol * CL Step LargCol
        Picture1.Line (XI, 170)-Step(LargCol, DDM(I)), , B
'        Picture1.Line (XI, 340)-(XI, 340 - DD(I))
'        Picture1.Line (XI, 340 - DD(I))-(XI + LARGCOL, 340 - DD(I))
'        Picture1.Line (XI + LARGCOL, 340 - DD(I))-(XI + LARGCOL, 340)
'        Picture1.Line (XI + LARGCOL, 340)-(XI, 340)
        I = I + 1
    Next XI

        Picture1.DrawStyle = 2
        For I = 0 To CL
            If FlagDatiAColori = True Then
'                Picture1.Line (460 * Sens(I), 340 - (1 - Spec(I)) * 340)-(460 * Sens(I + 1), 340 - (1 - Spec(I + 1)) * 340), QBColor(13)
                Picture1.Line (460 * (1 - Spec(I)), 340 - Sens(I) * 340)-(460 * (1 - Spec(I + 1)), 340 - Sens(I + 1) * 340), QBColor(13)
            Else
'                Picture1.Line (460 * Sens(I), 340 - (1 - Spec(I)) * 340)-(460 * Sens(I + 1), 340 - (1 - Spec(I + 1)) * 340), QBColor(0)
                Picture1.Line (460 * (1 - Spec(I)), 340 - Sens(I) * 340)-(460 * (1 - Spec(I + 1)), 340 - Sens(I + 1) * 340), QBColor(0)
            End If
        Next I
        
    Y5 = YSU: Y4 = Int((YIN + (YSU - YIN) / 2) + 0.5): Y3 = YIN: Y2 = Int((YIN + (YSU - YIN) / 2) + 0.5): Y1 = YSU
    X1 = XMI: X2 = XMI + (XMA - XMI) / 4: X3 = XMI + (XMA - XMI) / 2: X4 = XMA - (XMA - XMI) / 4: X5 = XMA
    
    Label13.Left = 330 - (Label13.Width / 2): Label13.Top = 27
    Label14.Left = 330 - (Label14.Width / 2): Label14.Top = 453
    Label15.Left = 100: Label15.Top = 61
    Label16.Left = 100: Label16.Top = 437
    Label17.Left = 215 - (Label17.Width / 2): Label17.Top = 437
    Label18.Left = 330 - (Label18.Width / 2): Label18.Top = 437
    Label19.Left = 445 - (Label19.Width / 2): Label19.Top = 437
    Label20.Left = 560 - Label20.Width: Label20.Top = 437
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
    Label13.Caption = Titolo5
    Label14.Caption = InAscisseBis
    Label15.Caption = InOrdinateBis
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
    
    FormBayes.CurrentX = 480: FormBayes.CurrentY = 60
    FormBayes.Print "Sensibilità"
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 408
    FormBayes.Print "0"
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 330
    FormBayes.Print Format(0.25, "0.00")
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 160
    FormBayes.Print Format(0.75, "0.00")
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 78
    FormBayes.Print Format(1, "0.00")
    
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 423
    FormBayes.Print "(1 - Specificità)"
    FormBayes.CurrentX = 100: FormBayes.CurrentY = 423
    FormBayes.Print "0"
    FormBayes.CurrentX = 195: FormBayes.CurrentY = 423
    FormBayes.Print Format(0.25, "0.00")
    FormBayes.CurrentX = 310: FormBayes.CurrentY = 423
    FormBayes.Print Format(0.5, "0.00")
    FormBayes.CurrentX = 425: FormBayes.CurrentY = 423
    FormBayes.Print Format(0.75, "0.00")
    FormBayes.CurrentX = 525: FormBayes.CurrentY = 423
    FormBayes.Print Format(1, "0.00")
    
    FormBayes.CurrentX = 580: FormBayes.CurrentY = 177
    FormBayes.Print "Curva ROC del test"
    FormBayes.DrawStyle = 2
        If FlagDatiAColori = True Then
            FormBayes.Line (580, 193)-(680, 193), QBColor(13)
        Else
            FormBayes.Line (580, 193)-(680, 193), QBColor(0)
        End If
    
    Label26.Visible = True
    Label26.Left = 574
    Label26.Top = Picture1.Top + (Picture1.Height / 3)
    Label26.Caption = "" & Chr(13)
    Label26.Caption = Label26.Caption & "   SANI" & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   [" & InAscisse & "]" & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   Prevalenza = " & Format(Prev, "0.000") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   MALATI" & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   [" & InOrdinate & "]" & "   " & Chr(13)
    Label26.Top = Picture1.Top + (Picture1.Height / 2) - (Label26.Height / 2)
    Flag1 = False: Flag2 = False: Flag3 = False: Flag4 = False: Flag5 = True
    Picture1.DrawStyle = 0
    FormBayes.DrawStyle = 0
    Screen.MousePointer = 1

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
    If Flag4 = True Then
        Text1.Text = Titolo4
    End If
    If Flag5 = True Then
        Text1.Text = Titolo5
    End If
    
    Text2.Text = InAscisseBis
    Text3.Text = InOrdinateBis
    MaskedEdit1.Text = Format(XMI, "+#####0.0####;-#####0.0####")
    MaskedEdit2.Text = Format(XMA, "+#####0.0####;-#####0.0####")
    MaskedEdit3.Text = Format(YMI, "+#####0.0####;-#####0.0####")
    MaskedEdit4.Text = Format(YMA, "+#####0.0####;-#####0.0####")
    MaskedEdit5.Text = Format(CL, "##")
    MaskedEdit6.Text = Format(Prev, "0.0##")
    Command4.SetFocus

End Sub

Private Sub GraficaSensibilita_Click()
    Dim I, XX, YIN, YSU, FX, FY, CX, CY
    Dim X1, X2, X3, X4, X5
    Dim Y1, Y2, Y3, Y4, Y5
    Dim FF, LargCol, XF, XI

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
    FormBayes.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
            Label26.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
            Label26.BackColor = &H80000005
        End If
    Picture1.Cls
    FormBayes.ScaleWidth = 800
    FormBayes.ScaleHeight = 480
    FormBayes.Refresh

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
    
    Picture1.Line (0, 170)-(459, 170)
    
    LargCol = 459 / CL
    
    FF = 170 / MX
    For I = 1 To CL
        DDS(I) = Int(CCS(I) * FF)
    Next I
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
        Picture1.Line (XI, 170)-Step(LargCol, -DDS(I)), , B
'        Picture1.Line (XI, 340)-(XI, 340 - DD(I))
'        Picture1.Line (XI, 340 - DD(I))-(XI + LARGCOL, 340 - DD(I))
'        Picture1.Line (XI + LARGCOL, 340 - DD(I))-(XI + LARGCOL, 340)
'        Picture1.Line (XI + LARGCOL, 340)-(XI, 340)
        I = I + 1
    Next XI
    
    FF = 169 / MX
    For I = 1 To CL
        DDM(I) = Int(CCM(I) * FF)
    Next I
    XF = 0
    I = 1
        If FlagDatiAColori = True Then
            Picture1.FillStyle = 0
            Picture1.FillColor = QBColor(11)
        Else
            Picture1.FillStyle = 3
            Picture1.FillColor = QBColor(0)
        End If
'    For XI = XF To XF + LARGCOL * (CL - 1) Step LARGCOL
    For XI = XF To XF + LargCol * CL Step LargCol
        Picture1.Line (XI, 170)-Step(LargCol, DDM(I)), , B
'        Picture1.Line (XI, 340)-(XI, 340 - DD(I))
'        Picture1.Line (XI, 340 - DD(I))-(XI + LARGCOL, 340 - DD(I))
'        Picture1.Line (XI + LARGCOL, 340 - DD(I))-(XI + LARGCOL, 340)
'        Picture1.Line (XI + LARGCOL, 340)-(XI, 340)
        I = I + 1
    Next XI

        XF = 0
        I = 0
        Picture1.DrawStyle = 2
        For XI = XF To XF + LargCol * CL Step LargCol
            If FlagDatiAColori = True Then
                Picture1.Line (XI, 340 - Sens(I) * 340)-(XI + LargCol, 340 - Sens(I + 1) * 340), QBColor(9)
            Else
                Picture1.Line (XI, 340 - Sens(I) * 340)-(XI + LargCol, 340 - Sens(I + 1) * 340), QBColor(0)
            End If
        I = I + 1
        Next XI
        
        XF = 0
        I = 0
        Picture1.DrawStyle = 4
        For XI = XF To XF + LargCol * CL Step LargCol
            Picture1.DrawStyle = 1
            If FlagDatiAColori = True Then
                Picture1.Line (XI, 340 - Spec(I) * 340)-(XI + LargCol, 340 - Spec(I + 1) * 340), QBColor(2)
            Else
                Picture1.Line (XI, 340 - Spec(I) * 340)-(XI + LargCol, 340 - Spec(I + 1) * 340), QBColor(0)
            End If
        I = I + 1
        Next XI
    

    Y5 = YSU: Y4 = Int((YIN + (YSU - YIN) / 2) + 0.5): Y3 = YIN: Y2 = Int((YIN + (YSU - YIN) / 2) + 0.5): Y1 = YSU
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
    Label13.Caption = Titolo2
    Label14.Caption = InAscisseBis
    Label15.Caption = InOrdinateBis
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
    
    FormBayes.CurrentX = 460: FormBayes.CurrentY = 60
    FormBayes.Print "Probabilità p"
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 408
    FormBayes.Print "0"
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 330
    FormBayes.Print Format(0.25, "0.00")
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 160
    FormBayes.Print Format(0.75, "0.00")
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 78
    FormBayes.Print Format(1, "0.00")
    FormBayes.CurrentX = 580: FormBayes.CurrentY = 177
    FormBayes.Print "Specificità"
    FormBayes.DrawStyle = 4
        If FlagDatiAColori = True Then
            FormBayes.Line (580, 193)-(680, 193), QBColor(2)
        Else
            FormBayes.Line (580, 193)-(680, 193), QBColor(0)
        End If
    FormBayes.CurrentX = 580: FormBayes.CurrentY = 305
    FormBayes.Print "Sensibilità"
    FormBayes.DrawStyle = 2
        If FlagDatiAColori = True Then
            FormBayes.Line (580, 320)-(680, 320), QBColor(9)
        Else
            FormBayes.Line (580, 320)-(680, 320), QBColor(0)
        End If
    
    Label26.Visible = True
    Label26.Left = 574
    Label26.Top = Picture1.Top + (Picture1.Height / 3)
    Label26.Caption = "" & Chr(13)
    Label26.Caption = Label26.Caption & "   SANI" & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   [" & InAscisse & "]" & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   Prevalenza = " & Format(Prev, "0.000") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   MALATI" & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   [" & InOrdinate & "]" & "   " & Chr(13)
    Label26.Top = Picture1.Top + (Picture1.Height / 2) - (Label26.Height / 2)
    Flag1 = False: Flag2 = True: Flag3 = False: Flag4 = False: Flag5 = False
    Picture1.DrawStyle = 0
    FormBayes.DrawStyle = 0
    Screen.MousePointer = 1

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
            IstogrammaDellaDistribuzione_Click
        End If
        If Flag2 = True Then
            GraficaSensibilita_Click
        End If
        If Flag3 = True Then
            GraficaValorePredittivo_Click
        End If
        If Flag4 = True Then
            GraficaEfficienzaDiagnostica_Click
        End If
        If Flag5 = True Then
            GraficaROC_Click
        End If

    End If

End Sub

Private Sub GraficaValorePredittivo_Click()
    Dim I, XX, YIN, YSU, FX, FY, CX, CY
    Dim X1, X2, X3, X4, X5
    Dim Y1, Y2, Y3, Y4, Y5
    Dim FF, LargCol, XF, XI

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
    FormBayes.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
            Label26.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
            Label26.BackColor = &H80000005
        End If
    Picture1.Cls
    FormBayes.ScaleWidth = 800
    FormBayes.ScaleHeight = 480
    FormBayes.Refresh

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
    
    Picture1.Line (0, 170)-(459, 170)
    
    LargCol = 459 / CL
    
    FF = 170 / MX
    For I = 1 To CL
        DDS(I) = Int(CCS(I) * FF)
    Next I
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
        Picture1.Line (XI, 170)-Step(LargCol, -DDS(I)), , B
'        Picture1.Line (XI, 340)-(XI, 340 - DD(I))
'        Picture1.Line (XI, 340 - DD(I))-(XI + LARGCOL, 340 - DD(I))
'        Picture1.Line (XI + LARGCOL, 340 - DD(I))-(XI + LARGCOL, 340)
'        Picture1.Line (XI + LARGCOL, 340)-(XI, 340)
        I = I + 1
    Next XI
    
    FF = 169 / MX
    For I = 1 To CL
        DDM(I) = Int(CCM(I) * FF)
    Next I
    XF = 0
    I = 1
        If FlagDatiAColori = True Then
            Picture1.FillStyle = 0
            Picture1.FillColor = QBColor(11)
        Else
            Picture1.FillStyle = 3
            Picture1.FillColor = QBColor(0)
        End If
'    For XI = XF To XF + LARGCOL * (CL - 1) Step LARGCOL
    For XI = XF To XF + LargCol * CL Step LargCol
        Picture1.Line (XI, 170)-Step(LargCol, DDM(I)), , B
'        Picture1.Line (XI, 340)-(XI, 340 - DD(I))
'        Picture1.Line (XI, 340 - DD(I))-(XI + LARGCOL, 340 - DD(I))
'        Picture1.Line (XI + LARGCOL, 340 - DD(I))-(XI + LARGCOL, 340)
'        Picture1.Line (XI + LARGCOL, 340)-(XI, 340)
        I = I + 1
    Next XI

        XF = 0
        I = 0
        Picture1.DrawStyle = 2
        For XI = XF To XF + LargCol * CL Step LargCol
            If FlagDatiAColori = True Then
                If VPPos(I) <> 0 And VPPos(I + 1) <> 0 Then Picture1.Line (XI, 340 - VPPos(I) * 340)-(XI + LargCol, 340 - VPPos(I + 1) * 340), QBColor(9)
            Else
                If VPPos(I) <> 0 And VPPos(I + 1) <> 0 Then Picture1.Line (XI, 340 - VPPos(I) * 340)-(XI + LargCol, 340 - VPPos(I + 1) * 340), QBColor(0)
            End If
        I = I + 1
        Next XI
        
        XF = 0
        I = 0
        Picture1.DrawStyle = 4
        For XI = XF To XF + LargCol * CL Step LargCol
            Picture1.DrawStyle = 1
            If FlagDatiAColori = True Then
                If VPNeg(I) <> 0 And VPNeg(I + 1) <> 0 Then Picture1.Line (XI, 340 - VPNeg(I) * 340)-(XI + LargCol, 340 - VPNeg(I + 1) * 340), QBColor(2)
            Else
                If VPNeg(I) <> 0 And VPNeg(I + 1) <> 0 Then Picture1.Line (XI, 340 - VPNeg(I) * 340)-(XI + LargCol, 340 - VPNeg(I + 1) * 340), QBColor(0)
            End If
        I = I + 1
        Next XI
    
    Y5 = YSU: Y4 = Int((YIN + (YSU - YIN) / 2) + 0.5): Y3 = YIN: Y2 = Int((YIN + (YSU - YIN) / 2) + 0.5): Y1 = YSU
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
    Label13.Caption = Titolo3
    Label14.Caption = InAscisseBis
    Label15.Caption = InOrdinateBis
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
    
    FormBayes.CurrentX = 460: FormBayes.CurrentY = 60
    FormBayes.Print "Probabilità p"
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 408
    FormBayes.Print "0"
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 330
    FormBayes.Print Format(0.25, "0.00")
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 160
    FormBayes.Print Format(0.75, "0.00")
    FormBayes.CurrentX = 569: FormBayes.CurrentY = 78
    FormBayes.Print Format(1, "0.00")
    FormBayes.CurrentX = 580: FormBayes.CurrentY = 177
    FormBayes.Print "VP test negativo"
    FormBayes.DrawStyle = 4
        If FlagDatiAColori = True Then
            FormBayes.Line (580, 193)-(680, 193), QBColor(2)
        Else
            FormBayes.Line (580, 193)-(680, 193), QBColor(0)
        End If
    FormBayes.CurrentX = 580: FormBayes.CurrentY = 305
    FormBayes.Print "VP test positivo"
    FormBayes.DrawStyle = 2
        If FlagDatiAColori = True Then
            FormBayes.Line (580, 320)-(680, 320), QBColor(9)
        Else
            FormBayes.Line (580, 320)-(680, 320), QBColor(0)
        End If
    
    Label26.Visible = True
    Label26.Left = 574
    Label26.Top = Picture1.Top + (Picture1.Height / 3)
    Label26.Caption = "" & Chr(13)
    Label26.Caption = Label26.Caption & "   SANI" & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   [" & InAscisse & "]" & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   Prevalenza = " & Format(Prev, "0.000") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   MALATI" & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   [" & InOrdinate & "]" & "   " & Chr(13)
    Label26.Top = Picture1.Top + (Picture1.Height / 2) - (Label26.Height / 2)
    Flag1 = False: Flag2 = False: Flag3 = True: Flag4 = False: Flag5 = False
    Picture1.DrawStyle = 0
    FormBayes.DrawStyle = 0
    Screen.MousePointer = 1

End Sub

Private Sub IstogrammaDellaDistribuzione_Click()
    Dim I, XX, YIN, YSU, FX, FY, CX, CY
    Dim X1, X2, X3, X4, X5
    Dim Y1, Y2, Y3, Y4, Y5
    Dim FF, LargCol, XF, XI

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
    FormBayes.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
            Label26.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
            Label26.BackColor = &H80000005
        End If
    Picture1.Cls
    FormBayes.ScaleWidth = 800
    FormBayes.ScaleHeight = 480
    FormBayes.Refresh

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
    
    Picture1.Line (0, 170)-(459, 170)
    
    LargCol = 459 / CL
    
    FF = 170 / MX
    For I = 1 To CL
        DDS(I) = Int(CCS(I) * FF)
    Next I
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
        Picture1.Line (XI, 170)-Step(LargCol, -DDS(I)), , B
'        Picture1.Line (XI, 340)-(XI, 340 - DD(I))
'        Picture1.Line (XI, 340 - DD(I))-(XI + LARGCOL, 340 - DD(I))
'        Picture1.Line (XI + LARGCOL, 340 - DD(I))-(XI + LARGCOL, 340)
'        Picture1.Line (XI + LARGCOL, 340)-(XI, 340)
        I = I + 1
    Next XI
    
    FF = 169 / MX
    For I = 1 To CL
        DDM(I) = Int(CCM(I) * FF)
    Next I
    XF = 0
    I = 1
        If FlagDatiAColori = True Then
            Picture1.FillStyle = 0
            Picture1.FillColor = QBColor(11)
        Else
            Picture1.FillStyle = 3
            Picture1.FillColor = QBColor(0)
        End If
'    For XI = XF To XF + LARGCOL * (CL - 1) Step LARGCOL
    For XI = XF To XF + LargCol * CL Step LargCol
        Picture1.Line (XI, 170)-Step(LargCol, DDM(I)), , B
'        Picture1.Line (XI, 340)-(XI, 340 - DD(I))
'        Picture1.Line (XI, 340 - DD(I))-(XI + LARGCOL, 340 - DD(I))
'        Picture1.Line (XI + LARGCOL, 340 - DD(I))-(XI + LARGCOL, 340)
'        Picture1.Line (XI + LARGCOL, 340)-(XI, 340)
        I = I + 1
    Next XI

    Y5 = YSU: Y4 = Int((YIN + (YSU - YIN) / 2) + 0.5): Y3 = YIN: Y2 = Int((YIN + (YSU - YIN) / 2) + 0.5): Y1 = YSU
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
    Label13.Caption = Titolo1
    Label14.Caption = InAscisseBis
    Label15.Caption = InOrdinateBis
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
    Label26.Caption = Label26.Caption & "   SANI" & Chr(13)
    Label26.Caption = Label26.Caption & "   [" & InAscisse & "]" & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   Dati = " & NumS & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   Minimo = " & Format(X(1), "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   1" & Chr(176) & " quartile = " & Format(PrimoQuartS, "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   2" & Chr(176) & " quartile = " & Format(MedianaS, "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   3" & Chr(176) & " quartile = " & Format(TerzoQuartS, "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   Massimo = " & Format(X(NumS), "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & Chr(13)
    Label26.Caption = Label26.Caption & "   Prevalenza = " & Format(Prev, "0.000") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & Chr(13)
    Label26.Caption = Label26.Caption & "   MALATI" & Chr(13)
    Label26.Caption = Label26.Caption & "   [" & InOrdinate & "]" & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   Dati = " & NumM & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   Minimo = " & Format(Y(1), "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   1" & Chr(176) & " quartile = " & Format(PrimoQuartM, "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   2" & Chr(176) & " quartile = " & Format(MedianaM, "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   3" & Chr(176) & " quartile = " & Format(TerzoQuartM, "#####0.0####") & "   " & Chr(13)
    Label26.Caption = Label26.Caption & "   Massimo = " & Format(Y(NumM), "#####0.0####") & "   " & Chr(13)
    Label26.Top = Picture1.Top + (Picture1.Height / 2) - (Label26.Height / 2)
    Flag1 = True: Flag2 = False: Flag3 = False: Flag4 = False: Flag5 = False
    Screen.MousePointer = 1

End Sub

Private Sub MenuFile_Click()

    FileStampa.Enabled = False
        If FlagGrafico = True Then
            If FormBayes.Width < Screen.Width Then
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

Private Sub StatisticaEfficienzaDiagnostica_Click()
    Dim I, K
    Dim Dummy As Double
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
    FormBayes.Cls
    FormBayes.Refresh
    
        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Statistica bayesiana, efficienza diagnostica del test (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File                      : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella                   : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Primo campione (sani)     : " & InAscisse & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Secondo campione (malati) : " & InOrdinate & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Prevalenza della malattia = " & Format(Prev, "0.000") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
                                                                    
        Text4.Text = Text4.Text & "Positivo se...  Sensibilità   Specificità   Efficienza " & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "--------------  ------------  ------------  ------------" & Chr(13) & Chr(10)

                For I = 0 To CL
                    If MedianaM <= MedianaS Then ' cut-off preceduto dal segno <
                        Text4.Text = Text4.Text & Format("<", "@@")
                    Else ' cut-off preceduto dal segno >
                        Text4.Text = Text4.Text & Format(">", "@@")
                    End If
                    Stringa = Format(Limite(I), "#####0.0####"): If Stringa = "" Then Stringa = "0"
                    Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@")
                    Stringa = Format(Sens(I), "#####0.00000"): If Stringa = "" Then Stringa = "0"
                    Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@@@")
                    Stringa = Format(Spec(I), "#####0.00000"): If Stringa = "" Then Stringa = "0"
                    Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@@@")
                    Dummy = Sens(I) - (1 - Spec(I))
                    If Dummy < 0 Then Dummy = 0
                    Stringa = Format(Dummy, "#####0.00000"): If Stringa = "" Then Stringa = "0"
                    Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@@@")
                    Text4.Text = Text4.Text & Chr(13) & Chr(10)
                Next I
            
        Text4.Text = Text4.Text & "--------------  ------------  ------------  ------------" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Statistica bayesiana, efficienza diagnostica del test (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        FlagGrafico = False
        Flag1 = False: Flag2 = False: Flag3 = False: Flag4 = False: Flag5 = False
        Screen.MousePointer = 1

End Sub

Private Sub StatisticaStatisticheDelleClassi_Click()
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
    FormBayes.Cls
    FormBayes.Refresh
    
        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Statistica bayesiana, statistiche delle classi (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File                      : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella                   : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Primo campione (sani)     : " & InAscisse & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Secondo campione (malati) : " & InOrdinate & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Prevalenza della malattia = " & Format(Prev, "0.000") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
                                                                    
        Text4.Text = Text4.Text & "                                         Sani           Malati  " & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                                    --------------  --------------" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Classe      da...         a...        VN      FP      VP      FN " & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "------  ------------  ------------  ------  ------  ------  ------" & Chr(13) & Chr(10)

        For I = 1 To CL
            Stringa = Format(I, "##")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@") & "  "
            Stringa = Format(LW + ((I - 1) * AM), "#####0.0####")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@") & "  "
            Stringa = Format(LW + (I * AM), "#####0.0####")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@") & "  "
            Stringa = Format(VN(I), "###0"): If Stringa = "" Then Stringa = "0"
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@") & "  "
            Stringa = Format(FP(I), "###0"): If Stringa = "" Then Stringa = "0"
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@") & "  "
            Stringa = Format(VP(I), "###0"): If Stringa = "" Then Stringa = "0"
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@") & "  "
            Stringa = Format(FN(I), "###0"): If Stringa = "" Then Stringa = "0"
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@") & "  "
            Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Next I
        
        Text4.Text = Text4.Text & "------  ------------  ------------  ------  ------  ------  ------" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "V = veri, F = falsi, N = negativi, P = positivi" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Statistica bayesiana, statistiche delle classi (*fine*)" & Chr(13) & Chr(10)
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        FlagGrafico = False
        Flag1 = False: Flag2 = False: Flag3 = False: Flag4 = False: Flag5 = False
        Screen.MousePointer = 1

End Sub

Private Sub StatisticaStatisticheElementari_Click()
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
    FormBayes.Cls
    FormBayes.Refresh
    
        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Statistica bayesiana, statistiche elementari (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File                      : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella                   : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Primo campione (sani)     : " & InAscisse & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Secondo campione (malati) : " & InOrdinate & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Sani   : numero dei dati  = " & Z1 & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "         minimo           = " & Format(X(1), "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "         primo quartile   = " & Format(PrimoQuartS, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "         secondo quartile = " & Format(MedianaS, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "         terzo quartile   = " & Format(TerzoQuartS, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "         massimo          = " & Format(X(Z1), "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Malati : numero dei dati  = " & Z2 & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "         minimo           = " & Format(Y(1), "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "         primo quartile   = " & Format(PrimoQuartM, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "         secondo quartile = " & Format(MedianaM, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "         terzo quartile   = " & Format(TerzoQuartM, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "         massimo          = " & Format(Y(Z2), "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Prevalenza della malattia = " & Format(Prev, "0.000") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Statistica bayesiana, statistiche elementari (*fine*)" & Chr(13) & Chr(10)
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        FlagGrafico = False
        Flag1 = False: Flag2 = False: Flag3 = False: Flag4 = False: Flag5 = False
        Screen.MousePointer = 1

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
    FormBayes.Cls
    FormBayes.Refresh
    
        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Statistica bayesiana, suddivisione dei dati in classi (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File                      : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella                   : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Primo campione (sani)     : " & InAscisse & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Secondo campione (malati) : " & InOrdinate & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Prevalenza della malattia = " & Format(Prev, "0.000") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
                                                                    
        Text4.Text = Text4.Text & "Classe      da...         a...      Sani    Malati" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "------  ------------  ------------  ------  ------" & Chr(13) & Chr(10)

        For I = 1 To CL
            Stringa = Format(I, "##")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@") & "  "
            Stringa = Format(LW + ((I - 1) * AM), "#####0.0####")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@") & "  "
            Stringa = Format(LW + (I * AM), "#####0.0####")
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@") & "  "
            Stringa = Format(CCS(I), "###0"): If Stringa = "" Then Stringa = "0"
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@") & "  "
            Stringa = Format(CCM(I), "###0"): If Stringa = "" Then Stringa = "0"
            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@") & "  "
            Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Next I
        
        Text4.Text = Text4.Text & "------  ------------  ------------  ------  ------" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Statistica bayesiana, suddivisione dei dati in classi (*fine*)" & Chr(13) & Chr(10)
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        FlagGrafico = False
        Flag1 = False: Flag2 = False: Flag3 = False: Flag4 = False: Flag5 = False
        Screen.MousePointer = 1

End Sub

Private Sub StatisticaValorePredittivo_Click()
    Dim I, K
    Dim Dummy As Double
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
    FormBayes.Cls
    FormBayes.Refresh
    
        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Statistica bayesiana, valore predittivo del test (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File                      : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella                   : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Primo campione (sani)     : " & InAscisse & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Secondo campione (malati) : " & InOrdinate & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Prevalenza della malattia = " & Format(Prev, "0.000") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
                                                                    
        Text4.Text = Text4.Text & "                                             Valore predittivo del test " & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                                            ----------------------------" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Positivo se...  Sensibilità   Specificità   test positivo  test negativo" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "--------------  ------------  ------------  -------------  -------------" & Chr(13) & Chr(10)

                For I = 0 To CL
                    If MedianaM <= MedianaS Then ' cut-off preceduto dal segno <
                        Text4.Text = Text4.Text & Format("<", "@@")
                    Else ' cut-off preceduto dal segno >
                        Text4.Text = Text4.Text & Format(">", "@@")
                    End If
                    Stringa = Format(Limite(I), "#####0.0####"): If Stringa = "" Then Stringa = "0"
                    Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@")
                    Stringa = Format(Sens(I), "#####0.00000"): If Stringa = "" Then Stringa = "0"
                    Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@@@")
                    Stringa = Format(Spec(I), "#####0.00000"): If Stringa = "" Then Stringa = "0"
                    Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@@@")
                        If VPPos(I) <> 0 Then
                            Stringa = Format(VPPos(I), "#####0.00000")
                            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@@@@")
                        Else
                            Stringa = "###"
                            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@@@@")
                        End If
                        If VPNeg(I) <> 0 Then
                            Stringa = Format(VPNeg(I), "#####0.00000")
                            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@@@@")
                        Else
                            Stringa = "###"
                            Text4.Text = Text4.Text & Format(Stringa, "@@@@@@@@@@@@@@@")
                        End If
                    Text4.Text = Text4.Text & Chr(13) & Chr(10)
                Next I
            
        Text4.Text = Text4.Text & "--------------  ------------  ------------  -------------  -------------" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Statistica bayesiana, valore predittivo del test (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        FlagGrafico = False
        Flag1 = False: Flag2 = False: Flag3 = False: Flag4 = False: Flag5 = False
        Screen.MousePointer = 1

End Sub

