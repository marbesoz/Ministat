VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormSovrapposizione 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Diagramma di sovrapposizione"
   ClientHeight    =   6270
   ClientLeft      =   1200
   ClientTop       =   1830
   ClientWidth     =   7365
   ForeColor       =   &H80000008&
   Icon            =   "FORMSOVR.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6270
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
      Left            =   4320
      ScaleHeight     =   345
      ScaleWidth      =   945
      TabIndex        =   33
      Top             =   5040
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
      TabIndex        =   25
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
      Height          =   3132
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
         TabIndex        =   17
         Top             =   2520
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
         Left            =   5160
         TabIndex        =   21
         Top             =   2520
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
         Left            =   720
         TabIndex        =   16
         Top             =   2520
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
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   19
      Top             =   1100
      Width           =   972
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   20
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
      TabIndex        =   18
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
      TabIndex        =   32
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
      TabIndex        =   31
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
      TabIndex        =   30
      Top             =   6000
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
      TabIndex        =   29
      Top             =   6000
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
      TabIndex        =   28
      Top             =   6000
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
      TabIndex        =   27
      Top             =   6000
      Width           =   972
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1200
      TabIndex        =   26
      Top             =   6000
      Width           =   972
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1200
      TabIndex        =   24
      Top             =   840
      Width           =   3972
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   23
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
      TabIndex        =   22
      Top             =   360
      Width           =   6492
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   3492
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
   Begin VB.Menu MenuGrafica 
      Caption         =   "&Grafica"
      Begin VB.Menu GraficaX 
         Caption         =   "&Diagramma di sovrapposizione"
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
Attribute VB_Name = "FormSovrapposizione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim XMI, XMA, YMI, YMA
    Dim XMI_Default, XMA_Default
    Dim YMI_Default, YMA_Default
    Dim DEVIATAT, MEDIAX, DEVIANZAX
    Dim Titolo
    Dim InAscisse, InOrdinate
    Dim Flag1, Flag2, Flag3
    Dim FlagStatistica, FlagGrafico
    Dim FlagSfondoGrigio
    Dim FlagDatiAColori
    Dim X()
    Dim NUMCOLONNE
    Dim NomeVariabile()
    Dim NumDati()
    Dim FineDati()

Private Sub AiutoCalcolatrice_Click()
    Dim Calcolatrice
    On Error GoTo ErroreCalcolatriceOtto:
    
    Calcolatrice = Shell("Calc.exe", 1)
        
Exit Sub

ErroreCalcolatriceOtto:
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
    On Error GoTo Sovrapposizione
        
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
            
            Titolo = Text1.Text
            InAscisse = Text2.Text
            InOrdinate = Text3.Text
            XMI = MaskedEdit1.Text
            XMA = MaskedEdit2.Text
            YMI = MaskedEdit3.Text
            YMA = MaskedEdit4.Text
            Label13.Caption = Titolo
            Label14.Caption = InAscisse
            Label15.Caption = InOrdinate

            GraficaX_Click
        
        Frame1.Visible = False
        Shape1.Visible = False
        Command5.Enabled = True

        End If

Sovrapposizione:
        MaskedEdit1.Text = Format(XMI_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit2.Text = Format(XMA_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit3.Text = Format(YMI_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit4.Text = Format(YMA_Default, "+#####0.0####;-#####0.0####")
        
End Sub

Private Sub Command4_Click()
    
    GraficaX_Click
    
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
'    ReDim X(0)
'    ReDim Y(0)
'    ReDim AA(0), SUP(0), INF(0)
    FormPrincipale.Show
    Unload FormSovrapposizione
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
End Sub

Private Sub FileStampa_Click()
    On Error GoTo ErroreFileStampaSovr

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

ErroreFileStampaSovr:
    Exit Sub

End Sub

Private Sub Form_Load()
    On Error GoTo GestioneErroriSovrapposizione
    Dim I, J, K, Riga, Dummy

    Screen.MousePointer = 11
    
    NUMCOLONNE = FormPrincipale!Grid1.SelEndCol - FormPrincipale!Grid1.SelStartCol + 1

    ReDim X(NUMCOLONNE, NUM_RIGHE - 1)
    ReDim NomeVariabile(NUMCOLONNE)
    ReDim NumDati(NUMCOLONNE)
    ReDim FineDati(NUMCOLONNE)

    AnnullatoCaricamentoForm = False
    FlagGrafico = False
    Picture2.Visible = False
    
' centra il form
    FormSovrapposizione.Top = Screen.Height / 2 - Height / 2
    FormSovrapposizione.Left = Screen.Width / 2 - Width / 2
' sposta in alto a sinistra
    Text4.Move 0, 0
' dimensionamento, ripreso nella sub resize
    Text4.Width = ScaleWidth
    Text4.Height = ScaleHeight
' inizializza
    Text4.Text = ""
    Text4.Enabled = False
' nome delle variabili
    
    FormPrincipale!Grid1.Row = 0
    K = 0
    For I = FormPrincipale!Grid1.SelStartCol To FormPrincipale!Grid1.SelEndCol
        K = K + 1
        FormPrincipale!Grid1.Col = I
        NomeVariabile(K) = FormPrincipale!Grid1.Text
    Next I

    Titolo = "Diagramma di sovrapposizione"
    InAscisse = NomeVariabile(1)
    InOrdinate = "Grandezza e unità in ordinate"
    FlagSfondoGrigio = True
    FlagDatiAColori = True

    FormGauge.Show
    FormGauge.Caption = "Diagramma di sovrapposizione"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' carica i dati nel vettore X()
    For I = 1 To NUMCOLONNE
        NumDati(I) = 0
        FineDati(I) = False
    Next I

    K = 0
    For I = FormPrincipale!Grid1.SelStartCol To FormPrincipale!Grid1.SelEndCol
        FormPrincipale!Grid1.Col = I
        K = K + 1
        FormGauge!Label6.Caption = "[" & K & "/" & NUMCOLONNE & "]"
        For Riga = 1 To NUM_RIGHE - 1
            If 100 * Riga / (NUM_RIGHE - 1) = Int(100 * Riga / (NUM_RIGHE - 1)) Then
                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If
        FormPrincipale!Grid1.Row = Riga
        
        If FineDati(K) = False Then
            If FormPrincipale!Grid1.Text <> "" Then
                NumDati(K) = NumDati(K) + 1
                X(K, Riga) = CDbl(-(-FormPrincipale!Grid1.Text))
            Else
                FineDati(K) = True
            End If
        End If
        
        Next Riga
    Next I

' controlli finali sui dati
' devono esserci almeno tre dati per ciascuna variabile
    Dummy = True
    For I = 1 To NUMCOLONNE
        If NumDati(I) < 3 Then Dummy = False
    Next I

    If Dummy = False Then
        Unload FormGauge
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        MsgBox "Per il diagramma di sovrapposizione" + Chr(10) + "sono necessari almeno tre dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
        Exit Sub
    End If

' deve esserci lo stesso numero di dati per ciascuna variabile
'    If Flag1 = False Then
'    Dummy = True
'    For I = 2 To NumColonne
'        If NumDati(I) <> NumDati(I - 1) Then Dummy = False
'    Next I
'
'    If Dummy = False Then
'        Unload FormGauge
'        AnnullatoCaricamentoForm = True
'        Screen.MousePointer = 1
'        MsgBox "Per il diagramma di sovrapposizione è necessario" + Chr(10) + "lo stesso numero di dati per ogni variabile.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
'        Exit Sub
'    End If
'    End If

' altrimenti...

        XMI = X(1, 1): XMA = X(1, 1)
        For I = 2 To NumDati(1)
            If X(1, I) < XMI Then XMI = X(1, I)
            If X(1, I) > XMA Then XMA = X(1, I)
        Next I
        XMI_Default = XMI: XMA_Default = XMA

        YMI = X(2, 1): YMA = X(2, 1)
        For K = 2 To NUMCOLONNE
            For I = 1 To NumDati(K)
                If X(K, I) < YMI Then YMI = X(K, I)
                If X(K, I) > YMA Then YMA = X(K, I)
            Next I
        Next K
        YMI_Default = YMI: YMA_Default = YMA

        Unload FormGauge
        Text4.Visible = True
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
        FormSovrapposizione.Show
        ' FormPrincipale.Hide
        GraficaX_Click
        Screen.MousePointer = 1

        Exit Sub

GestioneErroriSovrapposizione:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente: riprovare" + Chr(10) + "con un numero inferiore di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
        Exit Sub

End Sub

Private Sub Form_Resize()
    
    Text4.Width = ScaleWidth
    Text4.Height = ScaleHeight

    GraficaX_Click

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

        GraficaX_Click

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
        
    Text1.Text = Titolo
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

        GraficaX_Click

    End If

End Sub

Private Sub GraficaX_Click()
    On Error GoTo GestioneErroriGraficaSovrapposizione
    ReDim AA(460), INF(460), SUP(460)
    Dim I, K, XX, YIN, YSU, FX, FY, CX, CY, Z1
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
    Picture1.Visible = False
    Picture2.Visible = False
    Text4.Visible = False
    FormSovrapposizione.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
        End If
    Picture1.Cls
    FormSovrapposizione.ScaleWidth = 800
    FormSovrapposizione.ScaleHeight = 480
    FormSovrapposizione.Refresh

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

    For K = 2 To NUMCOLONNE
        Z1 = NumDati(1): If NumDati(K) < NumDati(1) Then Z1 = NumDati(K)
    For I = 1 To Z1
        CX = (X(1, I) - XMI) * FX ' : If CX < 0 Or CX > 459 Then GoTo 10565
        CY = 339 - (X(K, I) - YIN) * FY ' : If CY < 0 Or CY > 339 Then GoTo 10565
            If FlagDatiAColori = True Then
                If K = 2 Then
                    Picture1.Line (CX - 4, CY)-(CX + 4, CY), QBColor(12)
                    Picture1.Line (CX, CY - 4)-(CX, CY + 4), QBColor(12)
                End If
                If K = 3 Then
                    Picture1.Line (CX - 4, CY - 3)-(CX, CY + 3), QBColor(9)
                    Picture1.Line (CX, CY + 3)-(CX + 4, CY - 3), QBColor(9)
                    Picture1.Line (CX + 4, CY - 3)-(CX - 4, CY - 3), QBColor(9)
                End If
                If K = 4 Then
                    Picture1.Line (CX - 4, CY - 4)-(CX - 4, CY + 4), QBColor(13)
                    Picture1.Line (CX - 4, CY + 4)-(CX + 4, CY + 4), QBColor(13)
                    Picture1.Line (CX + 4, CY + 4)-(CX + 4, CY - 4), QBColor(13)
                    Picture1.Line (CX + 4, CY - 4)-(CX - 4, CY - 4), QBColor(13)
                End If
                If K = 5 Then
                    Picture1.Circle (CX, CY), 4, QBColor(2)
                End If
            Else
                If K = 2 Then
                    Picture1.Line (CX - 4, CY)-(CX + 4, CY), QBColor(0)
                    Picture1.Line (CX, CY - 4)-(CX, CY + 4), QBColor(0)
                End If
                If K = 3 Then
                    Picture1.Line (CX - 4, CY - 3)-(CX, CY + 3), QBColor(0)
                    Picture1.Line (CX, CY + 3)-(CX + 4, CY - 3), QBColor(0)
                    Picture1.Line (CX + 4, CY - 3)-(CX - 4, CY - 3), QBColor(0)
                End If
                If K = 4 Then
                    Picture1.Line (CX - 4, CY - 4)-(CX - 4, CY + 4), QBColor(0)
                    Picture1.Line (CX - 4, CY + 4)-(CX + 4, CY + 4), QBColor(0)
                    Picture1.Line (CX + 4, CY + 4)-(CX + 4, CY - 4), QBColor(0)
                    Picture1.Line (CX + 4, CY - 4)-(CX - 4, CY - 4), QBColor(0)
                End If
                If K = 5 Then
                    Picture1.Circle (CX, CY), 4, QBColor(0)
                End If
            
            End If
' 10565 :
    Next I
    
    Next K

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
    Label13.Caption = Titolo
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
    
        For K = 2 To NUMCOLONNE
            If FlagDatiAColori = True Then
                If K = 2 Then
                    CX = 576: CY = 200
                    FormSovrapposizione.Line (CX - 4, CY)-(CX + 4, CY), QBColor(12)
                    FormSovrapposizione.Line (CX, CY - 4)-(CX, CY + 4), QBColor(12)
                    FormSovrapposizione.CurrentX = 596: FormSovrapposizione.CurrentY = 193
                    FormSovrapposizione.Print NomeVariabile(K)
                End If
                If K = 3 Then
                    CX = 576: CY = 225
                    FormSovrapposizione.Line (CX - 4, CY - 3)-(CX, CY + 3), QBColor(9)
                    FormSovrapposizione.Line (CX, CY + 3)-(CX + 4, CY - 3), QBColor(9)
                    FormSovrapposizione.Line (CX + 4, CY - 3)-(CX - 4, CY - 3), QBColor(9)
                    FormSovrapposizione.CurrentX = 596: FormSovrapposizione.CurrentY = 218
                    FormSovrapposizione.Print NomeVariabile(K)
                End If
                If K = 4 Then
                    CX = 576: CY = 250
                    FormSovrapposizione.Line (CX - 4, CY - 4)-(CX - 4, CY + 4), QBColor(13)
                    FormSovrapposizione.Line (CX - 4, CY + 4)-(CX + 4, CY + 4), QBColor(13)
                    FormSovrapposizione.Line (CX + 4, CY + 4)-(CX + 4, CY - 4), QBColor(13)
                    FormSovrapposizione.Line (CX + 4, CY - 4)-(CX - 4, CY - 4), QBColor(13)
                    FormSovrapposizione.CurrentX = 596: FormSovrapposizione.CurrentY = 243
                    FormSovrapposizione.Print NomeVariabile(K)
                End If
                If K = 5 Then
                    CX = 576: CY = 275
                    FormSovrapposizione.Circle (CX, CY), 4, QBColor(2)
                    FormSovrapposizione.CurrentX = 596: FormSovrapposizione.CurrentY = 268
                    FormSovrapposizione.Print NomeVariabile(K)
                End If
            Else
                If K = 2 Then
                    CX = 576: CY = 200
                    FormSovrapposizione.Line (CX - 4, CY)-(CX + 4, CY), QBColor(0)
                    FormSovrapposizione.Line (CX, CY - 4)-(CX, CY + 4), QBColor(0)
                    FormSovrapposizione.CurrentX = 596: FormSovrapposizione.CurrentY = 193
                    FormSovrapposizione.Print NomeVariabile(K)
                End If
                If K = 3 Then
                    CX = 576: CY = 225
                    FormSovrapposizione.Line (CX - 4, CY - 3)-(CX, CY + 3), QBColor(0)
                    FormSovrapposizione.Line (CX, CY + 3)-(CX + 4, CY - 3), QBColor(0)
                    FormSovrapposizione.Line (CX + 4, CY - 3)-(CX - 4, CY - 3), QBColor(0)
                    FormSovrapposizione.CurrentX = 596: FormSovrapposizione.CurrentY = 218
                    FormSovrapposizione.Print NomeVariabile(K)
                End If
                If K = 4 Then
                    CX = 576: CY = 250
                    FormSovrapposizione.Line (CX - 4, CY - 4)-(CX - 4, CY + 4), QBColor(0)
                    FormSovrapposizione.Line (CX - 4, CY + 4)-(CX + 4, CY + 4), QBColor(0)
                    FormSovrapposizione.Line (CX + 4, CY + 4)-(CX + 4, CY - 4), QBColor(0)
                    FormSovrapposizione.Line (CX + 4, CY - 4)-(CX - 4, CY - 4), QBColor(0)
                    FormSovrapposizione.CurrentX = 596: FormSovrapposizione.CurrentY = 243
                    FormSovrapposizione.Print NomeVariabile(K)
                End If
                If K = 5 Then
                    CX = 576: CY = 275
                    FormSovrapposizione.Circle (CX, CY), 4, QBColor(0)
                    FormSovrapposizione.CurrentX = 596: FormSovrapposizione.CurrentY = 268
                    FormSovrapposizione.Print NomeVariabile(K)
                End If
            
            End If
        Next K

    Picture1.Visible = True
    FlagGrafico = True
    Screen.MousePointer = 1
    Exit Sub

GestioneErroriGraficaSovrapposizione:
    Screen.MousePointer = 1
    MsgBox "E' presente un errore logico" + Chr(10) + "nei dati da rappresentare.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
    Exit Sub

End Sub

Private Sub MenuFile_Click()

    FileStampa.Enabled = False
        If FlagStatistica = True Or FlagGrafico = True Then
            If FormSovrapposizione.Width < Screen.Width Then
                FileStampa.Enabled = True
            End If
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
End Sub

