VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormPrimaDopo 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Diagramma di appaiamento"
   ClientHeight    =   6210
   ClientLeft      =   1185
   ClientTop       =   1830
   ClientWidth     =   7365
   ForeColor       =   &H80000008&
   Icon            =   "FORMPRIM.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6210
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
      Left            =   4080
      ScaleHeight     =   345
      ScaleWidth      =   945
      TabIndex        =   22
      Top             =   4800
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
      TabIndex        =   15
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
      Height          =   3135
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   7095
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
         TabIndex        =   14
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
         TabIndex        =   18
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
         TabIndex        =   13
         Top             =   2520
         Width           =   1212
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         MaxLength       =   40
         TabIndex        =   6
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
      Begin MSMask.MaskEdBox MaskedEdit3 
         Height          =   285
         Left            =   3720
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "In ordinate:"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   3720
         TabIndex        =   5
         Top             =   960
         Width           =   1812
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
      TabIndex        =   16
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
      TabIndex        =   17
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
      TabIndex        =   7
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
      TabIndex        =   8
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
      TabIndex        =   4
      Top             =   5740
      Width           =   972
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1200
      TabIndex        =   21
      Top             =   840
      Width           =   2292
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   20
      Top             =   6120
      Width           =   2175
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
      TabIndex        =   19
      Top             =   480
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   3495
      Left            =   1200
      Top             =   720
      Width           =   7335
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
      Begin VB.Menu DiagrammaPrimaDopo 
         Caption         =   "&Diagramma di appaiamento"
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
Attribute VB_Name = "FormPrimaDopo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim X() As Double
    Dim Y() As Double
    Dim Riga, ColonnaX, ColonnaY, FineDati
    Dim Z1, Z2, A, B, C, D, e, F, G, H, I, L, M
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
    Dim NomePrimaVariabile, NomeSecondaVariabile

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
    On Error GoTo PrimaDopo
        
    FlagErroreScale = False
    MaskedEdit3.Text = FormattaStringa(MaskedEdit3.Text)
    MaskedEdit4.Text = FormattaStringa(MaskedEdit4.Text)
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
            InOrdinate = Text3.Text
            YMI = MaskedEdit3.Text
            YMA = MaskedEdit4.Text
            Label13.Caption = Titolo
            Label15.Caption = InOrdinate

            DiagrammaPrimaDopo_Click
        
        Frame1.Visible = False
        Shape1.Visible = False
        Command5.Enabled = True

        End If

PrimaDopo:
        MaskedEdit3.Text = Format(YMI_Default, "+#####0.0####;-#####0.0####")
        MaskedEdit4.Text = Format(YMA_Default, "+#####0.0####;-#####0.0####")

End Sub

Private Sub Command4_Click()
    
    DiagrammaPrimaDopo_Click
    
    Frame1.Visible = False
    Shape1.Visible = False
    Command5.Enabled = True

End Sub

Private Sub Command5_Click()
    
    MaskedEdit3.Text = Format(YMI_Default, "+#####0.0####;-#####0.0####")
    MaskedEdit4.Text = Format(YMA_Default, "+#####0.0####;-#####0.0####")

    Screen.MousePointer = 1
    
End Sub

Private Sub DiagrammaPrimaDopo_Click()
    On Error GoTo GestioneErroriGraficaPrimaDopo
    Dim I, XX, YIN, YSU, FX, FY, CXP, CYP, CXD, CYD
    Dim X1, X2, X3, X4, X5
    Dim Y1, Y2, Y3, Y4, Y5

    Screen.MousePointer = 11
    Shape1.Visible = False
    Frame1.Visible = False
    Label13.Visible = False
    Label15.Visible = False
    Label21.Visible = False
    Label22.Visible = False
    Label23.Visible = False
    Label24.Visible = False
    Label25.Visible = False
    Picture1.Visible = False
    Picture2.Visible = False
    Text4.Visible = False
    FormPrimaDopo.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
        End If
    Picture1.Cls
    FormPrimaDopo.ScaleWidth = 800
    FormPrimaDopo.ScaleHeight = 480
    FormPrimaDopo.Refresh

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
    
    
    FY = 339 / (YSU - YIN)

    For I = 1 To Z1
        CXP = 125
        CYP = 339 - (X(I) - YIN) * FY
        CXD = 350
        CYD = 339 - (Y(I) - YIN) * FY
            If FlagDatiAColori = True Then
                Picture1.FillStyle = 0
                Picture1.FillColor = QBColor(12)
                Picture1.Circle (CXP, CYP), 6, QBColor(0)
                Picture1.FillColor = QBColor(11)
                Picture1.Circle (CXD, CYD), 6, QBColor(0)
                Picture1.Line (CXP, CYP)-(CXD, CYD), QBColor(0)
            Else
                Picture1.FillStyle = 1
                Picture1.FillColor = QBColor(7)
                Picture1.Circle (CXP, CYP), 6, QBColor(0)
                Picture1.Circle (CXD, CYD), 6, QBColor(0)
                Picture1.Line (CXP, CYP)-(CXD, CYD), QBColor(0)
            End If
    Next I
    
    Y5 = YSU: Y4 = YSU - (YSU - YIN) / 4: Y3 = YIN + (YSU - YIN) / 2: Y2 = YIN + (YSU - YIN) / 4: Y1 = YIN
    
    Label13.Left = 330 - (Label13.Width / 2): Label13.Top = 27
    Label15.Left = 100: Label15.Top = 61
    Label21.Left = 90 - Label21.Width: Label21.Top = 409
    Label22.Left = 90 - Label22.Width: Label22.Top = 329
    Label23.Left = 90 - Label23.Width: Label23.Top = 244
    Label24.Left = 90 - Label24.Width: Label24.Top = 160
    Label25.Left = 90 - Label25.Width: Label25.Top = 79

    Label13.Visible = True
    Label15.Visible = True
    Label21.Visible = True
    Label22.Visible = True
    Label23.Visible = True
    Label24.Visible = True
    Label25.Visible = True

    FlagStatistica = False
    FlagGrafico = True
    Label13.Caption = Titolo
    Label15.Caption = InOrdinate
    Label21.Caption = Format(Y1, "#####0.0####;-#####0.0####")
    Label22.Caption = Format(Y2, "#####0.0####;-#####0.0####")
    Label23.Caption = Format(Y3, "#####0.0####;-#####0.0####")
    Label24.Caption = Format(Y4, "#####0.0####;-#####0.0####")
    Label25.Caption = Format(Y5, "#####0.0####;-#####0.0####")
    
    FormPrimaDopo.CurrentX = 596: FormPrimaDopo.CurrentY = 193
    FormPrimaDopo.Print "A = " & NomePrimaVariabile
    FormPrimaDopo.CurrentX = 596: FormPrimaDopo.CurrentY = 218
    FormPrimaDopo.Print "B = " & NomeSecondaVariabile
    FormPrimaDopo.CurrentX = 220: FormPrimaDopo.CurrentY = 440
    FormPrimaDopo.Print "A"
    FormPrimaDopo.CurrentX = 444: FormPrimaDopo.CurrentY = 440
    FormPrimaDopo.Print "B"

    Picture1.Visible = True
    Screen.MousePointer = 1
    Exit Sub

GestioneErroriGraficaPrimaDopo:
    Screen.MousePointer = 1
    MsgBox "E' presente un errore logico" + Chr(10) + "nei dati da rappresentare.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
    Exit Sub

End Sub

Private Sub FileEsci_Click()
    FormPrincipale.Show
    Unload FormPrimaDopo
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
End Sub

Private Sub FileStampa_Click()
    On Error GoTo ErroreFileStampaPrim

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

ErroreFileStampaPrim:
    Exit Sub

End Sub

Private Sub Form_Load()
    ReDim X(1 To NUM_RIGHE - 1) As Double
    ReDim Y(1 To NUM_RIGHE - 1) As Double

    Screen.MousePointer = 11
    AnnullatoCaricamentoForm = False
    FlagGrafico = False
    FlagSfondoGrigio = True
    FlagDatiAColori = True
    Picture2.Visible = False
    
' centra il form
    FormPrimaDopo.Top = Screen.Height / 2 - Height / 2
    FormPrimaDopo.Left = Screen.Width / 2 - Width / 2
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
    NomePrimaVariabile = FormPrincipale!Grid1.Text
    FormPrincipale!Grid1.Col = ColonnaY
    NomeSecondaVariabile = FormPrincipale!Grid1.Text
    InOrdinate = "Grandezza e unità in ordinate"
    Titolo = "Diagramma di appaiamento"
' carica i dati nei vettori X() e Y()
    FormGauge.Show
    FormGauge.Caption = "Diagramma di appaiamento"
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
' per la regressione devono esserci almeno due punti. ..
    If Z1 > 1 Then
' ...e i dati devono essere in numero uguale
        If Z1 = Z2 Then
            YMI = Y(1): YMA = Y(1)
            For I = 1 To Z1
                If X(I) < YMI Then YMI = X(I)
                If X(I) > YMA Then YMA = X(I)
                If Y(I) < YMI Then YMI = Y(I)
                If Y(I) > YMA Then YMA = Y(I)
            Next
            YMI_Default = YMI
            YMA_Default = YMA
            Unload FormGauge
            Text4.Visible = True
            Shape1.Visible = False
            Frame1.Visible = False
            Picture1.Visible = False
            FormPrincipale!Grid1.HighLight = True
            RipristinaSchermo
            FormPrimaDopo.Show
            ' FormPrincipale.Hide
            DiagrammaPrimaDopo_Click
            Screen.MousePointer = 1
            Exit Sub
        Else
            Unload FormGauge
            AnnullatoCaricamentoForm = True
            Screen.MousePointer = 1
            MsgBox "Per il diagramma di appaiamento" + Chr(10) + "le due variabili devono contenere" + Chr(10) + "lo stesso numero di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
            FormPrincipale!Grid1.HighLight = True
            RipristinaSchermo
        End If
    Else
        Unload FormGauge
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        MsgBox "Per il diagramma di appaiamento" + Chr(10) + "sono necessari almeno due dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
    End If

End Sub

Private Sub Form_Resize()
    
    Text4.Width = ScaleWidth
    Text4.Height = ScaleHeight

    DiagrammaPrimaDopo_Click

End Sub

Private Sub GraficaDatiAColori_Click()
    
    GraficaDatiAColori.Checked = Not GraficaDatiAColori.Checked
    FlagDatiAColori = Not FlagDatiAColori

    If FlagGrafico = True Then
        Picture1.Visible = False
        Label13.Visible = False
        Label15.Visible = False
        Label21.Visible = False
        Label22.Visible = False
        Label23.Visible = False
        Label24.Visible = False
        Label25.Visible = False

        DiagrammaPrimaDopo_Click

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
    Label15.Visible = False
    Label21.Visible = False
    Label22.Visible = False
    Label23.Visible = False
    Label24.Visible = False
    Label25.Visible = False
        
    Text1.Text = Titolo
    Text3.Text = InOrdinate
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
        Label15.Visible = False
        Label21.Visible = False
        Label22.Visible = False
        Label23.Visible = False
        Label24.Visible = False
        Label25.Visible = False

        DiagrammaPrimaDopo_Click

    End If

End Sub

Private Sub MenuFile_Click()

    FileStampa.Enabled = False
        If FlagGrafico = True Then
            If FormPrimaDopo.Width < Screen.Width Then
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

