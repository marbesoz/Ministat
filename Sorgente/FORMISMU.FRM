VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FormIstogrammiMultipli 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Diagramma a istogrammi multipli"
   ClientHeight    =   6465
   ClientLeft      =   1095
   ClientTop       =   1755
   ClientWidth     =   7365
   ForeColor       =   &H80000008&
   Icon            =   "FORMISMU.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6465
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
      Left            =   4440
      ScaleHeight     =   345
      ScaleWidth      =   945
      TabIndex        =   35
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
      TabIndex        =   30
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
      TabIndex        =   29
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
      TabIndex        =   28
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
      TabIndex        =   27
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
   Begin VB.Menu MenuGrafica 
      Caption         =   "&Grafica"
      Begin VB.Menu IstogrammaDellaDistribuzione 
         Caption         =   "&Diagramma a istogrammi multipli"
         Begin VB.Menu IstoCon 
            Caption         =   "&Con curve gaussiane"
         End
         Begin VB.Menu IstoSenza 
            Caption         =   "&Senza curve gaussiane"
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
            Caption         =   "&Istogrammi a colori"
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
Attribute VB_Name = "FormIstogrammiMultipli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim X() As Double
    Dim Y()
    Dim CC()
    Dim DD()
    Dim NumDati()
    Dim NomeVariabile()
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
    Dim SomX, SomXquad, SomXCubo, SomXquar
    Dim Media(), DevStnd()
    Dim CV, ErrStnd, NUMCOLONNE
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
    Dim FlagErroreScale, N, R, I, NewMX, Colonna
    On Error GoTo IstogrammiMultipli
        
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
            FormIstogrammiMultipli.Refresh
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
            
            For Colonna = 1 To NUMCOLONNE
            For I = 1 To CL
                CC(I, Colonna) = 0
            Next I
            Next Colonna

            For Colonna = 1 To NUMCOLONNE
            For N = 1 To NumDati(Colonna)
                R = CL
                If X(N, Colonna) = HI Then GoTo a13300
                R = Int((X(N, Colonna) - LW) / AM) + 1
a13300:
                CC(R, Colonna) = CC(R, Colonna) + 1
            Next N
            Next Colonna
            
            NewMX = CC(1, 1)
            For Colonna = 1 To NUMCOLONNE
                For I = 1 To CL
                    If CC(I, Colonna) > NewMX Then NewMX = CC(I, Colonna)
                Next I
            Next Colonna
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
        
IstogrammiMultipli:
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
    MaskedEdit5.Text = Format(Int(Sqr(Z1) + 1), "00")
    
    Screen.MousePointer = 1

End Sub

Private Sub FileEsci_Click()
    ReDim X(0, 0)
    ReDim Y(0, 0)
    ReDim CC(0, 0)
    ReDim DD(0, 0)
'    ReDim AA(0), SUP(0), INF(0)
    ReDim Centile(0)
    ReDim Zeta(0)
    FormPrincipale.Show
    Unload FormIstogrammiMultipli
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
End Sub

Private Sub FileStampa_Click()
    On Error GoTo ErroreFileStampaIstz

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

ErroreFileStampaIstz:
    Exit Sub

End Sub

Private Sub Form_Load()
    
    On Error GoTo GestioneErroriIsMu
    Dim D, I, J, L, M1, M2, ValoreInizialeD, Colonna, ColonnaIniziale
    Dim XX, N, R, K

    Screen.MousePointer = 11
    NUMCOLONNE = FormPrincipale!Grid1.SelEndCol - FormPrincipale!Grid1.SelStartCol + 1
    ColonnaIniziale = FormPrincipale!Grid1.SelStartCol

    ReDim X(NUM_RIGHE - 1, NUMCOLONNE) As Double
    ReDim Y(NUM_RIGHE - 1, NUMCOLONNE)
    ReDim CC(41, NUMCOLONNE)
    ReDim DD(41, NUMCOLONNE)
    ReDim NumDati(NUMCOLONNE)
    ReDim Media(NUMCOLONNE), DevStnd(NUMCOLONNE)
    ReDim NomeVariabile(NUMCOLONNE)

    AnnullatoCaricamentoForm = False
    FlagGrafico = False
    FlagStatistica = False
    FlagSfondoGrigio = True
    FlagDatiAColori = True
    FlagCon = False
    FlagSenza = False
    Picture2.Visible = False
    
    FormPrincipale!Grid1.Row = 0
    K = 0
    For I = FormPrincipale!Grid1.SelStartCol To FormPrincipale!Grid1.SelEndCol
        K = K + 1
        FormPrincipale!Grid1.Col = I
        NomeVariabile(K) = FormPrincipale!Grid1.Text
    Next I

' centra il form
    FormIstogrammiMultipli.Top = Screen.Height / 2 - Height / 2
    FormIstogrammiMultipli.Left = Screen.Width / 2 - Width / 2
' sposta in alto a sinistra
    Text4.Move 0, 0
' dimensionamento, ripreso nella sub resize
    Text4.Width = ScaleWidth
    Text4.Height = ScaleHeight
' inizializza
    Text4.Text = ""
    Text4.Enabled = False
' colonne in cui si trovano i dati
'    ColonnaX = FormPrincipale!Grid1.SelStartCol
' nome delle variabili
    FormPrincipale!Grid1.Row = 0
    FormPrincipale!Grid1.Col = ColonnaX
'    InAscisse = FormPrincipale!Grid1.Text
    InAscisse = "Grandezza e unità in ascisse"
    Titolo = "Diagramma a istogrammi multipli"
    InOrdinate = "Frequenza assoluta"
' carica i dati nel vettore X()
    FormGauge.Show
    FormGauge.Caption = "Diagramma a istogrammi multipli"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    FineDati = False
    
        For Colonna = 1 To NUMCOLONNE
            NumDati(Colonna) = 0
        Next Colonna

        Z1 = 0

        For Colonna = 1 To NUMCOLONNE
            FineDati = False
            
        For Riga = 1 To NUM_RIGHE - 1

            FormGauge!Label6.Caption = "[" & Colonna & "/" & NUMCOLONNE * 3 & "]"
            FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
            FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        FormPrincipale!Grid1.Row = Riga
        
        If FineDati = False Then
            FormPrincipale!Grid1.Col = ColonnaIniziale + Colonna - 1
            If FormPrincipale!Grid1.Text <> "" Then
                NumDati(Colonna) = NumDati(Colonna) + 1
                X(Riga, Colonna) = CDbl(-(-FormPrincipale!Grid1.Text))
            Else
                FineDati = True
            End If
        End If
        
        Next Riga

            If Z1 < NumDati(Colonna) Then Z1 = NumDati(Colonna)

        Next Colonna

' controlli finali sui dati
    If Z1 > 1 Then

            XMI = X(1, 1): XMA = X(1, 1)
            For Colonna = 1 To NUMCOLONNE
            For Riga = 1 To NumDati(Colonna)
                If X(Riga, Colonna) < XMI Then XMI = X(Riga, Colonna)
                If X(Riga, Colonna) > XMA Then XMA = X(Riga, Colonna)
            Next Riga
            Next Colonna
            
            XMI_Default = XMI
            XMA_Default = XMA
            If XMI = XMA Then GoTo GestioneErroriIsMu ' meglio così

' suddivisione in classi
            LW = XMI
            HI = XMA
            CL = Int(Sqr(Z1) + 1)
            AM = (HI - LW) / CL
            
            For Colonna = 1 To NUMCOLONNE
            FormGauge!Label6.Caption = "[" & NUMCOLONNE + Colonna & "/" & NUMCOLONNE * 3 & "]"
            For N = 1 To NumDati(Colonna)
                FormGauge.Label1.Caption = Int(100 * N / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * N / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                R = CL
                If X(N, Colonna) = HI Then GoTo 13300
                R = Int((X(N, Colonna) - LW) / AM) + 1
13300:
                CC(R, Colonna) = CC(R, Colonna) + 1
            Next N
            Next Colonna

            MX = CC(1, 1)
            For Colonna = 1 To NUMCOLONNE
                For I = 1 To CL
                    If CC(I, Colonna) > MX Then MX = CC(I, Colonna)
                Next I
            Next Colonna
            
            YMI = 0
            YMA = MX
            YMI_Default = YMI
            YMA_Default = YMA

            
            For Colonna = 1 To NUMCOLONNE
            
            FormGauge!Label6.Caption = "[" & NUMCOLONNE * 2 + Colonna & "/" & NUMCOLONNE * 3 & "]"
            
            SomX = 0: SomXquad = 0: SomXCubo = 0: SomXquar = 0

            For I = 1 To NumDati(Colonna)

                SomX = SomX + X(I, Colonna)
                SomXquad = SomXquad + X(I, Colonna) ^ 2
                SomXCubo = SomXCubo + X(I, Colonna) ^ 3
                SomXquar = SomXquar + X(I, Colonna) ^ 4

                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            Next I

            Media(Colonna) = SomX / NumDati(Colonna)
            DevStnd(Colonna) = Sqr((SomXquad - SomX ^ 2 / NumDati(Colonna)) / (NumDati(Colonna) - 1))
            Next Colonna

'            CV = DevStnd * 100 / Media
'            ErrStnd = DevStnd / Sqr(NumDati)
'
'            EmmeDue = (SomXquad - NumDati * Media ^ 2) / NumDati
'            EmmeTre = (SomXCubo - 3 * Media * SomXquad + 2 * NumDati * Media ^ 3) / NumDati
'            EmmeQua = (SomXquar - 4 * Media * SomXCubo + 6 * Media ^ 2 * SomXquad - 3 * NumDati * Media ^ 4) / NumDati
'            GiUno = EmmeTre / (EmmeDue * Sqr(EmmeDue))
'            GiDue = EmmeQua / EmmeDue ^ 2 - 3
'            EsseUno = Sqr(6 / NumDati)
'            EsseDue = Sqr(24 / NumDati)

            FormGauge.Label1.Caption = "100%"
            FormGauge!Gauge1.Value = 100
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            
            Unload FormGauge
            Text4.Visible = True
            Shape1.Visible = False
            Frame1.Visible = False
            Picture1.Visible = False
            FormPrincipale!Grid1.HighLight = True
            RipristinaSchermo
            FormIstogrammiMultipli.Show
            ' FormPrincipale.Hide
            IstoSenza_Click
            Screen.MousePointer = 1
            Exit Sub
    Else
        Unload FormGauge
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        MsgBox "Per il diagramma a istogrammi multipli" + Chr(10) + "sono necessari almeno due dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
    End If
    
    Exit Sub

GestioneErroriIsMu:
    Unload FormGauge
    AnnullatoCaricamentoForm = True
    Screen.MousePointer = 1
    MsgBox "E' presente un errore logico" + Chr(10) + "nei dati da rappresentare.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
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
    Dim FF, LargCol, XF, XI, K
    Dim SommaAree, ClassiValide, Altmed
    Dim Inizio, Fine, Ampiezza, Deviata, Ordinata, Colonna
    Dim Colore

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
    FormIstogrammiMultipli.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
        End If
    Picture1.Cls
    FormIstogrammiMultipli.ScaleWidth = 800
    FormIstogrammiMultipli.ScaleHeight = 480
    FormIstogrammiMultipli.Refresh

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
    For Colonna = 1 To NUMCOLONNE
    For I = 1 To CL + 1
        DD(I, Colonna) = 0
    Next I
    Next Colonna
    For Colonna = 1 To NUMCOLONNE
    For I = 1 To CL
        DD(I, Colonna) = Int(CC(I, Colonna) * FF)
    Next I
    Next Colonna

    LargCol = 459 / CL
    
    For Colonna = 1 To NUMCOLONNE
        
        If FlagDatiAColori = True Then
            Picture1.DrawStyle = 0
            Picture1.FillStyle = 0
            
            If Colonna = 1 Then
                Colore = 12
            End If
            If Colonna = 2 Then
                Colore = 9
            End If
            If Colonna = 3 Then
                Colore = 13
            End If
            If Colonna = 4 Then
                Colore = 2
            End If
        Else
            Picture1.DrawStyle = Colonna
            Picture1.FillStyle = 1
            Colore = 0
        End If
    
    Picture1.Line (0, 340)-(0, 340 - DD(1, Colonna)), QBColor(Colore)
    
     For I = 1 To CL
        If I > 1 Then
            If DD(I - 1, Colonna) = 0 Then
                Picture1.Line ((I - 1) * LargCol, 340)-((I - 1) * LargCol, 340 - DD(I, Colonna)), QBColor(Colore)
            End If
        End If
        Picture1.Line ((I - 1) * LargCol, 340 - DD(I, Colonna))-((I - 1) * LargCol + LargCol, 340 - DD(I, Colonna)), QBColor(Colore)
        Picture1.Line ((I - 1) * LargCol + LargCol, 340 - DD(I, Colonna))-((I - 1) * LargCol + LargCol, 340 - DD(I + 1, Colonna)), QBColor(Colore)
    Next I

'traccia gaussiana teorica
    
    Picture1.DrawStyle = 0
    Picture1.FillStyle = 0
    
    Inizio = (((Media(Colonna) - 4 * DevStnd(Colonna)) - LW) / (HI - LW)) * ((XF + (459 / CL) * CL) - XF)
    Fine = (((Media(Colonna) + 4 * DevStnd(Colonna)) - LW) / (HI - LW)) * ((XF + (459 / CL) * CL) - XF)
    
' calcola l'area dell'istogramma e deriva l'altezza della gaussiana
' assunta come un triangolo equilatero con uguale superficie
    SommaAree = 0
    For I = 1 To CL
        SommaAree = SommaAree + (DD(I, Colonna) * LargCol)
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
            Picture1.Line -(I, 340 - (Ordinata * Altmed)), QBColor(Colore)
        Next I

' fine tracciamento gaussiana teorica
    
    Next Colonna

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
    
        For K = 1 To NUMCOLONNE
        If FlagDatiAColori = True Then
            
            FormIstogrammiMultipli.DrawStyle = 0
            
            If K = 1 Then
                Colore = 12
            End If
            If K = 2 Then
                Colore = 9
            End If
            If K = 3 Then
                Colore = 13
            End If
            If K = 4 Then
                Colore = 2
            End If
        Else
            FormIstogrammiMultipli.DrawStyle = K
            Colore = 0
        End If
                If K = 1 Then
                    FormIstogrammiMultipli.CurrentX = 624: FormIstogrammiMultipli.CurrentY = 193
                    FormIstogrammiMultipli.Print NomeVariabile(K)
                    FormIstogrammiMultipli.Line (580, 197)-(616, 197), QBColor(Colore)
                End If
                If K = 2 Then
                    FormIstogrammiMultipli.CurrentX = 624: FormIstogrammiMultipli.CurrentY = 218
                    FormIstogrammiMultipli.Print NomeVariabile(K)
                    FormIstogrammiMultipli.Line (580, 222)-(616, 222), QBColor(Colore)
                End If
                If K = 3 Then
                    FormIstogrammiMultipli.CurrentX = 624: FormIstogrammiMultipli.CurrentY = 243
                    FormIstogrammiMultipli.Print NomeVariabile(K)
                    FormIstogrammiMultipli.Line (580, 247)-(616, 247), QBColor(Colore)
                End If
                If K = 4 Then
                    FormIstogrammiMultipli.CurrentX = 624: FormIstogrammiMultipli.CurrentY = 268
                    FormIstogrammiMultipli.Print NomeVariabile(K)
                    FormIstogrammiMultipli.Line (580, 272)-(616, 272), QBColor(Colore)
                End If
        Next K
    
    FlagCon = True
    FlagSenza = False
    Screen.MousePointer = 1

End Sub

Private Sub IstoSenza_Click()
    
    Dim I, XX, YIN, YSU, FX, FY, CX, CY
    Dim X1, X2, X3, X4, X5
    Dim Y1, Y2, Y3, Y4, Y5
    Dim FF, LargCol, XF, XI, K
    Dim SommaAree, ClassiValide, Altmed
    Dim Inizio, Fine, Ampiezza, Deviata, Ordinata, Colonna
    Dim Colore

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
    FormIstogrammiMultipli.Cls
        If FlagSfondoGrigio = True Then
            Picture1.BackColor = &H80000000
        Else
            Picture1.BackColor = &H80000005
        End If
    Picture1.Cls
    FormIstogrammiMultipli.ScaleWidth = 800
    FormIstogrammiMultipli.ScaleHeight = 480
    FormIstogrammiMultipli.Refresh

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
    For Colonna = 1 To NUMCOLONNE
    For I = 1 To CL + 1
        DD(I, Colonna) = 0
    Next I
    Next Colonna
    For Colonna = 1 To NUMCOLONNE
    For I = 1 To CL
        DD(I, Colonna) = Int(CC(I, Colonna) * FF)
    Next I
    Next Colonna

    LargCol = 459 / CL
    
    For Colonna = 1 To NUMCOLONNE

        If FlagDatiAColori = True Then
            Picture1.DrawStyle = 0
            Picture1.FillStyle = 0
            
            If Colonna = 1 Then
                Colore = 12
            End If
            If Colonna = 2 Then
                Colore = 9
            End If
            If Colonna = 3 Then
                Colore = 13
            End If
            If Colonna = 4 Then
                Colore = 2
            End If
        Else
            Picture1.DrawStyle = Colonna
            Picture1.FillStyle = 1
            Colore = 0
        End If
    
    Picture1.Line (0, 340)-(0, 340 - DD(I, Colonna)), QBColor(Colore)
    
     For I = 1 To CL
        If I > 1 Then
            If DD(I - 1, Colonna) = 0 Then
                Picture1.Line ((I - 1) * LargCol, 340)-((I - 1) * LargCol, 340 - DD(I, Colonna)), QBColor(Colore)
            End If
        End If
        Picture1.Line ((I - 1) * LargCol, 340 - DD(I, Colonna))-((I - 1) * LargCol + LargCol, 340 - DD(I, Colonna)), QBColor(Colore)
        Picture1.Line ((I - 1) * LargCol + LargCol, 340 - DD(I, Colonna))-((I - 1) * LargCol + LargCol, 340 - DD(I + 1, Colonna)), QBColor(Colore)
    Next I

'traccia gaussiana teorica
'
'    Picture1.DrawStyle = 0
'    Picture1.FillStyle = 0
'
'    Inizio = (((Media(Colonna) - 3 * DevStnd(Colonna)) - LW) / (HI - LW)) * ((XF + (459 / CL) * CL) - XF)
'    Fine = (((Media(Colonna) + 3 * DevStnd(Colonna)) - LW) / (HI - LW)) * ((XF + (459 / CL) * CL) - XF)
'
' calcola l'area dell'istogramma e deriva l'altezza della gaussiana
' assunta come un triangolo equilatero con uguale superficie
'    SommaAree = 0
'    For I = 1 To CL
'        SommaAree = SommaAree + (DD(I, Colonna) * LargCol)
'    Next I
'    Altmed = SommaAree * 2 / (Fine - Inizio)
'    Altmed = Altmed * 1.1 ' fattore di correzione
'
'        I = XF
'        Ampiezza = 4 / (Fine - Inizio)
'        Deviata = ((I - (XF + Inizio)) * Ampiezza) - 2
'        Ordinata = Exp(-Deviata * Deviata)
'        Picture1.PSet (I, 340 - (Ordinata * Altmed))
'
'        For I = XF + Inizio To XF + Fine
'            Ampiezza = 4 / (Fine - Inizio)
'            Deviata = ((I - (XF + Inizio)) * Ampiezza) - 2
'            Ordinata = Exp(-Deviata * Deviata)
'            Picture1.Line -(I, 340 - (Ordinata * Altmed)), QBColor(Colore)
'        Next I
'
' fine tracciamento gaussiana teorica
    
    Next Colonna

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
    
        For K = 1 To NUMCOLONNE
        If FlagDatiAColori = True Then
            
            FormIstogrammiMultipli.DrawStyle = 0
            
            If K = 1 Then
                Colore = 12
            End If
            If K = 2 Then
                Colore = 9
            End If
            If K = 3 Then
                Colore = 13
            End If
            If K = 4 Then
                Colore = 2
            End If
        Else
            FormIstogrammiMultipli.DrawStyle = K
            Colore = 0
        End If
                If K = 1 Then
                    FormIstogrammiMultipli.CurrentX = 624: FormIstogrammiMultipli.CurrentY = 193
                    FormIstogrammiMultipli.Print NomeVariabile(K)
                    FormIstogrammiMultipli.Line (580, 197)-(616, 197), QBColor(Colore)
                End If
                If K = 2 Then
                    FormIstogrammiMultipli.CurrentX = 624: FormIstogrammiMultipli.CurrentY = 218
                    FormIstogrammiMultipli.Print NomeVariabile(K)
                    FormIstogrammiMultipli.Line (580, 222)-(616, 222), QBColor(Colore)
                End If
                If K = 3 Then
                    FormIstogrammiMultipli.CurrentX = 624: FormIstogrammiMultipli.CurrentY = 243
                    FormIstogrammiMultipli.Print NomeVariabile(K)
                    FormIstogrammiMultipli.Line (580, 247)-(616, 247), QBColor(Colore)
                End If
                If K = 4 Then
                    FormIstogrammiMultipli.CurrentX = 624: FormIstogrammiMultipli.CurrentY = 268
                    FormIstogrammiMultipli.Print NomeVariabile(K)
                    FormIstogrammiMultipli.Line (580, 272)-(616, 272), QBColor(Colore)
                End If
        Next K
    
    FlagCon = False
    FlagSenza = True
    Screen.MousePointer = 1

End Sub

Private Sub MenuFile_Click()

    FileStampa.Enabled = False
        If FlagStatistica = True Or FlagGrafico = True Then
            If FormIstogrammiMultipli.Width < Screen.Width Then
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

