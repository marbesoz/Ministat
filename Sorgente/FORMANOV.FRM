VERSION 5.00
Begin VB.Form FormANOVA 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Analisi della varianza"
   ClientHeight    =   6090
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
   Icon            =   "FORMANOV.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6090
   ScaleWidth      =   7365
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
   Begin VB.Menu MenuStatistica 
      Caption         =   "&Statistica"
      Begin VB.Menu StatisticaRapportoTraVarianze 
         Caption         =   "&Rapporto tra varianze"
      End
      Begin VB.Menu StatisticaANOVAUno 
         Caption         =   "&Analisi della varianza a un fattore"
      End
      Begin VB.Menu StatisticaANOVADue 
         Caption         =   "A&nalisi della varianza a due fattori"
      End
      Begin VB.Menu StatisticaSeparatoreUno 
         Caption         =   "-"
      End
      Begin VB.Menu StatisticaDuplicati 
         Caption         =   "&Variabilità entro duplicati"
      End
      Begin VB.Menu StatisticaComponenti 
         Caption         =   "&Componenti della variabilità"
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
Attribute VB_Name = "FormANOVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim X()
    Dim NUMCOLONNE
    Dim NomeVariabile()
    Dim NumDati()
    Dim FineDati()
    Dim FlagStatistica, Flag1

Private Sub AiutoCalcolatrice_Click()
    Dim Calcolatrice
    On Error GoTo ErroreCalcolatriceUno:
    
    Calcolatrice = Shell("Calc.exe", 1)
        
Exit Sub

ErroreCalcolatriceUno:
    MsgBox "Calcolatrice non disponibile.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcolatrice"
Exit Sub


End Sub

Private Sub AiutoGuida_Click()
    
    FormPrincipale!CMDialog1.HelpFile = App.Path & "\ministat.hlp"
    FormPrincipale!CMDialog1.HelpCommand = cdlHelpContents
    FormPrincipale!CMDialog1.ShowHelp

End Sub

Private Sub FileEsci_Click()
    ReDim X(0, 0)
    FormPrincipale.Show
    Unload FormANOVA
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
End Sub

Private Sub FileStampa_Click()
    On Error GoTo ErroreFileStampaAnova

    If FlagStatistica = True Then
        Printer.Print Text4.Text
        Printer.EndDoc
    End If

    Exit Sub

ErroreFileStampaAnova:
    Exit Sub

End Sub

Private Sub Form_Load()
    On Error GoTo GestioneErroriANOVA
    Dim I, J, K, Riga, Dummy

    Screen.MousePointer = 11
    
    NUMCOLONNE = FormPrincipale!Grid1.SelEndCol - FormPrincipale!Grid1.SelStartCol + 1
    If NUMCOLONNE = 2 Then
        Flag1 = True
    Else
        Flag1 = False
    End If

    ReDim X(NUMCOLONNE, NUM_RIGHE - 1)
    ReDim NomeVariabile(NUMCOLONNE)
    ReDim NumDati(NUMCOLONNE)
    ReDim FineDati(NUMCOLONNE)

    AnnullatoCaricamentoForm = False
    FlagStatistica = False

' centra il form
    FormANOVA.Top = Screen.Height / 2 - Height / 2
    FormANOVA.Left = Screen.Width / 2 - Width / 2
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

    FormGauge.Show
    FormGauge.Caption = "Analisi della varianza"
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
' devono esserci almeno due dati per ciascuna variabile
    Dummy = True
    For I = 1 To NUMCOLONNE
        If NumDati(I) < 2 Then Dummy = False
    Next I

    If Dummy = False Then
        Unload FormGauge
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        MsgBox "Per l'analisi della varianza" + Chr(10) + "sono necessari almeno due dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Analisi della varianza non eseguibile"
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
        Exit Sub
    End If

' deve esserci lo stesso numero di dati per ciascuna variabile
    If Flag1 = False Then
    Dummy = True
    For I = 2 To NUMCOLONNE
        If NumDati(I) <> NumDati(I - 1) Then Dummy = False
    Next I

    If Dummy = False Then
        Unload FormGauge
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        MsgBox "Se le variabili sono più di due è necessario" + Chr(10) + "lo stesso numero di dati per ogni variabile.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Analisi della varianza non eseguibile"
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
        Exit Sub
    End If
    End If

' altrimenti...
        Unload FormGauge
        Text4.Visible = True
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
        FormANOVA.Show
        ' FormPrincipale.Hide
        Screen.MousePointer = 1

Exit Sub

GestioneErroriANOVA:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente: riprovare" + Chr(10) + "con un numero inferiore di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Analisi della varianza non eseguibile"
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
        Exit Sub

End Sub

Private Sub Form_Resize()
    
    Text4.Width = ScaleWidth
    Text4.Height = ScaleHeight

End Sub

Private Sub Label24_Click()

End Sub

Private Sub MenuFile_Click()

    FileStampa.Enabled = False
        If FlagStatistica = True Then
            FileStampa.Enabled = True
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

Private Sub MenuStatistica_Click()
    
    If Flag1 = True Then
        StatisticaRapportoTraVarianze.Enabled = True
        StatisticaANOVAUno.Enabled = True
        StatisticaANOVADue.Enabled = False
        StatisticaDuplicati.Enabled = True
        StatisticaComponenti.Enabled = False
    Else
        StatisticaRapportoTraVarianze.Enabled = False
        StatisticaANOVAUno.Enabled = True
        StatisticaANOVADue.Enabled = True
        StatisticaDuplicati.Enabled = False
        StatisticaComponenti.Enabled = True
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

Private Sub StatisticaANOVADue_Click()
    On Error GoTo GestioneErroriANOVADue
    Dim I, F, N1, N2, XX, PI, K, U
    Dim U2, S1, A, B, C, G, N, Q, R
    Dim T7, J2, A3, S2, J, P1
    Dim t, P2 As Double
    Dim Ri, Co, SO, SR, SC, SN
    Dim VR, VC, VN, Media, ST, N1A, N2A
    ReDim MR(NUMCOLONNE), MC(NumDati(1))

    Screen.MousePointer = 11
    Text4.Visible = False

    Cls
    
    FormGauge.Show
    FormGauge.Caption = "ANOVA a due fattori"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' ANALISI DELLA VARIANZA A 2 FATTORI

    Ri = NUMCOLONNE
    Co = NumDati(1)

        FormGauge!Label6.Caption = "[1/4]"
    SO = 0
    For I = 1 To Ri
        FormGauge.Label1.Caption = Int(100 * I / Ri) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / Ri)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    For J = 1 To Co
    SO = SO + X(I, J)
    Next J
    Next I
    Media = SO / (Ri * Co)

        FormGauge!Label6.Caption = "[2/4]"
    SO = 0
    For I = 1 To Ri
        FormGauge.Label1.Caption = Int(100 * I / Ri) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / Ri)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    For J = 1 To Co
    SO = SO + X(I, J)
    Next J
    MR(I) = SO / Co
    SO = 0
    Next I

        FormGauge!Label6.Caption = "[3/4]"
    SO = 0
    For J = 1 To Co
        FormGauge.Label1.Caption = Int(100 * J / Co) & "%"
        FormGauge!Gauge1.Value = Int(100 * J / Co)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    For I = 1 To Ri
    SO = SO + X(I, J)
    Next I
    MC(J) = SO / Ri
    SO = 0
    Next J
    
        FormGauge!Label6.Caption = "[4/4]"
    ST = 0
    For I = 1 To Ri
        FormGauge.Label1.Caption = Int(100 * I / Ri) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / Ri)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    For J = 1 To Co
    ST = ST + (X(I, J) - Media) * (X(I, J) - Media)
    Next J
    Next I
    
    SR = 0
    For I = 1 To Ri
    SR = SR + (MR(I) - Media) * (MR(I) - Media)
    Next I
    SR = Co * SR
    SC = 0
    For J = 1 To Co
    SC = SC + (MC(J) - Media) * (MC(J) - Media)
    Next J
    SC = Ri * SC
    SN = ST - SR - SC
    VR = SR / (Ri - 1)
    VC = SC / (Co - 1)
    VN = SN / ((Ri - 1) * (Co - 1))
    F = VR / VN: N1 = Ri - 1: N2 = (Ri - 1) * (Co - 1): GoSub b4200: N1A = N1: N2A = N2: P1 = P2
    F = VC / VN: N1 = Co - 1: N2 = (Ri - 1) * (Co - 1): GoSub b4200

        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Analisi della varianza a due fattori (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File         : " & NomeDelFile & Chr(13) & Chr(10)
            If NomeDellaTabella <> "" Then
                Text4.Text = Text4.Text & "Tabella      : " & NomeDellaTabella & Chr(13) & Chr(10)
            End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        For I = 1 To NUMCOLONNE
            If I >= 1 And I <= 9 Then
                Text4.Text = Text4.Text & "Variabile " & I & "  : " & NomeVariabile(I) & Chr(13) & Chr(10)
            End If
            If I = 10 Then
                Text4.Text = Text4.Text & "Variabile " & I & " : " & NomeVariabile(I) & Chr(13) & Chr(10)
            End If
        Next I
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Gradi di libertà del numeratore     = " & N1A & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Gradi di libertà del denominatore   = " & N2A & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Varianza spiegata (tra le colonne)  = " & Format(VR, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Varianza residua                    = " & Format(VN, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Test F (rapporto spiegata/residua)  = " & Format(VR / VN, "#####0.0####") & Chr(13) & Chr(10)
        If VR = 0 Then P1 = 1
        If P1 >= 0.00001 Then
            Text4.Text = Text4.Text & "Probabilità p                       = " & Format(P1, "#####0.0####") & Chr(13) & Chr(10)
        Else
            P1 = 0.00001
            Text4.Text = Text4.Text & "Probabilità p                       < " & Format(P1, "#####0.0####") & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Gradi di libertà del numeratore     = " & N1 & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Gradi di libertà del denominatore   = " & N2 & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Varianza spiegata (tra le righe)    = " & Format(VC, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Varianza residua                    = " & Format(VN, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Test F (rapporto spiegata/residua)  = " & Format(VC / VN, "#####0.0####") & Chr(13) & Chr(10)
        If VC = 0 Then P2 = 1
        If P2 >= 0.00001 Then
            Text4.Text = Text4.Text & "Probabilità p                       = " & Format(P2, "#####0.0####") & Chr(13) & Chr(10)
        Else
            P2 = 0.00001
            Text4.Text = Text4.Text & "Probabilità p                       < " & Format(P2, "#####0.0####") & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Analisi della varianza a due fattori (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        Screen.MousePointer = 1

Exit Sub

b4200:
b4210:
b4220:
    PI = 3.141592654: K = 1 / 3
b4230:
    U = (1 - 2 / (9 * N2)) * F ^ K - 1 + 2 / (9 * N1)
b4240:
    U = U / Sqr(Abs(2 / (9 * N1) + 2 / (9 * N2) * F ^ (2 * K)))
b4250:
    t = 0: If U >= 0 Then GoTo b4270
b4260:
    t = 1: U = Abs(U)
b4270:
    U2 = U * U
b4280:
    S1 = Exp(-U2 / 2) / Sqr(2 * PI)
b4290:
    I = 1: A = U: C = U
b4300:
    I = I + 2
b4310:
    G = C: B = A
b4320:
    B = A
b4330:
    A = B / I * U2
b4340:
    C = G + A
b4350:
    If C > G Then GoTo b4300
b4360:
    P2 = 0.5 - C * S1
b4370:
    If t <> 1 Then GoTo b4390
b4380:
    P2 = 1 - P2
b4390:
    Return

GestioneErroriANOVADue:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico dei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        Exit Sub

End Sub

Private Sub StatisticaANOVAUno_Click()
    On Error GoTo GestioneErroriANOVAUno
    Dim I, F, N1, N2, XX, PI, K, U
    Dim U2, S1, A, B, C, G, N, Q, R
    Dim T7, J2, A3, S2, J, P1
    Dim t, P2 As Double
    Dim SO, Media, ST, SS, SN, VS, VN
    ReDim MR(NUMCOLONNE)

    Screen.MousePointer = 11
    Text4.Visible = False

    Cls
' anche due sole variabili devono avere lo stesso numero di dati
    If NUMCOLONNE = 2 Then
        If NumDati(1) <> NumDati(2) Then
            Screen.MousePointer = 1
            MsgBox "Per l'ANOVA a un fattore le" + Chr(10) + "due variabili devono contenere" + Chr(10) + "lo stesso numero di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
            Exit Sub
        End If
    End If

    FormGauge.Show
    FormGauge.Caption = "ANOVA a un fattore"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' ANALISI DELLA VARIANZA A 1 FATTORE

    R = NUMCOLONNE
    N = NumDati(1)
    
        FormGauge!Label6.Caption = "[1/3]"
    SO = 0
    For I = 1 To R
        FormGauge.Label1.Caption = Int(100 * I / R) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / R)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    For J = 1 To N
    SO = SO + X(I, J)
    Next J
    Next I
    Media = SO / (N * R)
    
        FormGauge!Label6.Caption = "[2/3]"
    SO = 0
    For I = 1 To R
        FormGauge.Label1.Caption = Int(100 * I / R) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / R)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    For J = 1 To N
    SO = SO + X(I, J)
    Next J
    MR(I) = SO / N
    SO = 0
    Next I
    
        FormGauge!Label6.Caption = "[3/3]"
    ST = 0
    For I = 1 To R
        FormGauge.Label1.Caption = Int(100 * I / R) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / R)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    For J = 1 To N
    ST = ST + (X(I, J) - Media) * (X(I, J) - Media)
    Next J
    Next I
    
    SS = 0
    For I = 1 To R
    SS = SS + (MR(I) - Media) * (MR(I) - Media)
    Next I
    
    SS = N * SS
    SN = ST - SS
    VS = SS / (R - 1)
    VN = SN / (R * (N - 1))
    F = VS / VN: N1 = R - 1: N2 = R * (N - 1): GoSub a4200
    
        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Analisi della varianza a un fattore (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File         : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella      : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        For I = 1 To NUMCOLONNE
            If I >= 1 And I <= 9 Then
                Text4.Text = Text4.Text & "Variabile " & I & "  : " & NomeVariabile(I) & Chr(13) & Chr(10)
            End If
            If I = 10 Then
                Text4.Text = Text4.Text & "Variabile " & I & " : " & NomeVariabile(I) & Chr(13) & Chr(10)
            End If
        Next I
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        For I = 1 To NUMCOLONNE
            If I >= 1 And I <= 9 Then
                Text4.Text = Text4.Text & "Media variabile " & I & "                   = " & Format(MR(I), "#####0.0####") & Chr(13) & Chr(10)
            End If
            If I = 10 Then
                Text4.Text = Text4.Text & "Media variabile " & I & "                  = " & Format(MR(I), "#####0.0####") & Chr(13) & Chr(10)
            End If
        Next I
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Gradi di libertà del numeratore     = " & N1 & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Gradi di libertà del denominatore   = " & N2 & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Varianza spiegata (tra le medie)    = " & Format(VS, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Varianza residua                    = " & Format(VN, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Test F (rapporto spiegata/residua)  = " & Format(VS / VN, "#####0.0####") & Chr(13) & Chr(10)
        If VS = 0 Then P2 = 1
        If P2 >= 0.00001 Then
            Text4.Text = Text4.Text & "Probabilità p                       = " & Format(P2, "#####0.0####") & Chr(13) & Chr(10)
        Else
            P2 = 0.00001
            Text4.Text = Text4.Text & "Probabilità p                       < " & Format(P2, "#####0.0####") & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Analisi della varianza a un fattore (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        Screen.MousePointer = 1

Exit Sub

a4200:
a4210:
a4220:
    PI = 3.141592654: K = 1 / 3
a4230:
    U = (1 - 2 / (9 * N2)) * F ^ K - 1 + 2 / (9 * N1)
a4240:
    U = U / Sqr(Abs(2 / (9 * N1) + 2 / (9 * N2) * F ^ (2 * K)))
a4250:
    t = 0: If U >= 0 Then GoTo a4270
a4260:
    t = 1: U = Abs(U)
a4270:
    U2 = U * U
a4280:
    S1 = Exp(-U2 / 2) / Sqr(2 * PI)
a4290:
    I = 1: A = U: C = U
a4300:
    I = I + 2
a4310:
    G = C: B = A
a4320:
    B = A
a4330:
    A = B / I * U2
a4340:
    C = G + A
a4350:
    If C > G Then GoTo a4300
a4360:
    P2 = 0.5 - C * S1
a4370:
    If t <> 1 Then GoTo a4390
a4380:
    P2 = 1 - P2
a4390:
    Return

GestioneErroriANOVAUno:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico dei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        Exit Sub

End Sub

Private Sub StatisticaComponenti_Click()
    On Error GoTo GestioneErroriComponenti
    Dim SO, SJ, SD, SW, KK, NN
    Dim I, J, Media, MD, MW, DW
    Dim CD, DD, DT
    ReDim XI(NUMCOLONNE)

    Screen.MousePointer = 11
    Text4.Visible = False

    Cls
    
    FormGauge.Show
    FormGauge.Caption = "Componenti della variabilità"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' Componenti della variabilità

    KK = NUMCOLONNE
    NN = NumDati(1)
        
        FormGauge!Label6.Caption = "[1/3]"
    SO = 0: SJ = 0: SD = 0: SW = 0
    For I = 1 To KK
        FormGauge.Label1.Caption = Int(100 * I / KK) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / KK)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    For J = 1 To NN
    SO = SO + X(I, J)
    Next J
    Next I
    Media = SO / (NN * KK)

        FormGauge!Label6.Caption = "[2/3]"
    For I = 1 To KK
        FormGauge.Label1.Caption = Int(100 * I / KK) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / KK)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    For J = 1 To NN
    SJ = SJ + X(I, J)
    Next J
    XI(I) = SJ / NN
    SJ = 0
    Next I

    For I = 1 To KK
    SD = SD + (XI(I) - Media) ^ 2
    Next I
    MD = (NN * SD) / (KK - 1)
    
        FormGauge!Label6.Caption = "[3/3]"
    For I = 1 To KK
        FormGauge.Label1.Caption = Int(100 * I / KK) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / KK)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    For J = 1 To NN
    SW = SW + (X(I, J) - XI(I)) ^ 2
    Next J
    Next I
    
    MW = SW / (KK * (NN - 1))
    DW = Sqr(MW)
    CD = (MD - MW) / NN: If CD < 0 Then CD = 0
    DD = Sqr(CD)
    DT = Sqr(MW + CD)
        
        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Analisi delle componenti della variabilità (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File         : " & NomeDelFile & Chr(13) & Chr(10)
            If NomeDellaTabella <> "" Then
                Text4.Text = Text4.Text & "Tabella      : " & NomeDellaTabella & Chr(13) & Chr(10)
            End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        For I = 1 To NUMCOLONNE
            If I >= 1 And I <= 9 Then
                Text4.Text = Text4.Text & "Variabile " & I & "  : " & NomeVariabile(I) & Chr(13) & Chr(10)
            End If
            If I = 10 Then
                Text4.Text = Text4.Text & "Variabile " & I & " : " & NomeVariabile(I) & Chr(13) & Chr(10)
            End If
        Next I
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Media generale delle osservazioni      = " & Format(Media, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Deviazione standard : entro le colonne = " & Format(DW, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                      tra le colonne   = " & Format(DD, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                      totale           = " & Format(DT, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "CV %                : entro le colonne = " & Format(100 * DW / Media, "#####0.0####") & " %" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                      tra le colonne   = " & Format(100 * DD / Media, "#####0.0####") & " %" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                      totale           = " & Format(100 * DT / Media, "#####0.0####") & " %" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Analisi delle componenti della variabilità (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        Screen.MousePointer = 1

Exit Sub

GestioneErroriComponenti:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico dei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        Exit Sub

End Sub

Private Sub StatisticaDuplicati_Click()
    On Error GoTo GestioneErroriDuplicati
    Dim NP, NS, I, SO, SOD, SOQ
    Dim Media, DiffMedia, DevSt
    Screen.MousePointer = 11
    Text4.Visible = False

    Cls
    
    NP = NumDati(1): NS = NumDati(2)
    
    If NP = NS Then

    FormGauge.Show
    FormGauge.Caption = "Variabilità entro duplicati"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
    SO = 0: SOD = 0: SOQ = 0
    
    FormGauge!Label6.Caption = "[1/1]"
    For I = 1 To NP
        SO = SO + X(1, I) + X(2, I)
        SOD = SOD + (X(1, I) - X(2, I))
        SOQ = SOQ + (X(1, I) - X(2, I)) ^ 2
        FormGauge.Label1.Caption = Int(100 * I / NP) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / NP)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    Next I
        
        Media = SO / (NP * 2): DiffMedia = Abs(SOD / NP): DevSt = Sqr(SOQ / (NP * 2))

        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Analisi della varianza, variabilità entro duplicati (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File                        : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella                     : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Primo valore delle coppie   : " & NomeVariabile(1) & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Secondo valore delle coppie : " & NomeVariabile(2) & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Media dei valori                 = " & Format(Media, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Differenze : media               = " & Format(DiffMedia, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "             deviazione standard = " & Format(DevSt, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "             CV %                = " & Format(100 * DevSt / Media, "#####0.00") & " %" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Analisi della varianza, variabilità entro duplicati (*fine*)" & Chr(13) & Chr(10)

        Unload FormGauge
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        Screen.MousePointer = 1
        
        Else
            Screen.MousePointer = 1
            MsgBox "Per la variabilità entro duplicati" + Chr(10) + "le due variabili devono contenere" + Chr(10) + "lo stesso numero di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        End If

Exit Sub

GestioneErroriDuplicati:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico dei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        Exit Sub


End Sub

Private Sub StatisticaRapportoTraVarianze_Click()
    On Error GoTo GestioneErroriRapportoTraVarianze
    Dim NP, SO, SOQ, MP, SP, VP
    Dim MS, NS, SS, VS, LN, LD, VN, VD
    Dim I, F, N1, N2, XX, PI, K, U
    Dim U2, S1, A, B, C, G, N, Q, R
    Dim T7, J2, A3, S2, J, P1
    Dim t, P2 As Double

    Screen.MousePointer = 11
    Text4.Visible = False

    Cls
    
    FormGauge.Show
    FormGauge.Caption = "Rapporto tra varianze"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

2570:
    NP = NumDati(1): NS = NumDati(2)
2580:
    SO = 0: SOQ = 0
    
    FormGauge!Label6.Caption = "[1/2]"
    For I = 1 To NP
        SO = SO + X(1, I)
        SOQ = SOQ + X(1, I) * X(1, I)
        FormGauge.Label1.Caption = Int(100 * I / NP) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / NP)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    Next I
    MP = SO / NP: VP = (SOQ - SO * SO / NP) / (NP - 1): SP = Sqr(VP)
2590:
    SO = 0: SOQ = 0
    
    FormGauge!Label6.Caption = "[2/2]"
    For I = 1 To NS
        SO = SO + X(2, I)
        SOQ = SOQ + X(2, I) * X(2, I)
        FormGauge.Label1.Caption = Int(100 * I / NS) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / NS)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    Next I
    MS = SO / NS: VS = (SOQ - SO * SO / NS) / (NS - 1): SS = Sqr(VS)
        
        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Analisi della varianza, rapporto tra varianze (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File             : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella          : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Primo campione   : " & NomeVariabile(1) & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Secondo campione : " & NomeVariabile(2) & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Primo campione   : numero dei dati     = " & NP & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   media               = " & Format(MP, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   deviazione standard = " & Format(SP, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   CV %                = " & Format(100 * SP / MP, "#####0.00") & " %" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Secondo campione : numero dei dati     = " & NS & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   media               = " & Format(MS, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   deviazione standard = " & Format(SS, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   CV %                = " & Format(100 * SS / MS, "#####0.00") & " %" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
4000:
4010:
4020:
4050:
    LN = NP - 1: LD = NS - 1: VN = VP: VD = VS
4060:
    If VN < VD Then GoTo 4180
4070:
    F = VN / VD: N1 = LN: N2 = LD
4080:
    GoSub 4190
        
        Text4.Text = Text4.Text & "Gradi di libertà del numeratore        = " & LN & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Gradi di libertà del denominatore      = " & LD & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Varianza al numeratore                 = " & Format(VN, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Varianza al denominatore               = " & Format(VD, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Test F (rapporto tra le varianze)      = " & Format(VN / VD, "#####0.0####") & Chr(13) & Chr(10)
        If P2 >= 0.00001 Then
            Text4.Text = Text4.Text & "Probabilità p                          = " & Format(P2, "#####0.0####") & Chr(13) & Chr(10)
        Else
            P2 = 0.00001
            Text4.Text = Text4.Text & "Probabilità p                          < " & Format(P2, "#####0.0####") & Chr(13) & Chr(10)
        End If

4150:
    GoTo 5000
4180:
    XX = VN: VN = VD: VD = XX: XX = LN: LN = LD: LD = XX: GoTo 4070
4190:
4200:
4210:
4220:
    PI = 3.141592654: K = 1 / 3
4230:
    U = (1 - 2 / (9 * N2)) * F ^ K - 1 + 2 / (9 * N1)
4240:
    U = U / Sqr(Abs(2 / (9 * N1) + 2 / (9 * N2) * F ^ (2 * K)))
4250:
    t = 0: If U >= 0 Then GoTo 4270
4260:
    t = 1: U = Abs(U)
4270:
    U2 = U * U
4280:
    S1 = Exp(-U2 / 2) / Sqr(2 * PI)
4290:
    I = 1: A = U: C = U
4300:
    I = I + 2
4310:
    G = C: B = A
4320:
    B = A
4330:
    A = B / I * U2
4340:
    C = G + A
4350:
    If C > G Then GoTo 4300
4360:
    P2 = 0.5 - C * S1
4370:
    If t <> 1 Then GoTo 4390
4380:
    P2 = 1 - P2
4390:
    Return

5000:
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Analisi della varianza, rapporto tra varianze (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        Screen.MousePointer = 1

Exit Sub

GestioneErroriRapportoTraVarianze:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico dei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        Exit Sub

End Sub

