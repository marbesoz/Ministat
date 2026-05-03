VERSION 5.00
Begin VB.Form FormT 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Confronto tra medie"
   ClientHeight    =   6075
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
   Icon            =   "FORMT.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6075
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
      Begin VB.Menu StatisticaTapp 
         Caption         =   "Test t di Student per dati &appaiati"
      End
      Begin VB.Menu StatisticaTind 
         Caption         =   "Test t di Student per campioni &indipendenti"
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
Attribute VB_Name = "FormT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim X() As Double
    Dim Y() As Double
    Dim Riga, ColonnaX, ColonnaY, FineDatiX, FineDatiY
    Dim Z1, Z2
    Dim XMI, XMA, YMI, YMA
    Dim InAscisse, InOrdinate
    Dim FlagStatistica

Private Sub AiutoCalcolatrice_Click()
    Dim Calcolatrice
    On Error GoTo ErroreCalcolatriceNove:
    
    Calcolatrice = Shell("Calc.exe", 1)
        
Exit Sub

ErroreCalcolatriceNove:
    MsgBox "Calcolatrice non disponibile.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcolatrice"
Exit Sub


End Sub

Private Sub AiutoGuida_Click()

    FormPrincipale!CMDialog1.HelpFile = App.Path & "\ministat.hlp"
    FormPrincipale!CMDialog1.HelpCommand = cdlHelpContents
    FormPrincipale!CMDialog1.ShowHelp

End Sub

Private Sub FileEsci_Click()
    ReDim X(0)
    ReDim Y(0)
    ReDim AA(0), SUP(0), INF(0)
    FormPrincipale.Show
    Unload FormT
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
End Sub

Private Sub FileStampa_Click()
    On Error GoTo ErroreFileStampaT

    If FlagStatistica = True Then
        Printer.Print Text4.Text
        Printer.EndDoc
    End If

    Exit Sub
        
ErroreFileStampaT:
    Exit Sub

End Sub

Private Sub Form_Load()
    ReDim X(1 To NUM_RIGHE - 1) As Double
    ReDim Y(1 To NUM_RIGHE - 1) As Double
    Dim I

    Screen.MousePointer = 11
    AnnullatoCaricamentoForm = False
    FlagStatistica = False

' centra il form
    FormT.Top = Screen.Height / 2 - Height / 2
    FormT.Left = Screen.Width / 2 - Width / 2
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
' carica i dati nei vettori X() e Y()
    FormGauge.Show
    FormGauge.Caption = "Confronto tra medie"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    FineDatiX = False
    FineDatiY = False
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
' controlli finali sui dati
' per il test t devono esserci almeno due dati
    If Z1 >= 2 And Z2 >= 2 Then
        XMI = X(1): XMA = X(1)
        For I = 2 To Z1
            If X(I) < XMI Then XMI = X(I)
            If X(I) > XMA Then XMA = X(I)
        Next
        YMI = Y(1): YMA = Y(1)
        For I = 2 To Z2
            If Y(I) < YMI Then YMI = Y(I)
            If Y(I) > YMA Then YMA = Y(I)
        Next
        Unload FormGauge
        Text4.Visible = True
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
        FormT.Show
        ' FormPrincipale.Hide
            If Z1 = Z2 Then
                StatisticaTapp_Click
            Else
                StatisticaTind_Click
            End If
        Screen.MousePointer = 1
        Exit Sub
    Else
        Unload FormGauge
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        MsgBox "Per il confronto tra medie" + Chr(10) + "sono necessari almeno due dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Confronto tra medie non eseguibile"
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
    End If

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
    
    If Z1 = Z2 Then
        StatisticaTapp.Enabled = True
        StatisticaTind.Enabled = True
    Else
        StatisticaTapp.Enabled = False
        StatisticaTind.Enabled = True
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

Private Sub StatisticaTapp_Click()
    On Error GoTo GestioneErroriTapp
    Dim SX, SY, B, C, D, V, DM, N
    Dim I, J, K, PI, A3, J2, S1, S2
    Dim t, T7, P1, MinDiff, MaxDiff, MediaDiff
    Dim TT, PT  As Double

    Screen.MousePointer = 11
    Text4.Visible = False

    Cls
    
    FormGauge.Show
    FormGauge.Caption = "Test per dati appaiati"
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Caption = "0%"
    FormGauge!Gauge1.Value = 0
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    SX = 0: SY = 0: B = 0: C = 0
    MinDiff = X(1) - Y(1): MaxDiff = X(1) - Y(1)
    
    For I = 1 To Z1
        SX = SX + X(I): SY = SY + Y(I): D = X(I) - Y(I): B = B + D: C = C + D * D
        If MinDiff > X(I) - Y(I) Then MinDiff = X(I) - Y(I)
        If MaxDiff < X(I) - Y(I) Then MaxDiff = X(I) - Y(I)
        FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / Z1)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    Next I

    MediaDiff = B / Z1
3070:
    V = (C - B * B / Z1) / (Z1 - 1): DM = B / Z1
    
    If B <> 0 And V = 0 Then TT = 999999.99999: PT = 0: GoTo tt3300
    If B = 0 And V = 0 Then TT = 0: PT = 1: GoTo tt3300
    
    TT = (B / Z1) / Sqr(V / Z1)
    t = TT: N = Z1 - 1
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
        Text4.Text = Text4.Text & "Confronto tra medie, test t di Student per dati appaiati (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File                        : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella                     : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Primo valore delle coppie   : " & InAscisse & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Secondo valore delle coppie : " & InOrdinate & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Numero delle coppie di dati            = " & Z1 & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Differenze tra le coppie    : minima   = " & Format(MinDiff, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                              massima  = " & Format(MaxDiff, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                              media    = " & Format(MediaDiff, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Gradi di libertà                       = " & Z1 - 1 & Chr(13) & Chr(10)
        If TT < 999999.99999 Then
            Text4.Text = Text4.Text & "Test t di Student                      = " & Format(Abs(TT), "#####0.0####") & Chr(13) & Chr(10)
        Else
            Text4.Text = Text4.Text & "Test t di Student                      > 999999.99999" & Chr(13) & Chr(10)
        End If
        If PT >= 0.00001 Then
            Text4.Text = Text4.Text & "Probabilità p                          = " & Format(PT, "#####0.0####") & Chr(13) & Chr(10)
        Else
            PT = 0.00001
            Text4.Text = Text4.Text & "Probabilità p                          < " & Format(PT, "#####0.0####") & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Confronto tra medie, test t di Student per dati appaiati (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        Screen.MousePointer = 1

    Exit Sub

GestioneErroriTapp:
    If Erl = 3070 Then
        TT = 0: PT = 1: Resume tt3300
    Else
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico nei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
    End If
        
    Exit Sub

End Sub

Private Sub StatisticaTind_Click()
    On Error GoTo GestioneErroriTind
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
    FormGauge.Caption = "Test per campioni indipendenti"
    FormGauge!Label6.Caption = "[1/3]"
    FormGauge!Label1.Caption = "0%"
    FormGauge!Gauge1.Value = 0
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

2570:
    NP = Z1: NS = Z2
2580:
    SO = 0: SOQ = 0
    
    For I = 1 To NP
        SO = SO + X(I)
        SOQ = SOQ + X(I) * X(I)
        FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / Z1)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    Next I
    MP = SO / NP: VP = (SOQ - SO * SO / NP) / (NP - 1): SP = Sqr(VP)
2590:
    SO = 0: SOQ = 0
    
    FormGauge!Label6.Caption = "[2/3]"
    For I = 1 To NS
        SO = SO + Y(I)
        SOQ = SOQ + Y(I) * Y(I)
        FormGauge.Label1.Caption = Int(100 * I / Z2) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / Z2)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    Next I
    MS = SO / NS: VS = (SOQ - SO * SO / NS) / (NS - 1): SS = Sqr(VS)
        
        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Confronto tra medie, test t di Student per campioni indipendenti (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File             : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella          : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Primo campione   : " & InAscisse & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Secondo campione : " & InOrdinate & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Primo campione   : numero dei dati     = " & NP & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   minimo              = " & Format(XMI, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   massimo             = " & Format(XMA, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   media               = " & Format(MP, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   deviazione standard = " & Format(SP, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   CV %                = " & Format(100 * SP / MP, "#####0.00") & " %" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Secondo campione : numero dei dati     = " & NS & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   minimo              = " & Format(YMI, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   massimo             = " & Format(YMA, "#####0.0####") & Chr(13) & Chr(10)
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
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        FormGauge!Label6.Caption = "[3/3]"
        FormGauge.Label1.Caption = "30%"
        FormGauge!Gauge1.Value = 30
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

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
5010:
5020:
5050:
    I = (VP * (NP - 1) + VS * (NS - 1)) / (NP + NS - 2)
5060:
    t = (MP - MS) / Sqr(I / NP + I / NS): N = NP + NS - 2
5070:
    GoSub 6190
        
        Text4.Text = Text4.Text & "Media dei dati del Primo campione      = " & Format(MP, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Media dei dati del Secondo campione    = " & Format(MS, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Differenza tra le medie (P - S)        = " & Format((MP - MS), "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        
        Text4.Text = Text4.Text & "Gradi di libertà                       = " & N & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Test t per varianze omogenee           = " & Format(Abs(t), "#####0.0####") & Chr(13) & Chr(10)
        If P2 >= 0.00001 Then
            Text4.Text = Text4.Text & "Probabilità p                          = " & Format(P2, "#####0.0####") & Chr(13) & Chr(10)
        Else
            P2 = 0.00001
            Text4.Text = Text4.Text & "Probabilità p                          < " & Format(P2, "#####0.0####") & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        FormGauge.Label1.Caption = "60%"
        FormGauge!Gauge1.Value = 60
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

6000:
6010:
6020:
6050:
    Q = VP / NP: R = VS / NS
6060:
    t = (MP - MS) / Sqr(Q + R)
6070:
    N = ((Q + R) * (Q + R) / (Q * Q / (NP - 1) + R * R / (NS - 1))): N = Int(N + 0.5)
6075:
    GoSub 6190

        Text4.Text = Text4.Text & "Gradi di libertà                       = " & N & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Test t per varianze non omogenee       = " & Format(Abs(t), "#####0.0####") & Chr(13) & Chr(10)
        If P2 >= 0.00001 Then
            Text4.Text = Text4.Text & "Probabilità p                          = " & Format(P2, "#####0.0####") & Chr(13) & Chr(10)
        Else
            P2 = 0.00001
            Text4.Text = Text4.Text & "Probabilità p                          < " & Format(P2, "#####0.0####") & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Confronto tra medie, test t di Student per campioni indipendenti (*fine*)" & Chr(13) & Chr(10)
        
        
        FormGauge.Label1.Caption = "100%"
        FormGauge!Gauge1.Value = 100
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

6120:
    GoTo 6969
6190:
6200:
6210:
6220:
    PI = 3.141592654
6230:
    T7 = Atn(Abs(t) / Sqr(N))
6240:
    A3 = 0
6250:
    J2 = N - 2
6260:
    If N <> Int(N / 2) * 2 Then GoTo 6330
6270:
    For K = J2 To 2 Step -2
6280:
    S1 = 0: S2 = 0
6290:
    For I = K - 1 To 1 Step -2: S1 = S1 + Log(I): Next
6300:
    For J = K To 2 Step -2: S2 = S2 + Log(J): Next
6310:
    A3 = A3 + Exp(S1 - S2) * Cos(T7) ^ K: Next
6320:
    P1 = (A3 + 1) * Sin(T7): GoTo 6410
6330:
    If N <> 1 Then GoTo 6350
6340:
    P1 = 2 * T7 / PI: GoTo 6410
6350:
    For K = J2 To 2 Step -2
6360:
        S1 = 0: S2 = 0
6370:
        For I = K - 1 To 2 Step -2: S1 = S1 + Log(I): Next I
6380:
        For J = K To 1 Step -2: S2 = S2 + Log(J): Next J
6390:
        A3 = A3 + Exp(S1 - S2) * Cos(T7) ^ K
    Next K
6400:
    P1 = ((A3 + Cos(T7)) * Sin(T7) + T7) * 2 / PI
6410:
    P2 = 1 - P1
6420:
    Return

6969:
        
        Unload FormGauge
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        Screen.MousePointer = 1
        Exit Sub

GestioneErroriTind:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico nei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        Exit Sub

End Sub

