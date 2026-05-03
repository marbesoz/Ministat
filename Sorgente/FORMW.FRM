VERSION 5.00
Begin VB.Form FormW 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Confronto tra mediane"
   ClientHeight    =   6030
   ClientLeft      =   1230
   ClientTop       =   2040
   ClientWidth     =   7230
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
   Icon            =   "FORMW.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6030
   ScaleWidth      =   7230
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
         Caption         =   "Test di Wilcoxon per dati &appaiati"
      End
      Begin VB.Menu StatisticaTind 
         Caption         =   "Test di Wilcoxon per campioni &indipendenti"
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
Attribute VB_Name = "FormW"
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
    On Error GoTo ErroreCalcolatriceDieci:
    
    Calcolatrice = Shell("Calc.exe", 1)
        
Exit Sub

ErroreCalcolatriceDieci:
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
    Unload FormW
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
End Sub

Private Sub FileStampa_Click()
    On Error GoTo ErroreFileStampaW

    If FlagStatistica = True Then
        Printer.Print Text4.Text
        Printer.EndDoc
    End If

    Exit Sub

ErroreFileStampaW:
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
    FormW.Top = Screen.Height / 2 - Height / 2
    FormW.Left = Screen.Width / 2 - Width / 2
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
    FormGauge.Caption = "Confronto tra mediane"
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
        FormW.Show
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
        MsgBox "Per il confronto tra mediane" + Chr(10) + "sono necessari almeno due dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Confronto tra mediane non eseguibile"
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
    On Error GoTo GestioneErroriWapp
    ReDim XX(Z1) As Double
    ReDim YY(Z1) As Double
    ReDim AA(Z1) As Double
    ReDim EE(Z1) As Double
    ReDim FF(Z1) As Double
    ReDim PercX(Z1) As Double
    Dim I, D, J, L, XXX, YYY
    Dim ValoreInizialeD, Dummy, DummyXX
    Dim N, M1, M2, SP, SN, t, M, S
    Dim XXXX, U, PI, U2, P1, A, C, B, p, V
    Dim W, PE, MED
    Dim WI, PW As Double

    Screen.MousePointer = 11
    Text4.Visible = False

    Cls
    
    FormGauge.Show
    FormGauge.Caption = "Test per dati appaiati"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    FormGauge!Label6.Caption = "[1/4]"
    For I = 1 To Z1
                    
        XX(I) = X(I) - Y(I)

                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    Next I

' inizio ordinamento dati
wapx330:
    FormGauge!Label6.Caption = "[2/4]"
    D = 1
wapx340:
    D = 2 * D: If D <= Z1 Then GoTo wapx340
wapx350:
    D = Int((D - 1) / 2): If D = 0 Then GoTo wapx420
wapx360:
    For I = 1 To Z1 - D
wapx370:
        J = I
wapx380:
        L = J + D
        If XX(J) <= XX(L) Then GoTo wapx400
                    
wapx390:
            XXX = XX(J): XX(J) = XX(L): XX(L) = XXX

            If I > ValoreInizialeD And 100 * I / Z1 >= 1 Then
                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If
            J = J - D: If J > 0 Then GoTo wapx380
wapx400:
    Next I
wapx410:
    ValoreInizialeD = Z1 - D
    GoTo wapx350
wapx420:
' fine ordinamento dati
                
                FormGauge.Label1.Caption = "100%"
                FormGauge!Gauge1.Value = 100
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            FormGauge!Label6.Caption = "[3/4]"
            For I = 1 To Z1
                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                M1 = 0: M2 = 0
petx840:
                    For J = 1 To Z1
petx850:
                       If XX(J) > XX(I) Then GoTo petx890
petx860:
                       If XX(J) = XX(I) Then GoTo petx880
petx870:
                       M1 = M1 + 1: GoTo petx890
petx880:
                       M2 = M2 + 1
petx890:
                    Next J
petx900:
                PercX(I) = M1 + 0.5 * M2 + 0.5
            Next I

            For I = 1 To Z1
                PercX(I) = Format(100 * PercX(I) / (Z1 + 1), "##0.0####")
            Next I
            
            FormGauge.Label1.Caption = "100%"
            FormGauge!Gauge1.Value = 100
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
        
        Dummy = Format(50, "##0.0")
        For I = 1 To Z1 - 1
            If Dummy = PercX(I) Then
                DummyXX = XX(I)
                Exit For
            End If
            If Dummy = PercX(I + 1) Then
                DummyXX = XX(I + 1)
                Exit For
            End If
            If Dummy > PercX(I) And Dummy < PercX(I + 1) Then
                DummyXX = XX(I) + ((XX(I + 1) - XX(I)) * (Dummy - PercX(I)) / (PercX(I + 1) - PercX(I)))
                Exit For
            End If
        Next I

    MED = DummyXX

3050:
    For I = 1 To Z1: AA(I) = X(I) - Y(I): Next I
3060:
    N = 0: For I = 1 To Z1: If AA(I) = 0 Then GoTo 3080
3070:
    N = N + 1: EE(N) = AA(I)
3080:
    Next
3090:
    FormGauge!Label6.Caption = "[4/4]"
    For I = 1 To N
        
        FormGauge.Label1.Caption = Int(100 * I / N) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / N)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
        M1 = 0: M2 = 0
3100:
    For J = 1 To N
3110:
    If Abs(EE(J)) > Abs(EE(I)) Then GoTo 3150
3120:
    If Abs(EE(J)) = Abs(EE(I)) Then GoTo 3140
3130:
    M1 = M1 + 1: GoTo 3150
3140:
    M2 = M2 + 1
3150:
    Next
3160:
    FF(I) = (M1 + 0.5 * M2 + 0.5) * Sgn(EE(I))
3170:
    Next
3180:
    SP = 0: SN = 0: For I = 1 To N
3190:
    If FF(I) < 0 Then SN = SN + Abs(FF(I))
3200:
    If FF(I) > 0 Then SP = SP + FF(I)
3210:
    Next
3220:
    t = SP: If SN < SP Then t = SN
3230:
    M = N * (N + 1) / 4
3240:
    S = Sqr((2 * N + 1) * M / 6)
3250:
    WI = (M - t - 0.5) / S
3260:
' PROBABILITA' DI WI
3270:
    XXXX = WI
3280:
    M = 0
3290:
    S = 1
3300:
    U = Abs((XXXX - M) / S)
3310:
    PI = 3.141592654
3320:
    If U < 7 Then GoTo 3340
3330:
    PW = 0: GoTo 3400
3340:
    U2 = U * U
3350:
    P1 = Exp(-U2 / 2) / (XXXX * S * Sqr(2 * PI))
3360:
    I = 1: A = U: C = U
3370:
    I = I + 2
3380:
    D = C: B = A: A = B / I * U2: C = D + A: If C > D Then GoTo 3370
3390:
    PW = 2 * (0.5 - C * P1 * S * XXXX)
3400:
    
3480:
'    P = 50: P = P / 100: V = P * (z1 + 1): W = Int(V): I = V - W: PE = XX(W) + I * (XX(W + 1) - XX(W))
'3490 :
'    MEP = PE
'3560 :
'    P = 50: P = P / 100: V = P * (z1 + 1): W = Int(V): I = V - W: PE = YY(W) + I * (YY(W + 1) - YY(W))
'3570 :
'    MES = PE
'3600 :
'    DIFF = MEP - MES

        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Confronto tra mediane, test di Wilcoxon per dati appaiati (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File                        : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella                     : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Primo valore delle coppie   : " & InAscisse & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Secondo valore delle coppie : " & InOrdinate & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Numero delle coppie di dati             = " & Z1 & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Differenze tra le coppie   : minima     = " & Format(XX(1), "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                             massima    = " & Format(XX(Z1), "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                             mediana    = " & Format(MED, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        If MED = 0 Then WI = 0: PW = 1
        Text4.Text = Text4.Text & "Test di Wilcoxon                        = " & Format(WI, "#####0.0####") & Chr(13) & Chr(10)
        If PW >= 0.00001 Then
            Text4.Text = Text4.Text & "Probabilità p                           = " & Format(PW, "#####0.0####") & Chr(13) & Chr(10)
        Else
            PW = 0.00001
            Text4.Text = Text4.Text & "Probabilità p                           < " & Format(PW, "#####0.0####") & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Confronto tra mediane, test di Wilcoxon per dati appaiati (*fine*)" & Chr(13) & Chr(10)
        
        
        Unload FormGauge
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        Screen.MousePointer = 1

        Exit Sub

GestioneErroriWapp:
    If Erl = 3250 Then
        WI = 0: PW = 1: Resume 3480
    Else
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico nei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
    End If

    Exit Sub

End Sub

Private Sub StatisticaTind_Click()
    On Error GoTo GestioneErroriWind
    ReDim XX(Z1) As Double
    ReDim YY(Z2) As Double
    ReDim EE(Z1 + Z2) As Double
    ReDim FF(Z1 + Z2) As Double
    ReDim PercX(Z1) As Double
    ReDim PercY(Z2) As Double
    Dim I, D, J, L, XXX, YYY, ValoreInizialeD, DIFF
    Dim NP, MINP, MAXP, MEDP
    Dim NS, MINS, MAXS, MEDS
    Dim p, V, W, PE, M1, M2, SP, SS, t, N1, N2
    Dim t1, t2, M, S, XXXX, U, PI
    Dim U2, P1, A, B, C
    Dim Dummy, DummyXX, DummyYY As Double
    Dim WI, PW As Double

    Screen.MousePointer = 11
    Text4.Visible = False

    Cls
    
    FormGauge.Show
    FormGauge.Caption = "Test per campioni indipendenti"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    FormGauge!Label6.Caption = "[1/6]"
    For I = 1 To Z1
                    
        XX(I) = X(I)

                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    Next I

    FormGauge!Label6.Caption = "[2/6]"
    For I = 1 To Z2
                    
        YY(I) = Y(I)

                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    Next I

' inizio ordinamento dati
windx330:
    FormGauge!Label6.Caption = "[3/6]"
    D = 1
windx340:
    D = 2 * D: If D <= Z1 Then GoTo windx340
windx350:
    D = Int((D - 1) / 2): If D = 0 Then GoTo windx420
windx360:
    For I = 1 To Z1 - D
windx370:
        J = I
windx380:
        L = J + D
        If XX(J) <= XX(L) Then GoTo windx400
                    
windx390:
            XXX = XX(J): XX(J) = XX(L): XX(L) = XXX

            If I > ValoreInizialeD And 100 * I / Z1 >= 1 Then
                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If
            J = J - D: If J > 0 Then GoTo windx380
windx400:
    Next I
windx410:
    ValoreInizialeD = Z1 - D
    GoTo windx350
windx420:
' fine ordinamento dati
                
                FormGauge.Label1.Caption = "100%"
                FormGauge!Gauge1.Value = 100
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            FormGauge!Label6.Caption = "[4/6]"
            For I = 1 To Z1
                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                M1 = 0: M2 = 0
pesx840:
                    For J = 1 To Z1
pesx850:
                       If XX(J) > XX(I) Then GoTo pesx890
pesx860:
                       If XX(J) = XX(I) Then GoTo pesx880
pesx870:
                       M1 = M1 + 1: GoTo pesx890
pesx880:
                       M2 = M2 + 1
pesx890:
                    Next J
pesx900:
                PercX(I) = M1 + 0.5 * M2 + 0.5
            Next I

            For I = 1 To Z1
                PercX(I) = Format(100 * PercX(I) / (Z1 + 1), "##0.0####")
            Next I
            
            FormGauge.Label1.Caption = "100%"
            FormGauge!Gauge1.Value = 100
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' inizio ordinamento dati
windy330:
    FormGauge!Label6.Caption = "[5/6]"
    D = 1
windy340:
    D = 2 * D: If D <= Z2 Then GoTo windy340
windy350:
    D = Int((D - 1) / 2): If D = 0 Then GoTo windy420
windy360:
    For I = 1 To Z2 - D
windy370:
        J = I
windy380:
        L = J + D
        If YY(J) <= YY(L) Then GoTo windy400
                    
windy390:
            YYY = YY(J): YY(J) = YY(L): YY(L) = YYY

            If I > ValoreInizialeD And 100 * I / Z2 >= 1 Then
                FormGauge.Label1.Caption = Int(100 * I / Z2) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z2)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If
            J = J - D: If J > 0 Then GoTo windy380
windy400:
    Next I
windy410:
    ValoreInizialeD = Z2 - D
    GoTo windy350
windy420:
' fine ordinamento dati
                
                FormGauge.Label1.Caption = "100%"
                FormGauge!Gauge1.Value = 100
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            FormGauge!Label6.Caption = "[6/6]"
            For I = 1 To Z2
                FormGauge.Label1.Caption = Int(100 * I / Z2) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z2)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                M1 = 0: M2 = 0
pesy840:
                    For J = 1 To Z2
pesy850:
                       If YY(J) > YY(I) Then GoTo pesy890
pesy860:
                       If YY(J) = YY(I) Then GoTo pesy880
pesy870:
                       M1 = M1 + 1: GoTo pesy890
pesy880:
                       M2 = M2 + 1
pesy890:
                    Next J
pesy900:
                PercY(I) = M1 + 0.5 * M2 + 0.5
            Next I

            For I = 1 To Z2
                PercY(I) = Format(100 * PercY(I) / (Z2 + 1), "##0.0####")
            Next I
            
            FormGauge.Label1.Caption = "100%"
            FormGauge!Gauge1.Value = 100
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
    NP = Z1: NS = Z2

33:
    For I = 1 To NP: EE(I) = XX(I): Next
34:
    For I = NP + 1 To NP + NS: EE(I) = YY(I - NP): Next
35:
    GoSub 59
36:
'    P = 50: P = P / 100: V = P * (NP + 1): W = Int(V): I = V - W: PE = XX(W) + I * (XX(W + 1) - XX(W))
'37 :
'    MEDP = PE
        Dummy = Format(50, "##0.0")
        For I = 1 To Z1 - 1
            If Dummy = PercX(I) Then
                DummyXX = XX(I)
                Exit For
            End If
            If Dummy = PercX(I + 1) Then
                DummyXX = XX(I + 1)
                Exit For
            End If
            If Dummy > PercX(I) And Dummy < PercX(I + 1) Then
                DummyXX = XX(I) + ((XX(I + 1) - XX(I)) * (Dummy - PercX(I)) / (PercX(I + 1) - PercX(I)))
                Exit For
            End If
        Next I

    MEDP = DummyXX
38:
    MINP = XX(1)
39:
    MAXP = XX(NP)
'40 :
'    P = 50: P = P / 100: V = P * (NS + 1): W = Int(V): I = V - W: PE = YY(W) + I * (YY(W + 1) - YY(W))
'41 :
'    MEDS = PE
        Dummy = Format(50, "##0.0")
        For I = 1 To Z2 - 1
            If Dummy = PercY(I) Then
                DummyYY = YY(I)
                Exit For
            End If
            If Dummy = PercY(I + 1) Then
                DummyYY = YY(I + 1)
                Exit For
            End If
            If Dummy > PercY(I) And Dummy < PercY(I + 1) Then
                DummyYY = YY(I) + ((YY(I + 1) - YY(I)) * (Dummy - PercY(I)) / (PercY(I + 1) - PercY(I)))
                Exit For
            End If
        Next I

    MEDS = DummyYY

42:
    MINS = YY(1)
43:
    MAXS = YY(NS)
44:
    DIFF = MEDP - MEDS

58:
    GoTo 99

59:
' TEST DI WILCOXON PER CAMPIONI INDIPENDENTI
60:
    For I = 1 To NP + NS: M1 = 0: M2 = 0
61:
    For J = 1 To NP + NS
62:
    If EE(J) > EE(I) Then GoTo 66
63:
    If EE(J) = EE(I) Then GoTo 65
64:
    M1 = M1 + 1: GoTo 66
65:
    M2 = M2 + 1
66:
    Next
67:
    FF(I) = M1 + 0.5 * M2 + 0.5
68:
    Next
69:
    SP = 0: SS = 0
70:
    For I = 1 To NP: SP = SP + FF(I): Next
71:
    For I = NP + 1 To NP + NS: SS = SS + FF(I): Next
72:
    If NP <> NS Then GoTo 76
73:
    t = SP
74:
    If SP > SS Then t = SS
75:
    N1 = NP: N2 = NS: GoTo 81
76:
    If NP < NS Then t1 = SP: N1 = NP: N2 = NS
77:
    If NP > NS Then t1 = SS: N1 = NS: N2 = NP
78:
    t2 = N1 * (N1 + N2 + 1) - t1
79:
    If t1 < t2 Then t = t1
80:
    If t1 > t2 Then t = t2
81:
    M = N1 * (N1 + N2 + 1) / 2
82:
    S = Sqr(N2 * M / 6)
83:
    WI = (M - t - 0.5) / S
84:
' PROBABILITA' DI WI
85:
    XXXX = WI
86:
    M = 0
87:
    S = 1
88:
    U = Abs((XXXX - M) / S)
89:
    PI = 3.141592654
90:
    If U < 7 Then GoTo 92
91:
    PW = 0: GoTo 98
92:
    U2 = U * U

If (XXXX * S * Sqr(2 * PI)) = 0 Then WI = 0: PW = 1: GoTo 36

93:
    P1 = Exp(-U2 / 2) / (XXXX * S * Sqr(2 * PI))
94:
    I = 1: A = U: C = U
95:
    I = I + 2
96:
    D = C: B = A: A = B / I * U2: C = D + A: If C > D Then GoTo 95
97:
    PW = 2 * (0.5 - C * P1 * S * XXXX)
98:
    Return

99:
        
        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Confronto tra mediane, test di Wilcoxon per campioni indipendenti (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File             : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella          : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Primo campione   : " & InAscisse & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Secondo campione : " & InOrdinate & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Primo campione   : numero dei dati    = " & NP & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   minimo             = " & Format(MINP, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   massimo            = " & Format(MAXP, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   mediana            = " & Format(MEDP, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   range              = " & Format(Abs(MAXP - MINP), "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   range %            = " & Format(Abs(100 * (MAXP - MINP) / MEDP), "#####0.00") & " %" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Secondo campione : numero dei dati    = " & NS & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   minimo             = " & Format(MINS, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   massimo            = " & Format(MAXS, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   mediana            = " & Format(MEDS, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   range              = " & Format(Abs(MAXS - MINS), "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                   range %            = " & Format(Abs(100 * (MAXS - MINS) / MEDS), "#####0.00") & " %" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Mediana dei dati del Primo campione   = " & Format(MEDP, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Mediana dei dati del Secondo campione = " & Format(MEDS, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Differenza tra le mediane (P - S)     = " & Format(DIFF, "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        If DIFF = 0 Then WI = 0: PW = 1
        Text4.Text = Text4.Text & "Test di Wilcoxon                      = " & Format(WI, "#####0.0####") & Chr(13) & Chr(10)
        If PW >= 0.00001 Then
            Text4.Text = Text4.Text & "Probabilità p                         = " & Format(PW, "#####0.0####") & Chr(13) & Chr(10)
        Else
            PW = 0.00001
            Text4.Text = Text4.Text & "Probabilità p                         < " & Format(PW, "#####0.0####") & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Confronto tra mediane, test di Wilcoxon per campioni indipendenti (*fine*)" & Chr(13) & Chr(10)

        Unload FormGauge
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        Screen.MousePointer = 1
        Exit Sub

GestioneErroriWind:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico nei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        Exit Sub

End Sub

