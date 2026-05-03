VERSION 5.00
Begin VB.Form FormChi 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Test chi-quadrato per tabelle di contingenza"
   ClientHeight    =   6135
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
   Icon            =   "FORMCHI.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6135
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
      Begin VB.Menu StatisticaChi 
         Caption         =   "&Test chi-quadrato"
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
Attribute VB_Name = "FormChi"
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
    Unload FormChi
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
End Sub

Private Sub FileStampa_Click()
    On Error GoTo ErroreFileStampaChi

    If FlagStatistica = True Then
        Printer.Print Text4.Text
        Printer.EndDoc
    End If

    Exit Sub

ErroreFileStampaChi:
    Exit Sub

End Sub

Private Sub Form_Load()
    On Error GoTo GestioneErroriChi
    Dim I, J, K, Riga, Dummy

    Screen.MousePointer = 11
    
    NUMCOLONNE = FormPrincipale!Grid1.SelEndCol - FormPrincipale!Grid1.SelStartCol + 1

    ReDim X(NUMCOLONNE, NUM_RIGHE - 1)
    ReDim NomeVariabile(NUMCOLONNE)
    ReDim NumDati(NUMCOLONNE)
    ReDim FineDati(NUMCOLONNE)

    AnnullatoCaricamentoForm = False
    FlagStatistica = False

' centra il form
    FormChi.Top = Screen.Height / 2 - Height / 2
    FormChi.Left = Screen.Width / 2 - Width / 2
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
    FormGauge.Caption = "Test chi-quadrato"
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
        If NumDati(I) < 2 Or NumDati(I) > 10 Then Dummy = False
    Next I

    If Dummy = False Then
        Unload FormGauge
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        MsgBox "Per il test chi-quadrato è necessaria" + Chr(10) + "una tabella di almeno due righe per due" + Chr(10) + "colonne e al massimo di dieci righe per" + Chr(10) + "dieci colonne.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Test chi-quadrato non eseguibile"
        RipristinaSchermo
        Exit Sub
    End If

' deve esserci lo stesso numero di dati per ciascuna variabile
    Dummy = True
    For I = 2 To NUMCOLONNE
        If NumDati(I) <> NumDati(I - 1) Then Dummy = False
    Next I

    If Dummy = False Then
        Unload FormGauge
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        MsgBox "Per il test chi-quadrato" + Chr(10) + "le colonne devono contenere" + Chr(10) + "lo stesso numero di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Test chi-quadrato non eseguibile"
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
        Exit Sub
    End If

' altrimenti...
        Unload FormGauge
        Text4.Visible = True
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
        FormChi.Show
        ' FormPrincipale.Hide
        StatisticaChi_Click
        Screen.MousePointer = 1

Exit Sub

GestioneErroriChi:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente: riprovare" + Chr(10) + "con un numero inferiore di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Test chi-quadrato non eseguibile"
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

Private Sub ModificaCopia_Click()
    
    Clipboard.Clear
    Clipboard.SetText Text4.SelText

End Sub

Private Sub ModificaTaglia_Click()
    
    Clipboard.Clear
    Clipboard.SetText Text4.SelText
    Text4.SelText = ""

End Sub

Private Sub StatisticaChi_Click()
    On Error GoTo GestioneErroriChiQuad
    ReDim F(10), G(10), H(10)
    Dim I, M, N, O, PI, J, Q2, Q1, c2
    Dim J1, S1, Q3, Z1, Q, W1, A, R1
    Dim R2, R3

    Screen.MousePointer = 11
    Text4.Visible = False

    Cls

    N = NUMCOLONNE
    M = NumDati(1)

    For I = 1 To N
    F(I) = 0: H(I) = 0
    Next I
    O = 0: PI = 3.14159
    
    For I = 1 To M
    G(I) = 0
    For J = 1 To N
    G(I) = G(I) + X(J, I)
    Next J
    O = O + G(I)
    For J = 1 To N
    F(J) = F(J) + X(J, I)
    H(J) = H(J) + X(J, I) * X(J, I) / G(I)
    Next J
    Next I
    
    Q2 = 0
    For I = 1 To N
    Q2 = Q2 + H(I) / F(I)
    Next I
    Q2 = O * (Q2 - 1)
    Q1 = (N - 1) * (M - 1)
    c2 = Sqr(Q2)
    J1 = Int(Q1 / 2)
    S1 = Exp(-Q2 / 2) / Sqr(2 * PI)
    If Q1 <> J1 * 2 Then GoTo 3400
    Q3 = 1
    For I = 1 To (Q1 - 2) / 2
    Z1 = 1
    For J = 2 To 2 * I Step 2
    Z1 = Z1 * J
    Next J
    Q3 = Q3 + c2 ^ (2 * I) / Z1
    Next I
    Q3 = S1 * Q3 * Sqr(2 * PI)
    GoTo 3580
3400:
    Q3 = 0
    For I = 1 To (Q1 - 1) / 2
    Z1 = 1
    For J = 1 To 2 * I - 1 Step 2
    Z1 = Z1 * J
    Next J
    Q3 = Q3 + c2 ^ (2 * I - 1) / Z1
    Next I
    I = 1
    J1 = Abs(c2): Z1 = Abs(c2)
3500:
    I = I + 2
    Q = Z1
    W1 = J1
    J1 = W1 / I * Q2
    Z1 = Q + J1
    If Z1 > Q Then GoTo 3500
    A = 0.5 - Z1 * S1
    Q3 = 2 * (S1 * Q3 + A)
3580:
    Q = Sqr(Q2 / (O + Q2))
    R1 = 0: R2 = 0
    For J = 1 To M
    For I = 1 To N
    c2 = H(J) * F(I) / N
    If c2 > 1 Then GoTo 3650
    R1 = R1 + 1
3650:
    If c2 >= 5 Then GoTo 3670
    R2 = R2 + 1
3670:
    Next I
3680 Next J
    R3 = 100 * (R2 / (M * N))

        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Test chi-quadrato per tabelle di contingenza (*inizio*)" & Chr(13) & Chr(10)
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
        Text4.Text = Text4.Text & "Test chi-quadrato              = " & Format(Q2, "#####0.0####") & Chr(13) & Chr(10)
        If Q3 >= 0.00001 Then
            Text4.Text = Text4.Text & "Probabilità p                  = " & Format(Q3, "#####0.0####") & Chr(13) & Chr(10)
        Else
            Q3 = 0.00001
            Text4.Text = Text4.Text & "Probabilità p                  < " & Format(Q3, "#####0.0####") & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Frequenze attese inferiori a 1 = " & Format(R1, "#####0") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Frequenze attese inferiori a 5 = " & Format(R2, "#####0") & Chr(13) & Chr(10)
        If R3 <> 0 Then
            Text4.Text = Text4.Text & "Frequenze attese inferiori a 5 = " & Format(R3, "#####0.0####") & " %" & Chr(13) & Chr(10)
        Else
            Text4.Text = Text4.Text & "Frequenze attese inferiori a 5 = " & Format(R3, "#####0") & " %" & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Test chi-quadrato per tabelle di contingenza (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        Screen.MousePointer = 1

        Exit Sub

GestioneErroriChiQuad:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico dei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        Exit Sub

End Sub

