VERSION 5.00
Begin VB.Form FormMultipla 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Regressione lineare multipla"
   ClientHeight    =   6045
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
   Icon            =   "FORMREMU.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6045
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
      Begin VB.Menu StatisticaMultipla 
         Caption         =   "&Regressione lineare multipla"
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
Attribute VB_Name = "FormMultipla"
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
    Dim II()
    Dim S()
    Dim t()
    Dim A()
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
    Unload FormMultipla
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
End Sub

Private Sub FileStampa_Click()
    On Error GoTo ErroreFileStampaRemu

    If FlagStatistica = True Then
        Printer.Print Text4.Text
        Printer.EndDoc
    End If

    Exit Sub
        
ErroreFileStampaRemu:
    Exit Sub

End Sub

Private Sub Form_Load()
    On Error GoTo GestioneErroriMultipla
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
    FormMultipla.Top = Screen.Height / 2 - Height / 2
    FormMultipla.Left = Screen.Width / 2 - Width / 2
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
    FormGauge.Caption = "Regressione lineare multipla"
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
        MsgBox "Per la regressione multipla" + Chr(10) + "sono necessari almeno due dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Regressione multipla non eseguibile"
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
        MsgBox "Per la regressione multipla" + Chr(10) + "le variabili devono contenere" + Chr(10) + "lo stesso numero di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Regressione multipla non eseguibile"
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
        FormMultipla.Show
        ' FormPrincipale.Hide
        StatisticaMultipla_Click
        Screen.MousePointer = 1

Exit Sub

GestioneErroriMultipla:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente: riprovare" + Chr(10) + "con un numero inferiore di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Regressione multipla non eseguibile"
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

Private Sub StatisticaMultipla_Click()
    On Error GoTo GestioneErroriStatisticaMultipla
    
    Screen.MousePointer = 11
    
    ReDim Preserve X(NUMCOLONNE, NumDati(1))
    ReDim II(NumDati(1) + 1)
    ReDim S(NumDati(1) + 1)
    ReDim t(NumDati(1) + 1)
    ReDim A(NumDati(1) + 1, NumDati(1) + 2)
    Dim I, V, N, J, K, L, B, Z, p
    Dim R, Corr

    Text4.Visible = False

    Cls
    
    FormGauge.Show
    FormGauge.Caption = "Regressione lineare multipla"
    FormGauge!Label6.Caption = "[1/2]"
    FormGauge!Label1.Caption = "0%"
    FormGauge!Gauge1.Value = 0
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Regressione lineare multipla (*inizio*)" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "File         : " & NomeDelFile & Chr(13) & Chr(10)
        If NomeDellaTabella <> "" Then
            Text4.Text = Text4.Text & "Tabella      : " & NomeDellaTabella & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        For I = 1 To NUMCOLONNE - 1
            Text4.Text = Text4.Text & "Variabile indipendente " & I & " : " & NomeVariabile(I) & Chr(13) & Chr(10)
        Next I
        If NUMCOLONNE < 10 Then
            Text4.Text = Text4.Text & "Variabile dipendente y   : " & NomeVariabile(NUMCOLONNE) & Chr(13) & Chr(10)
        Else
            Text4.Text = Text4.Text & "Variabile dipendente y  : " & NomeVariabile(NUMCOLONNE) & Chr(13) & Chr(10)
        End If
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Polinomio di equazione y = a + b x + b x + .... + b x" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "                                1 1   2 2          n n" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)

' calcolo regressione lineare multipla

    V = NUMCOLONNE - 1
    N = NumDati(1)
    
    For I = 1 To NumDati(1) + 1
        II(I) = 0: S(I) = 0: t(I) = 0
        For J = 1 To NumDati(1) + 2
            A(I, J) = 0
        Next J
    Next I

    II(1) = 1
    For I = 1 To N
        FormGauge.Label1.Caption = Int(100 * I / N) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / N)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
        For J = 1 To V + 1
            II(J + 1) = X(J, I)
        Next J
        For K = 1 To V + 1
            For L = 1 To V + 2
                A(K, L) = A(K, L) + II(K) * II(L)
                S(K) = A(K, V + 2)
            Next L
        Next K
        S(V + 2) = S(V + 2) + II(V + 2) * II(V + 2)
        
    Next I

    For I = 2 To V + 1
        t(I) = A(1, I)
    Next I

    FormGauge!Label6.Caption = "[2/2]"
    For I = 1 To V + 1
        FormGauge.Label1.Caption = Int(100 * I / (V + 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / (V + 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
        J = I
3320:
        If A(J, I) <> 0 Then GoTo 3370
        J = J + 1
        If J <= V + 1 Then GoTo 3320
        Text4.Text = Text4.Text & "Attenzione: non esiste una soluzione unica !" & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Regressione lineare multipla (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        Screen.MousePointer = 1
        
        Exit Sub

3370:
        For K = 1 To V + 2
            B = A(I, K)
            A(I, K) = A(J, K)
            A(J, K) = B
        Next K
        Z = 1 / A(I, I)
        For K = 1 To V + 2
            A(I, K) = Z * A(I, K)
        Next K
        For J = 1 To V + 1
            If J = I Then GoTo 3520
            Z = -A(J, I)
                For K = 1 To V + 2
                    A(J, K) = A(J, K) + Z * A(I, K)
                Next K
3520:
        Next J
    Next I

        Text4.Text = Text4.Text & "Numero dei dati                                        = " & NumDati(1) & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Intercetta a                                           = " & Format(A(1, V + 2), "#######0.0#########") & Chr(13) & Chr(10)
        For I = 2 To V + 1
            Text4.Text = Text4.Text & "Coefficiente angolare b della variabile indipendente " & I - 1 & " = " & Format(A(I, V + 2), "#######0.0#########") & Chr(13) & Chr(10)
        Next I
        p = 0
        For I = 2 To V + 1
            p = p + A(I, V + 2) * (S(I) - t(I) * S(1) / N)
        Next I
        R = S(V + 2) - S(1) * S(1) / N
        Z = R - p
        L = N - V - 1
        I = p / R: Corr = I
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Coefficiente di correlazione multipla                  = " & Format(Sqr(Corr), "#####0.0####") & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Errore standard della stima                            = " & Format(Sqr(Abs(Z / L)), "#####0.0####") & Chr(13) & Chr(10)

        Text4.Text = Text4.Text & Chr(13) & Chr(10)
        Text4.Text = Text4.Text & "Regressione lineare multipla (*fine*)" & Chr(13) & Chr(10)
        
        Unload FormGauge
        
        Text4.Visible = True
        Text4.Enabled = True
        FlagStatistica = True
        Screen.MousePointer = 1

    Exit Sub

GestioneErroriStatisticaMultipla:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico nei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        Exit Sub

End Sub

