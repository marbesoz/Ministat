VERSION 5.00
Begin VB.Form FormCluster 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Analisi della somiglianza"
   ClientHeight    =   6285
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
   Icon            =   "FORMCLUS.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6285
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
      Begin VB.Menu StatisticaAnalisiDellaSomiglianza 
         Caption         =   "&Analisi della somiglianza"
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
Attribute VB_Name = "FormCluster"
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
    Dim Deviate()          ' matrice dati standardizzati (z)
    Dim Distanze()              ' matrice delle distanze
    Dim Cluster()           ' matrice dei cluster formatisi
    Dim Caso()                        ' identificativo del caso
    Dim Distanza()                  ' distanza alla quale si forma il cluster
    Dim DummyVettore()                    ' vettore appoggio per scambio variabili
    Dim FlagStatistica, Flag1
    Dim PassiEffettivi

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
    Unload FormCluster
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
End Sub

Private Sub FileStampa_Click()
    On Error GoTo GestioneErroriStampaCluster
    Dim Risposta, FlagRigaVuota, Tabulatore, Z1
    Dim Riga, Colonna, Posizione1, Posizione2
    Dim ContaPagine, ContaRighe, SpessoreLinea, I
    
    If FlagStatistica = True Then
    
    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Stampa risultati"
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Caption = "0%"
    FormGauge!Gauge1.Value = 0
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    ContaPagine = 1
    ContaRighe = 1
    Posizione1 = Printer.CurrentY
    Printer.Print "Ministat 1.1 "; Tab(32); Now; Tab(66); "pag. "; Tab(71); Trim(Str$(ContaPagine))
    Posizione2 = Printer.CurrentY
    SpessoreLinea = Posizione2 - Posizione1
    Printer.Print
    Printer.Print "Analisi della somiglianza (*inizio*)"
    Printer.Print
        Printer.Print "File         : " & NomeDelFile
    If NomeDellaTabella <> "" Then
        Printer.Print "Tabella      : " & NomeDellaTabella
    End If
    Printer.Print
        For I = 1 To NUMCOLONNE
            If NUMCOLONNE < 10 Then
                Printer.Print "Variabile " & I & "  : " & NomeVariabile(I)
            Else
                Printer.Print "Variabile " & I & " : " & NomeVariabile(I)
            End If
        Next I
    Printer.Print
    Printer.Print "Numero dei dati  = " & NumDati(1)
    Printer.Print
    Printer.Print "Distanza (z)      Cluster                Casi"
    Printer.Print "------------      -------                ----"

    For Colonna = 1 To PassiEffettivi
        FormGauge.Label1.Caption = Int(100 * Colonna / PassiEffettivi) & "%"
        FormGauge!Gauge1.Value = Int(100 * Colonna / PassiEffettivi)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
        
        If Printer.ScaleHeight - Printer.CurrentY < (2 * SpessoreLinea) Then
            Printer.NewPage
            ContaPagine = ContaPagine + 1
            Printer.Print "Ministat 1.1 "; Tab(32); Now; Tab(66); "pag. "; Tab(71); Trim(Str$(ContaPagine))
            Printer.Print
            Printer.Print "Distanza (z)      Cluster                Casi"
            Printer.Print "------------      -------                ----"
        End If
        Printer.Print Tab(7); Format(Format(Distanza(Colonna), "#0.0##"), "@@@@@@");
    For Riga = 1 To NumDati(1)

        If Printer.ScaleHeight - Printer.CurrentY < (2 * SpessoreLinea) Then
            Printer.NewPage
            ContaPagine = ContaPagine + 1
            Printer.Print "Ministat 1.1 "; Tab(32); Now; Tab(66); "pag. "; Tab(71); Trim(Str$(ContaPagine))
            Printer.Print
            Printer.Print "Distanza (z)      Cluster                Casi"
            Printer.Print "------------      -------                ----"
        End If

            If Cluster(Colonna, Riga) <> 0 Then
                Printer.Print Tab(20); Format(Cluster(Colonna, Riga), "@@@@@@");
                Printer.Print Tab(40); Format(Caso(Riga), "@@@@@@");
                Printer.Print
            End If
            ContaRighe = ContaRighe + 1

    Next Riga
    
    Printer.Print
    Next Colonna
    
    Printer.Print "Analisi della somiglianza (*fine*)"
    Printer.EndDoc
    Unload FormGauge
    Screen.MousePointer = 1

Exit Sub

GestioneErroriStampaCluster:
    Printer.EndDoc ' non si sa mai !
    Unload FormGauge
    Screen.MousePointer = 1
    MsgBox "Errore durante la stampa: per favore riprovare.", MB_ICONEXCLAMATION, "Stampa dei risultati"
    Exit Sub

    End If

End Sub

Private Sub Form_Load()
    On Error GoTo GestioneErroriCluster
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
    FormCluster.Top = Screen.Height / 2 - Height / 2
    FormCluster.Left = Screen.Width / 2 - Width / 2
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
    FormGauge.Caption = "Analisi della somiglianza"
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
        If NumDati(I) < 3 Then Dummy = False
    Next I

    If Dummy = False Then
        Unload FormGauge
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        MsgBox "Per l'analisi della somiglianza" + Chr(10) + "sono necessari almeno tre dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Analisi della somiglianza non eseguibile"
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
        MsgBox "Per l'analisi della somiglianza" + Chr(10) + "le variabili devono contenere" + Chr(10) + "lo stesso numero di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Analisi della somiglianza non eseguibile"
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
        FormCluster.Show
        ' FormPrincipale.Hide
        StatisticaAnalisiDellaSomiglianza_Click
        Screen.MousePointer = 1

        Exit Sub

GestioneErroriCluster:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente: riprovare" + Chr(10) + "con un numero inferiore di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Analisi della somiglianza non eseguibile"
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

Private Sub StatisticaAnalisiDellaSomiglianza_Click()
    On Error GoTo GestioneErroriStatisticaCluster
    ReDim Preserve X(NUMCOLONNE, NumDati(1))
    Dim I, N, Variabili, Casi, MaxCasi
    Dim Co, Ri, SoX, SoXq, Media, DevSt, Somma
    Dim Var, MaxPassi, TotPassi
    Dim Passo, Min, Colonna, Riga, Dummy, R
    Dim Flag, D, J, L, C, Flag1, Flag2
    Dim Rig ' locale, per non toccare Rig globale !
    
    Screen.MousePointer = 11
    Text4.Visible = False

    Cls
    
    FormGauge.Show
    FormGauge.Caption = "Analisi della somiglianza"
    FormGauge!Label6.Caption = "[1/6]"
    FormGauge!Label1.Caption = "0%"
    FormGauge!Gauge1.Value = 0
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' calcolo cluster

                                          ' colonne = varibili, righe = casi
    Variabili = NUMCOLONNE                ' numero delle variabili/files
    Casi = NumDati(1)                     ' numero dei casi per variabile/file

    MaxCasi = Casi * (Casi - 1) / 2      ' massimo teorico distanza valide

    ReDim Deviate(Variabili, Casi)          ' matrice dati standardizzati (z)
    ReDim Distanze(Casi, Casi)              ' matrice delle distanze
    ReDim Cluster(MaxCasi, Casi)           ' matrice dei cluster formatisi
    ReDim Caso(Casi)                        ' identificativo del caso
    ReDim Distanza(MaxCasi)                  ' distanza alla quale si forma il cluster
    ReDim DummyVettore(MaxCasi)                    ' vettore appoggio per scambio variabili

    For I = 1 To Casi                      ' identifica i casi con progressivo
        FormGauge.Label1.Caption = Int(100 * I / Casi) & "%"
        FormGauge!Gauge1.Value = Int(100 * I / Casi)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
        Caso(I) = I
    Next I

' TRASFORMA IN DEVIATA NORMALE STANDARDIZZATA (z)

    For Co = 1 To Variabili

    SoX = 0: SoXq = 0
    For Ri = 1 To Casi
        SoX = SoX + X(Co, Ri): SoXq = SoXq + (X(Co, Ri)) ^ 2
    Next Ri
    Media = SoX / Casi: DevSt = (Sqr(SoXq - (SoX * SoX / Casi)) / (Casi - 1))
    For Ri = 1 To Casi
      X(Co, Ri) = (X(Co, Ri) - Media) / DevSt
    Next Ri

    Next Co

' CALCOLA LA MATRICE DELLE DISTANZE

    FormGauge!Label6.Caption = "[2/6]"
    For Co = 1 To Casi
        FormGauge.Label1.Caption = Int(100 * Co / Casi) & "%"
        FormGauge!Gauge1.Value = Int(100 * Co / Casi)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
 
    For Ri = 1 To Casi
    Somma = 0
        For Var = 1 To Variabili
            Somma = Somma + (X(Var, Co) - X(Var, Ri)) ^ 2
        Next Var
        Distanze(Co, Ri) = Sqr(Somma)
    Next Ri

    Next Co

' CALCOLA LA MATRICE DEI CLUSTER

    MaxPassi = Casi * (Casi - 1) / 2
    TotPassi = 0
    PassiEffettivi = 0

    FormGauge!Label6.Caption = "[3/6]"
    For Passo = 1 To MaxPassi
        FormGauge.Label1.Caption = Int(100 * Passo / MaxPassi) & "%"
        FormGauge!Gauge1.Value = Int(100 * Passo / MaxPassi)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    TotPassi = TotPassi + 1
    PassiEffettivi = PassiEffettivi + 1

    Min = 1E+38
    For Co = 1 To Casi - 1
      For Ri = Co + 1 To Casi
        If Distanze(Co, Ri) <> 1E+38 And Distanze(Co, Ri) < Min Then
          Min = Distanze(Co, Ri)
          Colonna = Co
          Riga = Ri
        End If
      Next Ri
    Next Co
    Distanza(Passo) = Distanze(Colonna, Riga)
    Distanze(Colonna, Riga) = 1E+38
   
    If Cluster(Passo, Riga) = 0 And Cluster(Passo, Colonna) = 0 Then
      Cluster(Passo, Riga) = Passo
      Cluster(Passo, Colonna) = Passo
    End If
   
    If Cluster(Passo, Riga) <> 0 And Cluster(Passo, Colonna) = 0 Then
      Dummy = Cluster(Passo, Riga)
      For R = 1 To Casi
        If Cluster(Passo, R) = Dummy Then
          Cluster(Passo, R) = Passo
        End If
      Next R
      Cluster(Passo, Colonna) = Passo
    End If
   
    If Cluster(Passo, Colonna) <> 0 And Cluster(Passo, Riga) = 0 Then
      Dummy = Cluster(Passo, Colonna)
      For R = 1 To Casi
        If Cluster(Passo, R) = Dummy Then
          Cluster(Passo, R) = Passo
        End If
      Next R
      Cluster(Passo, Riga) = Passo
    End If
   
    If Cluster(Passo, Colonna) <> 0 And Cluster(Passo, Riga) <> 0 Then
      Dummy = Cluster(Passo, Colonna)
      For R = 1 To Casi
        If Cluster(Passo, R) = Dummy Then
          Cluster(Passo, R) = Passo
        End If
      Next R
      Dummy = Cluster(Passo, Riga)
      For R = 1 To Casi
        If Cluster(Passo, R) = Dummy Then
          Cluster(Passo, R) = Passo
        End If
      Next R
    End If
   
    If Passo > 1 Then
      Flag = 0                           ' controlla se due cluster sono uguali
        For R = 1 To Casi
          If Cluster(Passo, R) = Cluster(Passo - 1, R) Or Cluster(Passo, R) = Cluster(Passo - 1, R) + 1 Then

          Else
            Flag = 1
          End If
        Next R
    End If

    If Passo > 1 Then
      If Flag = 0 Then                 ' se sono uguali ignora il secondo !
        Passo = Passo - 1
        PassiEffettivi = PassiEffettivi - 1  ' ovviamente...
      End If
    End If

    If Passo < MaxPassi Then           ' traspone il cluster al posto seguente
      For R = 1 To Casi
        Cluster(Passo + 1, R) = Cluster(Passo, R)
      Next R
    End If

    If TotPassi = MaxPassi Then        ' altrimenti loop infinito !!
      Exit For
    End If

    Next Passo

' ORDINA I CLUSTER

    FormGauge!Label6.Caption = "[4/6]"
300 For Co = 1 To PassiEffettivi
        FormGauge.Label1.Caption = Int(100 * Co / PassiEffettivi) & "%"
        FormGauge!Gauge1.Value = Int(100 * Co / PassiEffettivi)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
330 D = 1
340 D = 2 * D: If D <= Casi Then GoTo 340
350 D = Int((D - 1) / 2): If D = 0 Then GoTo 420
360 For I = 1 To Casi - D
370 J = I
380 L = J + D: If Cluster(Co, J) <= Cluster(Co, L) Then GoTo 400
382 For C = 1 To PassiEffettivi
384 Dummy = Cluster(C, J): Cluster(C, J) = Cluster(C, L): Cluster(C, L) = Dummy
386 Next C
388 Dummy = Caso(J): Caso(J) = Caso(L): Caso(L) = Dummy
390 J = J - D: If J > 0 Then GoTo 380
400 Next I
410 GoTo 350
420 Next Co

' CONTINUA FINCHE' NON HA SISTEMATO TUTTO

    While Flag1 = 0 Or Flag2 = 1         ' necessaria sistemata !

' PRIMA SISTEMATA

    FormGauge!Label6.Caption = "[5/6]"
    For Co = 1 To PassiEffettivi
        FormGauge.Label1.Caption = Int(100 * Co / PassiEffettivi) & "%"
        FormGauge!Gauge1.Value = Int(100 * Co / PassiEffettivi)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    For R = 1 To Casi - 2
    For Ri = R + 2 To Casi

      If Cluster(Co, R) <> 0 And Cluster(Co, Ri) <> 0 Then
      If Cluster(Co, R) = Cluster(Co, Ri) Then

        Flag1 = 0
          For Rig = R + 1 To Ri - 1
            If Cluster(Co, Rig) <> 0 Then Flag1 = 1
          Next Rig

          If Flag1 = 0 Then
            For C = 1 To PassiEffettivi
              Dummy = Cluster(C, R + 1): Cluster(C, R + 1) = Cluster(C, Ri): Cluster(C, Ri) = Dummy
            Next C
              Dummy = Caso(R + 1): Caso(R + 1) = Caso(Ri): Caso(Ri) = Dummy
          End If

      End If
      End If
    Next Ri
    Next R

    Next Co

' SECONDA SISTEMATA

    FormGauge!Label6.Caption = "[6/6]"
    For Co = 1 To PassiEffettivi
        FormGauge.Label1.Caption = Int(100 * Co / PassiEffettivi) & "%"
        FormGauge!Gauge1.Value = Int(100 * Co / PassiEffettivi)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    For R = 1 To Casi - 2
    For Ri = R + 2 To Casi

      If Cluster(Co, R) <> 0 And Cluster(Co, Ri) <> 0 Then
      If Cluster(Co, R) = Cluster(Co, Ri) Then
       
        Flag2 = 0
          For Riga = R + 1 To Ri - 1
            If Cluster(Co, Riga) <> Cluster(Co, R) Then Flag2 = 1
          Next Riga

        If Flag2 = 1 Then
       
        For C = 1 To PassiEffettivi
          DummyVettore(C) = Cluster(C, Ri)
            For Rig = Ri To R + 1 Step -1
              Cluster(C, Rig) = Cluster(C, Rig - 1)
            Next Rig
          Cluster(C, R + 1) = DummyVettore(C)
        Next C
         
          Dummy = Caso(Ri)          ' sistema identificativo caso !
            For Rig = Ri To R + 1 Step -1
              Caso(Rig) = Caso(Rig - 1)
            Next Rig
          Caso(R + 1) = Dummy
       
        End If

      End If
      End If
    Next Ri
    Next R

    Next Co

' VERIFICA SE CONTINUARE

    Flag1 = 0
    Flag2 = 0

    For Co = 1 To PassiEffettivi

    For R = 1 To Casi - 2
    For Ri = R + 2 To Casi

      If Cluster(Co, R) <> 0 And Cluster(Co, Ri) <> 0 Then
      If Cluster(Co, R) = Cluster(Co, Ri) Then

        For Rig = R + 1 To Ri - 1
          If Cluster(Co, Rig) <> 0 Then Flag1 = 1
        Next Rig
        For Riga = R + 1 To Ri - 1
          If Cluster(Co, Riga) <> Cluster(Co, R) Then Flag2 = 1
        Next Riga

      End If
      End If
    Next Ri
    Next R

    Next Co

    Wend

' RICALCOLA LA MATRICE DELLE DISTANZE

    For Co = 1 To Casi
 
    For Ri = 1 To Casi
    Somma = 0
      For Var = 1 To Variabili
        Somma = Somma + (X(Var, Co) - X(Var, Ri)) ^ 2
      Next Var
        Distanze(Co, Ri) = Sqr(Somma)
    Next Ri

    Next Co

' presenta risultati
61616:
    Text4.Text = "Ministat 1.1 (32 bit)           " & Now & Chr(13) & Chr(10)
    Text4.Text = Text4.Text & Chr(13) & Chr(10)
    Text4.Text = Text4.Text & "Analisi della somiglianza (*inizio*)" & Chr(13) & Chr(10)
    Text4.Text = Text4.Text & Chr(13) & Chr(10)
    Text4.Text = Text4.Text & "File         : " & NomeDelFile & Chr(13) & Chr(10)
    If NomeDellaTabella <> "" Then
        Text4.Text = Text4.Text & "Tabella      : " & NomeDellaTabella & Chr(13) & Chr(10)
    End If
    Text4.Text = Text4.Text & Chr(13) & Chr(10)
    For I = 1 To NUMCOLONNE
    If NUMCOLONNE < 10 Then
        Text4.Text = Text4.Text & "Variabile " & I & "  : " & NomeVariabile(I) & Chr(13) & Chr(10)
    Else
        Text4.Text = Text4.Text & "Variabile " & I & " : " & NomeVariabile(I) & Chr(13) & Chr(10)
    End If
    Next I
    Text4.Text = Text4.Text & Chr(13) & Chr(10)
    Text4.Text = Text4.Text & "Numero dei dati  = " & NumDati(1) & Chr(13) & Chr(10)
    Text4.Text = Text4.Text & Chr(13) & Chr(10)

    Text4.Text = Text4.Text & "  Casi     Cluster" & Chr(13) & Chr(10)

    For R = 1 To Casi

    Text4.Text = Text4.Text & Format(Caso(R), "@@@@@@")

    For C = 1 To PassiEffettivi
    If Cluster(C, R) <> 0 Then
        Text4.Text = Text4.Text & Format(Cluster(C, R), "@@@@@@")
    Else
        Text4.Text = Text4.Text & Format("0", "@@@@@@")
    End If
    Next C
    Text4.Text = Text4.Text & Chr(13) & Chr(10)

    Next R

    Text4.Text = Text4.Text & Chr(13) & Chr(10)
    Text4.Text = Text4.Text & "      "

    For C = 1 To PassiEffettivi
    If Distanza(C) <> 0 Then
        Text4.Text = Text4.Text & Format(Format(Distanza(C), "#0.0##"), "@@@@@@")
    Else
        Text4.Text = Text4.Text & Format("0", "@@@@@@")
    End If
    Next C
  
    Text4.Text = Text4.Text & Chr(13) & Chr(10)
    Text4.Text = Text4.Text & "  Distanza euclidea (z)" & Chr(13) & Chr(10)

    Text4.Text = Text4.Text & Chr(13) & Chr(10)
    Text4.Text = Text4.Text & "Analisi della somiglianza (*fine*)" & Chr(13) & Chr(10)
61669:
    Unload FormGauge
        
    Text4.Visible = True
    Text4.Enabled = True
    FlagStatistica = True
    Screen.MousePointer = 1

Exit Sub

GestioneErroriStatisticaCluster:
        If Erl = 61616 Then GoTo 61669
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente o errore logico dei dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcoli non eseguibili"
        Exit Sub

End Sub

