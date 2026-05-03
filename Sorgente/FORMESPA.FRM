VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FormEsploraPar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Statistiche esplorative parametriche"
   ClientHeight    =   5820
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
   Icon            =   "FORMESPA.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5820
   ScaleWidth      =   7365
   Begin MSGrid.Grid Grid1 
      Height          =   1215
      Left            =   2400
      TabIndex        =   0
      Top             =   1800
      Width           =   3015
      _Version        =   65536
      _ExtentX        =   5318
      _ExtentY        =   2143
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&File"
      Begin VB.Menu FileEsci 
         Caption         =   "&Esci"
      End
   End
   Begin VB.Menu MenuOpzioni 
      Caption         =   "&Opzioni"
      Begin VB.Menu OpzioniNumeroDecimali 
         Caption         =   "&Numero di decimali"
         Begin VB.Menu OpzioniNessunDecimale 
            Caption         =   "&Nessuno"
         End
         Begin VB.Menu OpzionUnDecimale 
            Caption         =   "&Un decimale"
         End
         Begin VB.Menu OpzioniDueDecimali 
            Caption         =   "&Due decimali"
         End
         Begin VB.Menu OpzioniTreDecimali 
            Caption         =   "&Tre decimali"
         End
         Begin VB.Menu OpzioniQuattroDecimali 
            Caption         =   "&Quattro decimali"
         End
         Begin VB.Menu OpzioniCinqueDecimali 
            Caption         =   "&Cinque decimali"
         End
      End
      Begin VB.Menu OpzioniSeparatoreUno 
         Caption         =   "-"
      End
      Begin VB.Menu OpzioniGriglia 
         Caption         =   "&Griglia"
         Checked         =   -1  'True
      End
      Begin VB.Menu OpzioniSeparatoreDue 
         Caption         =   "-"
      End
      Begin VB.Menu OpzioniRipristinaAltezzaRighe 
         Caption         =   "Ripristina &altezza righe"
      End
      Begin VB.Menu OpzioniRipristinaLarghezzaColonne 
         Caption         =   "Ripristina &larghezza colonne"
      End
   End
   Begin VB.Menu MenuGuida 
      Caption         =   "&?"
      Begin VB.Menu AiutoGuida 
         Caption         =   "&Guida"
      End
      Begin VB.Menu GuidaCalcolatrice 
         Caption         =   "&Calcolatrice"
      End
   End
End
Attribute VB_Name = "FormEsploraPar"
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
    Dim X() As Double
    Dim Y() As Double
    Dim NUMCOLONNE
    Dim NomeVariabile()
    Dim NumDati()
    Dim FineDati()
    Dim Minimo(), Massimo(), PrimoQ(), SecondoQ(), TerzoQ()

Private Sub AiutoGuida_Click()

    FormPrincipale!CMDialog1.HelpFile = App.Path & "\ministat.hlp"
    FormPrincipale!CMDialog1.HelpCommand = cdlHelpContents
    FormPrincipale!CMDialog1.ShowHelp

End Sub

Private Sub FileEsci_Click()
    
    FormPrincipale.Show
    Unload FormEsploraPar
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo

End Sub

Private Sub Form_Load()
    On Error GoTo GestioneErroriEsploraPar
    Dim I, J, K, Riga, Z1, Dummy As Double
    Dim ValoreInizialeD, D, L, XX, M1, M2
    Dim ColonneGrid, RigheGrid, Colonna
    Dim SomX, SomXquad, SomXCubo, SomXquar, NumeroDati
    Dim Media, DevStnd, CV, ErrStnd
    Dim EmmeDue, EmmeTre, EmmeQua
    Dim GiUno, GiDue, EsseUno, EsseDue

    Screen.MousePointer = 11
    
    ColonneGrid = NUM_COLONNE
    RigheGrid = 41
    NUMCOLONNE = NUM_COLONNE - 1

    ReDim X(NUMCOLONNE, NUM_RIGHE - 1) As Double
    ReDim Y(NUMCOLONNE, NUM_RIGHE - 1) As Double
    ReDim NomeVariabile(NUMCOLONNE)
    ReDim NumDati(NUMCOLONNE)
    ReDim FineDati(NUMCOLONNE)
    ReDim Minimo(NUMCOLONNE)
    ReDim Massimo(NUMCOLONNE)
    ReDim PrimoQ(NUMCOLONNE)
    ReDim SecondoQ(NUMCOLONNE)
    ReDim TerzoQ(NUMCOLONNE)

    AnnullatoCaricamentoForm = False
    FlagGrafico = False

' centra il form
    FormEsploraPar.Top = Screen.Height / 2 - Height / 2
    FormEsploraPar.Left = Screen.Width / 2 - Width / 2
' numero di righe e colonne dell'oggetto Grid1
    FormEsploraPar!Grid1.Rows = RigheGrid
    FormEsploraPar!Grid1.Cols = ColonneGrid
' sposta in alto a sinistra
    FormEsploraPar!Grid1.Move 0, 0
' dimensionamento, ripreso nella sub resize
    FormEsploraPar!Grid1.Width = ScaleWidth
    FormEsploraPar!Grid1.Height = ScaleHeight
' definisce larghezza colonne
    FormEsploraPar!Grid1.ColWidth(0) = 2500
    FormEsploraPar!Grid1.Row = 0
        For Colonna = 1 To ColonneGrid - 1
            FormEsploraPar!Grid1.ColWidth(Colonna) = LARGHEZZA_COLONNE
        Next Colonna
' definisce altezza righe
    For Riga = 0 To RigheGrid - 1
        FormEsploraPar!Grid1.RowHeight(Riga) = 255
    Next Riga
' allineamento colonne
    FormEsploraPar!Grid1.FixedAlignment(0) = 0 ' intestazioni righe a sinistra
        For Colonna = 1 To ColonneGrid - 1
            FormEsploraPar!Grid1.FixedAlignment(Colonna) = 2 ' nome variabili centrato
            FormEsploraPar!Grid1.ColAlignment(Colonna) = 1 ' dati a destra
        Next Colonna

' nome delle variabili
    
    FormPrincipale!Grid1.Row = 0
    FormEsploraPar!Grid1.Row = 0
    K = 0
    For I = 1 To NUM_COLONNE - 1
        K = K + 1
        FormPrincipale!Grid1.Col = I
        FormEsploraPar!Grid1.Col = I
        FormEsploraPar!Grid1.Text = FormPrincipale!Grid1.Text
        NomeVariabile(K) = FormPrincipale!Grid1.Text
    Next I

' legende
    
    FormEsploraPar!Grid1.Col = 0
    
    FormEsploraPar!Grid1.Row = 2
    FormEsploraPar!Grid1.Text = "Numero dei dati"
    FormEsploraPar!Grid1.Row = 4
    FormEsploraPar!Grid1.Text = "Media"
    FormEsploraPar!Grid1.Row = 5
    FormEsploraPar!Grid1.Text = "Deviazione standard"
    FormEsploraPar!Grid1.Row = 6
    FormEsploraPar!Grid1.Text = "Coefficiente di variazione"
    FormEsploraPar!Grid1.Row = 7
    FormEsploraPar!Grid1.Text = "Errore standard della media"

    FormEsploraPar!Grid1.Row = 9
    FormEsploraPar!Grid1.Text = "Coefficiente di asimmetria"
    FormEsploraPar!Grid1.Row = 10
    FormEsploraPar!Grid1.Text = "Deviazione standard"
    FormEsploraPar!Grid1.Row = 11
    FormEsploraPar!Grid1.Text = "Rapporto"

    FormEsploraPar!Grid1.Row = 13
    FormEsploraPar!Grid1.Text = "Coefficiente di curtosi"
    FormEsploraPar!Grid1.Row = 14
    FormEsploraPar!Grid1.Text = "Deviazione standard"
    FormEsploraPar!Grid1.Row = 15
    FormEsploraPar!Grid1.Text = "Rappporto"
        
    FormEsploraPar!Grid1.Row = 17
    FormEsploraPar!Grid1.Text = "Valore minimo"
    FormEsploraPar!Grid1.Row = 18
    FormEsploraPar!Grid1.Text = "Primo quartile"
    FormEsploraPar!Grid1.Row = 19
    FormEsploraPar!Grid1.Text = "Secondo quartile"
    FormEsploraPar!Grid1.Row = 20
    FormEsploraPar!Grid1.Text = "Terzo quartile"
    FormEsploraPar!Grid1.Row = 21
    FormEsploraPar!Grid1.Text = "Valore massimo"
    
    FormGauge.Show
    FormGauge.Caption = "Statistiche esplorative parametriche"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' carica i dati nel vettore X()
    For I = 1 To NUMCOLONNE
        NumDati(I) = 0
        FineDati(I) = False
    Next I

    K = 0
    For I = 1 To NUM_COLONNE - 1
        FormPrincipale!Grid1.Col = I
        K = K + 1
        FormGauge!Label6.Caption = "[" & K & "/" & NUMCOLONNE * 4 & "]"
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

' ordina i dati nel vettore X()
    
    For K = 1 To NUMCOLONNE

    ValoreInizialeD = 0
    Z1 = NumDati(K)

' inizio ordinamento dati
pq330:
    FormGauge!Label6.Caption = "[" & NUMCOLONNE + K & "/" & NUMCOLONNE * 4 & "]"
    D = 1
pq340:
    D = 2 * D: If D <= Z1 Then GoTo pq340
pq350:
    D = Int((D - 1) / 2): If D = 0 Then GoTo pq420
pq360:
    For I = 1 To Z1 - D
pq370:
        J = I
pq380:
        L = J + D
        If X(K, J) <= X(K, L) Then GoTo pq400
                    
pq390:
            XX = X(K, J): X(K, J) = X(K, L): X(K, L) = XX

            If I > ValoreInizialeD And 100 * I / Z1 >= 1 Then
                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If
            J = J - D: If J > 0 Then GoTo pq380
pq400:
    Next I
pq410:
    ValoreInizialeD = Z1 - D
    GoTo pq350
pq420:
' fine ordinamento dati
                
                FormGauge.Label1.Caption = "100%"
                FormGauge!Gauge1.Value = 100
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    Next K
            
    For K = 1 To NUMCOLONNE
        Z1 = NumDati(K)

        FormGauge!Label6.Caption = "[" & NUMCOLONNE * 2 + K & "/" & NUMCOLONNE * 4 & "]"
            For I = 1 To Z1
                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                M1 = 0: M2 = 0
peq840:
                    For J = 1 To Z1
peq850:
                       If X(K, J) > X(K, I) Then GoTo peq890
peq860:
                       If X(K, J) = X(K, I) Then GoTo peq880
peq870:
                       M1 = M1 + 1: GoTo peq890
peq880:
                       M2 = M2 + 1
peq890:
                    Next J
peq900:
                Y(K, I) = M1 + 0.5 * M2 + 0.5
            Next I

            For I = 1 To Z1
                Y(K, I) = Format(100 * Y(K, I) / (Z1 + 1), "##0.0####")
            Next I

    Next K

' controlli finali sui dati
' devono esserci almeno tre dati per ciascuna variabile
'    Dummy = True
'    For I = 1 To NumColonne
'        If NumDati(I) < 3 Then Dummy = False
'    Next I
'
'    If Dummy = False Then
'        Unload FormGauge
'        AnnullatoCaricamentoForm = True
'        Screen.MousePointer = 1
'        MsgBox "Per il diagramma dei quartili e del" + Chr(10) + "range sono necessari almeno tre punti.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
'        RipristinaSchermo
'        Exit Sub
'    End If

' altrimenti continua

        YMI = X(1, 1): YMA = X(1, 1)
        
        For K = 1 To NUMCOLONNE
        If NumDati(K) > 2 Then

            Minimo(K) = X(K, 1): Massimo(K) = X(K, NumDati(K))
            For I = 1 To NumDati(K)
                If X(K, I) < YMI Then YMI = X(K, I)
                If X(K, I) > YMA Then YMA = X(K, I)
            Next I
        
        End If
        Next K
        YMI_Default = YMI: YMA_Default = YMA

        For K = 1 To NUMCOLONNE
        If NumDati(K) > 2 Then
            
            Z1 = NumDati(K)

            SomX = 0: SomXquad = 0: SomXCubo = 0: SomXquar = 0: NumeroDati = 0
    
            FormGauge!Label6.Caption = "[" & NUMCOLONNE * 3 + K & "/" & NUMCOLONNE * 4 & "]"
            For I = 1 To Z1
                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
        
                SomX = SomX + X(K, I)
                SomXquad = SomXquad + X(K, I) ^ 2
                SomXCubo = SomXCubo + X(K, I) ^ 3
                SomXquar = SomXquar + X(K, I) ^ 4
                NumeroDati = NumeroDati + 1
            
            Next I

            Media = SomX / NumeroDati
            DevStnd = Sqr((SomXquad - SomX ^ 2 / NumeroDati) / (NumeroDati - 1))
            CV = DevStnd * 100 / Media
            ErrStnd = DevStnd / Sqr(NumeroDati)
    
            EmmeDue = (SomXquad - NumeroDati * Media ^ 2) / NumeroDati
            EmmeTre = (SomXCubo - 3 * Media * SomXquad + 2 * NumeroDati * Media ^ 3) / NumeroDati
            EmmeQua = (SomXquar - 4 * Media * SomXCubo + 6 * Media ^ 2 * SomXquad - 3 * NumeroDati * Media ^ 4) / NumeroDati
            GiUno = EmmeTre / (EmmeDue * Sqr(EmmeDue))
            GiDue = EmmeQua / EmmeDue ^ 2 - 3
            EsseUno = Sqr(6 / NumeroDati)
            EsseDue = Sqr(24 / NumeroDati)

        FormEsploraPar!Grid1.Col = K
        
        FormEsploraPar!Grid1.Row = 2
        FormEsploraPar!Grid1.Text = NumDati(K)
        FormEsploraPar!Grid1.Row = 4
        FormEsploraPar!Grid1.Text = Format(CVar(Media), "######0.0######")
        FormEsploraPar!Grid1.Row = 5
        FormEsploraPar!Grid1.Text = Format(CVar(DevStnd), "######0.0######")
        FormEsploraPar!Grid1.Row = 6
        FormEsploraPar!Grid1.Text = Format(CVar(CV), "######0.0######")
        FormEsploraPar!Grid1.Row = 7
        FormEsploraPar!Grid1.Text = Format(CVar(ErrStnd), "######0.0######")

        FormEsploraPar!Grid1.Row = 9
        FormEsploraPar!Grid1.Text = Format(CVar(GiUno), "######0.0######")
        FormEsploraPar!Grid1.Row = 10
        FormEsploraPar!Grid1.Text = Format(CVar(EsseUno), "######0.0######")
        FormEsploraPar!Grid1.Row = 11
        FormEsploraPar!Grid1.Text = Format(CVar(Abs(GiUno / EsseUno)), "######0.0######")

        FormEsploraPar!Grid1.Row = 13
        FormEsploraPar!Grid1.Text = Format(CVar(GiDue), "######0.0######")
        FormEsploraPar!Grid1.Row = 14
        FormEsploraPar!Grid1.Text = Format(CVar(EsseDue), "######0.0######")
        FormEsploraPar!Grid1.Row = 15
        FormEsploraPar!Grid1.Text = Format(CVar(Abs(GiDue / EsseDue)), "######0.0######")
        
        FormEsploraPar!Grid1.Row = 17
        FormEsploraPar!Grid1.Text = Format(CVar(Minimo(K)), "######0.0######")
        FormEsploraPar!Grid1.Row = 18
        FormEsploraPar!Grid1.Text = Format(CVar(Media - (0.674 * DevStnd)), "######0.0######")
        FormEsploraPar!Grid1.Row = 19
        FormEsploraPar!Grid1.Text = Format(CVar(Media), "######0.0######")
        FormEsploraPar!Grid1.Row = 20
        FormEsploraPar!Grid1.Text = Format(CVar(Media + (0.674 * DevStnd)), "######0.0######")
        FormEsploraPar!Grid1.Row = 21
        FormEsploraPar!Grid1.Text = Format(CVar(Massimo(K)), "######0.0######")

        End If
        Next K

        Unload FormGauge
        Grid1.Visible = True
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
        FormEsploraPar.Show
        ' FormPrincipale.Hide
        Grid1.HighLight = False
        FormEsploraPar!Grid1.Col = 1
        FormEsploraPar!Grid1.Row = 1
        FormEsploraPar.SetFocus
        Screen.MousePointer = 1

        Exit Sub

GestioneErroriEsploraPar:
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente: riprovare" + Chr(10) + "con un numero inferiore di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
        Exit Sub

End Sub

Private Sub Form_Resize()
    
    Grid1.Width = ScaleWidth
    Grid1.Height = ScaleHeight

End Sub

Private Sub GuidaCalcolatrice_Click()
    
    Dim Calcolatrice
    On Error GoTo ErroreCalcolatriceEsploraNonPar:
    
    Calcolatrice = Shell("Calc.exe", 1)
        
Exit Sub

ErroreCalcolatriceEsploraNonPar:
    MsgBox "Calcolatrice non disponibile.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcolatrice"
Exit Sub


End Sub

Private Sub OpzioniCinqueDecimali_Click()
    
    Dim StringaDiFormattazione As String, Risposta
    Dim Riga, Colonna

    Risposta = MsgBox("In seguito all'operazione di arrotondamento" + Chr(10) + "i decimali in eccesso saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Arrotonda a cinque decimali")
        If Risposta = IDNO Then
            Exit Sub
            FormEsploraPar!Grid1.SetFocus
        End If

' limita a 12 i caratteri, inclusa la virgola
    StringaDiFormattazione = "#####0.00000"

    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Arrotonda a cinque decimali"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' arrotonda
    For Colonna = 1 To NUM_COLONNE - 1
    FormEsploraPar!Grid1.Col = Colonna
    
    For Riga = 1 To 40

        FormGauge.Label1.Caption = Int(100 * Riga / 40) & "%"
        FormGauge!Label6.Caption = "[" & Colonna & "/" & NUM_COLONNE - 1 & "]"
        FormGauge!Gauge1.Value = Int(100 * Riga / 40)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' 12 caratteri inclusa la virgola !
            FormEsploraPar!Grid1.Row = Riga
                If FormEsploraPar!Grid1.Text <> "" Then
'                If CDbl(FormEsploraPar!Grid1.Text) <> Int(CDbl(FormEsploraPar!Grid1.Text)) Then
                If Riga <> 2 Then
                    If Format(CVar(FormEsploraPar!Grid1.Text), StringaDiFormattazione) <> 0 Then
                        FormEsploraPar!Grid1.Text = Format(CVar(FormEsploraPar!Grid1.Text), StringaDiFormattazione)
                    Else
                        FormEsploraPar!Grid1.Text = Format(CVar(FormEsploraPar!Grid1.Text), "0")
                    End If
                 End If
                End If
    Next Riga
    Next Colonna

    Unload FormGauge
    FormEsploraPar!Grid1.Col = 1
    FormEsploraPar!Grid1.Row = 1
    FormEsploraPar!Grid1.Refresh
    FormEsploraPar!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

Private Sub OpzioniDueDecimali_Click()
    
    Dim StringaDiFormattazione As String, Risposta
    Dim Riga, Colonna

    Risposta = MsgBox("In seguito all'operazione di arrotondamento" + Chr(10) + "i decimali in eccesso saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Arrotonda a due decimali")
        If Risposta = IDNO Then
            Exit Sub
            FormEsploraPar!Grid1.SetFocus
        End If

' limita a 12 i caratteri, inclusa la virgola
    StringaDiFormattazione = "########0.00"

    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Arrotonda a due decimali"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' arrotonda
    For Colonna = 1 To NUM_COLONNE - 1
    FormEsploraPar!Grid1.Col = Colonna
    
    For Riga = 1 To 40

        FormGauge.Label1.Caption = Int(100 * Riga / 40) & "%"
        FormGauge!Label6.Caption = "[" & Colonna & "/" & NUM_COLONNE - 1 & "]"
        FormGauge!Gauge1.Value = Int(100 * Riga / 40)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' 12 caratteri inclusa la virgola !
            FormEsploraPar!Grid1.Row = Riga
                If FormEsploraPar!Grid1.Text <> "" Then
'                If CDbl(FormEsploraPar!Grid1.Text) <> Int(CDbl(FormEsploraPar!Grid1.Text)) Then
                If Riga <> 2 Then
                    If Format(CVar(FormEsploraPar!Grid1.Text), StringaDiFormattazione) <> 0 Then
                        FormEsploraPar!Grid1.Text = Format(CVar(FormEsploraPar!Grid1.Text), StringaDiFormattazione)
                    Else
                        FormEsploraPar!Grid1.Text = Format(CVar(FormEsploraPar!Grid1.Text), "0")
                    End If
                 End If
                End If
    Next Riga
    Next Colonna

    Unload FormGauge
    FormEsploraPar!Grid1.Col = 1
    FormEsploraPar!Grid1.Row = 1
    FormEsploraPar!Grid1.Refresh
    FormEsploraPar!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

Private Sub OpzioniGriglia_Click()
    
    OpzioniGriglia.Checked = Not OpzioniGriglia.Checked
    FormEsploraPar!Grid1.GridLines = Not FormEsploraPar!Grid1.GridLines
    FormEsploraPar!Grid1.SetFocus

End Sub

Private Sub OpzioniNessunDecimale_Click()
    
    Dim StringaDiFormattazione As String, Risposta
    Dim Riga, Colonna

    Risposta = MsgBox("In seguito all'operazione di arrotondamento" + Chr(10) + "i decimali in eccesso saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Arrotonda a nessun decimale")
        If Risposta = IDNO Then
            Exit Sub
            FormEsploraPar!Grid1.SetFocus
        End If

' limita a 12 i caratteri, inclusa la virgola
    StringaDiFormattazione = "###########0"

    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Arrotonda a nessun decimale"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' arrotonda
    For Colonna = 1 To NUM_COLONNE - 1
    FormEsploraPar!Grid1.Col = Colonna
    
    For Riga = 1 To 40

        FormGauge.Label1.Caption = Int(100 * Riga / 40) & "%"
        FormGauge!Label6.Caption = "[" & Colonna & "/" & NUM_COLONNE - 1 & "]"
        FormGauge!Gauge1.Value = Int(100 * Riga / 40)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' 12 caratteri inclusa la virgola !
            FormEsploraPar!Grid1.Row = Riga
                If FormEsploraPar!Grid1.Text <> "" Then
'                If CDbl(FormEsploraPar!Grid1.Text) <> Int(CDbl(FormEsploraPar!Grid1.Text)) Then
                If Riga <> 2 Then
                    If Format(CVar(FormEsploraPar!Grid1.Text), StringaDiFormattazione) <> 0 Then
                        FormEsploraPar!Grid1.Text = Format(CVar(FormEsploraPar!Grid1.Text), StringaDiFormattazione)
                    Else
                        FormEsploraPar!Grid1.Text = Format(CVar(FormEsploraPar!Grid1.Text), "0")
                    End If
                 End If
                End If
    Next Riga
    Next Colonna

    Unload FormGauge
    FormEsploraPar!Grid1.Col = 1
    FormEsploraPar!Grid1.Row = 1
    FormEsploraPar!Grid1.Refresh
    FormEsploraPar!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

Private Sub OpzioniQuattroDecimali_Click()
    
    Dim StringaDiFormattazione As String, Risposta
    Dim Riga, Colonna

    Risposta = MsgBox("In seguito all'operazione di arrotondamento" + Chr(10) + "i decimali in eccesso saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Arrotonda a quattro decimali")
        If Risposta = IDNO Then
            Exit Sub
            FormEsploraPar!Grid1.SetFocus
        End If

' limita a 12 i caratteri, inclusa la virgola
    StringaDiFormattazione = "######0.0000"

    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Arrotonda a quattro decimali"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' arrotonda
    For Colonna = 1 To NUM_COLONNE - 1
    FormEsploraPar!Grid1.Col = Colonna
    
    For Riga = 1 To 40

        FormGauge.Label1.Caption = Int(100 * Riga / 40) & "%"
        FormGauge!Label6.Caption = "[" & Colonna & "/" & NUM_COLONNE - 1 & "]"
        FormGauge!Gauge1.Value = Int(100 * Riga / 40)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' 12 caratteri inclusa la virgola !
            FormEsploraPar!Grid1.Row = Riga
                If FormEsploraPar!Grid1.Text <> "" Then
'                If CDbl(FormEsploraPar!Grid1.Text) <> Int(CDbl(FormEsploraPar!Grid1.Text)) Then
                If Riga <> 2 Then
                    If Format(CVar(FormEsploraPar!Grid1.Text), StringaDiFormattazione) <> 0 Then
                        FormEsploraPar!Grid1.Text = Format(CVar(FormEsploraPar!Grid1.Text), StringaDiFormattazione)
                    Else
                        FormEsploraPar!Grid1.Text = Format(CVar(FormEsploraPar!Grid1.Text), "0")
                    End If
                 End If
                End If
    Next Riga
    Next Colonna

    Unload FormGauge
    FormEsploraPar!Grid1.Col = 1
    FormEsploraPar!Grid1.Row = 1
    FormEsploraPar!Grid1.Refresh
    FormEsploraPar!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

Private Sub OpzioniRipristinaAltezzaRighe_Click()
    
    Dim Riga
    
    Screen.MousePointer = 11
    
    FormPrincipale!Grid1.HighLight = False

    For Riga = 0 To 40
        FormEsploraPar!Grid1.RowHeight(Riga) = 255
    Next Riga

    Screen.MousePointer = 1

End Sub

Private Sub OpzioniRipristinaLarghezzaColonne_Click()
    
    Dim Colonna
    
    Screen.MousePointer = 11
    
    FormEsploraPar!Grid1.ColWidth(0) = 2500
    FormEsploraPar!Grid1.Row = 0
        For Colonna = 1 To NUM_COLONNE - 1
            FormEsploraPar!Grid1.ColWidth(Colonna) = LARGHEZZA_COLONNE
        Next Colonna
    
    Screen.MousePointer = 1

End Sub

Private Sub OpzioniTreDecimali_Click()
    
    Dim StringaDiFormattazione As String, Risposta
    Dim Riga, Colonna

    Risposta = MsgBox("In seguito all'operazione di arrotondamento" + Chr(10) + "i decimali in eccesso saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Arrotonda a tre decimali")
        If Risposta = IDNO Then
            Exit Sub
            FormEsploraPar!Grid1.SetFocus
        End If

' limita a 12 i caratteri, inclusa la virgola
    StringaDiFormattazione = "######0.000"

    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Arrotonda a tre decimali"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' arrotonda
    For Colonna = 1 To NUM_COLONNE - 1
    FormEsploraPar!Grid1.Col = Colonna
    
    For Riga = 1 To 40

        FormGauge.Label1.Caption = Int(100 * Riga / 40) & "%"
        FormGauge!Label6.Caption = "[" & Colonna & "/" & NUM_COLONNE - 1 & "]"
        FormGauge!Gauge1.Value = Int(100 * Riga / 40)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' 12 caratteri inclusa la virgola !
            FormEsploraPar!Grid1.Row = Riga
                If FormEsploraPar!Grid1.Text <> "" Then
'                If CDbl(FormEsploraPar!Grid1.Text) <> Int(CDbl(FormEsploraPar!Grid1.Text)) Then
                If Riga <> 2 Then
                    If Format(CVar(FormEsploraPar!Grid1.Text), StringaDiFormattazione) <> 0 Then
                        FormEsploraPar!Grid1.Text = Format(CVar(FormEsploraPar!Grid1.Text), StringaDiFormattazione)
                    Else
                        FormEsploraPar!Grid1.Text = Format(CVar(FormEsploraPar!Grid1.Text), "0")
                    End If
                 End If
                End If
    Next Riga
    Next Colonna

    Unload FormGauge
    FormEsploraPar!Grid1.Col = 1
    FormEsploraPar!Grid1.Row = 1
    FormEsploraPar!Grid1.Refresh
    FormEsploraPar!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

Private Sub OpzionUnDecimale_Click()
    
    Dim StringaDiFormattazione As String, Risposta
    Dim Riga, Colonna

    Risposta = MsgBox("In seguito all'operazione di arrotondamento" + Chr(10) + "i decimali in eccesso saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Arrotonda a un decimale")
        If Risposta = IDNO Then
            Exit Sub
            FormEsploraPar!Grid1.SetFocus
        End If

' limita a 12 i caratteri, inclusa la virgola
    StringaDiFormattazione = "#########0.0"

    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Arrotonda a un decimale"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' arrotonda
    For Colonna = 1 To NUM_COLONNE - 1
    FormEsploraPar!Grid1.Col = Colonna
    
    For Riga = 1 To 40

        FormGauge.Label1.Caption = Int(100 * Riga / 40) & "%"
        FormGauge!Label6.Caption = "[" & Colonna & "/" & NUM_COLONNE - 1 & "]"
        FormGauge!Gauge1.Value = Int(100 * Riga / 40)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' 12 caratteri inclusa la virgola !
            FormEsploraPar!Grid1.Row = Riga
                If FormEsploraPar!Grid1.Text <> "" Then
'                If CDbl(FormEsploraPar!Grid1.Text) <> Int(CDbl(FormEsploraPar!Grid1.Text)) Then
                If Riga <> 2 Then
                    If Format(CVar(FormEsploraPar!Grid1.Text), StringaDiFormattazione) <> 0 Then
                        FormEsploraPar!Grid1.Text = Format(CVar(FormEsploraPar!Grid1.Text), StringaDiFormattazione)
                    Else
                        FormEsploraPar!Grid1.Text = Format(CVar(FormEsploraPar!Grid1.Text), "0")
                    End If
                 End If
                End If
    Next Riga
    Next Colonna

    Unload FormGauge
    FormEsploraPar!Grid1.Col = 1
    FormEsploraPar!Grid1.Row = 1
    FormEsploraPar!Grid1.Refresh
    FormEsploraPar!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

