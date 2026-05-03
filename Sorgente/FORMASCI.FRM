VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form FormASCII 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Importa file ASCII"
   ClientHeight    =   6225
   ClientLeft      =   1125
   ClientTop       =   1875
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
   Icon            =   "FORMASCI.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6225
   ScaleWidth      =   7365
   Begin MSGrid.Grid Grid2 
      Height          =   1695
      Left            =   3960
      TabIndex        =   1
      Top             =   840
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   2990
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
   Begin MSGrid.Grid Grid1 
      Height          =   1695
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
      _ExtentY        =   2990
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
      Begin VB.Menu FileImporta 
         Caption         =   "&Importa"
      End
      Begin VB.Menu FileImportaDalleColonne 
         Caption         =   "Importa &dalle colonne"
      End
      Begin VB.Menu FileImportaNellaColonna 
         Caption         =   "Importa &nella colonna"
      End
      Begin VB.Menu FileSeparatoreUno 
         Caption         =   "-"
      End
      Begin VB.Menu FileEsci 
         Caption         =   "&Esci"
      End
   End
   Begin VB.Menu MenuOpzioni 
      Caption         =   "&Opzioni"
      Begin VB.Menu OpzioniGriglia 
         Caption         =   "&Griglia"
         Checked         =   -1  'True
      End
      Begin VB.Menu OpzioniSeparatoreUno 
         Caption         =   "-"
      End
      Begin VB.Menu OpzioniRipristinaAltezzaRighe 
         Caption         =   "Ripristina &altezza righe"
      End
      Begin VB.Menu OpzioniRipristinaLarghezzaColonne 
         Caption         =   "Ripristina &larghezza colonne"
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
Attribute VB_Name = "FormASCII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim DallaColonna, AllaColonna
    Dim RigheFileASCII

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
    On Error GoTo GestioneErroriFileEsci
    Dim Riga, Colonna, PrimaParte, SecondaParte, Risposta
    Dim Eureka, Posizione, NumeroDummy, Dummy, I
    Dim StringaDummy As String

    FormASCII!Grid1.HighLight = False
    FormASCII!Grid2.HighLight = False
    FormASCII!Grid1.Visible = False
    FormASCII!Grid2.Visible = True
    FormASCII!Grid2.Refresh
    
    Risposta = MsgBox("Conversione dei dati del file." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON1, "Importa file ASCII")
    If Risposta = IDNO Then
        FormGauge!Label6.Caption = "[1/1]"
        InizializzaFile
        FormPrincipale.Show
        Unload FormASCII
        RipristinaSchermo
        NomeDellaTabella = ""
        NomeDelFile = ""
        FormPrincipale.Caption = Logo
        Exit Sub
    Else

    Risposta = MsgBox("La prima riga contiene" + Chr(10) + "i nomi delle variabili ?", MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON1, "Importa file ASCII")
    If Risposta = IDNO Then

1234
    Screen.MousePointer = 11
    
    NumeroDummy = 69.69
    StringaDummy = Format(NumeroDummy, "##.##")

    FormGauge.Show
    FormGauge.Caption = "Conversione dei dati"
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
        For Riga = 1 To RigheFileASCII
            FormGauge.Label1.Caption = Int(100 * Riga / RigheFileASCII) & "%"
            FormGauge!Gauge1.Value = Int(100 * Riga / RigheFileASCII)
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            FormASCII!Grid2.Row = Riga
            FormPrincipale!Grid1.Row = Riga
                For Colonna = 1 To NUM_COLONNE - 1
                    FormASCII!Grid2.Col = Colonna
                    Dummy = FormASCII!Grid2.Text
                    FormPrincipale!Grid1.Col = Colonna

' e qui comincia la sinfonia...
                
                If StringaDummy = "69,69" Then
                Eureka = False
                    For I = 1 To Len(Dummy)
                        If Mid(Dummy, I, 1) = "." Then
                            Posizione = I
                            Eureka = True
                        End If
                    Next I
                    If Eureka = True Then
                        PrimaParte = Left(Dummy, Posizione - 1)
                        SecondaParte = Right(Dummy, Len(Dummy) - Posizione)
                        FormPrincipale!Grid1.Text = PrimaParte & "," & SecondaParte
                    Else
                        FormPrincipale!Grid1.Text = Dummy
                    End If
                End If
                
                If StringaDummy = "69.69" Then
                Eureka = False
                    For I = 1 To Len(Dummy)
                        If Mid(Dummy, I, 1) = "," Then
                            Posizione = I
                            Eureka = True
                        End If
                    Next I
                    If Eureka = True Then
                        PrimaParte = Left(Dummy, Posizione - 1)
                        SecondaParte = Right(Dummy, Len(Dummy) - Posizione)
                        FormPrincipale!Grid1.Text = PrimaParte & "." & SecondaParte
                    Else
                        FormPrincipale!Grid1.Text = Dummy
                    End If
                End If
' ...che qui finisce !
5678
' genera un errore se il dato non è convertibile in numerico !!!!
                    If FormPrincipale!Grid1.Text <> "" Then
                        Dummy = CDbl(FormPrincipale!Grid1.Text)
                    End If
                Next Colonna
        Next Riga
    
    Unload FormGauge
    FormPrincipale.Show
    Unload FormASCII
    
    RipristinaSchermo
    NomeDellaTabella = ""
    FormPrincipale.Caption = Logo & " [" & NomeDelFile & "]"
    
    Screen.MousePointer = 1
    
    Exit Sub
    
    Else
    
    Screen.MousePointer = 11
    
    NumeroDummy = 69.69
    StringaDummy = Format(NumeroDummy, "##.##")

    FormGauge.Show
    FormGauge.Caption = "Conversione dei dati"
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
        For Riga = 1 To RigheFileASCII
            FormGauge.Label1.Caption = Int(100 * Riga / RigheFileASCII) & "%"
            FormGauge!Gauge1.Value = Int(100 * Riga / RigheFileASCII)
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            FormASCII!Grid2.Row = Riga
            FormPrincipale!Grid1.Row = Riga - 1
                For Colonna = 1 To NUM_COLONNE - 1
                    FormASCII!Grid2.Col = Colonna
                    Dummy = FormASCII!Grid2.Text
                    FormPrincipale!Grid1.Col = Colonna

' e qui comincia la sinfonia...
                
                If StringaDummy = "69,69" Then
                Eureka = False
                    For I = 1 To Len(Dummy)
                        If Mid(Dummy, I, 1) = "." Then
                            Posizione = I
                            Eureka = True
                        End If
                    Next I
                    If Eureka = True Then
                        PrimaParte = Left(Dummy, Posizione - 1)
                        SecondaParte = Right(Dummy, Len(Dummy) - Posizione)
                        FormPrincipale!Grid1.Text = PrimaParte & "," & SecondaParte
                    Else
                        FormPrincipale!Grid1.Text = Dummy
                    End If
                End If
                
                If StringaDummy = "69.69" Then
                Eureka = False
                    For I = 1 To Len(Dummy)
                        If Mid(Dummy, I, 1) = "," Then
                            Posizione = I
                            Eureka = True
                        End If
                    Next I
                    If Eureka = True Then
                        PrimaParte = Left(Dummy, Posizione - 1)
                        SecondaParte = Right(Dummy, Len(Dummy) - Posizione)
                        FormPrincipale!Grid1.Text = PrimaParte & "." & SecondaParte
                    Else
                        FormPrincipale!Grid1.Text = Dummy
                    End If
                End If
' ...che qui finisce !
5679
' genera un errore se il dato non è convertibile in numerico !!!!
                    If Riga > 1 Then
                    If FormPrincipale!Grid1.Text <> "" Then
                        Dummy = CDbl(FormPrincipale!Grid1.Text)
                    End If
                    End If
                Next Colonna
        Next Riga
    
    Unload FormGauge
    FormPrincipale.Show
    Unload FormASCII
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
    NomeDellaTabella = ""
    FormPrincipale.Caption = Logo & " [" & NomeDelFile & "]"
    
    Screen.MousePointer = 1

    Exit Sub
    
GestioneErroriFileEsci:
    If Erl = 5678 Or Erl = 5679 Then
        Unload FormGauge
        FileImporta.Enabled = True
        FormASCII!Grid2.HighLight = False
        FormASCII!Grid2.Row = 0
        Dummy = FormASCII!Grid2.Text
        Screen.MousePointer = 1
        MsgBox "Dato non convertibile:" & Chr(10) & "errore alla riga " & Riga & "," & Chr(10) & "colonna " & Dummy & ".", MB_ICONSTOP, "Conversione dei dati"
        FormASCII!Grid2.Col = Colonna
        FormASCII!Grid2.Row = Riga
        Exit Sub
    Else
        Unload FormGauge
        FormASCII!Grid2.HighLight = False
        Screen.MousePointer = 1
        MsgBox "Errore: file non convertibile", MB_ICONSTOP, "Conversione dei dati"
        FormPrincipale.Show
        Unload FormASCII
        FormPrincipale!Grid1.HighLight = True
        RipristinaSchermo
        NomeDellaTabella = ""
        FormPrincipale.Caption = Logo
        Exit Sub
    End If

    Exit Sub

    End If ' nomi delle variabili ?
    End If ' conferma Esci o rinuncia !

End Sub

Private Sub FileImporta_Click()
    
    Screen.MousePointer = 11

' deseleziona
    FormASCII!Grid2.HighLight = False
    FormASCII!Grid2.SelEndCol = FormASCII!Grid2.SelStartCol
    FormASCII!Grid2.SelEndRow = FormASCII!Grid2.SelStartRow
    FormASCII!Grid2.Row = 1
    FormASCII!Grid2.Col = FormASCII!Grid2.SelStartCol
    FormASCII!Grid2.HighLight = True
' torna a grid1
    FormASCII!Grid2.Visible = False
    FileImporta.Enabled = False
    FormASCII!Grid1.Visible = True
    FormASCII!Grid1.HighLight = False
    FormASCII!Grid1.Row = 1
    FormASCII!Grid1.Col = DallaColonna
    FormASCII!Grid1.HighLight = True
    
    Screen.MousePointer = 1

End Sub

Private Sub FileImportaDalleColonne_Click()

    Screen.MousePointer = 11

' memorizza
    DallaColonna = FormASCII!Grid1.SelStartCol
    AllaColonna = FormASCII!Grid1.SelEndCol
' deseleziona
    FormASCII!Grid1.HighLight = False
    FormASCII!Grid1.SelEndCol = FormASCII!Grid1.SelStartCol
    FormASCII!Grid1.SelEndRow = FormASCII!Grid1.SelStartRow
    FormASCII!Grid1.HighLight = True
    
    If (AllaColonna - DallaColonna) > 11 Then
        FormASCII!Grid1.HighLight = False
        Screen.MousePointer = 1
        MsgBox "Non possono essere importate" + Chr(10) + "più di 12 colonne per volta.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Importa dalle colonne"
        FormASCII!Grid1.HighLight = True
        Exit Sub
    End If

    FormASCII!Grid1.Visible = False
    FormASCII!Grid2.Visible = True

    Screen.MousePointer = 1

End Sub

Private Sub FileImportaNellaColonna_Click()
    Dim Riga, Colonna

    Screen.MousePointer = 11

    FormGauge.Show
    FormGauge.Caption = "Importa come Variabile " & FormASCII!Grid2.SelStartCol
    FormGauge!Label6.Caption = "[1/2]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
    FormASCII!Grid2.Col = FormASCII!Grid2.SelStartCol
        For Riga = 1 To RigheFileASCII
            FormGauge.Label1.Caption = Int(100 * Riga / RigheFileASCII) & "%"
            FormGauge!Gauge1.Value = Int(100 * Riga / RigheFileASCII)
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            FormASCII!Grid2.Row = Riga
            FormASCII!Grid1.Row = Riga
            FormASCII!Grid2.Text = "" ' per prudenza !
                For Colonna = DallaColonna To AllaColonna
                    FormASCII!Grid1.Col = Colonna
                    FormASCII!Grid2.Text = FormASCII!Grid2.Text + FormASCII!Grid1.Text
                Next Colonna
             FormASCII!Grid2.Text = Trim(FormASCII!Grid2.Text)
        Next Riga
' vuota celle in caso di sovrascrittura
    FormGauge!Label6.Caption = "[2/2]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
        For Riga = RigheFileASCII + 1 To NUM_RIGHE - 1
            FormGauge.Label1.Caption = Int(100 * (Riga - RigheFileASCII) / (NUM_RIGHE - 1 - RigheFileASCII)) & "%"
            FormGauge!Gauge1.Value = Int(100 * (Riga - RigheFileASCII) / (NUM_RIGHE - 1 - RigheFileASCII))
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                FormASCII!Grid2.Row = Riga
                FormASCII!Grid2.Text = ""
        Next Riga

    Unload FormGauge

' deseleziona
    FormASCII!Grid2.HighLight = False
    FormASCII!Grid2.SelEndCol = FormASCII!Grid2.SelStartCol
    FormASCII!Grid2.SelEndRow = FormASCII!Grid2.SelStartRow
    FormASCII!Grid2.Row = 1
    FormASCII!Grid2.Col = FormASCII!Grid2.SelStartCol
    FormASCII!Grid2.HighLight = True
    
    FileImporta.Enabled = True
    
    Screen.MousePointer = 1

End Sub

Private Sub Form_Load()
    On Error GoTo GestioneErroriFileImportaFileASCII
    Dim Riga, Colonna
    Dim Dummy As String

    Screen.MousePointer = 11
    FormASCII.Caption = "Importa file ASCII [" & NomeDelFile & "]"
    
    FormGauge.Show
    FormGauge.Caption = "Lettura file ASCII"
    FormGauge!Label6.Caption = "[2/2]"
    FormGauge.Label1.Caption = "0%"
    FormGauge!Gauge1.Value = 0
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' centra il form
    FormASCII.Top = Screen.Height / 2 - Height / 2
    FormASCII.Left = Screen.Width / 2 - Width / 2

' numero di righe e colonne dell'oggetto Grid1
    FormASCII!Grid1.Rows = NUM_RIGHE
    FormASCII!Grid1.Cols = 1 + ((NUM_COLONNE - 1) * 14)
' sposta in alto a sinistra
    FormASCII!Grid1.Move 0, 0
' dimensionamento, ripreso nella sub resize
    FormASCII!Grid1.Width = ScaleWidth
    FormASCII!Grid1.Height = ScaleHeight
' definisce larghezza colonne
    FormASCII!Grid1.ColWidth(0) = 574
    FormASCII!Grid1.Row = 0
        For Colonna = 1 To (NUM_COLONNE - 1) * 14
            FormASCII!Grid1.ColWidth(Colonna) = LARGHEZZA_COLONNE / 10
        Next Colonna
' definisce altezza righe
    For Riga = 0 To NUM_RIGHE - 1
        FormASCII!Grid1.RowHeight(Riga) = 255
    Next Riga
' intestazione righe
    FormASCII!Grid1.Col = 0
        For Riga = 1 To NUM_RIGHE - 1
            FormASCII!Grid1.Row = Riga
            FormASCII!Grid1.Text = Riga
        Next Riga
' intestazione colonne
    FormASCII!Grid1.Row = 0
        For Colonna = 1 To (NUM_COLONNE - 1) * 14
            FormASCII!Grid1.Col = Colonna
            FormASCII!Grid1.Text = Colonna
        Next Colonna
' allineamento colonne
    FormASCII!Grid1.FixedAlignment(0) = 1 ' intestazioni righe a destra
        For Colonna = 1 To NUM_COLONNE - 1
            FormASCII!Grid1.FixedAlignment(Colonna) = 2 ' nome variabili centrato
            FormASCII!Grid1.ColAlignment(Colonna) = 1 ' dati a destra
        Next Colonna

' numero di righe e colonne dell'oggetto Grid2
    FormASCII!Grid2.Rows = NUM_RIGHE
    FormASCII!Grid2.Cols = NUM_COLONNE
' sposta in alto a sinistra
    FormASCII!Grid2.Move 0, 0
' dimensionamento, ripreso nella sub resize
    FormASCII!Grid2.Width = ScaleWidth
    FormASCII!Grid2.Height = ScaleHeight
' definisce larghezza colonne
    FormASCII!Grid2.ColWidth(0) = 574
    FormASCII!Grid2.Row = 0
        For Colonna = 1 To NUM_COLONNE - 1
            FormASCII!Grid2.ColWidth(Colonna) = LARGHEZZA_COLONNE
        Next Colonna
' definisce altezza righe
    For Riga = 0 To NUM_RIGHE - 1
        FormASCII!Grid2.RowHeight(Riga) = 255
    Next Riga
' intestazione righe
    FormASCII!Grid2.Col = 0
        For Riga = 1 To NUM_RIGHE - 1
            FormASCII!Grid2.Row = Riga
            FormASCII!Grid2.Text = Riga
        Next Riga
' intestazione colonne
    FormASCII!Grid2.Row = 0
        For Colonna = 1 To NUM_COLONNE - 1
            FormASCII!Grid2.Col = Colonna
            FormASCII!Grid2.Text = "Variabile " & Colonna
        Next Colonna
' allineamento colonne
    FormASCII!Grid2.FixedAlignment(0) = 1 ' intestazioni righe a destra
        For Colonna = 1 To NUM_COLONNE - 1
            FormASCII!Grid2.FixedAlignment(Colonna) = 2 ' nome variabili centrato
            FormASCII!Grid2.ColAlignment(Colonna) = 1 ' dati a destra
        Next Colonna

' termine inizializzazioni e inizio import file ASCII

' conta righe del file ASCII
    RigheFileASCII = 0

    Open NomeDelFile For Input As #1
        Do While Not EOF(1)
            Line Input #1, Dummy
            RigheFileASCII = RigheFileASCII + 1
        Loop

    Close #1
    
    If RigheFileASCII > NUM_RIGHE - 1 Then RigheFileASCII = NUM_RIGHE - 1

' importa file in Grid1
    Open NomeDelFile For Input As #1
        For Riga = 1 To RigheFileASCII
            FormGauge.Label1.Caption = Int(100 * Riga / RigheFileASCII) & "%"
            FormGauge!Gauge1.Value = Int(100 * Riga / RigheFileASCII)
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            
            Line Input #1, Dummy
            If Len(Dummy) > (NUM_COLONNE - 1) * 14 Then Dummy = Left(Dummy, (NUM_COLONNE - 1) * 14)
            FormASCII!Grid1.Row = Riga
                For Colonna = 1 To Len(Dummy)
                    FormASCII!Grid1.Col = Colonna
                    FormASCII!Grid1.Text = Mid(Dummy, Colonna, 1)
                Next Colonna
        Next Riga
    
    Close #1

    Unload FormGauge
    FormASCII!Grid1.Visible = True
    FormASCII!Grid2.Visible = False
    FormASCII!Grid1.HighLight = False
    FormASCII!Grid1.Col = 1
    FormASCII!Grid1.Row = 1
'    FormASCII!Grid1.HighLight = True
'    FormASCII!Grid2.HighLight = False
    FormASCII!Grid2.Col = 1
    FormASCII!Grid2.Row = 1
    FormASCII!Grid2.HighLight = True
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
    FormASCII.Show
    ' FormPrincipale.Hide
    FileImporta.Enabled = False
    Screen.MousePointer = 1
    
    Exit Sub

GestioneErroriFileImportaFileASCII:
    Unload FormGauge
    Screen.MousePointer = 1
    MsgBox "Errore su disc(hett)o," & Chr(10) & "o errore nella struttura" & Chr(10) & "o nei driver del file.", MB_ICONSTOP, "Importa file dBase III"
    FileEsci_Click
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo
    Exit Sub

End Sub

Private Sub Form_Resize()

    FormASCII!Grid1.Width = ScaleWidth
    FormASCII!Grid1.Height = ScaleHeight
    FormASCII!Grid2.Width = ScaleWidth
    FormASCII!Grid2.Height = ScaleHeight

End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    FormASCII!Grid1.HighLight = True

End Sub

Private Sub Grid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    FormASCII.SetFocus

End Sub

Private Sub Grid1_SelChange()

' nessuna selezione
    If (FormASCII!Grid1.SelStartRow = FormASCII!Grid1.SelEndRow) And (FormASCII!Grid1.SelStartCol = FormASCII!Grid1.SelEndCol) Then
        Exit Sub
    End If
' selezionate una o più colonne
    If FormASCII!Grid1.SelStartRow = 1 And FormASCII!Grid1.SelEndRow = NUM_RIGHE - 1 Then
        Exit Sub
    End If
' selezione non ammessa
    FormASCII!Grid1.HighLight = False
    FormASCII!Grid1.SelEndCol = FormASCII!Grid1.SelStartCol
    FormASCII!Grid1.SelEndRow = FormASCII!Grid1.SelStartRow
    FormASCII!Grid1.Col = FormASCII!Grid1.SelStartCol
    FormASCII!Grid1.Row = FormASCII!Grid1.SelStartRow

End Sub

Private Sub Grid2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    FormASCII!Grid2.HighLight = True

End Sub

Private Sub Grid2_SelChange()

' nessuna selezione
    If (FormASCII!Grid2.SelStartRow = FormASCII!Grid2.SelEndRow) And (FormASCII!Grid2.SelStartCol = FormASCII!Grid2.SelEndCol) Then
        Exit Sub
    End If
' selezionate una colonna
    If (FormASCII!Grid2.SelStartRow = 1 And FormASCII!Grid2.SelEndRow = NUM_RIGHE - 1) And (FormASCII!Grid2.SelStartCol = FormASCII!Grid2.SelEndCol) Then
        Exit Sub
    End If
' selezione non ammessa
    FormASCII!Grid2.HighLight = False
    FormASCII!Grid2.SelEndCol = FormASCII!Grid2.SelStartCol
    FormASCII!Grid2.SelEndRow = FormASCII!Grid2.SelStartRow
    FormASCII!Grid2.Col = FormASCII!Grid2.SelStartCol
    FormASCII!Grid2.Row = FormASCII!Grid2.SelStartRow

End Sub

Private Sub MenuFile_Click()

    FileImportaDalleColonne.Enabled = False
    FileImportaNellaColonna.Enabled = False
    
    If FormASCII!Grid1.SelStartRow = 1 And FormASCII!Grid1.SelEndRow = NUM_RIGHE - 1 Then
        FileImportaDalleColonne.Enabled = True
        FileImportaNellaColonna.Enabled = False
    End If
    If (FormASCII!Grid2.SelStartRow = 1 And FormASCII!Grid2.SelEndRow = NUM_RIGHE - 1) And (FormASCII!Grid2.SelStartCol = FormASCII!Grid2.SelEndCol) Then
        FileImportaDalleColonne.Enabled = False
        FileImportaNellaColonna.Enabled = True
    End If

End Sub

Private Sub OpzioniGriglia_Click()
    
    OpzioniGriglia.Checked = Not OpzioniGriglia.Checked
    
    If FormASCII!Grid1.Visible = True Then
        FormASCII!Grid1.GridLines = Not FormASCII!Grid1.GridLines
    Else
        FormASCII!Grid2.GridLines = Not FormASCII!Grid2.GridLines
    End If

End Sub

Private Sub OpzioniRipristinaAltezzaRighe_Click()
    Dim Riga
    
    Screen.MousePointer = 11
    
    If FormASCII!Grid1.Visible = True Then
        FormASCII!Grid1.HighLight = False
        FormASCII!Grid1.SelEndCol = FormASCII!Grid1.SelStartCol
        FormASCII!Grid1.SelEndRow = FormASCII!Grid1.SelStartRow
        FormASCII!Grid1.Col = FormASCII!Grid1.SelStartCol
        FormASCII!Grid1.Row = FormASCII!Grid1.SelStartRow
            For Riga = 0 To NUM_RIGHE - 1
                FormASCII!Grid1.RowHeight(Riga) = 255
            Next Riga
        FormASCII!Grid1.SetFocus
    Else
        FormASCII!Grid2.HighLight = False
        FormASCII!Grid2.SelEndCol = FormASCII!Grid2.SelStartCol
        FormASCII!Grid2.SelEndRow = FormASCII!Grid2.SelStartRow
        FormASCII!Grid2.Col = FormASCII!Grid2.SelStartCol
        FormASCII!Grid2.Row = FormASCII!Grid2.SelStartRow
            For Riga = 0 To NUM_RIGHE - 1
                FormASCII!Grid2.RowHeight(Riga) = 255
            Next Riga
        FormASCII!Grid2.SetFocus
    End If

    Screen.MousePointer = 1

'

' ripristina schermo
'    FormPrincipale!Grid1.Row = Rig
'    FormPrincipale!Grid1.Col = Col
'    FormPrincipale!Grid1.TopRow = RigSup
'    FormPrincipale!Grid1.LeftCol = ColSin
'    FormPrincipale!Grid1.SetFocus


End Sub

Private Sub OpzioniRipristinaLarghezzaColonne_Click()
    Dim Colonna
    
    Screen.MousePointer = 11

    If FormASCII!Grid1.Visible = True Then
        FormASCII!Grid1.HighLight = False
        FormASCII!Grid1.SelEndCol = FormASCII!Grid1.SelStartCol
        FormASCII!Grid1.SelEndRow = FormASCII!Grid1.SelStartRow
        FormASCII!Grid1.Col = FormASCII!Grid1.SelStartCol
        FormASCII!Grid1.Row = FormASCII!Grid1.SelStartRow
        FormASCII!Grid1.ColWidth(0) = 574
'        FormASCII!Grid1.Row = 0
            For Colonna = 1 To (NUM_COLONNE - 1) * 14
               FormASCII!Grid1.ColWidth(Colonna) = LARGHEZZA_COLONNE / 10
            Next Colonna
        FormASCII!Grid1.SetFocus
    Else
        FormASCII!Grid2.HighLight = False
        FormASCII!Grid2.SelEndCol = FormASCII!Grid2.SelStartCol
        FormASCII!Grid2.SelEndRow = FormASCII!Grid2.SelStartRow
        FormASCII!Grid2.Col = FormASCII!Grid2.SelStartCol
        FormASCII!Grid2.Row = FormASCII!Grid2.SelStartRow
        FormASCII!Grid2.ColWidth(0) = 574
'        FormASCII!Grid2.Row = 0
            For Colonna = 1 To NUM_COLONNE - 1
               FormASCII!Grid2.ColWidth(Colonna) = LARGHEZZA_COLONNE
            Next Colonna
        FormASCII!Grid2.SetFocus
    End If

    Screen.MousePointer = 1

'    FormPrincipale!Grid1.HighLight = False
    

' ripristina schermo
'    FormPrincipale!Grid1.Row = Rig
'    FormPrincipale!Grid1.Col = Col
'    FormPrincipale!Grid1.TopRow = RigSup
'    FormPrincipale!Grid1.LeftCol = ColSin
'    FormPrincipale!Grid1.SetFocus


End Sub

