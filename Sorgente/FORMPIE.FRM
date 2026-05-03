VERSION 5.00
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form FormPie 
   AutoRedraw      =   -1  'True
   Caption         =   "Diagramma a torta"
   ClientHeight    =   6195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8460
   Icon            =   "FormPie.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin GraphLib.Graph Graph1 
      Height          =   1215
      Left            =   1200
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   2143
      _StockProps     =   96
      BorderStyle     =   1
      Background      =   7
      Foreground      =   0
      GraphType       =   1
      PrintStyle      =   1
      RandomData      =   1
      ColorData       =   0
      ExtraData       =   0
      ExtraData[]     =   0
      FontFamily      =   4
      FontSize        =   4
      FontSize[0]     =   200
      FontSize[1]     =   150
      FontSize[2]     =   100
      FontSize[3]     =   100
      FontStyle       =   4
      FontStyle[0]    =   2
      GraphData       =   0
      GraphData[]     =   0
      LabelText       =   0
      LegendText      =   0
      PatternData     =   0
      SymbolData      =   0
      XPosData        =   0
      XPosData[]      =   0
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
   Begin VB.Menu Grafica 
      Caption         =   "&Grafica"
      Begin VB.Menu GraficaTorta 
         Caption         =   "&Diagramma a torta"
         Begin VB.Menu GraficaTorta2D 
            Caption         =   "&2D"
         End
         Begin VB.Menu GraficaTorta3D 
            Caption         =   "&3D"
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
            Caption         =   "&Dati a colori"
            Checked         =   -1  'True
         End
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
Attribute VB_Name = "FormPie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim X() As Double
    Dim NUMCOLONNE As Integer
    Dim NomeVariabile() As String
    Dim NumDati As Integer
    Dim FlagSfondoGrigio
    Dim FlagDatiAColori
    Dim Flag2D, Flag3D
    
Private Sub AiutoCalcolatrice_Click()
    Dim Calcolatrice
    On Error GoTo ErroreCalcolatriceBoh:
    
    Calcolatrice = Shell("Calc.exe", 1)
        
Exit Sub

ErroreCalcolatriceBoh:
    MsgBox "Calcolatrice non disponibile.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcolatrice"
Exit Sub

End Sub

Private Sub AiutoGuida_Click()

    FormPrincipale!CMDialog1.HelpFile = App.Path & "\ministat.hlp"
    FormPrincipale!CMDialog1.HelpCommand = cdlHelpContents
    FormPrincipale!CMDialog1.ShowHelp

End Sub

Private Sub File_Click()

End Sub

Private Sub FileEsci_Click()
    
    Unload FormPie
    FormPrincipale.Show
    FormPrincipale!Grid1.SelEndCol = FormPrincipale!Grid1.SelStartCol
    FormPrincipale!Grid1.SelEndRow = FormPrincipale!Grid1.SelStartRow
    FormPrincipale!Grid1.Col = FormPrincipale!Grid1.SelStartCol
    
End Sub

Private Sub FilseStampa_Click()
    On Error GoTo ErroreFileStampaTorta

    If Flag2D = True Or Flag3D = True Then
        
        FormPie!Graph1.Width = ScaleWidth * 0.8
        FormPie!Graph1.Height = ScaleHeight * 0.8
        
        If FlagDatiAColori = True Then
            FormPie!Graph1.Palette = 0
        Else
            FormPie!Graph1.Palette = 2
        End If
        
        FormPie!Graph1.PrintStyle = 1
        FormPie!Graph1.DrawMode = 5
        
        PrintForm
        
        FormPie!Graph1.Width = ScaleWidth
        FormPie!Graph1.Height = ScaleHeight
            
            If Flag2D = True Then
                FormPie!Graph1.DrawMode = 0
                GraficaTorta2D_Click
            End If
            If Flag3D = True Then
                FormPie!Graph1.DrawMode = 0
                GraficaTorta3D_Click
            End If
    
    End If

    Exit Sub

ErroreFileStampaTorta:
    Exit Sub

End Sub

Private Sub FileStampa_Click()

    If Flag2D = True Or Flag3D = True Then
        PrintForm
    End If
    
End Sub

Private Sub Form_Load()
    On Error GoTo GestioneErroriTorta
    Dim I, K, Riga, Dummy, FineDati

    Screen.MousePointer = 11
    FormPie!Graph1.Visible = False
    FlagSfondoGrigio = True
    FlagDatiAColori = True
    Flag2D = False
    Flag3D = False
    
    If FormPrincipale!Grid1.SelStartCol <> FormPrincipale!Grid1.SelEndCol Then
    If FormPrincipale!Grid1.SelStartRow = FormPrincipale!Grid1.SelEndRow Then
        
    NUMCOLONNE = FormPrincipale!Grid1.SelEndCol - FormPrincipale!Grid1.SelStartCol + 1

    ReDim X(NUMCOLONNE) As Double
    ReDim NomeVariabile(NUMCOLONNE) As String

    AnnullatoCaricamentoForm = False

' centra il form
    FormPie.Top = Screen.Height / 2 - Height / 2
    FormPie.Left = Screen.Width / 2 - Width / 2
' sposta in alto a sinistra
    FormPie!Graph1.Move 0, 0
' dimensionamento, ripreso nella sub resize
    FormPie!Graph1.Width = ScaleWidth
    FormPie!Graph1.Height = ScaleHeight

' nome delle variabili
    
    K = FormPrincipale!Grid1.SelEndCol - FormPrincipale!Grid1.SelStartCol + 1
    FormPrincipale!Grid1.Row = FormPrincipale!Grid1.SelStartRow
    
    NumDati = 0
    
    For I = 1 To K
        If FineDati = False Then
            FormPrincipale!Grid1.Col = I
            If FormPrincipale!Grid1.Text <> "" Then
                NumDati = NumDati + 1
            Else
                FineDati = True
            End If
        End If
    Next I
    
    FormPrincipale!Grid1.Row = 0
    For I = 1 To NumDati
        FormPrincipale!Grid1.Col = I
        NomeVariabile(I) = Left(FormPrincipale!Grid1.Text, 20)
    Next I

' carica i dati nel vettore X()

    FormPrincipale!Grid1.Row = FormPrincipale!Grid1.SelStartRow
    
    For I = 1 To NumDati
        FormPrincipale!Grid1.Col = I
        X(I) = CDbl(-(-FormPrincipale!Grid1.Text))
    Next I

    If NumDati < 2 Then
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        MsgBox "Per il diagramma a torta" + Chr(10) + "sono necessari almeno 2 dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
        FormPrincipale!Grid1.SelEndCol = FormPrincipale!Grid1.SelStartCol
        FormPrincipale!Grid1.SelEndRow = FormPrincipale!Grid1.SelStartRow
        FormPrincipale!Grid1.Col = FormPrincipale!Grid1.SelStartCol
        Exit Sub
    Else
        FormPrincipale!Grid1.SelEndCol = FormPrincipale!Grid1.SelStartCol
        FormPrincipale!Grid1.SelEndRow = FormPrincipale!Grid1.SelStartRow
        FormPrincipale!Grid1.Col = FormPrincipale!Grid1.SelStartCol
        FormPie.Show
        ' FormPrincipale.Hide
        GraficaTorta2D_Click
        Screen.MousePointer = 1
        Exit Sub
    End If
    
    End If 'selezionata una sola riga
    End If 'selezionate più colonne
    
        MsgBox "Per il diagramma a torta" + Chr(10) + "selezionare una riga di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
        AnnullatoCaricamentoForm = True
        Screen.MousePointer = 1
        FormPrincipale!Grid1.SelEndCol = FormPrincipale!Grid1.SelStartCol
        FormPrincipale!Grid1.SelEndRow = FormPrincipale!Grid1.SelStartRow
        FormPrincipale!Grid1.Col = FormPrincipale!Grid1.SelStartCol
        Exit Sub

GestioneErroriTorta:
        Screen.MousePointer = 1
        MsgBox "Memoria insufficiente: riprovare" + Chr(10) + "con un numero inferiore di dati.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Dati non rappresentabili"
        FormPrincipale!Grid1.SelEndCol = FormPrincipale!Grid1.SelStartCol
        FormPrincipale!Grid1.SelEndRow = FormPrincipale!Grid1.SelStartRow
        FormPrincipale!Grid1.Col = FormPrincipale!Grid1.SelStartCol
        Exit Sub

End Sub

Private Sub Form_Resize()
    
    FormPie!Graph1.Width = ScaleWidth
    FormPie!Graph1.Height = ScaleHeight
        
        If Flag2D = True Then
        FormPie!Graph1.DrawMode = 0
        GraficaTorta2D_Click
    End If

    If Flag3D = True Then
        FormPie!Graph1.DrawMode = 0
        GraficaTorta3D_Click
    End If

End Sub

Private Sub GraficaDatiAColori_Click()
    
    GraficaDatiAColori.Checked = Not GraficaDatiAColori.Checked
    FlagDatiAColori = Not FlagDatiAColori

    If Flag2D = True Then
        FormPie!Graph1.DrawMode = 0
        GraficaTorta2D_Click
    End If

    If Flag3D = True Then
        FormPie!Graph1.DrawMode = 0
        GraficaTorta3D_Click
    End If

End Sub

Private Sub GraficaSfondoGrigio_Click()
    
    GraficaSfondoGrigio.Checked = Not GraficaSfondoGrigio.Checked
    FlagSfondoGrigio = Not FlagSfondoGrigio

    If Flag2D = True Then
        FormPie!Graph1.DrawMode = 0
        GraficaTorta2D_Click
    End If

    If Flag3D = True Then
        FormPie!Graph1.DrawMode = 0
        GraficaTorta3D_Click
    End If

End Sub

Private Sub GraficaTorta2D_Click()
    Dim I As Integer
    
    FormPie!Graph1.NumPoints = NumDati
    
    If FlagDatiAColori = True Then
        
        GraficaSfondoGrigio.Enabled = True
    
        If FlagSfondoGrigio = True Then
            FormPie!Graph1.Background = 7
        Else
            FormPie!Graph1.Background = 15
        End If
        
        FormPie!Graph1.DrawStyle = 1
        FormPie!Graph1.Palette = 0
        FormPie!Graph1.DataReset = 2
        FormPie!Graph1.DataReset = 6
        
        For I = 1 To NumDati
            FormPie!Graph1.ThisPoint = I
            FormPie!Graph1.GraphData = X(I)
            FormPie!Graph1.ThisPoint = I
            FormPie!Graph1.LabelText = NomeVariabile(I)
            FormPie!Graph1.ThisPoint = I
            FormPie!Graph1.ColorData = 5 + I
                If I + 5 = 7 Then
                    FormPie!Graph1.ThisPoint = I
                    FormPie!Graph1.ColorData = 3
                End If
                If I + 5 = 8 Then
                    FormPie!Graph1.ThisPoint = I
                    FormPie!Graph1.ColorData = 4
                End If
            FormPie!Graph1.ThisPoint = I
            FormPie!Graph1.PatternData = 0
        Next I
        
    Else
            
            GraficaSfondoGrigio.Enabled = False
            FormPie!Graph1.Background = 15
            FormPie!Graph1.DrawStyle = 1
            FormPie!Graph1.Palette = 2
            FormPie!Graph1.DataReset = 2
            FormPie!Graph1.DataReset = 6
            
        For I = 1 To NumDati
            FormPie!Graph1.ThisPoint = I
            FormPie!Graph1.GraphData = X(I)
            FormPie!Graph1.ThisPoint = I
            FormPie!Graph1.LabelText = NomeVariabile(I)
            FormPie!Graph1.ThisPoint = I
            FormPie!Graph1.ColorData = 7
            FormPie!Graph1.ThisPoint = I
            FormPie!Graph1.PatternData = 20 + I
        Next I
        
    End If

    FormPie!Graph1.GraphType = 1    ' grafico 2D
    FormPie!Graph1.ThisPoint = 1
    FormPie!Graph1.ExtraData = 1
    FormPie!Graph1.DrawMode = 2
    FormPie!Graph1.Visible = True
    
    Flag2D = True
    Flag3D = False
    
End Sub

Private Sub GraficaTorta3D_Click()
    Dim I As Integer
    
    FormPie!Graph1.NumPoints = NumDati
    
    If FlagDatiAColori = True Then
        
        GraficaSfondoGrigio.Enabled = True
    
        If FlagSfondoGrigio = True Then
            FormPie!Graph1.Background = 7
        Else
            FormPie!Graph1.Background = 15
        End If
        
        FormPie!Graph1.DrawStyle = 1
        FormPie!Graph1.Palette = 0
        FormPie!Graph1.DataReset = 2
        FormPie!Graph1.DataReset = 6
        
        For I = 1 To NumDati
            FormPie!Graph1.ThisPoint = I
            FormPie!Graph1.GraphData = X(I)
            FormPie!Graph1.ThisPoint = I
            FormPie!Graph1.LabelText = NomeVariabile(I)
            FormPie!Graph1.ThisPoint = I
            FormPie!Graph1.ColorData = 5 + I
                If I + 5 = 7 Then
                    FormPie!Graph1.ThisPoint = I
                    FormPie!Graph1.ColorData = 3
                End If
                If I + 5 = 8 Then
                    FormPie!Graph1.ThisPoint = I
                    FormPie!Graph1.ColorData = 4
                End If
            FormPie!Graph1.ThisPoint = I
            FormPie!Graph1.PatternData = 0
        Next I
        
    Else
            
            GraficaSfondoGrigio.Enabled = False
            FormPie!Graph1.Background = 15
            FormPie!Graph1.DrawStyle = 1
            FormPie!Graph1.Palette = 2
            FormPie!Graph1.DataReset = 2
            FormPie!Graph1.DataReset = 6
            
        For I = 1 To NumDati
            FormPie!Graph1.ThisPoint = I
            FormPie!Graph1.GraphData = X(I)
            FormPie!Graph1.ThisPoint = I
            FormPie!Graph1.LabelText = NomeVariabile(I)
            FormPie!Graph1.ThisPoint = I
            FormPie!Graph1.ColorData = 7
            FormPie!Graph1.ThisPoint = I
            FormPie!Graph1.PatternData = 20 + I
        Next I
        
    End If

    FormPie!Graph1.GraphType = 2    ' grafico 3D
    FormPie!Graph1.ThisPoint = 1
    FormPie!Graph1.ExtraData = 1
    FormPie!Graph1.DrawMode = 2
    FormPie!Graph1.Visible = True
    
    Flag2D = False
    Flag3D = True
    
End Sub

Private Sub MenuFile_Click()
    
    FileStampa.Enabled = False
        If Flag2D = True Or Flag3D = True Then
            If FormPie.Width < Screen.Width Then
                FileStampa.Enabled = True
            End If
        End If

End Sub
