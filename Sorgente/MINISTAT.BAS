Attribute VB_Name = "MINISTAT1"
Option Explicit
' nome del file in uso e logo
Global NomeDelFile As String
Global NomeDelFileMinistat As String
Global NomeDellaTabella As String
Global Logo As String
Global NumeroPerCalcoli
Global InseritaNuovaColonna
Global NomeDellaVariabile
Global RigaDiPartenza
Global FileSalvatoInPrecedenza
Global AnnullatoCaricamentoForm
Global BarraDegliStrumenti
' riga e colonna del puntatore in Grid1
Global Rig As Single, Col As Single
' riga superiore e colonna sinistra schermo in Grid1
Global RigSup As Single, ColSin As Single
Global Const NUM_RIGHE = 1001
Global Const NUM_COLONNE = 11
Global Const LARGHEZZA_COLONNE = 1447

' MsgBox parameters
Global Const MB_OK = 0                 ' OK button only
Global Const MB_OKCANCEL = 1           ' OK and Cancel buttons
Global Const MB_ABORTRETRYIGNORE = 2   ' Abort, Retry, and Ignore buttons
Global Const MB_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
Global Const MB_YESNO = 4              ' Yes and No buttons
Global Const MB_RETRYCANCEL = 5        ' Retry and Cancel buttons

Global Const MB_ICONSTOP = 16          ' Critical message
Global Const MB_ICONQUESTION = 32      ' Warning query
Global Const MB_ICONEXCLAMATION = 48   ' Warning message
Global Const MB_ICONINFORMATION = 64   ' Information message

Global Const MB_APPLMODAL = 0          ' Application Modal Message Box
Global Const MB_DEFBUTTON1 = 0         ' First button is default
Global Const MB_DEFBUTTON2 = 256       ' Second button is default
Global Const MB_DEFBUTTON3 = 512       ' Third button is default
Global Const MB_SYSTEMMODAL = 4096      'System Modal

' MsgBox return values
Global Const IDOK = 1                  ' OK button pressed
Global Const IDCANCEL = 2              ' Cancel button pressed
Global Const IDABORT = 3               ' Abort button pressed
Global Const IDRETRY = 4               ' Retry button pressed
Global Const IDIGNORE = 5              ' Ignore button pressed
Global Const IDYES = 6                 ' Yes button pressed
Global Const IDNO = 7                  ' No button pressed

'   Help engine declarations.
 
'  Commands to pass WinHelp()
Global Const HELP_CONTEXT = &H1     ' Display topic identified by number in Data
Global Const HELP_QUIT = &H2        ' Terminate help
Global Const HELP_INDEX = &H3       ' Display index
Global Const HELP_HELPONHELP = &H4  ' Display help on using help
Global Const HELP_SETINDEX = &H5    ' Set an alternate Index for help file with more than one index
Global Const HELP_KEY = &H101       ' Display topic for keyword in Data
Global Const HELP_MULTIKEY = &H201  ' Lookup keyword in alternate table and display topic

Declare Function WinHelp Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, dwData As Any) As Integer
 
Type MULTIKEYHELP
    mkSize As Integer
    mkKeylist As String * 1
    szKeyphrase As String * 253
End Type

Function FormattaStringa(DummyIn As String)
    Dim I, Posizione, PrimaParte, SecondaParte
    Dim NumeroDummy, Eureka
    Dim StringaDummy As String
        
    NumeroDummy = 69.69
    StringaDummy = Format(NumeroDummy, "##.##")
    
' e qui comincia la sinfonia...
                
                If StringaDummy = "69,69" Then
                Eureka = False
                    For I = 1 To Len(DummyIn)
                        If Mid(DummyIn, I, 1) = "." Then
                            Posizione = I
                            Eureka = True
                        End If
                    Next I
                    If Eureka = True Then
                        PrimaParte = Left(DummyIn, Posizione - 1)
                        SecondaParte = Right(DummyIn, Len(DummyIn) - Posizione)
                        FormattaStringa = PrimaParte & "," & SecondaParte
                    Else
                        FormattaStringa = DummyIn
                    End If
                End If
                
                If StringaDummy = "69.69" Then
                Eureka = False
                    For I = 1 To Len(DummyIn)
                        If Mid(DummyIn, I, 1) = "," Then
                            Posizione = I
                            Eureka = True
                        End If
                    Next I
                    If Eureka = True Then
                        PrimaParte = Left(DummyIn, Posizione - 1)
                        SecondaParte = Right(DummyIn, Len(DummyIn) - Posizione)
                        FormattaStringa = PrimaParte & "." & SecondaParte
                    Else
                        FormattaStringa = DummyIn
                    End If
                End If
' ...che qui finisce !

End Function
Sub ApriFile(NomeDelFile As String)
    On Error GoTo GestioneErroriApriFile
    Dim Linea As String
    Dim Riga, Colonna, I, Posizione, PrimaParte, SecondaParte
    Dim NumeroDummy, Dummy, Eureka
    Dim StringaDummy As String
    
    NumeroDummy = 69.69
    StringaDummy = Format(NumeroDummy, "##.##")

    Screen.MousePointer = 11
    Open NomeDelFile For Input As #1
    
    FormGauge.Show
    FormGauge.Caption = "Apri file Ministat 1.1"
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' recupera i nomi delle variabili nella riga zero
    FormPrincipale!Grid1.Row = 0
    Line Input #1, Linea

        For Colonna = 1 To NUM_COLONNE - 1
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = Trim(Left(Mid(Linea, ((Colonna - 1) * 14) + 1, 14), 12))
        Next Colonna

' recupera i dati nelle righe successive
    For Riga = 1 To NUM_RIGHE - 1

        FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        FormPrincipale!Grid1.Row = Riga
                
        Line Input #1, Linea

        For Colonna = 1 To NUM_COLONNE - 1
            FormPrincipale!Grid1.Col = Colonna

            If Mid(Linea, ((Colonna - 1) * 14) + 1, 14) = "              " Then
                FormPrincipale!Grid1.Text = ""
            Else
                Dummy = Trim(CVar(Mid(Linea, ((Colonna - 1) * 14) + 1, 14)))
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
            End If
        Next Colonna
    Next Riga

    Close #1

    Unload FormGauge
    RipristinaSchermo
    NomeDellaTabella = ""
    FormPrincipale.Caption = Logo & " [" & NomeDelFile & "]"
    Screen.MousePointer = 1

Exit Sub

GestioneErroriApriFile:
    MsgBox "Errore su disc(hett)o:" & Chr(10) & "per favore riprovare.", MB_ICONSTOP, "Apri file Ministat 1.1"
    NomeDelFile = ""
    FormPrincipale.Caption = Logo
    Unload FormGauge
    RipristinaSchermo
    Screen.MousePointer = 1
    Exit Sub

End Sub

Sub DeselezionaColonna()
    
    FormPrincipale!Grid1.SelEndRow = FormPrincipale!Grid1.SelStartRow
' funziona anche se sono state selezionate più colonne !
    FormPrincipale!Grid1.SelEndCol = FormPrincipale!Grid1.SelStartCol
    
    FormPrincipale!Grid1.HighLight = False

End Sub

Sub EsportaFormatoDelimitato(NomeDelFile As String)
    On Error GoTo GestioneErroriEsportaFormatoDelimitato
    Dim Risposta As Integer
    Dim Riga As Single, Colonna As Single

    FormPrincipale.Refresh
    If Dir(NomeDelFile) <> "" Then
        Risposta = MsgBox("Sarà sovrascritto il file preesistente" + Chr(10) + NomeDelFile + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Esporta file ASCII formato delimitato")
            If Risposta = IDNO Then
                Exit Sub
            End If
    End If

    Screen.MousePointer = 11
    FormPrincipale.Refresh
    Open NomeDelFile For Output As #1
    
    FormGauge.Show
    FormGauge.Caption = "Esporta file ASCII formato delimitato"
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
 
 ' non salva i nomi delle variabili della riga 0 !
    For Riga = RigaDiPartenza To NUM_RIGHE - 1

        FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        FormPrincipale!Grid1.Row = Riga
            For Colonna = 1 To NUM_COLONNE - 2
                FormPrincipale!Grid1.Col = Colonna
                Print #1, Chr(34) & Left(FormPrincipale!Grid1.Text, 12) & Chr(34) & Chr(44);
            Next Colonna
                Colonna = NUM_COLONNE - 1
                FormPrincipale!Grid1.Col = Colonna
                Print #1, Chr(34) & Left(FormPrincipale!Grid1.Text, 12) & Chr(34)
        
    Next Riga

    Close #1

'    FileSalvatoInPrecedenza = True
    Unload FormGauge
    RipristinaSchermo
'    FormPrincipale.Caption = Logo & " [" & NomeDelFile & "]"
    Screen.MousePointer = 1

Exit Sub

GestioneErroriEsportaFormatoDelimitato:
    MsgBox "Errore su disc(hett)o:" & Chr(10) & "per favore riprovare.", MB_ICONSTOP, "Esporta file ASCII formato delimitato"
'    NomeDelFile = ""
'    FormPrincipale.Caption = Logo
    Unload FormGauge
    RipristinaSchermo
    Screen.MousePointer = 1
    Exit Sub

End Sub

Sub EsportaFormatoFisso(NomeDelFile As String)
    On Error GoTo GestioneErroriEsportaFormatoFisso
    Dim Risposta As Integer
    Dim Riga As Single, Colonna As Single

    FormPrincipale.Refresh
    If Dir(NomeDelFile) <> "" Then
        Risposta = MsgBox("Sarà sovrascritto il file preesistente" + Chr(10) + NomeDelFile + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Esporta file ASCII formato fisso")
            If Risposta = IDNO Then
                Exit Sub
            End If
    End If

    Screen.MousePointer = 11
    FormPrincipale.Refresh
    Open NomeDelFile For Output As #1
    
    FormGauge.Show
    FormGauge.Caption = "Esporta file ASCII formato fisso"
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
 
 ' non salva i nomi delle variabili della riga 0 !
    For Riga = RigaDiPartenza To NUM_RIGHE - 1

        FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        FormPrincipale!Grid1.Row = Riga
            For Colonna = 1 To NUM_COLONNE - 1
                FormPrincipale!Grid1.Col = Colonna
                Print #1, Left(FormPrincipale!Grid1.Text, 12),
            Next Colonna
                Print #1, ""
        
    Next Riga

    Close #1

'    FileSalvatoInPrecedenza = True
    Unload FormGauge
    RipristinaSchermo
'    FormPrincipale.Caption = Logo & " [" & NomeDelFile & "]"
    Screen.MousePointer = 1

Exit Sub

GestioneErroriEsportaFormatoFisso:
    MsgBox "Errore su disc(hett)o:" & Chr(10) & "per favore riprovare.", MB_ICONSTOP, "Esporta file ASCII formato fisso"
'    NomeDelFile = ""
'    FormPrincipale.Caption = Logo
    Unload FormGauge
    RipristinaSchermo
    Screen.MousePointer = 1
    Exit Sub

End Sub

Sub FormattaNumero(StringaDiFormattazione As String)
    Dim Riga, Colonna

    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge!Label6.Caption = "[1/1]"
    If StringaDiFormattazione = "###########0" Then
        FormGauge.Caption = "Arrotonda a nessun decimale"
    End If
    If StringaDiFormattazione = "#########0.0" Then
        FormGauge.Caption = "Arrotonda a un decimale"
    End If
    If StringaDiFormattazione = "########0.00" Then
        FormGauge.Caption = "Arrotonda a due decimali"
    End If
    If StringaDiFormattazione = "#######0.000" Then
        FormGauge.Caption = "Arrotonda a tre decimali"
    End If
    If StringaDiFormattazione = "######0.0000" Then
        FormGauge.Caption = "Arrotonda a quattro decimali"
    End If
    If StringaDiFormattazione = "#####0.00000" Then
        FormGauge.Caption = "Arrotonda a cinque decimali"
    End If
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
' arrotonda
    For Riga = 1 To NUM_RIGHE - 1

        FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' 12 caratteri inclusa la virgola !
            FormPrincipale!Grid1.Row = Riga
                If FormPrincipale!Grid1.Text <> "" Then
                    If Format(CVar(FormPrincipale!Grid1.Text), StringaDiFormattazione) <> 0 Then
                        FormPrincipale!Grid1.Text = Format(CVar(FormPrincipale!Grid1.Text), StringaDiFormattazione)
                    Else
                        FormPrincipale!Grid1.Text = Format(CVar(FormPrincipale!Grid1.Text), "0")
                    End If
                End If
    Next Riga
    
    Unload FormGauge
    RipristinaSchermo
    DeselezionaColonna
    Screen.MousePointer = 1

End Sub

Sub InizializzaFile()
    Dim Riga As Single, Colonna As Single
    
    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Vuota la tabella dati"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' intestazione colonne
    FormPrincipale!Grid1.Row = 0
        For Colonna = 1 To NUM_COLONNE - 1
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Variabile " & Colonna
        Next Colonna
    
    For Riga = 1 To NUM_RIGHE - 1

        FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        FormPrincipale!Grid1.Row = Riga
            For Colonna = 1 To NUM_COLONNE - 1
                FormPrincipale!Grid1.Col = Colonna
                FormPrincipale!Grid1.Text = ""
            Next Colonna
    Next Riga

    Unload FormGauge
    Screen.MousePointer = 1
End Sub

Sub InserisciColonna()
    Dim Riga, Colonna, Risposta, I, Dummy
    
    InseritaNuovaColonna = False
    Risposta = MsgBox("I dati della colonna 10 saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Inserisci colonna in posizione " & FormPrincipale!Grid1.SelStartCol)
        If Risposta = IDNO Then
            RipristinaSchermo
            DeselezionaColonna
            Exit Sub
        End If
    
    Screen.MousePointer = 11
'    FormPrincipale.Refresh
    Colonna = FormPrincipale!Grid1.SelStartCol
    FormGauge.Show
    FormGauge.Caption = "Inserisci colonna in posizione " & Colonna
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
' il caso generale delle prime 9 colonne
    If Colonna < NUM_COLONNE - 1 Then
    
    For Riga = 0 To NUM_RIGHE - 1

        FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    For I = NUM_COLONNE - 1 To Colonna + 1 Step -1
        FormPrincipale!Grid1.Row = Riga
        FormPrincipale!Grid1.Col = I - 1
        Dummy = FormPrincipale!Grid1.Text
        FormPrincipale!Grid1.Text = ""
        FormPrincipale!Grid1.Col = I
        FormPrincipale!Grid1.Text = Dummy
    Next I

    Next Riga
    End If

' il caso particolare della colonna 10
    If Colonna = NUM_COLONNE - 1 Then
    
    For Riga = 0 To NUM_RIGHE - 1

        FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        FormPrincipale!Grid1.Row = Riga
        FormPrincipale!Grid1.Text = ""

    Next Riga
    End If

    Unload FormGauge
    InseritaNuovaColonna = True
    FormPrincipale!Grid1.SelEndCol = FormPrincipale!Grid1.SelStartCol
'    FormPrincipale.Refresh
    Screen.MousePointer = 1

End Sub

Sub Main()
    
    SetDataAccessOption 1, App.Path & "\MINISTAT.INI"

    Screen.MousePointer = 11
    BarraDegliStrumenti = True
    FormCaricamentoProgramma.Show
    FormCaricamentoProgramma.Refresh
    Load FormPrincipale
    Unload FormCaricamentoProgramma
    FormPrincipale.Show
    Screen.MousePointer = 0

End Sub

Sub RipristinaSchermo()

    FormPrincipale!Grid1.Row = Rig
    FormPrincipale!Grid1.Col = Col
    FormPrincipale!Grid1.TopRow = RigSup
    FormPrincipale!Grid1.LeftCol = ColSin
    FormPrincipale!Grid1.SetFocus

End Sub

Sub SalvaFile(NomeDelFile As String)
    On Error GoTo GestioneErroriSalvaFile
    Dim Risposta As Integer
    Dim Riga As Single, Colonna As Single

    If Dir(NomeDelFile) <> "" Then
        Risposta = MsgBox("Sarà sovrascritto il file preesistente" + Chr(10) + NomeDelFile + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Salva file Ministat 1.1")
            If Risposta = IDNO Then
                RipristinaSchermo
                Exit Sub
            End If
    End If

    Screen.MousePointer = 11
    Open NomeDelFile For Output As #1
    
    FormGauge.Show
    FormGauge.Caption = "Salva file Ministat 1.1"
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
 
 ' salva anche i nomi delle variabili della riga 0 !
    For Riga = 0 To NUM_RIGHE - 1

        FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        FormPrincipale!Grid1.Row = Riga
            For Colonna = 1 To NUM_COLONNE - 1
                FormPrincipale!Grid1.Col = Colonna
                Print #1, Left(FormPrincipale!Grid1.Text, 12),
            Next Colonna
                Print #1, ""
        
    Next Riga

    Close #1

    FileSalvatoInPrecedenza = True
    Unload FormGauge
    RipristinaSchermo
    NomeDellaTabella = ""
    FormPrincipale.Caption = Logo & " [" & NomeDelFile & "]"
    NomeDelFileMinistat = NomeDelFile
    Screen.MousePointer = 1

Exit Sub

GestioneErroriSalvaFile:
    MsgBox "Errore su disc(hett)o:" & Chr(10) & "per favore riprovare.", MB_ICONSTOP, "Salva file Ministat 1.1"
'    NomeDelFile = ""
'    FormPrincipale.Caption = Logo
    Unload FormGauge
    RipristinaSchermo
    Screen.MousePointer = 1
    Exit Sub

End Sub

