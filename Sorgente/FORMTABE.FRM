VERSION 5.00
Begin VB.Form FormTabelleAccess 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1935
   ClientLeft      =   1800
   ClientTop       =   2385
   ClientWidth     =   6195
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
   Icon            =   "FORMTABE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1935
   ScaleWidth      =   6195
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton CommandAnnulla 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Annulla"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton CommandOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   $"FORMTABE.frx":030A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2640
      TabIndex        =   4
      Top             =   480
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Tabelle disponibili:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "FormTabelleAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandAnnulla_Click()
    
    NomeDellaTabella = ""
    Unload FormTabelleAccess
    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo

End Sub

Private Sub CommandOK_Click()
    
    If List1.ListIndex <> -1 Then
        NomeDellaTabella = List1.List(List1.ListIndex)
    End If
    
    Unload FormTabelleAccess
'    FormPrincipale!Grid1.HighLight = True
    RipristinaSchermo

End Sub

Private Sub Form_Activate()
    CommandAnnulla.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo GestioneErroriFormTabelleAccess
    Dim I

' centra il form
    FormTabelleAccess.Top = Screen.Height / 2 - Height / 2
    FormTabelleAccess.Left = Screen.Width / 2 - Width / 2
' titolo = nome del file Access aperto !
    FormTabelleAccess.Caption = NomeDelFile
' carica nomi tabelle nella finestra
    For I = 0 To FormPrincipale!Data1.Database.TableDefs.Count - 1
        If Left(FormPrincipale!Data1.Database.TableDefs(I).Name, 4) <> "MSys" Then
            List1.AddItem FormPrincipale!Data1.Database.TableDefs(I).Name
        End If
    Next I
    Screen.MousePointer = 1
Exit Sub

GestioneErroriFormTabelleAccess:
    Unload FormTabelleAccess
    MsgBox "Errore nella struttura" & Chr(10) & "o nei driver del file.", MB_ICONSTOP, "Importa file Access 1.x"
    Screen.MousePointer = 1
    RipristinaSchermo
Exit Sub

End Sub

Private Sub List1_DblClick()
    
    If List1.ListIndex <> -1 Then
        NomeDellaTabella = List1.List(List1.ListIndex)
    End If
    
    Unload FormTabelleAccess
    RipristinaSchermo

End Sub

