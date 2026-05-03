VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FormPrincipale 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   ClientHeight    =   6165
   ClientLeft      =   255
   ClientTop       =   1950
   ClientWidth     =   9105
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
   Icon            =   "FORMPRIN.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6165
   ScaleWidth      =   9105
   WindowState     =   2  'Maximized
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5400
      Visible         =   0   'False
      Width           =   3495
   End
   Begin Threed.SSCommand CommandEsploraNonPar 
      Height          =   615
      Left            =   0
      TabIndex        =   40
      ToolTipText     =   " Statistiche esplorative parametriche "
      Top             =   1200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":030A
   End
   Begin Threed.SSCommand CommandEsploraPar 
      Height          =   615
      Left            =   5400
      TabIndex        =   39
      ToolTipText     =   " Statistiche esplorative non parametriche "
      Top             =   1200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":0624
   End
   Begin Threed.SSCommand CommandGuida 
      Height          =   615
      Left            =   3600
      TabIndex        =   38
      ToolTipText     =   " Consente di consultare la Guida di Ministat 1.1 "
      Top             =   0
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":093E
   End
   Begin Threed.SSCommand CommandIstogrammiMultipli 
      Height          =   615
      Left            =   6600
      TabIndex        =   37
      ToolTipText     =   " Diagramma a istogrammi multipli "
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":0C58
   End
   Begin Threed.SSCommand CommandErroreStandard 
      Height          =   615
      Left            =   7800
      TabIndex        =   36
      ToolTipText     =   " Diagramma dei limiti di confidenza "
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":0F72
   End
   Begin Threed.SSCommand CommandStampaColonne6 
      Height          =   615
      Left            =   2400
      TabIndex        =   35
      ToolTipText     =   " Stampa i dati delle ultime cinque variabili "
      Top             =   0
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":128C
   End
   Begin Threed.SSCommand CommandStampaColonne1 
      Height          =   615
      Left            =   1800
      TabIndex        =   34
      ToolTipText     =   " Stampa i dati delle prime cinque variabili "
      Top             =   0
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":15A6
   End
   Begin Threed.SSCommand CommandBarreVerticali 
      Height          =   615
      Left            =   6000
      TabIndex        =   33
      ToolTipText     =   " Diagramma a barre verticali "
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":18C0
   End
   Begin Threed.SSCommand CommandAltman 
      Height          =   615
      Left            =   3600
      TabIndex        =   32
      ToolTipText     =   " Diagramma di Altman "
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":1BDA
   End
   Begin Threed.SSCommand CommandOrdinaDecrescente 
      Height          =   615
      Left            =   6000
      TabIndex        =   31
      ToolTipText     =   " Ordina i dati in ordine decrescente "
      Top             =   0
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":1EF4
   End
   Begin Threed.SSCommand CommandOrdinaCrescente 
      Height          =   615
      Left            =   5400
      TabIndex        =   30
      ToolTipText     =   "  Ordina i dati in ordine crescente  "
      Top             =   0
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":220E
   End
   Begin Threed.SSCommand CommandDeviataZeta 
      Height          =   615
      Left            =   7200
      TabIndex        =   29
      ToolTipText     =   "  Calcola per ciascun dato la deviata normale standardizzata z  "
      Top             =   0
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "z"
   End
   Begin Threed.SSCommand CommandNuovo 
      Height          =   612
      Left            =   0
      TabIndex        =   28
      ToolTipText     =   " Crea un nuovo file in formato Ministat 1.1 "
      Top             =   0
      Width           =   612
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":2528
   End
   Begin Threed.SSCommand CommandNumera 
      Height          =   615
      Left            =   4800
      TabIndex        =   27
      ToolTipText     =   " Numera i dati in ordine progressivo "
      Top             =   0
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FORMPRIN.frx":2842
   End
   Begin Threed.SSCommand CommandRegNon 
      Height          =   615
      Left            =   7200
      TabIndex        =   26
      ToolTipText     =   " Regressione lineare non parametrica "
      Top             =   1200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":2B5C
   End
   Begin Threed.SSCommand CommandConfrontoTraMediane 
      Height          =   615
      Left            =   6600
      TabIndex        =   25
      ToolTipText     =   " Confronto tra mediane "
      Top             =   1200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":2E76
   End
   Begin Threed.SSCommand CommandStatNon 
      Height          =   615
      Left            =   6000
      TabIndex        =   24
      ToolTipText     =   " Statistiche elementari non parametriche "
      Top             =   1200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":3190
   End
   Begin Threed.SSCommand CommandPercentili 
      Height          =   615
      Left            =   8400
      TabIndex        =   23
      ToolTipText     =   " Calcola per ciascun dato il percentile "
      Top             =   0
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "P"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand CommandRango 
      Height          =   615
      Left            =   7800
      TabIndex        =   22
      ToolTipText     =   " Calcola per ciascun dato il rango "
      Top             =   0
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "Rango"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand CommandTestChiQuadrato 
      Height          =   615
      Left            =   7800
      TabIndex        =   21
      ToolTipText     =   " Test chi-quadrato "
      Top             =   1200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":34AA
   End
   Begin Threed.SSCommand CommandAnalisiDellaSomiglianza 
      Height          =   615
      Left            =   8400
      TabIndex        =   20
      ToolTipText     =   " Analisi della somiglianza "
      Top             =   1200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":37C4
   End
   Begin Threed.SSCommand CommandStatisticaBayesiana 
      Height          =   615
      Left            =   4200
      TabIndex        =   19
      ToolTipText     =   " Statistica bayesiana "
      Top             =   1200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":3ADE
   End
   Begin Threed.SSCommand CommandRegressioneMultipla 
      Height          =   615
      Left            =   3600
      TabIndex        =   18
      ToolTipText     =   " Regressione lineare multipla "
      Top             =   1200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":3DF8
   End
   Begin Threed.SSCommand CommandRegressionePolinomiale 
      Height          =   615
      Left            =   3000
      TabIndex        =   17
      ToolTipText     =   " Regressione polinomiale "
      Top             =   1200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":4112
   End
   Begin Threed.SSCommand CommandRegressioneLineare 
      Height          =   615
      Left            =   2400
      TabIndex        =   16
      ToolTipText     =   " Regressione lineare parametrica "
      Top             =   1200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":442C
   End
   Begin Threed.SSCommand CommandAnalisiDEllaVarianza 
      Height          =   615
      Left            =   1800
      TabIndex        =   15
      ToolTipText     =   " Analisi della varianza "
      Top             =   1200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":4746
   End
   Begin Threed.SSCommand CommandConfrontoTraMedie 
      Height          =   615
      Left            =   1200
      TabIndex        =   14
      ToolTipText     =   " Confronto tra medie "
      Top             =   1200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":4A60
   End
   Begin Threed.SSCommand CommandQuartiliERange 
      Height          =   615
      Left            =   8400
      TabIndex        =   13
      ToolTipText     =   " Diagramma dei quartili e del range "
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":4D7A
   End
   Begin Threed.SSCommand CommandDistribuzione 
      Height          =   615
      Left            =   7200
      TabIndex        =   12
      ToolTipText     =   " Diagramma di distribuzione "
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":5094
   End
   Begin Threed.SSCommand CommandLineaSpezzata 
      Height          =   615
      Left            =   5400
      TabIndex        =   11
      ToolTipText     =   " Diagramma a linee spezzate "
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":53AE
   End
   Begin Threed.SSCommand CommandSovrapposizione 
      Height          =   615
      Left            =   4800
      TabIndex        =   10
      ToolTipText     =   " Diagramma di sovrapposizione "
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":56C8
   End
   Begin Threed.SSCommand CommandAssociazione 
      Height          =   615
      Left            =   4200
      TabIndex        =   9
      ToolTipText     =   " Diagramma di associazione "
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":59E2
   End
   Begin Threed.SSCommand CommandDispersione 
      Height          =   615
      Left            =   3000
      TabIndex        =   8
      ToolTipText     =   " Diagramma di dispersione "
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":5CFC
   End
   Begin Threed.SSCommand CommandAppaiamento 
      Height          =   615
      Left            =   2400
      TabIndex        =   7
      ToolTipText     =   " Diagramma di appaiamento "
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":6016
   End
   Begin Threed.SSCommand CommandConfronto 
      Height          =   615
      Left            =   1800
      TabIndex        =   6
      ToolTipText     =   " Diagramma di confronto "
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":6330
   End
   Begin Threed.SSCommand CommandCumulativa 
      Height          =   615
      Left            =   1200
      TabIndex        =   5
      ToolTipText     =   " Distribuzione cumulativa "
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":664A
   End
   Begin Threed.SSCommand CommandIstogramma 
      Height          =   615
      Left            =   600
      TabIndex        =   4
      ToolTipText     =   " Istogramma della distribuzione "
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":6964
   End
   Begin Threed.SSCommand CommandSalva 
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   " Salva il file in formato Ministat 1.1 "
      Top             =   0
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":6C7E
   End
   Begin Threed.SSCommand CommandStatisticheElementari 
      Height          =   615
      Left            =   600
      TabIndex        =   2
      ToolTipText     =   " Statistiche elementari parametriche "
      Top             =   1200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":6F98
   End
   Begin Threed.SSCommand CommandApri 
      Height          =   615
      Left            =   600
      TabIndex        =   1
      ToolTipText     =   " Apre un file in formato Ministat 1.1 "
      Top             =   0
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":72B2
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   600
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      CancelError     =   -1  'True
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin MSGrid.Grid Grid1 
      Height          =   1095
      Left            =   5280
      TabIndex        =   0
      ToolTipText     =   " Aprire un file Ministat 1.1, importare un file Acces, ASCII, dBase, Excel, o immettere i dati da tastiera "
      Top             =   4800
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   1931
      _StockProps     =   77
      BackColor       =   -2147483643
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
   Begin Threed.SSCommand CommandTorta 
      Height          =   615
      Left            =   0
      TabIndex        =   41
      ToolTipText     =   " Diagramma a torta "
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      Picture         =   "FORMPRIN.frx":7704
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&File"
      Begin VB.Menu FileNuovo 
         Caption         =   "&Nuovo"
      End
      Begin VB.Menu FileApri 
         Caption         =   "&Apri..."
      End
      Begin VB.Menu FileSeparatoreUno 
         Caption         =   "-"
      End
      Begin VB.Menu FileImportaFile 
         Caption         =   "&Importa file"
         Begin VB.Menu FileImportaFileAccess 
            Caption         =   "&Access"
         End
         Begin VB.Menu FileImportaFileASCII 
            Caption         =   "A&SCII"
         End
         Begin VB.Menu FileImportaFileDbase 
            Caption         =   "&dBase III"
         End
         Begin VB.Menu FileImportaFileExcel80 
            Caption         =   "&Excel"
         End
      End
      Begin VB.Menu FileEsporta 
         Caption         =   "Esp&orta file"
         Begin VB.Menu FileEsportaFisso 
            Caption         =   "ASCII formato &fisso"
            Begin VB.Menu EsportaFissoCon 
               Caption         =   "&Con il nome delle variabili"
            End
            Begin VB.Menu EsportaFissoSenza 
               Caption         =   "&Senza il nome delle variabili"
            End
         End
         Begin VB.Menu FileEsportaDelimitato 
            Caption         =   "ASCII formato &delimitato"
            Begin VB.Menu EsportaDelimitatoCon 
               Caption         =   "&Con il nome delle variabili"
            End
            Begin VB.Menu EsportaDelimitatoSenza 
               Caption         =   "&Senza il nome delle variabili"
            End
         End
      End
      Begin VB.Menu FileSeparatoreDue 
         Caption         =   "-"
      End
      Begin VB.Menu FileSalva 
         Caption         =   "&Salva"
      End
      Begin VB.Menu FileSalvaConNome 
         Caption         =   "Sal&va con nome..."
      End
      Begin VB.Menu FileCancella 
         Caption         =   "&Cancella"
      End
      Begin VB.Menu FileSeparatoreTre 
         Caption         =   "-"
      End
      Begin VB.Menu FileStampa 
         Caption         =   "S&tampa"
         Begin VB.Menu FileStampaColonne1 
            Caption         =   "Colonne &1-5"
         End
         Begin VB.Menu FileStampaColonne6 
            Caption         =   "Colonne &6-10"
         End
      End
      Begin VB.Menu FileSeparatoreQuattro 
         Caption         =   "-"
      End
      Begin VB.Menu FileEsci 
         Caption         =   "&Esci"
      End
   End
   Begin VB.Menu MenuModifica 
      Caption         =   "&Modifica"
      Begin VB.Menu ModificaTagliaColonna 
         Caption         =   "&Taglia colonna"
      End
      Begin VB.Menu ModificaCopiaColonna 
         Caption         =   "&Copia colonna"
      End
      Begin VB.Menu ModificaIncollaColonna 
         Caption         =   "&Incolla colonna"
      End
      Begin VB.Menu ModificaAnnullaIncolla 
         Caption         =   "&Annulla incolla"
      End
      Begin VB.Menu ModificaSeparatoreUno 
         Caption         =   "-"
      End
      Begin VB.Menu ModificaCella 
         Caption         =   "C&ella"
         Begin VB.Menu ModificaVuotaCella 
            Caption         =   "&Vuota"
         End
         Begin VB.Menu ModificaEliminaCella 
            Caption         =   "&Elimina"
         End
         Begin VB.Menu ModificaInserisciCella 
            Caption         =   "&Inserisci"
         End
      End
      Begin VB.Menu ModificaColonna 
         Caption         =   "C&olonna"
         Begin VB.Menu ModificaVuotaColonna 
            Caption         =   "&Vuota"
         End
         Begin VB.Menu ModificaEliminaColonna 
            Caption         =   "&Elimina"
         End
         Begin VB.Menu ModificaInserisciColonna 
            Caption         =   "&Inserisci"
         End
      End
      Begin VB.Menu ModificaRiga 
         Caption         =   "&Riga"
         Begin VB.Menu ModificaVuotaRiga 
            Caption         =   "&Vuota"
         End
         Begin VB.Menu ModificaEliminaRiga 
            Caption         =   "&Elimina"
         End
         Begin VB.Menu ModificaInserisciRiga 
            Caption         =   "&Inserisci"
         End
      End
      Begin VB.Menu ModificaSeparatoreDue 
         Caption         =   "-"
      End
      Begin VB.Menu ModificaUnisciColonne 
         Caption         =   "&Unisci dati"
      End
      Begin VB.Menu ModificaScambiaColonne 
         Caption         =   "&Scambia colonne"
      End
      Begin VB.Menu ModificaSeparatoreTre 
         Caption         =   "-"
      End
      Begin VB.Menu ModificaOrdina 
         Caption         =   "Or&dina dati"
         Begin VB.Menu ModificaOrdinaCrescente 
            Caption         =   "In ordine &crescente"
         End
         Begin VB.Menu ModificaOrdinaDecrescente 
            Caption         =   "In ordine &decrescente"
         End
      End
   End
   Begin VB.Menu MenuOpzioni 
      Caption         =   "&Opzioni"
      Begin VB.Menu OpzioniNomeDellaVariabile 
         Caption         =   "&Nome della variabile..."
      End
      Begin VB.Menu OpzioniNumeroDiDecimali 
         Caption         =   "N&umero di decimali"
         Begin VB.Menu OpzioniNessunDecimale 
            Caption         =   "&Nessuno"
         End
         Begin VB.Menu OpzioniUnDecimale 
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
      Begin VB.Menu OpzioniBarraDegliStrumenti 
         Caption         =   "&Barra degli strumenti"
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
   Begin VB.Menu MenuCalcoli 
      Caption         =   "&Calcoli"
      Begin VB.Menu CalcoliSomma 
         Caption         =   "&Somma..."
      End
      Begin VB.Menu CalcoliSottrai 
         Caption         =   "S&ottrai..."
      End
      Begin VB.Menu CalcoliMoltiplica 
         Caption         =   "&Moltiplica..."
      End
      Begin VB.Menu CalcoliDividi 
         Caption         =   "&Dividi..."
      End
      Begin VB.Menu CalcoliSeparatoreUno 
         Caption         =   "-"
      End
      Begin VB.Menu CalcoliQuadrato 
         Caption         =   "&Quadrato"
      End
      Begin VB.Menu CalcoliRadice 
         Caption         =   "&Radice"
      End
      Begin VB.Menu CalcoliLogaritmo 
         Caption         =   "&Logaritmo"
         Begin VB.Menu CalcoliLogaritmoNaturale 
            Caption         =   "&Naturale"
         End
         Begin VB.Menu CalcoliLogaritmoDecimale 
            Caption         =   "&Decimale"
         End
      End
      Begin VB.Menu CalcoliSeparatoreDue 
         Caption         =   "-"
      End
      Begin VB.Menu CalcoliDifferenza 
         Caption         =   "D&ifferenza"
         Begin VB.Menu CalcoliDifferenzaAssoluta 
            Caption         =   "&Assoluta"
         End
         Begin VB.Menu CalcoliDifferenzaPercentuale 
            Caption         =   "&Percentuale"
         End
      End
      Begin VB.Menu CalcoliMedia 
         Caption         =   "M&edia"
         Begin VB.Menu CalcoliMediaAritmetica 
            Caption         =   "&Aritmetica"
         End
         Begin VB.Menu CalcoliMediaGeometrica 
            Caption         =   "&Geometrica"
         End
      End
      Begin VB.Menu CalcoliRapporto 
         Caption         =   "R&apporto"
      End
      Begin VB.Menu CalcoliSeparatoreTre 
         Caption         =   "-"
      End
      Begin VB.Menu CalcoliCambiaSegno 
         Caption         =   "&Cambia il segno"
      End
      Begin VB.Menu CalcoliValoreAssoluto 
         Caption         =   "&Valore assoluto"
      End
      Begin VB.Menu CalcoliSeparatoreQuattro 
         Caption         =   "-"
      End
      Begin VB.Menu CalcoliNumera 
         Caption         =   "N&umera"
      End
      Begin VB.Menu CalcoliDeviataZeta 
         Caption         =   "Deviata &z"
      End
      Begin VB.Menu CalcoliRango 
         Caption         =   "Ra&ngo"
      End
      Begin VB.Menu CalcoliPercentili 
         Caption         =   "&Percentili"
      End
   End
   Begin VB.Menu MenuGrafica 
      Caption         =   "&Grafica"
      Begin VB.Menu GraficaTorta 
         Caption         =   "&Torta"
      End
      Begin VB.Menu GraficaSeparatoreZero 
         Caption         =   "-"
      End
      Begin VB.Menu GraficaIstogramma 
         Caption         =   "&Istogramma"
      End
      Begin VB.Menu GraficaCumulativa 
         Caption         =   "&Cumulativa"
      End
      Begin VB.Menu GraficaSeparatoreUno 
         Caption         =   "-"
      End
      Begin VB.Menu GraficaConfronto 
         Caption         =   "C&onfronto"
      End
      Begin VB.Menu GraficaPrimaDopo 
         Caption         =   "&Appaiamento"
      End
      Begin VB.Menu GraficaDispersione 
         Caption         =   "&Dispersione"
      End
      Begin VB.Menu GraficaAltman 
         Caption         =   "Alt&man"
      End
      Begin VB.Menu GraficaSeparatoreDue 
         Caption         =   "-"
      End
      Begin VB.Menu GraficaAssociazione 
         Caption         =   "A&ssociazione"
      End
      Begin VB.Menu GraficaSovrapposizione 
         Caption         =   "So&vrapposizione"
      End
      Begin VB.Menu GraficaLineaSpezzata 
         Caption         =   "&Linea spezzata"
      End
      Begin VB.Menu GraficaBarreVerticali 
         Caption         =   "&Barre verticali"
      End
      Begin VB.Menu GraficaSeparatoreTre 
         Caption         =   "-"
      End
      Begin VB.Menu GraficaIstogrammiMultipli 
         Caption         =   "Istogrammi &multipli"
      End
      Begin VB.Menu GraficaDistribuzione 
         Caption         =   "Dis&tribuzione"
      End
      Begin VB.Menu GraficaErroreStandard 
         Caption         =   "Limiti di con&fidenza"
      End
      Begin VB.Menu GraficaQuartili 
         Caption         =   "&Quartili e range"
      End
   End
   Begin VB.Menu MenuStatistica 
      Caption         =   "&Statistica"
      Begin VB.Menu StatisticaEsploraPar 
         Caption         =   "&Statistiche esplorative"
      End
      Begin VB.Menu StatisticaStatisticheElementari 
         Caption         =   "S&tatistiche elementari"
      End
      Begin VB.Menu StatisticaSeparatoreUno 
         Caption         =   "-"
      End
      Begin VB.Menu StatisticaT 
         Caption         =   "&Confronto tra medie"
      End
      Begin VB.Menu StatisticaAnova 
         Caption         =   "&Analisi della varianza"
      End
      Begin VB.Menu StatisticaSeparatoreDue 
         Caption         =   "-"
      End
      Begin VB.Menu StatisticaRegressioneLineare 
         Caption         =   "&Regressione lineare"
      End
      Begin VB.Menu StatisticaRegressionePolinomiale 
         Caption         =   "R&egressione polinomiale"
      End
      Begin VB.Menu StatisticaRegressioneMultipla 
         Caption         =   "Regressione &multipla"
      End
      Begin VB.Menu StatisticaSeparatoreTre 
         Caption         =   "-"
      End
      Begin VB.Menu StatisticaBayesiana 
         Caption         =   "Statistica &bayesiana"
      End
      Begin VB.Menu StatisticaSeparatoreCinque 
         Caption         =   "-"
      End
      Begin VB.Menu StatisticaNonPArametrica 
         Caption         =   "Statistica &non parametrica"
         Begin VB.Menu StatisticaEsploraNonPar 
            Caption         =   "&Statistiche esplorative"
         End
         Begin VB.Menu StatisticaNonParametricaCumulativa 
            Caption         =   "S&tatistiche elementari"
         End
         Begin VB.Menu StatisticaNonParaSeparatoreDue 
            Caption         =   "-"
         End
         Begin VB.Menu StatisticaNonParametricaW 
            Caption         =   "&Confronto tra mediane"
         End
         Begin VB.Menu StatisticaNonParaSeparatoreTre 
            Caption         =   "-"
         End
         Begin VB.Menu StatisticaNonParametricaReglin 
            Caption         =   "&Regressione lineare"
         End
         Begin VB.Menu StatisticaNonParaSeparatoreUno 
            Caption         =   "-"
         End
         Begin VB.Menu StatisticaNonParametricaChi 
            Caption         =   "T&est chi-quadrato"
         End
         Begin VB.Menu StatisticaSeparatoreQuattro 
            Caption         =   "-"
         End
         Begin VB.Menu StatisticaCluster 
            Caption         =   "&Analisi della somiglianza"
         End
      End
   End
   Begin VB.Menu Aiuto 
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
      Begin VB.Menu AiutoSeparatoreDue 
         Caption         =   "-"
      End
      Begin VB.Menu AiutoInformazioniSu 
         Caption         =   "&Informazioni su..."
      End
   End
End
Attribute VB_Name = "FormPrincipale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim ImmessaVirgolaPrecedentemente
    Dim PosizioneDellaVirgola
    Dim QualcosaNelRitaglio
    Dim Ritaglio()
    Dim ColonnaSelezionataInPrecedenza

Private Sub Aiuto_Click()

' se non è stato selezionato nulla
    If (Grid1.SelStartCol = Grid1.SelEndCol) And (Grid1.SelStartRow = Grid1.SelEndRow) Then
    FormPrincipale!Grid1.HighLight = False
    End If

End Sub

Private Sub AiutoCalcolatrice_Click()
    
    Dim Calcolatrice
    On Error GoTo ErroreCalcolatriceCinque:

    If FormPrincipale!Grid1.SelEndCol = FormPrincipale!Grid1.SelStartCol And FormPrincipale!Grid1.SelEndRow = FormPrincipale!Grid1.SelStartRow Then
        FormPrincipale!Grid1.HighLight = False
    End If
    
    Calcolatrice = Shell("Calc.exe", 1)

'    FormPrincipale!Grid1.SetFocus
        
Exit Sub

ErroreCalcolatriceCinque:
    MsgBox "Calcolatrice non disponibile.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcolatrice"
Exit Sub

End Sub

Private Sub AiutoGuida_Click()
    
'   Dim Dummy
'   Dummy = WinHelp(FormPrincipale.hWnd, App.Path & "\ministat.hlp", HELP_INDEX, "")
        
    FormPrincipale!CMDialog1.HelpFile = App.Path & "\ministat.hlp"
    FormPrincipale!CMDialog1.HelpCommand = cdlHelpContents
    FormPrincipale!CMDialog1.ShowHelp
    
End Sub

Private Sub AiutoInformazioniSu_Click()
    FormInfo.Show 1
End Sub

Private Sub CalcoliCambiaSegno_Click()
    Dim Riga, Colonna, Dummy
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    If Colonna = NUM_COLONNE - 1 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    
    If Colonna < NUM_COLONNE - 1 Then
        FormPrincipale!Grid1.Row = 0
        NomeDellaVariabile = FormPrincipale!Grid1.Text
        MsgBox "Verrà cambiato il segno" + Chr(10) + "(+/-) della variabile.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "Cambia il segno di " & NomeDellaVariabile
        
        FormGauge!Label6.Caption = "[1/2]"
        InserisciColonna
            
        If InseritaNuovaColonna = True Then
            
            FormPrincipale.Refresh
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            FormGauge.Show
            FormGauge.Caption = "Cambia il segno di " & NomeDellaVariabile
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
            For Riga = 1 To NUM_RIGHE - 1

                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            FormPrincipale!Grid1.Row = Riga
            FormPrincipale!Grid1.Col = Colonna + 1
            Dummy = FormPrincipale!Grid1.Text
                If FormPrincipale!Grid1.Text <> "" Then
                    FormPrincipale!Grid1.Col = Colonna
                    FormPrincipale!Grid1.Text = -Dummy
                End If
            Next Riga
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Cambia segno"
        End If
    End If
        
        Unload FormGauge
        RipristinaSchermo
        DeselezionaColonna
        FormPrincipale!Grid1.SetFocus
        Screen.MousePointer = 1

End Sub

Private Sub CalcoliDeviataZeta_Click()
    Dim Riga, Colonna, X, M1, M2, Z1, Media, DevSt
    ReDim A(1 To NUM_RIGHE - 1) As Double
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    If Colonna = NUM_COLONNE - 1 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    
    If Colonna < NUM_COLONNE - 1 Then
        FormPrincipale!Grid1.Row = 0
        NomeDellaVariabile = FormPrincipale!Grid1.Text
        MsgBox "Verrà calcolato la deviata z" + Chr(10) + "(deviata normale standardizzata)" + Chr(10) + "corrispondente a ciascun dato.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "Deviata z di " & NomeDellaVariabile
        
        FormGauge!Label6.Caption = "[1/2]"
        InserisciColonna
            
        If InseritaNuovaColonna = True Then
            
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            FormGauge.Show
            FormGauge.Caption = "Deviata z di " & NomeDellaVariabile
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge.Label1.Caption = "0%"
            FormGauge!Gauge1.Value = 0
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
            Z1 = 0: M1 = 0: M2 = 0
            FormPrincipale!Grid1.Col = Colonna + 1
            For Riga = 1 To NUM_RIGHE - 1
                FormPrincipale!Grid1.Row = Riga
                If FormPrincipale!Grid1.Text <> "" Then
                    Z1 = Z1 + 1
                    X = CDbl(FormPrincipale!Grid1.Text)
                    A(Riga) = X
                    M1 = M1 + X
                    M2 = M2 + X ^ 2
                Else
                    Exit For
                End If
            Next Riga

            Media = M1 / Z1
            DevSt = Sqr((M2 - (M1 ^ 2 / Z1)) / (Z1 - 1))

            FormPrincipale!Grid1.Col = Colonna
            For Riga = 1 To Z1
                    
                FormGauge.Label1.Caption = Int(100 * Riga / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Text = (A(Riga) - Media) / DevSt
            Next Riga

            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Text = "Deviata z"
        End If
    End If
        
        ReDim A(0)
        Unload FormGauge
        RipristinaSchermo
        DeselezionaColonna
        FormPrincipale!Grid1.SetFocus
        Screen.MousePointer = 1

End Sub

Private Sub CalcoliDifferenzaAssoluta_Click()
    Dim Riga, Colonna, Dato1, Dato2, ColonnaDato1, ColonnaDato2
    Dim NomeDellaVariabile1, NomeDellaVariabile2
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    ColonnaDato1 = Colonna
    ColonnaDato2 = Colonna + 1
    
    If Colonna > (NUM_COLONNE - 1) - 2 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    If Colonna <= (NUM_COLONNE - 1) - 2 Then
    
        FormPrincipale!Grid1.Row = 0
        FormPrincipale!Grid1.Col = ColonnaDato1
        NomeDellaVariabile1 = FormPrincipale!Grid1.Text
        FormPrincipale!Grid1.Col = ColonnaDato2
        NomeDellaVariabile2 = FormPrincipale!Grid1.Text
        MsgBox "Verrà calcolata la differenza" + Chr(10) + "con il segno tra le due variabili.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "(" & NomeDellaVariabile1 & " - " & NomeDellaVariabile2 & ")"
            
        FormGauge!Label6.Caption = "[1/2]"
        InserisciColonna
            
        If InseritaNuovaColonna = True Then
        
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            ColonnaDato1 = ColonnaDato1 + 1
            ColonnaDato2 = ColonnaDato2 + 1
            FormGauge.Show
            FormGauge.Caption = NomeDellaVariabile1 & " - " & NomeDellaVariabile2
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            For Riga = 1 To NUM_RIGHE - 1

                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Col = ColonnaDato1
                Dato1 = FormPrincipale!Grid1.Text
                FormPrincipale!Grid1.Col = ColonnaDato2
                Dato2 = FormPrincipale!Grid1.Text
                    If Dato1 <> "" And Dato2 <> "" Then
                        FormPrincipale!Grid1.Col = Colonna
                        FormPrincipale!Grid1.Text = Dato1 - Dato2
                    End If
            Next Riga
        
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Differenza"
            FormPrincipale!Grid1.Row = 1
            Unload FormGauge
            DeselezionaColonna
            RipristinaSchermo
            InseritaNuovaColonna = False
            FormPrincipale!Grid1.SetFocus
            Screen.MousePointer = 1
        Else
            DeselezionaColonna
            RipristinaSchermo
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
    End If

End Sub

Private Sub CalcoliDifferenzaPercentuale_Click()
    Dim Riga, Colonna, Dato1, Dato2, ColonnaDato1, ColonnaDato2
    Dim NomeDellaVariabile1, NomeDellaVariabile2
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    ColonnaDato1 = Colonna
    ColonnaDato2 = Colonna + 1
    
    If Colonna > (NUM_COLONNE - 1) - 2 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    If Colonna <= (NUM_COLONNE - 1) - 2 Then
    
        FormPrincipale!Grid1.Row = 0
        FormPrincipale!Grid1.Col = ColonnaDato1
        NomeDellaVariabile1 = FormPrincipale!Grid1.Text
        FormPrincipale!Grid1.Col = ColonnaDato2
        NomeDellaVariabile2 = FormPrincipale!Grid1.Text
        MsgBox "Verrà calcolata la differenza tra le due variabili, con il" + Chr(10) + "segno, espressa come percentuale della loro media.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "100 * (" & NomeDellaVariabile1 & " - " & NomeDellaVariabile2 & ") / ((" & NomeDellaVariabile1 & " + " & NomeDellaVariabile2 & ") / 2)"
        
        FormGauge!Label6.Caption = "[1/2]"
        InserisciColonna
            
        If InseritaNuovaColonna = True Then
        
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            ColonnaDato1 = ColonnaDato1 + 1
            ColonnaDato2 = ColonnaDato2 + 1
            FormGauge.Show
            FormGauge.Caption = "100 * (" & NomeDellaVariabile1 & " - " & NomeDellaVariabile2 & ") / ((" & NomeDellaVariabile1 & " + " & NomeDellaVariabile2 & ") / 2)"
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            For Riga = 1 To NUM_RIGHE - 1

                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Col = ColonnaDato1
                Dato1 = FormPrincipale!Grid1.Text
                FormPrincipale!Grid1.Col = ColonnaDato2
                Dato2 = FormPrincipale!Grid1.Text
                    If Dato1 <> "" And Dato2 <> "" Then
                        FormPrincipale!Grid1.Col = Colonna
                        FormPrincipale!Grid1.Text = 100 * (Dato1 - Dato2) / ((Dato1 - (-Dato2)) / 2)
                    End If
            Next Riga
        
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Differenza %"
            FormPrincipale!Grid1.Row = 1
            Unload FormGauge
            DeselezionaColonna
            RipristinaSchermo
            InseritaNuovaColonna = False
            FormPrincipale!Grid1.SetFocus
            Screen.MousePointer = 1
        Else
            DeselezionaColonna
            RipristinaSchermo
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
    End If

End Sub

Private Sub CalcoliDivide_Click()
    Dim Riga, Colonna, Dummy
    On Error GoTo GestioneErroriCalcoliDivide

    Colonna = FormPrincipale!Grid1.SelStartCol
    If Colonna < 10 Then
    
        FormPrincipale!Grid1.Row = 0
        NomeDellaVariabile = FormPrincipale!Grid1.Text
        FormCalcoli.Caption = "Divide " & NomeDellaVariabile & " per..."
        FormCalcoli.Show 1
    
        If NumeroPerCalcoli <> 0 Then
            InserisciColonna
            
        If InseritaNuovaColonna = True Then
        
            FormPrincipale.Refresh
            FormGauge.Show

            For Riga = 1 To NUM_RIGHE - 1
                    If 100 * Riga / (NUM_RIGHE - 1) = Int(100 * Riga / (NUM_RIGHE - 1)) Then
                        FormGauge.Label1.Caption = "Calcoli completati al " & 100 * Riga / (NUM_RIGHE - 1) & "%"
                        FormGauge!Gauge1.Value = 100 * Riga / (NUM_RIGHE - 1)
                        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                    End If
                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Col = Colonna + 1
                Dummy = FormPrincipale!Grid1.Text
                    If FormPrincipale!Grid1.Text <> "" Then
                        FormPrincipale!Grid1.Col = Colonna
                        FormPrincipale!Grid1.Text = Dummy / NumeroPerCalcoli
                    End If
            Next Riga
        
            DeselezionaColonna
            RipristinaSchermo
            InseritaNuovaColonna = False
        Else
            Exit Sub
        End If
        End If
    End If
    Exit Sub

GestioneErroriCalcoliDivide:
    Exit Sub

End Sub

Private Sub CalcoliDividi_Click()
    Dim Riga, Colonna, Dummy
    On Error GoTo GestioneErroriCalcoliDividi

    Colonna = FormPrincipale!Grid1.SelStartCol
    If Colonna = NUM_COLONNE - 1 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    If Colonna < NUM_COLONNE - 1 Then
    
        FormPrincipale!Grid1.Row = 0
        NomeDellaVariabile = FormPrincipale!Grid1.Text
        FormCalcoli.Caption = "Dividi " & NomeDellaVariabile & " per numero"
        FormCalcoli.Show 1
    
        If NumeroPerCalcoli <> 0 Then
            FormGauge!Label6.Caption = "[1/2]"
            InserisciColonna
            
        If InseritaNuovaColonna = True Then
        
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            FormGauge.Show
            FormGauge.Caption = "Dividi " & NomeDellaVariabile & " per " & Format(NumeroPerCalcoli, "#####0.0####")
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            For Riga = 1 To NUM_RIGHE - 1
                    
                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Col = Colonna + 1
                Dummy = FormPrincipale!Grid1.Text
                    If FormPrincipale!Grid1.Text <> "" Then
                        FormPrincipale!Grid1.Col = Colonna
                        FormPrincipale!Grid1.Text = Dummy / NumeroPerCalcoli
                    End If
            Next Riga
        
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Dividi"
            Unload FormGauge
            RipristinaSchermo
            DeselezionaColonna
            InseritaNuovaColonna = False
            FormPrincipale!Grid1.SetFocus
            Screen.MousePointer = 1
        Else
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
        End If
    End If
Exit Sub

GestioneErroriCalcoliDividi:
            Unload FormGauge
            RipristinaSchermo
            DeselezionaColonna
            InseritaNuovaColonna = False
            FormPrincipale!Grid1.SetFocus
            Screen.MousePointer = 1
            Exit Sub

End Sub

Private Sub CalcoliLogaritmoDecimale_Click()
    Dim Riga, Colonna, Dummy
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    If Colonna = NUM_COLONNE - 1 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    
    If Colonna < NUM_COLONNE - 1 Then
        FormPrincipale!Grid1.Row = 0
        NomeDellaVariabile = FormPrincipale!Grid1.Text
        MsgBox "Verrà calcolato il logaritmo" + Chr(10) + "in base (10) della variabile.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "Logaritmo decimale di " & NomeDellaVariabile
        
        FormGauge!Label6.Caption = "[1/2]"
        InserisciColonna
            
        If InseritaNuovaColonna = True Then
            
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            FormGauge.Show
            FormGauge.Caption = "Logaritmo decimale di " & NomeDellaVariabile
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
            For Riga = 1 To NUM_RIGHE - 1

                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            FormPrincipale!Grid1.Row = Riga
            FormPrincipale!Grid1.Col = Colonna + 1
            Dummy = FormPrincipale!Grid1.Text
                If FormPrincipale!Grid1.Text <> "" Then
                    FormPrincipale!Grid1.Col = Colonna
                    FormPrincipale!Grid1.Text = Log(Dummy) / Log(10)
                End If
            Next Riga
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Log(10)"
        End If
    End If
        
        Unload FormGauge
        RipristinaSchermo
        DeselezionaColonna
        FormPrincipale!Grid1.SetFocus
        Screen.MousePointer = 1

End Sub

Private Sub CalcoliLogaritmoNaturale_Click()
    Dim Riga, Colonna, Dummy
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    If Colonna = NUM_COLONNE - 1 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    
    If Colonna < NUM_COLONNE - 1 Then
        FormPrincipale!Grid1.Row = 0
        NomeDellaVariabile = FormPrincipale!Grid1.Text
        MsgBox "Verrà calcolato il logaritmo" + Chr(10) + "in base (e) della variabile.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "Logaritmo naturale di " & NomeDellaVariabile
        
        FormGauge!Label6.Caption = "[1/2]"
        InserisciColonna
            
        If InseritaNuovaColonna = True Then
            
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            FormGauge.Show
            FormGauge.Caption = "Logaritmo naturale di " & NomeDellaVariabile
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
            For Riga = 1 To NUM_RIGHE - 1
                
                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            FormPrincipale!Grid1.Row = Riga
            FormPrincipale!Grid1.Col = Colonna + 1
            Dummy = FormPrincipale!Grid1.Text
                If FormPrincipale!Grid1.Text <> "" Then
                    FormPrincipale!Grid1.Col = Colonna
                    FormPrincipale!Grid1.Text = Log(Dummy)
                End If
            Next Riga
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Log(e)"
        End If
    End If
        
        Unload FormGauge
        RipristinaSchermo
        DeselezionaColonna
        FormPrincipale!Grid1.SetFocus
        Screen.MousePointer = 1

End Sub

Private Sub CalcoliMediaAritmetica_Click()
    Dim Riga, Colonna, Dato1, Dato2, ColonnaDato1, ColonnaDato2
    Dim NomeDellaVariabile1, NomeDellaVariabile2
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    ColonnaDato1 = Colonna
    ColonnaDato2 = Colonna + 1
    
    If Colonna > (NUM_COLONNE - 1) - 2 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    If Colonna <= (NUM_COLONNE - 1) - 2 Then
    
        FormPrincipale!Grid1.Row = 0
        FormPrincipale!Grid1.Col = ColonnaDato1
        NomeDellaVariabile1 = FormPrincipale!Grid1.Text
        FormPrincipale!Grid1.Col = ColonnaDato2
        NomeDellaVariabile2 = FormPrincipale!Grid1.Text
        MsgBox "Verrà calcolata la media aritmetica" + Chr(10) + "tra le due variabili.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "(" & NomeDellaVariabile1 & " + " & NomeDellaVariabile2 & ") / 2"
            
        FormGauge!Label6.Caption = "[1/2]"
        InserisciColonna
            
        If InseritaNuovaColonna = True Then
        
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            ColonnaDato1 = ColonnaDato1 + 1
            ColonnaDato2 = ColonnaDato2 + 1
            FormGauge.Show
            FormGauge.Caption = "(" & NomeDellaVariabile1 & " + " & NomeDellaVariabile2 & ") / 2"
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            For Riga = 1 To NUM_RIGHE - 1

                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Col = ColonnaDato1
                Dato1 = FormPrincipale!Grid1.Text
                FormPrincipale!Grid1.Col = ColonnaDato2
                Dato2 = FormPrincipale!Grid1.Text
                    If Dato1 <> "" And Dato2 <> "" Then
                        FormPrincipale!Grid1.Col = Colonna
                        FormPrincipale!Grid1.Text = (Dato1 - -Dato2) / 2
                    End If
            Next Riga
        
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Media aritm."
            FormPrincipale!Grid1.Row = 1
            Unload FormGauge
            DeselezionaColonna
            RipristinaSchermo
            InseritaNuovaColonna = False
            FormPrincipale!Grid1.SetFocus
            Screen.MousePointer = 1
        Else
            DeselezionaColonna
            RipristinaSchermo
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
    End If

End Sub

Private Sub CalcoliMediaGeometrica_Click()
    Dim Riga, Colonna, Dato1, Dato2, ColonnaDato1, ColonnaDato2
    Dim NomeDellaVariabile1, NomeDellaVariabile2
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    ColonnaDato1 = Colonna
    ColonnaDato2 = Colonna + 1
    
    If Colonna > (NUM_COLONNE - 1) - 2 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    If Colonna <= (NUM_COLONNE - 1) - 2 Then
    
        FormPrincipale!Grid1.Row = 0
        FormPrincipale!Grid1.Col = ColonnaDato1
        NomeDellaVariabile1 = FormPrincipale!Grid1.Text
        FormPrincipale!Grid1.Col = ColonnaDato2
        NomeDellaVariabile2 = FormPrincipale!Grid1.Text
        MsgBox "Verrà calcolata la media geometrica" + Chr(10) + "tra le due variabili.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "Radice di (" & NomeDellaVariabile1 & " * " & NomeDellaVariabile2 & ")"
            
        FormGauge!Label6.Caption = "[1/2]"
        InserisciColonna
            
        If InseritaNuovaColonna = True Then
        
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            ColonnaDato1 = ColonnaDato1 + 1
            ColonnaDato2 = ColonnaDato2 + 1
            FormGauge.Show
            FormGauge.Caption = "Radice di (" & NomeDellaVariabile1 & " * " & NomeDellaVariabile2 & ")"
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            For Riga = 1 To NUM_RIGHE - 1

                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Col = ColonnaDato1
                Dato1 = FormPrincipale!Grid1.Text
                FormPrincipale!Grid1.Col = ColonnaDato2
                Dato2 = FormPrincipale!Grid1.Text
                    If Dato1 <> "" And Dato2 <> "" Then
                        FormPrincipale!Grid1.Col = Colonna
                        FormPrincipale!Grid1.Text = Sqr(Dato1 * Dato2)
                    End If
            Next Riga
        
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Media geome."
            FormPrincipale!Grid1.Row = 1
            Unload FormGauge
            DeselezionaColonna
            RipristinaSchermo
            InseritaNuovaColonna = False
            FormPrincipale!Grid1.SetFocus
            Screen.MousePointer = 1
        Else
            DeselezionaColonna
            RipristinaSchermo
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
    End If

End Sub

Private Sub CalcoliMoltiplica_Click()
    Dim Riga, Colonna, Dummy
    On Error GoTo GestioneErroriCalcoliMoltiplica

    Colonna = FormPrincipale!Grid1.SelStartCol
    If Colonna = NUM_COLONNE - 1 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    If Colonna < NUM_COLONNE - 1 Then
    
        FormPrincipale!Grid1.Row = 0
        NomeDellaVariabile = FormPrincipale!Grid1.Text
        FormCalcoli.Caption = "Moltiplica " & NomeDellaVariabile & " per numero"
        FormCalcoli.Show 1
    
        If NumeroPerCalcoli <> 0 Then
            FormGauge!Label6.Caption = "[1/2]"
            InserisciColonna
            
        If InseritaNuovaColonna = True Then
        
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            FormGauge.Show
            FormGauge.Caption = "Moltiplica " & NomeDellaVariabile & " per " & Format(NumeroPerCalcoli, "#####0.0####")
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            For Riga = 1 To NUM_RIGHE - 1
                    
                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Col = Colonna + 1
                Dummy = FormPrincipale!Grid1.Text
                    If FormPrincipale!Grid1.Text <> "" Then
                        FormPrincipale!Grid1.Col = Colonna
                        FormPrincipale!Grid1.Text = Dummy * NumeroPerCalcoli
                    End If
            Next Riga
        
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Moltiplica"
            Unload FormGauge
            RipristinaSchermo
            DeselezionaColonna
            InseritaNuovaColonna = False
            FormPrincipale!Grid1.SetFocus
            Screen.MousePointer = 1
        Else
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
        End If
    End If
Exit Sub

GestioneErroriCalcoliMoltiplica:
            Unload FormGauge
            RipristinaSchermo
            DeselezionaColonna
            InseritaNuovaColonna = False
            FormPrincipale!Grid1.SetFocus
            Screen.MousePointer = 1
            Exit Sub

End Sub

Private Sub CalcoliNumera_Click()
    Dim Riga, Colonna, J
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    If Colonna = NUM_COLONNE - 1 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    
    If Colonna < NUM_COLONNE - 1 Then
        FormPrincipale!Grid1.Row = 0
        NomeDellaVariabile = FormPrincipale!Grid1.Text
        MsgBox "Verrà riportato il numero di" + Chr(10) + "posizione di ciascuna riga.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "Numera le righe"
        
        FormGauge!Label6.Caption = "[1/2]"
        InserisciColonna
            
        If InseritaNuovaColonna = True Then
            
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            FormGauge.Show
            FormGauge.Caption = "Numera le righe"
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
            FormPrincipale!Grid1.Col = Colonna
            
            For Riga = 1 To NUM_RIGHE - 1
                
                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            FormPrincipale!Grid1.Row = Riga

' controlla su tutte le colonne !
            For J = 1 To NUM_COLONNE - 1
            FormPrincipale!Grid1.Col = J
                If FormPrincipale!Grid1.Text <> "" Then
                    FormPrincipale!Grid1.Col = Colonna
                    FormPrincipale!Grid1.Text = Riga
                    Exit For
                End If
            Next J
' fine controllo su tutte le colonne

            Next Riga
            
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Numero riga"
        End If
    End If
        
        Unload FormGauge
        RipristinaSchermo
        DeselezionaColonna
        FormPrincipale!Grid1.SetFocus
        Screen.MousePointer = 1

End Sub

Private Sub CalcoliPercentili_Click()
    Dim Riga, Colonna, I, J, M1, M2, Z1
    ReDim Rank(1 To NUM_RIGHE - 1) As Double
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    If Colonna = NUM_COLONNE - 1 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    
    If Colonna < NUM_COLONNE - 1 Then
        FormPrincipale!Grid1.Row = 0
        NomeDellaVariabile = FormPrincipale!Grid1.Text
        MsgBox "Verrà calcolato il percentile" + Chr(10) + "corrispondente a ciascun dato.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "Percentili di " & NomeDellaVariabile
        
        FormGauge!Label6.Caption = "[1/2]"
        InserisciColonna
            
        If InseritaNuovaColonna = True Then
            
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            FormGauge.Show
            FormGauge.Caption = "Percentili di " & NomeDellaVariabile
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge.Label1.Caption = "0%"
            FormGauge!Gauge1.Value = 0
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
            Z1 = 0
            FormPrincipale!Grid1.Col = Colonna + 1
            For Riga = 1 To NUM_RIGHE - 1
                FormPrincipale!Grid1.Row = Riga
                If FormPrincipale!Grid1.Text <> "" Then
                    Z1 = Z1 + 1
                    Rank(Riga) = CDbl(FormPrincipale!Grid1.Text)
                Else
                    Exit For
                End If
            Next Riga

            FormPrincipale!Grid1.Col = Colonna
            For Riga = 1 To Z1
                    
                FormGauge.Label1.Caption = Int(100 * Riga / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                M1 = 0: M2 = 0
pe840:
                    For J = 1 To Z1
pe850:
                       If Rank(J) > Rank(Riga) Then GoTo pe890
pe860:
                       If Rank(J) = Rank(Riga) Then GoTo pe880
pe870:
                       M1 = M1 + 1: GoTo pe890
pe880:
                       M2 = M2 + 1
pe890:
                    Next J
pe900:
                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Text = M1 + 0.5 * M2 + 0.5
            Next Riga

            For Riga = 1 To Z1
                FormPrincipale!Grid1.Row = Riga
                    FormPrincipale!Grid1.Text = Format(100 * FormPrincipale!Grid1.Text / (Z1 + 1), "##0.0####")
            Next Riga

            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Text = "Percentili"
        End If
    End If
        
        ReDim Rank(0)
        Unload FormGauge
        RipristinaSchermo
        DeselezionaColonna
        FormPrincipale!Grid1.SetFocus
        Screen.MousePointer = 1

End Sub

Private Sub CalcoliQuadrato_Click()
    Dim Riga, Colonna, Dummy
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    If Colonna = NUM_COLONNE - 1 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    
    If Colonna < NUM_COLONNE - 1 Then
        FormPrincipale!Grid1.Row = 0
        NomeDellaVariabile = FormPrincipale!Grid1.Text
        MsgBox "Verrà calcolato il" + Chr(10) + "quadrato della variabile.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "Quadrato di " & NomeDellaVariabile
        
        FormGauge!Label6.Caption = "[1/2]"
        InserisciColonna
            
        If InseritaNuovaColonna = True Then
            
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            FormGauge.Show
            FormGauge.Caption = "Quadrato di " & NomeDellaVariabile
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
            For Riga = 1 To NUM_RIGHE - 1
                
                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
            FormPrincipale!Grid1.Row = Riga
            FormPrincipale!Grid1.Col = Colonna + 1
            Dummy = FormPrincipale!Grid1.Text
                If FormPrincipale!Grid1.Text <> "" Then
                    FormPrincipale!Grid1.Col = Colonna
                    FormPrincipale!Grid1.Text = Dummy * Dummy
                End If
            Next Riga
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Quadrato"
        End If
    End If
       
        Unload FormGauge
        RipristinaSchermo
        DeselezionaColonna
        FormPrincipale!Grid1.SetFocus
        Screen.MousePointer = 1

End Sub

Private Sub CalcoliRadice_Click()
    Dim Riga, Colonna, Dummy
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    If Colonna = NUM_COLONNE - 1 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    
    If Colonna < NUM_COLONNE - 1 Then
        FormPrincipale!Grid1.Row = 0
        NomeDellaVariabile = FormPrincipale!Grid1.Text
        MsgBox "Verrà calcolata la radice" + Chr(10) + "quadrata della variabile.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "Radice di " & NomeDellaVariabile
        
        FormGauge!Label6.Caption = "[1/2]"
        InserisciColonna
            
        If InseritaNuovaColonna = True Then
            
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            FormGauge.Show
            FormGauge.Caption = "Radice di " & NomeDellaVariabile
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
            For Riga = 1 To NUM_RIGHE - 1
                
                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            FormPrincipale!Grid1.Row = Riga
            FormPrincipale!Grid1.Col = Colonna + 1
            Dummy = FormPrincipale!Grid1.Text
                If FormPrincipale!Grid1.Text <> "" Then
                    FormPrincipale!Grid1.Col = Colonna
                    FormPrincipale!Grid1.Text = Sqr(Dummy)
                End If
            Next Riga
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Radice"
        End If
    End If
        
        Unload FormGauge
        RipristinaSchermo
        DeselezionaColonna
        FormPrincipale!Grid1.SetFocus
        Screen.MousePointer = 1

End Sub

Private Sub CalcoliRango_Click()
    Dim Riga, Colonna, I, J, M1, M2, Z1
    ReDim Rank(1 To NUM_RIGHE - 1) As Double
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    If Colonna = NUM_COLONNE - 1 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    
    If Colonna < NUM_COLONNE - 1 Then
        FormPrincipale!Grid1.Row = 0
        NomeDellaVariabile = FormPrincipale!Grid1.Text
        MsgBox "Verrà calcolato il rango" + Chr(10) + "corrispondente a ciascun dato.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "Rango di " & NomeDellaVariabile
        
        FormGauge!Label6.Caption = "[1/2]"
        InserisciColonna
            
        If InseritaNuovaColonna = True Then
            
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            FormGauge.Show
            FormGauge.Caption = "Rango di " & NomeDellaVariabile
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge.Label1.Caption = "0%"
            FormGauge!Gauge1.Value = 0
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
            Z1 = 0
            FormPrincipale!Grid1.Col = Colonna + 1
            For Riga = 1 To NUM_RIGHE - 1
                FormPrincipale!Grid1.Row = Riga
                If FormPrincipale!Grid1.Text <> "" Then
                    Z1 = Z1 + 1
                    Rank(Riga) = CDbl(FormPrincipale!Grid1.Text)
                Else
                    Exit For
                End If
            Next Riga

            FormPrincipale!Grid1.Col = Colonna
            For Riga = 1 To Z1

                FormGauge.Label1.Caption = Int(100 * Riga / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

                M1 = 0: M2 = 0
p840:
                    For J = 1 To Z1
p850:
                       If Rank(J) > Rank(Riga) Then GoTo p890
p860:
                       If Rank(J) = Rank(Riga) Then GoTo p880
p870:
                       M1 = M1 + 1: GoTo p890
p880:
                       M2 = M2 + 1
p890:
                    Next J
p900:
                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Text = M1 + 0.5 * M2 + 0.5
            Next Riga
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Text = "Rango"
        End If
    End If
        
        ReDim Rank(0)
        Unload FormGauge
        RipristinaSchermo
        DeselezionaColonna
        FormPrincipale!Grid1.SetFocus
        Screen.MousePointer = 1

End Sub

Private Sub CalcoliRapporto_Click()
    Dim Riga, Colonna, Dato1, Dato2, ColonnaDato1, ColonnaDato2
    Dim NomeDellaVariabile1, NomeDellaVariabile2
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    ColonnaDato1 = Colonna
    ColonnaDato2 = Colonna + 1
    
    If Colonna > (NUM_COLONNE - 1) - 2 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    If Colonna <= (NUM_COLONNE - 1) - 2 Then
    
        FormPrincipale!Grid1.Row = 0
        FormPrincipale!Grid1.Col = ColonnaDato1
        NomeDellaVariabile1 = FormPrincipale!Grid1.Text
        FormPrincipale!Grid1.Col = ColonnaDato2
        NomeDellaVariabile2 = FormPrincipale!Grid1.Text
        MsgBox "Verrà calcolato il rapporto" + Chr(10) + "tra le due variabili.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "(" & NomeDellaVariabile1 & " / " & NomeDellaVariabile2 & ")"
            
        FormGauge!Label6.Caption = "[1/2]"
        InserisciColonna
            
        If InseritaNuovaColonna = True Then
        
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            ColonnaDato1 = ColonnaDato1 + 1
            ColonnaDato2 = ColonnaDato2 + 1
            FormGauge.Show
            FormGauge.Caption = "(" & NomeDellaVariabile1 & " / " & NomeDellaVariabile2 & ")"
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            For Riga = 1 To NUM_RIGHE - 1

                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Col = ColonnaDato1
                Dato1 = FormPrincipale!Grid1.Text
                FormPrincipale!Grid1.Col = ColonnaDato2
                Dato2 = FormPrincipale!Grid1.Text
                    If Dato1 <> "" And Dato2 <> "" Then
                        FormPrincipale!Grid1.Col = Colonna
                        FormPrincipale!Grid1.Text = Dato1 / Dato2
                    End If
            Next Riga
        
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Rapporto"
            FormPrincipale!Grid1.Row = 1
            Unload FormGauge
            DeselezionaColonna
            RipristinaSchermo
            InseritaNuovaColonna = False
            FormPrincipale!Grid1.SetFocus
            Screen.MousePointer = 1
        Else
            DeselezionaColonna
            RipristinaSchermo
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
    End If

End Sub

Private Sub CalcoliSomma_Click()
    Dim Riga, Colonna, Dummy
    On Error GoTo GestioneErroriCalcoliSomma

    Colonna = FormPrincipale!Grid1.SelStartCol
    If Colonna = NUM_COLONNE - 1 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    If Colonna < NUM_COLONNE - 1 Then
    
        FormPrincipale!Grid1.Row = 0
        NomeDellaVariabile = FormPrincipale!Grid1.Text
        FormCalcoli.Caption = "Somma numero a " & NomeDellaVariabile
        FormCalcoli.Show 1
    
        If NumeroPerCalcoli <> 0 Then
            FormGauge!Label6.Caption = "[1/2]"
            InserisciColonna
            
        If InseritaNuovaColonna = True Then
        
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            FormGauge.Show
            FormGauge.Caption = "Somma " & Format(NumeroPerCalcoli, "#####0.0####") & " a " & NomeDellaVariabile
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            For Riga = 1 To NUM_RIGHE - 1
                
                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Col = Colonna + 1
                Dummy = FormPrincipale!Grid1.Text
                    If FormPrincipale!Grid1.Text <> "" Then
                        FormPrincipale!Grid1.Col = Colonna
                        FormPrincipale!Grid1.Text = Dummy - (-(NumeroPerCalcoli))
                    End If
            Next Riga
        
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Somma"
            Unload FormGauge
            RipristinaSchermo
            DeselezionaColonna
            InseritaNuovaColonna = False
            FormPrincipale!Grid1.SetFocus
            Screen.MousePointer = 1
        Else
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
        End If
    End If
Exit Sub

GestioneErroriCalcoliSomma:
            Unload FormGauge
            RipristinaSchermo
            DeselezionaColonna
            InseritaNuovaColonna = False
            FormPrincipale!Grid1.SetFocus
            Screen.MousePointer = 1
            Exit Sub

End Sub

Private Sub CalcoliSottrai_Click()
    Dim Riga, Colonna, Dummy
    On Error GoTo GestioneErroriCalcoliSottrai

    Colonna = FormPrincipale!Grid1.SelStartCol
    If Colonna = NUM_COLONNE - 1 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    If Colonna < NUM_COLONNE - 1 Then
    
        FormPrincipale!Grid1.Row = 0
        NomeDellaVariabile = FormPrincipale!Grid1.Text
        FormCalcoli.Caption = "Sottrai numero a " & NomeDellaVariabile
        FormCalcoli.Show 1
    
        If NumeroPerCalcoli <> 0 Then
            FormGauge!Label6.Caption = "[1/2]"
            InserisciColonna
            
        If InseritaNuovaColonna = True Then
        
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            FormGauge.Show
            FormGauge.Caption = "Sottrai " & Format(NumeroPerCalcoli, "#####0.0####") & " a " & NomeDellaVariabile
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            For Riga = 1 To NUM_RIGHE - 1
                    
                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Col = Colonna + 1
                Dummy = FormPrincipale!Grid1.Text
                    If FormPrincipale!Grid1.Text <> "" Then
                        FormPrincipale!Grid1.Col = Colonna
                        FormPrincipale!Grid1.Text = Dummy - NumeroPerCalcoli
                    End If
            Next Riga
        
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Sottrai"
            Unload FormGauge
            RipristinaSchermo
            DeselezionaColonna
            InseritaNuovaColonna = False
            FormPrincipale!Grid1.SetFocus
            Screen.MousePointer = 1
        Else
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
        End If
    End If
Exit Sub

GestioneErroriCalcoliSottrai:
            Unload FormGauge
            RipristinaSchermo
            DeselezionaColonna
            InseritaNuovaColonna = False
            FormPrincipale!Grid1.SetFocus
            Screen.MousePointer = 1
            Exit Sub

End Sub

Private Sub CalcoliValoreAssoluto_Click()
    Dim Riga, Colonna, Dummy
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    If Colonna = NUM_COLONNE - 1 Then
        MsgBox "I calcoli non possono essere" + Chr(10) + "effettuati oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Seleziona colonna per i calcoli"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    
    If Colonna < NUM_COLONNE - 1 Then
        FormPrincipale!Grid1.Row = 0
        NomeDellaVariabile = FormPrincipale!Grid1.Text
        MsgBox "Verrà calcolato il valore" + Chr(10) + "assoluto della variabile.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "Valore assoluto di " & NomeDellaVariabile
        
        FormGauge!Label6.Caption = "[1/2]"
        InserisciColonna
            
        If InseritaNuovaColonna = True Then
            
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            FormGauge.Show
            FormGauge.Caption = "Valore assoluto di " & NomeDellaVariabile
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
            For Riga = 1 To NUM_RIGHE - 1

                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            FormPrincipale!Grid1.Row = Riga
            FormPrincipale!Grid1.Col = Colonna + 1
            Dummy = FormPrincipale!Grid1.Text
                If FormPrincipale!Grid1.Text <> "" Then
                    FormPrincipale!Grid1.Col = Colonna
                    FormPrincipale!Grid1.Text = Abs(Dummy)
                End If
            Next Riga
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Valore ass."
        End If
    End If
        
        Unload FormGauge
        RipristinaSchermo
        DeselezionaColonna
        FormPrincipale!Grid1.SetFocus
        Screen.MousePointer = 1

End Sub

Private Sub Cancella_Click()
    Dim Risposta

    FormPrincipale!Grid1.HighLight = False

    Risposta = MsgBox("Il file " + NomeDelFile + Chr(10) + "sarà cancellato," + " e tutti i dati saranno persi. Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Cancella il file dal disco")
        If Risposta = IDNO Then
            Exit Sub
        End If

    InizializzaFile
    Kill NomeDelFile

    NomeDelFile = ""
    FormPrincipale.Caption = Logo
    
    Unload FormGauge

    Grid1.Row = 1
    Grid1.Col = 1
    Grid1.TopRow = 1
    Grid1.LeftCol = 1

End Sub

Private Sub CommandAltman_Click()

    GraficaAltman_Click

End Sub

Private Sub CommandAnalisiDellaSomiglianza_Click()

    StatisticaCluster_Click

End Sub

Private Sub CommandAnalisiDEllaVarianza_Click()

    StatisticaAnova_Click

End Sub

Private Sub CommandAppaiamento_Click()

    GraficaPrimaDopo_Click

End Sub

Private Sub CommandApri_Click()
    
    FileApri_Click

End Sub

Private Sub CommandAssociazione_Click()

    GraficaAssociazione_Click

End Sub

Private Sub CommandBarreVerticali_Click()

    GraficaBarreVerticali_Click

End Sub

Private Sub CommandCalcolatrice_Click()
    
    Dim Calcolatrice
    On Error GoTo ErroreCalcolatriceCinqueBis:

    If FormPrincipale!Grid1.SelEndCol = FormPrincipale!Grid1.SelStartCol And FormPrincipale!Grid1.SelEndRow = FormPrincipale!Grid1.SelStartRow Then
        FormPrincipale!Grid1.HighLight = False
    End If
    
    Calcolatrice = Shell("Calc.exe", 1)

'    FormPrincipale!Grid1.SetFocus
        
Exit Sub

ErroreCalcolatriceCinqueBis:
    MsgBox "Calcolatrice non disponibile.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Calcolatrice"
Exit Sub

End Sub

Private Sub CommandConfronto_Click()

    GraficaConfronto_Click

End Sub

Private Sub CommandConfrontoTraMediane_Click()

    StatisticaNonParametricaW_Click

End Sub

Private Sub CommandConfrontoTraMedie_Click()

    StatisticaT_Click

End Sub

Private Sub CommandCumulativa_Click()

    GraficaCumulativa_Click

End Sub

Private Sub CommandDeviataZeta_Click()
    
    CalcoliDeviataZeta_Click

End Sub

Private Sub CommandDispersione_Click()

    GraficaDispersione_Click

End Sub

Private Sub CommandDistribuzione_Click()

    GraficaDistribuzione_Click

End Sub

Private Sub CommandErroreStandard_Click()

    GraficaErroreStandard_Click

End Sub

Private Sub CommandEsploraNonPar_Click()

    StatisticaEsploraPar_Click

End Sub

Private Sub CommandEsploraPar_Click()

    StatisticaEsploraNonPar_Click

End Sub

Private Sub CommandGuida_Click()
    
    Dim Dummy

'   Dummy = WinHelp(FormPrincipale.hWnd, App.Path & "\ministat.hlp", HELP_INDEX, "")
    
    CMDialog1.HelpFile = App.Path & "\ministat.hlp"
    CMDialog1.HelpCommand = cdlHelpContents
    CMDialog1.ShowHelp
    
    FormPrincipale!Grid1.SetFocus
    FormPrincipale!Grid1.HighLight = True

End Sub

Private Sub CommandIstogramma_Click()
    
        GraficaIstogramma_Click

End Sub

Private Sub CommandIstogrammiMultipli_Click()

    GraficaIstogrammiMultipli_Click

End Sub

Private Sub CommandLineaSpezzata_Click()

    GraficaLineaSpezzata_Click

End Sub

Private Sub CommandNumera_Click()

    CalcoliNumera_Click

End Sub

Private Sub CommandNuovo_Click()

    FileNuovo_Click

End Sub

Private Sub CommandOrdinaCrescente_Click()

    ModificaOrdinaCrescente_Click

End Sub

Private Sub CommandOrdinaDecrescente_Click()

    ModificaOrdinaDecrescente_Click

End Sub

Private Sub CommandPercentili_Click()

    CalcoliPercentili_Click

End Sub

Private Sub CommandQuartiliERange_Click()

    GraficaQuartili_Click

End Sub

Private Sub CommandRango_Click()

    CalcoliRango_Click

End Sub

Private Sub CommandRegNon_Click()

    StatisticaNonParametricaReglin_Click

End Sub

Private Sub CommandRegressioneLineare_Click()

    StatisticaRegressioneLineare_Click

End Sub

Private Sub CommandRegressioneMultipla_Click()

    StatisticaRegressioneMultipla_Click

End Sub

Private Sub CommandRegressionePolinomiale_Click()

    StatisticaRegressionePolinomiale_Click

End Sub

Private Sub CommandSalva_Click()

    If FileSalvatoInPrecedenza = True And LCase(Right(FormPrincipale.Caption, 5)) = ".mi1]" Then
        FileSalva_Click
    Else
        FileSalvaConNome_Click
    End If

    RipristinaSchermo

End Sub

Private Sub CommandSovrapposizione_Click()

    GraficaSovrapposizione_Click

End Sub

Private Sub CommandStampaColonne1_Click()

    FileStampaColonne1_Click

End Sub

Private Sub CommandStampaColonne6_Click()

    FileStampaColonne6_Click

End Sub

Private Sub CommandStatisticaBayesiana_Click()

    StatisticaBayesiana_Click

End Sub

Private Sub CommandStatisticheElementari_Click()

    StatisticaStatisticheElementari_Click

End Sub

Private Sub CommandStatNon_Click()

    StatisticaNonParametricaCumulativa_Click

End Sub

Private Sub CommandTestChiQuadrato_Click()

    StatisticaNonParametricaChi_Click

End Sub

Private Sub CommandZeta_Click()
    
    CalcoliDeviataZeta_Click

End Sub

Private Sub CommandTorta_Click()
    
        GraficaTorta_Click

End Sub

Private Sub EsportaDelimitatoCon_Click()
    On Error GoTo GestioneErroriDelimitatoCon
    Dim Dummy, TrovatoIlPunto, Posizione

    FormPrincipale!Grid1.HighLight = False
    FormPrincipale.Refresh

    CMDialog1.Filter = "Files ASCII (*.prn) |*.prn|"
    CMDialog1.FilterIndex = 1
    CMDialog1.DialogTitle = "Esporta file ASCII formato delimitato con il nome delle variabili"
    CMDialog1.filename = "*.prn"
    CMDialog1.Action = 2 ' l'errore (*) si verifica qui !
    NomeDelFile = CMDialog1.filename
' estensione del file forzata a .txt
    FormPrincipale.Refresh
    For Dummy = 1 To Len(NomeDelFile)
        If Mid(NomeDelFile, Dummy, 1) = "." Then
            TrovatoIlPunto = True
            Posizione = Dummy
        End If
    Next Dummy
' nessuna estensione
    If TrovatoIlPunto = False Then
        NomeDelFile = NomeDelFile & ".prn"
' per prudenza
    Else
        If LCase(Right(NomeDelFile, 4)) <> ".prn" Then
            MsgBox "Il file deve avere" & Chr(10) & ".prn come estensione.", MB_ICONSTOP, "Esporta file ASCII formato delimitato con il nome delle variabili"
            Exit Sub
        End If
    End If
    
    RigaDiPartenza = 0
    EsportaFormatoDelimitato (NomeDelFile)
    FormPrincipale!Grid1.SetFocus

Exit Sub

' routine errori locale necessaria per intercettare l'errore (*)
' che si verifica se si esce con un ESC o con un ANNULLA dalla
' finestra di dialogo per la selezione del nome del file
GestioneErroriDelimitatoCon:
    FormPrincipale!Grid1.SetFocus
    Exit Sub

End Sub

Private Sub EsportaDelimitatoSenza_Click()
    On Error GoTo GestioneErroriDelimitatoSenza
    Dim Dummy, TrovatoIlPunto, Posizione

    FormPrincipale!Grid1.HighLight = False
    FormPrincipale.Refresh

    CMDialog1.Filter = "Files ASCII (*.txt) |*.txt|"
    CMDialog1.FilterIndex = 1
    CMDialog1.DialogTitle = "Esporta file ASCII formato delimitato senza il nome delle variabili"
    CMDialog1.filename = "*.txt"
    CMDialog1.Action = 2 ' l'errore (*) si verifica qui !
    NomeDelFile = CMDialog1.filename
' estensione del file forzata a .txt
    FormPrincipale.Refresh
    For Dummy = 1 To Len(NomeDelFile)
        If Mid(NomeDelFile, Dummy, 1) = "." Then
            TrovatoIlPunto = True
            Posizione = Dummy
        End If
    Next Dummy
' nessuna estensione
    If TrovatoIlPunto = False Then
        NomeDelFile = NomeDelFile & ".txt"
' per prudenza
    Else
        If LCase(Right(NomeDelFile, 4)) <> ".txt" Then
            MsgBox "Il file deve avere" & Chr(10) & ".txt come estensione.", MB_ICONSTOP, "Esporta file ASCII formato delimitato senza il nome delle variabili"
            Exit Sub
        End If
    End If
    
    RigaDiPartenza = 1
    EsportaFormatoDelimitato (NomeDelFile)
    FormPrincipale!Grid1.SetFocus

Exit Sub

' routine errori locale necessaria per intercettare l'errore (*)
' che si verifica se si esce con un ESC o con un ANNULLA dalla
' finestra di dialogo per la selezione del nome del file
GestioneErroriDelimitatoSenza:
    FormPrincipale!Grid1.SetFocus
    Exit Sub

End Sub

Private Sub EsportaFissoCon_Click()
    On Error GoTo GestioneErroriFissoCon
    Dim Dummy, TrovatoIlPunto, Posizione

    FormPrincipale!Grid1.HighLight = False
    FormPrincipale.Refresh

    CMDialog1.Filter = "Files ASCII (*.prn) |*.prn|"
    CMDialog1.FilterIndex = 1
    CMDialog1.DialogTitle = "Esporta file ASCII formato fisso con il nome delle variabili"
    CMDialog1.filename = "*.prn"
    CMDialog1.Action = 2 ' l'errore (*) si verifica qui !
    NomeDelFile = CMDialog1.filename
' estensione del file forzata a .txt
    FormPrincipale.Refresh
    For Dummy = 1 To Len(NomeDelFile)
        If Mid(NomeDelFile, Dummy, 1) = "." Then
            TrovatoIlPunto = True
            Posizione = Dummy
        End If
    Next Dummy
' nessuna estensione
    If TrovatoIlPunto = False Then
        NomeDelFile = NomeDelFile & ".prn"
' per prudenza
    Else
        If LCase(Right(NomeDelFile, 4)) <> ".prn" Then
            MsgBox "Il file deve avere" & Chr(10) & ".prn come estensione.", MB_ICONSTOP, "Esporta file ASCII formato fisso con il nome delle variabili"
            Exit Sub
        End If
    End If
    
    RigaDiPartenza = 0
    EsportaFormatoFisso (NomeDelFile)
    FormPrincipale!Grid1.SetFocus

Exit Sub

' routine errori locale necessaria per intercettare l'errore (*)
' che si verifica se si esce con un ESC o con un ANNULLA dalla
' finestra di dialogo per la selezione del nome del file
GestioneErroriFissoCon:
    FormPrincipale!Grid1.SetFocus
    Exit Sub

End Sub

Private Sub EsportaFissoSenza_Click()
    On Error GoTo GestioneErroriFissoSenza
    Dim Dummy, TrovatoIlPunto, Posizione

    FormPrincipale!Grid1.HighLight = False
    FormPrincipale.Refresh

    CMDialog1.Filter = "Files ASCII (*.txt) |*.txt|"
    CMDialog1.FilterIndex = 1
    CMDialog1.DialogTitle = "Esporta file ASCII formato fisso senza il nome delle variabili"
    CMDialog1.filename = "*.txt"
    CMDialog1.Action = 2 ' l'errore (*) si verifica qui !
    NomeDelFile = CMDialog1.filename
' estensione del file forzata a .txt
    FormPrincipale.Refresh
    For Dummy = 1 To Len(NomeDelFile)
        If Mid(NomeDelFile, Dummy, 1) = "." Then
            TrovatoIlPunto = True
            Posizione = Dummy
        End If
    Next Dummy
' nessuna estensione
    If TrovatoIlPunto = False Then
        NomeDelFile = NomeDelFile & ".txt"
' per prudenza
    Else
        If LCase(Right(NomeDelFile, 4)) <> ".txt" Then
            MsgBox "Il file deve avere" & Chr(10) & ".txt come estensione.", MB_ICONSTOP, "Esporta file ASCII formato fisso senza il nome delle variabili"
            Exit Sub
        End If
    End If
    
    RigaDiPartenza = 1
    EsportaFormatoFisso (NomeDelFile)
    FormPrincipale!Grid1.SetFocus

Exit Sub

' routine errori locale necessaria per intercettare l'errore (*)
' che si verifica se si esce con un ESC o con un ANNULLA dalla
' finestra di dialogo per la selezione del nome del file
GestioneErroriFissoSenza:
    FormPrincipale!Grid1.SetFocus
    Exit Sub

End Sub

Private Sub FileApri_Click()
    On Error GoTo GestioneErroriFileApri
    Dim Risposta

    FormPrincipale!Grid1.ToolTipText = ""
    FormPrincipale!Grid1.HighLight = False
    FormPrincipale.Refresh
    
    If NomeDelFile <> "" Then
    Risposta = MsgBox("I dati non salvati saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Apri file Ministat 1.1")
        If Risposta = IDNO Then
            RipristinaSchermo
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
    End If
    
    CMDialog1.Filter = "Files Ministat 1.1 (*.mi1) |*.mi1|"
    CMDialog1.FilterIndex = 1
    CMDialog1.DialogTitle = "Apri file Ministat 1.1"
    CMDialog1.filename = "*.mi1"
    CMDialog1.Action = 1 ' l'errore (*) su verifica qui !
    NomeDelFile = CMDialog1.filename

    If LCase(Right(NomeDelFile, 4)) = ".mi1" Then

        Screen.MousePointer = 11
        FormPrincipale.Refresh
        ApriFile (NomeDelFile)

        FileSalvatoInPrecedenza = True
        Grid1.Row = 1
        Grid1.Col = 1
        Grid1.TopRow = 1
        Grid1.LeftCol = 1
        Grid1_Click
        NomeDelFileMinistat = NomeDelFile
        FormPrincipale!Grid1.SetFocus

    Else
        MsgBox "Il file deve avere" & Chr(10) & ".mi1 come estensione.", MB_ICONSTOP, "Apri file Ministat 1.1"
        FormPrincipale!Grid1.SetFocus
    End If

Exit Sub

' routine errori locale necessaria per intercettare l'errore (*)
' che si verifica se si esce con un ESC o con un ANNULLA dalla
' finestra di dialogo per la selezione del nome del file
GestioneErroriFileApri:
    RipristinaSchermo
    FormPrincipale!Grid1.SetFocus
    Exit Sub

End Sub

Private Sub FileCancella_Click()
    On Error GoTo GestioneErroriFileCancella
    Dim Risposta

    FormPrincipale!Grid1.HighLight = False
    FormPrincipale.Refresh

    Risposta = MsgBox("Saranno persi tutti di dati del file " + Chr(10) + NomeDelFileMinistat + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Cancella il file dal disco")
        If Risposta = IDNO Then
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If

    InizializzaFile
    Kill NomeDelFileMinistat

    FileSalvatoInPrecedenza = False
    NomeDelFile = ""
    NomeDellaTabella = ""
    FormPrincipale.Caption = Logo
    Grid1.Row = 1
    Grid1.Col = 1
    Grid1.TopRow = 1
    Grid1.LeftCol = 1
    Grid1_Click
    FormPrincipale!Grid1.SetFocus

Exit Sub

GestioneErroriFileCancella:
    MsgBox "Errore su disc(hett)o:" & Chr(10) & "per favore riprovare.", MB_ICONSTOP, "Cancella file Ministat 1.1"
    RipristinaSchermo
    FormPrincipale!Grid1.SetFocus
    Exit Sub

End Sub

Private Sub FileEsci_Click()
    Dim Risposta
    
    FormPrincipale!Grid1.HighLight = False
    FormPrincipale.Refresh
    
    Risposta = MsgBox("I dati non salvati saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Termina esecuzione " & Logo)
        If Risposta = IDNO Then
            FormPrincipale.Refresh
            Exit Sub
        Else
            End
        End If

End Sub

Private Sub FileImportaFileAccess_Click()
    On Error GoTo GestioneErroriFileImportaFileAccess
    Dim Risposta, I, J, Riga, Colonna
    Dim NumeroCampi, NumeroRecord, Posizione

    FormPrincipale!Grid1.ToolTipText = ""
    FormPrincipale!Grid1.HighLight = False
    FormPrincipale.Refresh
    NomeDellaTabella = ""
    
    If NomeDelFile <> "" Then
    Risposta = MsgBox("I dati non salvati saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Importa file Access")
        If Risposta = IDNO Then
            FormPrincipale!Grid1.SetFocus
        Exit Sub
        End If
    End If
    
    CMDialog1.Filter = "Files Access (*.mdb) |*.mdb|"
    CMDialog1.FilterIndex = 1
    CMDialog1.DialogTitle = "Importa file Access"
    CMDialog1.filename = "*.mdb"
222:
    CMDialog1.Action = 1 ' l'errore (*) si verifica qui !
    NomeDelFile = CMDialog1.filename
    
    If LCase(Right(NomeDelFile, 4)) = ".mdb" Then

    Screen.MousePointer = 11
        FormPrincipale.Refresh
        Data1.Connect = ""
    Data1.DatabaseName = NomeDelFile
69:
    Data1.Refresh

    FormTabelleAccess.Show 1

    If NomeDellaTabella = "" Then
        Data1.Recordset.Close
        Exit Sub
    
    Else
    
    Data1.RecordSource = NomeDellaTabella
    Data1.Refresh
    
    FormPrincipale!Grid1.Row = 0
    FormPrincipale!Grid1.Col = 0
    FormPrincipale.Refresh
    Screen.MousePointer = 1
    
    FormGauge!Label6.Caption = "[1/2]"
    InizializzaFile

    Screen.MousePointer = 11
    FormPrincipale.Refresh
    FormGauge.Show
    FormGauge.Caption = "Importa file Access"
    FormGauge!Label6.Caption = "[2/2]"
    FormGauge.Label1.Caption = "0%"
    FormGauge!Gauge1.Value = 0
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
    NumeroCampi = Data1.Recordset.Fields.Count
        If NumeroCampi > (NUM_COLONNE - 1) Then NumeroCampi = NUM_COLONNE - 1

    NumeroRecord = 0

    Grid1.Row = 0
    For I = 0 To NumeroCampi - 1
        Grid1.Col = I + 1
        Grid1.Text = Data1.Database.TableDefs(Data1.RecordSource).Fields(I).Name
    Next I
    If NumeroCampi < (NUM_COLONNE - 1) Then
        For I = NumeroCampi + 1 To NUM_COLONNE - 1
            Grid1.Col = I
            Grid1.Text = ""
        Next I
    End If
    FormGauge.Label1.Caption = "2%"
    FormGauge!Gauge1.Value = 2
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    Data1.Recordset.MoveFirst
    NumeroRecord = 0
        Do While Data1.Recordset.EOF = False
            NumeroRecord = NumeroRecord + 1
            Data1.Recordset.MoveNext
        Loop
            If NumeroRecord > NUM_RIGHE - 1 Then
                NumeroRecord = NUM_RIGHE - 1
            End If

        Data1.Recordset.MoveFirst
        
        For I = 1 To NumeroRecord
            If Int(100 * I / NumeroRecord) > 2 Then
                FormGauge.Label1.Caption = Int(100 * I / NumeroRecord) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / NumeroRecord)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If

        Grid1.Row = I
            For J = 0 To NumeroCampi - 1
                Grid1.Col = J + 1
'                If Data1.Recordset(Data1.Database.TableDefs(Data1.RecordSource).Fields(j).Name) <> "#NULL" Then
                 If Data1.Database.TableDefs(Data1.RecordSource).Fields(J).Type = 7 Then
                    Grid1.Text = Data1.Recordset(Data1.Database.TableDefs(Data1.RecordSource).Fields(J).Name)
                 End If
'                End If
            Next J
        Data1.Recordset.MoveNext

        Next I
    
    Data1.Recordset.Close
    
    FileSalvatoInPrecedenza = False
    Unload FormGauge
    FormPrincipale.Caption = Logo & " [" & NomeDelFile & "][tabella " & NomeDellaTabella & "]"
    Grid1.Row = 1
    Grid1.Col = 1
    Grid1.TopRow = 1
    Grid1.LeftCol = 1
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1

    End If ' if NomeDellaTabella <> ""

    Else
        MsgBox "E' possibile importare solamente" & Chr(10) & "files con .mdb come estensione.", MB_ICONSTOP, "Importa file Access"
        FormPrincipale!Grid1.SetFocus
    End If ' se estensione = ".MDB"

    Exit Sub

' routine errori locale necessaria per intercettare l'errore (*)
' che si verifica se si esce con un ESC o con un ANNULLA dalla
' finestra di dialogo per la selezione del nome del file
GestioneErroriFileImportaFileAccess:
    If Erl = 69 Then
        Resume Next
    Else
        If Erl = 222 Then
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        Else
            Unload FormGauge
            Screen.MousePointer = 1
            MsgBox "Errore su disc(hett)o," & Chr(10) & "o errore nella struttura" & Chr(10) & "o nei driver del file.", MB_ICONSTOP, "Importa file Access"
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
    End If

End Sub

Private Sub FileImportaFileASCII_Click()
    On Error GoTo GestioneErroriFileImportaFileASCII
    Dim Risposta

    FormPrincipale!Grid1.ToolTipText = ""
    FormPrincipale!Grid1.HighLight = False
    FormPrincipale.Refresh
    
    If NomeDelFile <> "" Then
    Risposta = MsgBox("I dati non salvati saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Importa file ASCII")
        If Risposta = IDNO Then
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
    End If
    
    CMDialog1.Filter = "Files ASCII (*.txt) |*.txt|Files ASCII (*.prn) |*.prn|Files ASCII (*.doc) |*.doc|Tutti i files (*.*) |*.*|"
    CMDialog1.FilterIndex = 1
    CMDialog1.DialogTitle = "Importa file ASCII"
    CMDialog1.filename = "*.txt"
6111:
    CMDialog1.Action = 1 ' l'errore (*) si verifica qui !
6112:
    NomeDelFile = CMDialog1.filename
    
    If LCase(Right(NomeDelFile, 4)) <> ".mi1" Then
    
    FormPrincipale.Refresh
    FormGauge!Label6.Caption = "[1/2]"
    InizializzaFile
    Load FormASCII
    
    Else
        MsgBox "L'estensione .mi1 non" & Chr(10) & "può essere utilizzata.", MB_ICONSTOP, "Importa file ASCII"
        FormPrincipale!Grid1.SetFocus
    End If ' se estensione <> ".mi1"

Exit Sub

' routine errori locale necessaria per intercettare l'errore (*)
' che si verifica se si esce con un ESC o con un ANNULLA dalla
' finestra di dialogo per la selezione del nome del file
GestioneErroriFileImportaFileASCII:
    If Erl = 6111 Then
        FormPrincipale!Grid1.SetFocus
        Exit Sub
    Else
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Errore su disc(hett)o," & Chr(10) & "o errore nella struttura" & Chr(10) & "o nei driver del file.", MB_ICONSTOP, "Importa file ASCII"
        FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If


End Sub

Private Sub FileImportaFileDbase_Click()
    
    On Error GoTo GestioneErroriFileImportaFileDbase
    Dim Risposta, I, J, Riga, Colonna
    Dim NumeroCampi, NumeroRecord, Posizione

    FormPrincipale!Grid1.ToolTipText = ""
    FormPrincipale!Grid1.HighLight = False
    FormPrincipale.Refresh
    
    If NomeDelFile <> "" Then
    Risposta = MsgBox("I dati non salvati saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Importa file dBase III")
        If Risposta = IDNO Then
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
    End If
    
    CMDialog1.Filter = "Files dBase III (*.dbf) |*.dbf|"
    CMDialog1.FilterIndex = 1
    CMDialog1.DialogTitle = "Importa file dBase III"
    CMDialog1.filename = "*.dbf"
111:
    CMDialog1.Action = 1 ' l'errore (*) si verifica qui !
112:
    NomeDelFile = CMDialog1.filename
    
    If LCase(Right(NomeDelFile, 4)) = ".dbf" Then
    FormPrincipale.Refresh
    FormGauge!Label6.Caption = "[1/2]"
    InizializzaFile

    Screen.MousePointer = 11
    For I = Len(NomeDelFile) To 1 Step -1
        If Mid(NomeDelFile, I, 1) = "\" Then
            Posizione = I
            Exit For
        End If
    Next I
    
    FormPrincipale.Refresh
    FormGauge.Show
    FormGauge.Caption = "Importa file dBase III"
    FormGauge!Label6.Caption = "[2/2]"
    FormGauge!Label1.Caption = "0%"
    FormGauge!Gauge1.Value = 0
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
    Data1.Connect = "dBASE III"
    Data1.DatabaseName = Left(NomeDelFile, Posizione - 1)
    Data1.RecordSource = Mid(NomeDelFile, Posizione + 1, Len(NomeDelFile) - Posizione - 4)
    Data1.Refresh

    NumeroCampi = Data1.Recordset.Fields.Count
        If NumeroCampi > (NUM_COLONNE - 1) Then NumeroCampi = NUM_COLONNE - 1

    NumeroRecord = 0

    Grid1.Row = 0
    For I = 0 To NumeroCampi - 1
        Grid1.Col = I + 1
        Grid1.Text = Data1.Database.TableDefs(Data1.RecordSource).Fields(I).Name
    Next I
    If NumeroCampi < (NUM_COLONNE - 1) Then
        For I = NumeroCampi + 1 To NUM_COLONNE - 1
            Grid1.Col = I
            Grid1.Text = ""
        Next I
    End If
    
    FormGauge.Label1.Caption = "2%"
    FormGauge!Gauge1.Value = 2
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    Data1.Recordset.MoveFirst
    NumeroRecord = 0
        Do While Data1.Recordset.EOF = False
            NumeroRecord = NumeroRecord + 1
            Data1.Recordset.MoveNext
        Loop
            If NumeroRecord > NUM_RIGHE - 1 Then
                NumeroRecord = NUM_RIGHE - 1
            End If

        Data1.Recordset.MoveFirst
        
        For I = 1 To NumeroRecord
            If Int(100 * I / NumeroRecord) > 2 Then
                FormGauge.Label1.Caption = Int(100 * I / NumeroRecord) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / NumeroRecord)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If

        Grid1.Row = I
            For J = 0 To NumeroCampi - 1
                Grid1.Col = J + 1
                If Data1.Recordset(Data1.Database.TableDefs(Data1.RecordSource).Fields(J).Name) <> "#NULL" Then
                 If Data1.Database.TableDefs(Data1.RecordSource).Fields(J).Type = 7 Then
                    Grid1.Text = Data1.Recordset(Data1.Database.TableDefs(Data1.RecordSource).Fields(J).Name)
                 End If
                End If
            Next J
        Data1.Recordset.MoveNext

        Next I
    
    Data1.Recordset.Close
    
    FileSalvatoInPrecedenza = False
    Unload FormGauge
    NomeDellaTabella = ""
    FormPrincipale.Caption = Logo & " [" & NomeDelFile & "]"
    Grid1.Row = 1
    Grid1.Col = 1
    Grid1.TopRow = 1
    Grid1.LeftCol = 1
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1

    Else
        MsgBox "E' possibile importare solamente" & Chr(10) & "files con .dbf come estensione.", MB_ICONSTOP, "Importa file dBase III"
        FormPrincipale!Grid1.SetFocus
    End If ' se estensione = ".DBF"

Exit Sub

' routine errori locale necessaria per intercettare l'errore (*)
' che si verifica se si esce con un ESC o con un ANNULLA dalla
' finestra di dialogo per la selezione del nome del file
GestioneErroriFileImportaFileDbase:
    If Erl = 111 Then
        FormPrincipale!Grid1.SetFocus
        Exit Sub
    Else
        Unload FormGauge
        Screen.MousePointer = 1
        MsgBox "Errore su disc(hett)o," & Chr(10) & "o errore nella struttura" & Chr(10) & "o nei driver del file.", MB_ICONSTOP, "Importa file dBase III"
        FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If

End Sub

Private Sub FileImportaFileExcel40_Click()

End Sub

Private Sub FileImportaFileExcel50_Click()

End Sub

Private Sub FileNuovo_Click()
    Dim Risposta

    FormPrincipale!Grid1.ToolTipText = ""
    FormPrincipale!Grid1.HighLight = False
    FormPrincipale.Refresh

    Risposta = MsgBox("I dati non salvati saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Nuovo file Ministat 1.1")
        If Risposta = IDNO Then
            RipristinaSchermo
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If

    FormGauge!Label6.Caption = "[1/1]"
    InizializzaFile

    FileSalvatoInPrecedenza = False
    NomeDelFile = ""
    NomeDellaTabella = ""
    FormPrincipale.Caption = Logo
    Grid1.Row = 1
    Grid1.Col = 1
    Grid1.TopRow = 1
    Grid1.LeftCol = 1
    Grid1_Click
    FormPrincipale!Grid1.SetFocus

End Sub

Private Sub FileSalva_Click()
     
    FormPrincipale!Grid1.HighLight = False
    FormPrincipale.Refresh

    SalvaFile (NomeDelFileMinistat)
    FormPrincipale!Grid1.SetFocus

End Sub

Private Sub FileSalvaConNome_Click()
    On Error GoTo GestioneErroriFileSalvaConNome
    Dim Dummy, TrovatoIlPunto, Posizione

    FormPrincipale!Grid1.ToolTipText = ""
    FormPrincipale!Grid1.HighLight = False
    FormPrincipale.Refresh

    CMDialog1.Filter = "Files Ministat 1.1 (*.mi1) |*.mi1|"
    CMDialog1.FilterIndex = 1
    CMDialog1.DialogTitle = "Salva file Ministat 1.1 con nome"
    CMDialog1.filename = "*.mi1"
    CMDialog1.Action = 2 ' l'errore (*) si verifica qui !
    NomeDelFile = CMDialog1.filename
' estensione del file forzata a .mi1
    For Dummy = 1 To Len(NomeDelFile)
        If Mid(NomeDelFile, Dummy, 1) = "." Then
            TrovatoIlPunto = True
            Posizione = Dummy
        End If
    Next Dummy
' nessuna estensione
    If TrovatoIlPunto = False Then
        NomeDelFile = NomeDelFile & ".mi1"
' per prudenza
    Else
        If Right(NomeDelFile, 4) <> ".mi1" Then
            MsgBox "Il file deve avere" & Chr(10) & ".mi1 come estensione.", MB_ICONSTOP, "Salva file Ministat 1.1 con nome"
            FormPrincipale!Grid1.SetFocus
            FormPrincipale.Refresh
            Exit Sub
        End If
    End If
    
    SalvaFile (NomeDelFile)
    FormPrincipale!Grid1.SetFocus

Exit Sub

' routine errori locale necessaria per intercettare l'errore (*)
' che si verifica se si esce con un ESC o con un ANNULLA dalla
' finestra di dialogo per la selezione del nome del file
GestioneErroriFileSalvaConNome:
    FormPrincipale!Grid1.SetFocus
    Exit Sub

End Sub

Private Sub FileStampaColonne1_Click()
    On Error GoTo GestioneErroriStampaColonne1
    Dim Risposta, FlagRigaVuota, Tabulatore, Z1
    Dim Riga, Colonna, Posizione1, Posizione2
    Dim ContaPagine, ContaRighe, SpessoreLinea

    FormPrincipale!Grid1.HighLight = False
    FormPrincipale.Refresh
    
    If Dir(NomeDelFile) <> "" Then
        Risposta = MsgBox("Stampa dati del file" + Chr(10) + NomeDelFileMinistat + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, "Stampa dati colonne 1-5")
            If Risposta = IDNO Then
                FormPrincipale!Grid1.SetFocus
                Exit Sub
            End If
    End If

    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Stampa dati colonne 1-5"
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Caption = "0%"
    FormGauge!Gauge1.Value = 0
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    Z1 = 0
    For Riga = 1 To NUM_RIGHE - 1
        FormPrincipale!Grid1.Row = Riga
' controlla se la riga è vuota
        FlagRigaVuota = True
        For Colonna = 1 To (NUM_COLONNE - 1) / 2
            FormPrincipale!Grid1.Col = Colonna
            If FormPrincipale!Grid1.Text <> "" Then
                FlagRigaVuota = False
            End If
        Next Colonna
        If FlagRigaVuota = False Then
            Z1 = Z1 + 1
        Else
            Exit For
        End If
    Next Riga
    
    If Z1 = 0 Then Z1 = 1

    ContaPagine = 1
    ContaRighe = 1
    Posizione1 = Printer.CurrentY
    Printer.Print "Ministat 1.1 "; Tab(32); Now; Tab(66); "pag. "; Tab(71); Trim(Str$(ContaPagine)) + "[1-5]"
    Posizione2 = Printer.CurrentY
    SpessoreLinea = Posizione2 - Posizione1
    Printer.Print
    Printer.Print "File: " & NomeDelFileMinistat
    Printer.Print
    Printer.Print Tab(1); "Riga";
    FormPrincipale!Grid1.Row = 0
    Tabulatore = 7
        For Colonna = 1 To (NUM_COLONNE - 1) / 2
           FormPrincipale!Grid1.Col = Colonna
            Printer.Print Tab(Tabulatore); Format(Left(FormPrincipale!Grid1.Text, 12), "@@@@@@@@@@@@");
            Tabulatore = Tabulatore + 14
        Next Colonna
    Printer.Print ""
    Printer.Print

    For Riga = 1 To NUM_RIGHE - 1

        If Printer.ScaleHeight - Printer.CurrentY < (2 * SpessoreLinea) Then
            Printer.NewPage
            ContaPagine = ContaPagine + 1
            Printer.Print "Ministat 1.1 "; Tab(32); Now; Tab(66); "pag. "; Tab(71); Trim(Str$(ContaPagine)) + "[1-5]"
            Printer.Print
            Printer.Print "File: " & NomeDelFileMinistat
            Printer.Print
            Printer.Print Tab(1); "Riga";
            FormPrincipale!Grid1.Row = 0
            Tabulatore = 7
                For Colonna = 1 To (NUM_COLONNE - 1) / 2
                   FormPrincipale!Grid1.Col = Colonna
                   Printer.Print Tab(Tabulatore); Format(Left(FormPrincipale!Grid1.Text, 12), "@@@@@@@@@@@@");
                   Tabulatore = Tabulatore + 14
                Next Colonna
            Printer.Print ""
            Printer.Print
        End If

        FormGauge.Label1.Caption = Int(100 * (Riga - 1) / Z1) & "%"
        FormGauge!Gauge1.Value = Int(100 * (Riga - 1) / Z1)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        FormPrincipale!Grid1.Row = Riga
' controlla se la riga è vuota
            FlagRigaVuota = True
            For Colonna = 1 To (NUM_COLONNE - 1) / 2
                FormPrincipale!Grid1.Col = Colonna
                If FormPrincipale!Grid1.Text <> "" Then
                    FlagRigaVuota = False
                End If
            Next Colonna
' se la riga non è vuota allora stampa
            If FlagRigaVuota = False Then
                Printer.Print Tab(1); Format(ContaRighe, "@@@@");
                Tabulatore = 7
                For Colonna = 1 To (NUM_COLONNE - 1) / 2
                    FormPrincipale!Grid1.Col = Colonna
                    Printer.Print Tab(Tabulatore); Format(Left(FormPrincipale!Grid1.Text, 12), "@@@@@@@@@@@@");
                    Tabulatore = Tabulatore + 14
                Next Colonna
                Printer.Print ""
                ContaRighe = ContaRighe + 1
            Else
                Exit For
            End If

    Next Riga

    FormGauge!Label1.Caption = "100%"
    FormGauge!Gauge1.Value = 100
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    Printer.EndDoc
    Unload FormGauge
    RipristinaSchermo
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1

Exit Sub

GestioneErroriStampaColonne1:
    MsgBox "Errore durante la stampa:" & Chr(10) & "per favore riprovare.", MB_ICONEXCLAMATION, "Stampa dati colonne 1-5"
    Unload FormGauge
    RipristinaSchermo
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1
    Exit Sub

End Sub

Private Sub FileStampaColonne6_Click()
    On Error GoTo GestioneErroriStampaColonne6
    Dim Risposta, FlagRigaVuota, Tabulatore, Z1
    Dim Riga, Colonna, Posizione1, Posizione2
    Dim ContaPagine, ContaRighe, SpessoreLinea

    FormPrincipale!Grid1.HighLight = False
    FormPrincipale.Refresh

    If Dir(NomeDelFile) <> "" Then
        Risposta = MsgBox("Stampa dati del file" + Chr(10) + NomeDelFileMinistat + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, "Stampa dati colonne 6-10")
            If Risposta = IDNO Then
                FormPrincipale!Grid1.SetFocus
                Exit Sub
            End If
    End If

    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Stampa dati colonne 6-10"
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Caption = "0%"
    FormGauge!Gauge1.Value = 0
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    Z1 = 0
    For Riga = 1 To NUM_RIGHE - 1
        FormPrincipale!Grid1.Row = Riga
' controlla se la riga è vuota
        FlagRigaVuota = True
        For Colonna = ((NUM_COLONNE - 1) / 2) + 1 To NUM_COLONNE - 1
            FormPrincipale!Grid1.Col = Colonna
            If FormPrincipale!Grid1.Text <> "" Then
                FlagRigaVuota = False
            End If
        Next Colonna
        If FlagRigaVuota = False Then
            Z1 = Z1 + 1
        Else
            Exit For
        End If
    Next Riga
    
    If Z1 = 0 Then Z1 = 1

    ContaPagine = 1
    ContaRighe = 1
    Posizione1 = Printer.CurrentY
    Printer.Print "Ministat 1.1 "; Tab(32); Now; Tab(66); "pag. "; Tab(71); Trim(Str$(ContaPagine)) + "[6-10]"
    Posizione2 = Printer.CurrentY
    SpessoreLinea = Posizione2 - Posizione1
    Printer.Print
    Printer.Print "File: " & NomeDelFileMinistat
    Printer.Print
    Printer.Print Tab(1); "Riga";
    FormPrincipale!Grid1.Row = 0
    Tabulatore = 7
        For Colonna = ((NUM_COLONNE - 1) / 2) + 1 To NUM_COLONNE - 1
        FormPrincipale!Grid1.Col = Colonna
            Printer.Print Tab(Tabulatore); Format(Left(FormPrincipale!Grid1.Text, 12), "@@@@@@@@@@@@");
            Tabulatore = Tabulatore + 14
        Next Colonna
    Printer.Print ""
    Printer.Print

    For Riga = 1 To NUM_RIGHE - 1

        If Printer.ScaleHeight - Printer.CurrentY < (2 * SpessoreLinea) Then
            Printer.NewPage
            ContaPagine = ContaPagine + 1
            Printer.Print "Ministat 1.1 "; Tab(32); Now; Tab(66); "pag. "; Tab(71); Trim(Str$(ContaPagine)) + "[6-10]"
            Printer.Print
            Printer.Print "File: " & NomeDelFileMinistat
            Printer.Print
            Printer.Print Tab(1); "Riga";
            FormPrincipale!Grid1.Row = 0
            Tabulatore = 7
                For Colonna = ((NUM_COLONNE - 1) / 2) + 1 To NUM_COLONNE - 1
                   FormPrincipale!Grid1.Col = Colonna
                   Printer.Print Tab(Tabulatore); Format(Left(FormPrincipale!Grid1.Text, 12), "@@@@@@@@@@@@");
                   Tabulatore = Tabulatore + 14
                Next Colonna
            Printer.Print ""
            Printer.Print
        End If

        FormGauge.Label1.Caption = Int(100 * (Riga - 1) / Z1) & "%"
        FormGauge!Gauge1.Value = Int(100 * (Riga - 1) / Z1)
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        FormPrincipale!Grid1.Row = Riga
' controlla se la riga è vuota
            FlagRigaVuota = True
            For Colonna = ((NUM_COLONNE - 1) / 2) + 1 To NUM_COLONNE - 1
                FormPrincipale!Grid1.Col = Colonna
                If FormPrincipale!Grid1.Text <> "" Then
                    FlagRigaVuota = False
                End If
            Next Colonna
' se la riga non è vuota allora stampa
            If FlagRigaVuota = False Then
                Printer.Print Tab(1); Format(ContaRighe, "@@@@");
                Tabulatore = 7
                For Colonna = ((NUM_COLONNE - 1) / 2) + 1 To NUM_COLONNE - 1
                    FormPrincipale!Grid1.Col = Colonna
                    Printer.Print Tab(Tabulatore); Format(Left(FormPrincipale!Grid1.Text, 12), "@@@@@@@@@@@@");
                    Tabulatore = Tabulatore + 14
                Next Colonna
                Printer.Print ""
                ContaRighe = ContaRighe + 1
            Else
                Exit For
            End If

    Next Riga

    FormGauge!Label1.Caption = "100%"
    FormGauge!Gauge1.Value = 100
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    Printer.EndDoc
    Unload FormGauge
    RipristinaSchermo
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1

Exit Sub

GestioneErroriStampaColonne6:
    MsgBox "Errore durante la stampa:" & Chr(10) & "per favore riprovare.", MB_ICONEXCLAMATION, "Stampa dati colonne 6-10"
    Unload FormGauge
    RipristinaSchermo
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1
    Exit Sub

End Sub

Private Sub FileImportaFileExcel80_Click()

    On Error GoTo GestioneErroriFileImportaFileExcel80
    Dim Risposta, I, J, Riga, Colonna
    Dim NumeroCampi, NumeroRecord, Posizione

    FormPrincipale!Grid1.ToolTipText = ""
    FormPrincipale!Grid1.HighLight = False
    FormPrincipale.Refresh
    NomeDellaTabella = ""
    
    If NomeDelFile <> "" Then
    Risposta = MsgBox("I dati non salvati saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Importa file Excel 8.0")
        If Risposta = IDNO Then
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
    End If
    
    CMDialog1.Filter = "Files Excel (*.xls) |*.xls|"
    CMDialog1.FilterIndex = 1
    CMDialog1.DialogTitle = "Importa file Excel 8.0"
    CMDialog1.filename = "*.xls"
222:
    CMDialog1.Action = 1 ' l'errore (*) si verifica qui !
    NomeDelFile = CMDialog1.filename
    
    If LCase(Right(NomeDelFile, 4)) = ".xls" Then

    Screen.MousePointer = 11
    FormPrincipale.Refresh
    Data1.Connect = "Excel 8.0"
    Data1.DatabaseName = NomeDelFile
69:
    Data1.Refresh

    FormTabelleAccess.Show 1

    If NomeDellaTabella = "" Then
        Data1.Recordset.Close
        Exit Sub
    
    Else
    
    Data1.RecordSource = NomeDellaTabella
    Data1.Refresh
    
    FormPrincipale!Grid1.Row = 0
    FormPrincipale!Grid1.Col = 0
    FormPrincipale.Refresh
    Screen.MousePointer = 1
    
    FormGauge!Label6.Caption = "[1/2]"
    InizializzaFile

    Screen.MousePointer = 11
    FormPrincipale.Refresh
    FormGauge.Show
    FormGauge.Caption = "Importa file Excel 8.0"
    FormGauge!Label6.Caption = "[2/2]"
    FormGauge.Label1.Caption = "0%"
    FormGauge!Gauge1.Value = 0
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    
    NumeroCampi = Data1.Recordset.Fields.Count
        If NumeroCampi > (NUM_COLONNE - 1) Then NumeroCampi = NUM_COLONNE - 1

    NumeroRecord = 0

    Grid1.Row = 0
    For I = 0 To NumeroCampi - 1
        Grid1.Col = I + 1
        Grid1.Text = Data1.Database.TableDefs(Data1.RecordSource).Fields(I).Name
'       Grid1.Text = Data1.Recordset(Data1.Database.TableDefs(Data1.RecordSource).Fields(I).Name)
    Next I
    If NumeroCampi < (NUM_COLONNE - 1) Then
        For I = NumeroCampi + 1 To NUM_COLONNE - 1
            Grid1.Col = I
            Grid1.Text = ""
        Next I
    End If
    FormGauge.Label1.Caption = "2%"
    FormGauge!Gauge1.Value = 2
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    Data1.Recordset.MoveFirst
    NumeroRecord = 0
        Do While Data1.Recordset.EOF = False
            NumeroRecord = NumeroRecord + 1
            Data1.Recordset.MoveNext
        Loop
            If NumeroRecord > NUM_RIGHE - 1 Then
                NumeroRecord = NUM_RIGHE - 1
            End If

        Data1.Recordset.MoveFirst
        
        For I = 1 To NumeroRecord
            If Int(100 * I / NumeroRecord) > 2 Then
                FormGauge.Label1.Caption = Int(100 * I / NumeroRecord) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / NumeroRecord)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If

        Grid1.Row = I
            For J = 0 To NumeroCampi - 1
                Grid1.Col = J + 1
'                If Data1.Recordset(Data1.Database.TableDefs(Data1.RecordSource).Fields(j).Name) <> "#NULL" Then
                 If Data1.Database.TableDefs(Data1.RecordSource).Fields(J).Type = 7 Then
                    Grid1.Text = Data1.Recordset(Data1.Database.TableDefs(Data1.RecordSource).Fields(J).Name)
                 End If
'                End If
            Next J
        Data1.Recordset.MoveNext

        Next I
    
    Data1.Recordset.Close
    
    FileSalvatoInPrecedenza = False
    Unload FormGauge
    FormPrincipale.Caption = Logo & " [" & NomeDelFile & "][tabella " & NomeDellaTabella & "]"
    Grid1.Row = 1
    Grid1.Col = 1
    Grid1.TopRow = 1
    Grid1.LeftCol = 1
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1

    End If ' if NomeDellaTabella <> ""

    Else
        MsgBox "E' possibile importare solamente" & Chr(10) & "files con .xls come estensione.", MB_ICONSTOP, "Importa file Excel 8.0"
        FormPrincipale!Grid1.SetFocus
    End If ' se estensione = ".MDB"

    Exit Sub

' routine errori locale necessaria per intercettare l'errore (*)
' che si verifica se si esce con un ESC o con un ANNULLA dalla
' finestra di dialogo per la selezione del nome del file
GestioneErroriFileImportaFileExcel80:
    If Erl = 69 Then
        Resume Next
    Else
        If Erl = 222 Then
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        Else
            Unload FormGauge
            Screen.MousePointer = 1
            MsgBox "Errore su disc(hett)o," & Chr(10) & "o errore nella struttura" & Chr(10) & "o nei driver del file.", MB_ICONSTOP, "Importa file Excel 8.0"
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
    End If

End Sub

Private Sub Form_Load()
    Dim Riga, Colonna

' cambia la directory a quella di esecuzione del programma
    ChDrive App.Path
    ChDir App.Path
' nessun file in uso
    NomeDelFile = ""
    NomeDellaTabella = ""
    Logo = "Ministat 1.1 (32 bit)"
    FormPrincipale.Caption = Logo
' niente nel ritaglio
    QualcosaNelRitaglio = False
' serve, serve...
    FileSalvatoInPrecedenza = False
    AnnullatoCaricamentoForm = False
' centra il form
    FormPrincipale.Top = Screen.Height / 2 - Height / 2
    FormPrincipale.Left = Screen.Width / 2 - Width / 2
' inizializza variabili globali per Grid1
    Rig = 1
    Col = 1
    RigSup = 1
    ColSin = 1
' numero di righe e colonne dell'oggetto Grid1
    Grid1.Rows = NUM_RIGHE
    Grid1.Cols = NUM_COLONNE
' sposta in alto a sinistra
    If BarraDegliStrumenti = True Then
        Grid1.Move 0, 1800
    Else
        Grid1.Move 0, 0
    End If
' dimensionamento, ripreso nella sub resize
    If BarraDegliStrumenti = True Then
        Grid1.Width = ScaleWidth
        Grid1.Height = ScaleHeight - 1800
    Else
        Grid1.Width = ScaleWidth
        Grid1.Height = ScaleHeight
    End If
' definisce larghezza colonne
    Grid1.ColWidth(0) = 574
    Grid1.Row = 0
        For Colonna = 1 To NUM_COLONNE - 1
            Grid1.ColWidth(Colonna) = LARGHEZZA_COLONNE
        Next Colonna
' definisce altezza righe
    For Riga = 0 To NUM_RIGHE - 1
        Grid1.RowHeight(Riga) = 255
    Next Riga
' intestazione righe
    Grid1.Col = 0
        For Riga = 1 To NUM_RIGHE - 1
            Grid1.Row = Riga
            Grid1.Text = Riga
        Next Riga
' intestazione colonne
    Grid1.Row = 0
        For Colonna = 1 To NUM_COLONNE - 1
            Grid1.Col = Colonna
            Grid1.Text = "Variabile " & Colonna
        Next Colonna
' allineamento colonne
    Grid1.FixedAlignment(0) = 1 ' intestazioni righe a destra
        For Colonna = 1 To NUM_COLONNE - 1
            Grid1.FixedAlignment(Colonna) = 2 ' nome variabili centrato
            Grid1.ColAlignment(Colonna) = 1 ' dati a destra
        Next Colonna

' inizializza pulsanti
    If BarraDegliStrumenti = True Then
    
    CommandApri.Enabled = True
    CommandSalva.Enabled = True
    CommandNuovo.Enabled = True
    CommandStampaColonne1.Enabled = False
    CommandStampaColonne6.Enabled = False
    
    CommandNumera.Enabled = False
    CommandOrdinaCrescente.Enabled = False
    CommandOrdinaDecrescente.Enabled = False
    CommandRango.Enabled = False
    CommandPercentili.Enabled = False
    CommandDeviataZeta.Enabled = False
    
    CommandIstogramma.Enabled = False
    CommandCumulativa.Enabled = False
    CommandConfronto.Enabled = False
    CommandAppaiamento.Enabled = False
    CommandDispersione.Enabled = False
    CommandAltman.Enabled = False
    CommandAssociazione.Enabled = False
    CommandSovrapposizione.Enabled = False
    CommandLineaSpezzata.Enabled = False
    CommandBarreVerticali.Enabled = False
    CommandIstogrammiMultipli.Enabled = False
    CommandDistribuzione.Enabled = False
    CommandErroreStandard.Enabled = False
    CommandQuartiliERange.Enabled = False
    
    CommandStatisticheElementari.Enabled = False
    CommandConfrontoTraMedie.Enabled = False
    CommandAnalisiDEllaVarianza.Enabled = False
    CommandRegressioneLineare.Enabled = False
    CommandRegressionePolinomiale.Enabled = False
    CommandRegressioneMultipla.Enabled = False
    CommandStatisticaBayesiana.Enabled = False
    CommandAnalisiDellaSomiglianza.Enabled = False
    CommandTestChiQuadrato.Enabled = False
    CommandStatNon.Enabled = False
    CommandConfrontoTraMediane.Enabled = False
    CommandRegNon.Enabled = False

    CommandApri.BevelWidth = 3
    CommandSalva.BevelWidth = 3
    CommandNuovo.BevelWidth = 3
    CommandStampaColonne1.BevelWidth = 1
    CommandStampaColonne6.BevelWidth = 1
    
    CommandNumera.BevelWidth = 1
    CommandOrdinaCrescente.BevelWidth = 1
    CommandOrdinaDecrescente.BevelWidth = 1
    CommandRango.BevelWidth = 1
    CommandPercentili.BevelWidth = 1
    CommandDeviataZeta.BevelWidth = 1
    
    CommandIstogramma.BevelWidth = 1
    CommandCumulativa.BevelWidth = 1
    CommandConfronto.BevelWidth = 1
    CommandAppaiamento.BevelWidth = 1
    CommandDispersione.BevelWidth = 1
    CommandAltman.BevelWidth = 1
    CommandAssociazione.BevelWidth = 1
    CommandSovrapposizione.BevelWidth = 1
    CommandLineaSpezzata.BevelWidth = 1
    CommandBarreVerticali.BevelWidth = 1
    CommandIstogrammiMultipli.BevelWidth = 1
    CommandDistribuzione.BevelWidth = 1
    CommandErroreStandard.BevelWidth = 1
    CommandQuartiliERange.BevelWidth = 1
    
    CommandStatisticheElementari.BevelWidth = 1
    CommandConfrontoTraMedie.BevelWidth = 1
    CommandAnalisiDEllaVarianza.BevelWidth = 1
    CommandRegressioneLineare.BevelWidth = 1
    CommandRegressionePolinomiale.BevelWidth = 1
    CommandRegressioneMultipla.BevelWidth = 1
    CommandStatisticaBayesiana.BevelWidth = 1
    CommandAnalisiDellaSomiglianza.BevelWidth = 1
    CommandTestChiQuadrato.BevelWidth = 1
    CommandStatNon.BevelWidth = 1
    CommandConfrontoTraMediane.BevelWidth = 1
    CommandRegNon.BevelWidth = 1

    End If

' dummy
    Grid1.HighLight = False
    Grid1.Col = 1
    Grid1.Row = 1
    Grid1.HighLight = True

End Sub

Private Sub Form_Resize()

    If BarraDegliStrumenti = True Then
        Grid1.Width = ScaleWidth
        If ScaleHeight - 1800 > 0 Then
            Grid1.Height = ScaleHeight - 1800
        Else
            Grid1.Height = ScaleHeight
        End If
    Else
        Grid1.Width = ScaleWidth
        Grid1.Height = ScaleHeight
    End If

End Sub

Private Sub GraficaAltman_Click()

'    Screen.MousePointer = 11
    Load FormAltman
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormAltman
        End If


End Sub

Private Sub GraficaAssociazione_Click()
    
'    Screen.MousePointer = 11
    Load FormAssociazione
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormAssociazione
        End If

End Sub

Private Sub GraficaBarreVerticali_Click()
    
'    Screen.MousePointer = 11
    Load FormBarreVerticali
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormBarreVerticali
        End If

End Sub

Private Sub GraficaConfronto_Click()

'    Screen.MousePointer = 11
    Load FormConfronto
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormConfronto
        End If

End Sub

Private Sub GraficaCumulativa_Click()
    
'    Screen.MousePointer = 11
    Load FormCumuGrafica
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormCumuGrafica
        End If

End Sub

Private Sub GraficaDispersione_Click()

'    Screen.MousePointer = 11
    Load FormDispersione
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormDispersione
        End If

End Sub

Private Sub GraficaDistribuzione_Click()
    
'    Screen.MousePointer = 11
    Load FormDistribuzione
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormDistribuzione
        End If

End Sub

Private Sub GraficaErroreStandard_Click()
    
'    Screen.MousePointer = 11
    Load FormErroreStandard
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormErroreStandard
        End If

End Sub

Private Sub GraficaIstogramma_Click()
    
'    Screen.MousePointer = 11
    Load FormIstoGrafica
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormIstoGrafica
        End If

End Sub

Private Sub GraficaIstogrammiMultipli_Click()
    
'    Screen.MousePointer = 11
    Load FormIstogrammiMultipli
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormIstogrammiMultipli
        End If

End Sub

Private Sub GraficaLineaSpezzata_Click()
    
'    Screen.MousePointer = 11
    Load FormSpezzata
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormSpezzata
        End If

End Sub

Private Sub GraficaPrimaDopo_Click()

'    Screen.MousePointer = 11
    Load FormPrimaDopo
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormPrimaDopo
        End If

End Sub

Private Sub GraficaQuartili_Click()
    
'    Screen.MousePointer = 11
    Load FormQuartili
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormQuartili
        End If

End Sub

Private Sub GraficaSovrapposizione_Click()
    
'    Screen.MousePointer = 11
    Load FormSovrapposizione
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormSovrapposizione
        End If

End Sub

Private Sub GraficaTorta_Click()

'    Screen.MousePointer = 11
    Load FormPie
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormPie
        End If

End Sub

Private Sub Grid1_Click()

    FormPrincipale.Refresh
'    FormPrincipale!Grid1.ToolTipText = ""
    
' memorizza coordinate del puntatore !
    Rig = FormPrincipale!Grid1.Row
    Col = FormPrincipale!Grid1.Col
    RigSup = FormPrincipale!Grid1.TopRow
    ColSin = FormPrincipale!Grid1.LeftCol

End Sub

Private Sub Grid1_DblClick()
    Dim Dummy
    
    If Grid1.Text <> "" Then
        Dummy = Grid1.Text
        Grid1.Text = -Grid1.Text
        If Dummy > 0 Then
            MsgBox "Il segno del dato è stato cambiato" + Chr(13) + "da positivo (+) a negativo (-).", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "Cambia il segno (dato: " & Dummy & ")"
        End If
        If Dummy < 0 Then
            MsgBox "Il segno del dato è stato cambiato" + Chr(13) + "da negativo (-) a positivo (+).", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "Cambia il segno (dato: " & Dummy & ")"
        End If
    End If

End Sub

Private Sub Grid1_GotFocus()

    FormPrincipale.Refresh
    
    If (FormPrincipale!Grid1.SelStartRow = FormPrincipale!Grid1.SelEndRow) And (FormPrincipale!Grid1.SelStartCol = FormPrincipale!Grid1.SelEndCol) Then
    If FileSalvatoInPrecedenza = True Then
        If LCase(Right(FormPrincipale.Caption, 5)) = ".mi1]" Then
            CommandStampaColonne1.Enabled = True
            CommandStampaColonne6.Enabled = True
            CommandStampaColonne1.BevelWidth = 3
            CommandStampaColonne6.BevelWidth = 3
        Else
            CommandStampaColonne1.Enabled = False
            CommandStampaColonne6.Enabled = False
            CommandStampaColonne1.BevelWidth = 1
            CommandStampaColonne6.BevelWidth = 1
        End If
    Else
        CommandStampaColonne1.Enabled = False
        CommandStampaColonne6.Enabled = False
        CommandStampaColonne1.BevelWidth = 1
        CommandStampaColonne6.BevelWidth = 1
    End If
    Else
        CommandStampaColonne1.Enabled = False
        CommandStampaColonne6.Enabled = False
        CommandStampaColonne1.BevelWidth = 1
        CommandStampaColonne6.BevelWidth = 1
    End If

End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Dim NumeroDummy
    Dim StringaDummy As String
    
    FormPrincipale!Grid1.ToolTipText = ""
    
NumeroDummy = 69.69
StringaDummy = Format(NumeroDummy, "##.##")

If StringaDummy = "69,69" Then

' se la prima cifra introdotta è stato uno 0
' allora DEVE essere introdotta una virgola !
' (la sola eccezione ammessa riguarda il backspace)
    If Grid1.Text = "0" And KeyAscii <> 44 Then
        If KeyAscii = 8 Then
            Grid1.Text = ""
        Else
            Exit Sub
        End If
    End If
' cifre da 0 a 9
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Grid1.Text = Grid1.Text + Chr(KeyAscii)
    End If
' virgola
    If ImmessaVirgolaPrecedentemente = False Then
        If Len(Grid1.Text) > 6 Then ' massimo sei cifre prima della virgola
            Grid1.Text = Left(Grid1.Text, 6)
        End If
        If KeyAscii = 44 Then
            Grid1.Text = Grid1.Text + Chr(KeyAscii)
            ImmessaVirgolaPrecedentemente = True
            PosizioneDellaVirgola = Len(Grid1.Text)
        End If
    End If
' backspace
    If KeyAscii = 8 And Len(Grid1.Text) > 0 Then
        If Right(Grid1.Text, 1) = "," Then
            ImmessaVirgolaPrecedentemente = False
        End If
        Grid1.Text = Left(Grid1.Text, Len(Grid1.Text) - 1)
    End If
' limita a cinque i decimali !
    If ImmessaVirgolaPrecedentemente = True Then
        If Len(Grid1.Text) - PosizioneDellaVirgola > 5 Then
            Grid1.Text = Left(Grid1.Text, Len(Grid1.Text) - 1)
        End If
    End If

End If

If StringaDummy = "69.69" Then

' se la prima cifra introdotta è stato uno 0
' allora DEVE essere introdotta una virgola !
' (la sola eccezione ammessa riguarda il backspace)
    If Grid1.Text = "0" And KeyAscii <> 46 Then
        If KeyAscii = 8 Then
            Grid1.Text = ""
        Else
            Exit Sub
        End If
    End If
' cifre da 0 a 9
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Grid1.Text = Grid1.Text + Chr(KeyAscii)
    End If
' punto
    If ImmessaVirgolaPrecedentemente = False Then
        If Len(Grid1.Text) > 6 Then ' massimo sei cifre prima della virgola
            Grid1.Text = Left(Grid1.Text, 6)
        End If
        If KeyAscii = 46 Then
            Grid1.Text = Grid1.Text + Chr(KeyAscii)
            ImmessaVirgolaPrecedentemente = True
            PosizioneDellaVirgola = Len(Grid1.Text)
        End If
    End If
' backspace
    If KeyAscii = 8 And Len(Grid1.Text) > 0 Then
        If Right(Grid1.Text, 1) = "," Then
            ImmessaVirgolaPrecedentemente = False
        End If
        Grid1.Text = Left(Grid1.Text, Len(Grid1.Text) - 1)
    End If
' limita a cinque i decimali !
    If ImmessaVirgolaPrecedentemente = True Then
        If Len(Grid1.Text) - PosizioneDellaVirgola > 5 Then
            Grid1.Text = Left(Grid1.Text, Len(Grid1.Text) - 1)
        End If
    End If

End If

End Sub

Private Sub Grid1_LostFocus()
    
    If ColonnaSelezionataInPrecedenza = False Then
        FormPrincipale!Grid1.HighLight = False
    End If

End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then ' solo il tasto sinistro !

' ripristina
    FormPrincipale!Grid1.HighLight = True

    End If

End Sub

Private Sub Grid1_RowColChange()
    
    ImmessaVirgolaPrecedentemente = False

End Sub

Private Sub Grid1_SelChange()

' memorizza coordinate del puntatore !
    Rig = FormPrincipale!Grid1.Row
    Col = FormPrincipale!Grid1.Col
    RigSup = FormPrincipale!Grid1.TopRow
    ColSin = FormPrincipale!Grid1.LeftCol

' nessuna selezione
    If (FormPrincipale!Grid1.SelStartRow = FormPrincipale!Grid1.SelEndRow) And (FormPrincipale!Grid1.SelStartCol = FormPrincipale!Grid1.SelEndCol) Then
    If BarraDegliStrumenti = True Then
    
        CommandApri.Enabled = True
        CommandSalva.Enabled = True
        CommandNuovo.Enabled = True
    
        CommandEsploraPar.Enabled = True
        CommandTorta.Enabled = False
        CommandEsploraNonPar.Enabled = True
        
        CommandApri.BevelWidth = 3
        CommandSalva.BevelWidth = 3
        CommandNuovo.BevelWidth = 3
        
        CommandEsploraPar.BevelWidth = 3
        CommandTorta.BevelWidth = 1
        CommandEsploraNonPar.BevelWidth = 3
    
    If FileSalvatoInPrecedenza = True Then
        If LCase(Right(FormPrincipale.Caption, 5)) = ".mi1]" Then
            CommandStampaColonne1.Enabled = True
            CommandStampaColonne6.Enabled = True
            CommandStampaColonne1.BevelWidth = 3
            CommandStampaColonne6.BevelWidth = 3
        Else
            CommandStampaColonne1.Enabled = False
            CommandStampaColonne6.Enabled = False
            CommandStampaColonne1.BevelWidth = 1
            CommandStampaColonne6.BevelWidth = 1
        End If
    Else
        CommandStampaColonne1.Enabled = False
        CommandStampaColonne6.Enabled = False
        CommandStampaColonne1.BevelWidth = 1
        CommandStampaColonne6.BevelWidth = 1
    End If

    If ColonnaSelezionataInPrecedenza = True Then
        
        CommandNumera.Enabled = False
        CommandOrdinaCrescente.Enabled = False
        CommandOrdinaDecrescente.Enabled = False
        CommandRango.Enabled = False
        CommandPercentili.Enabled = False
        CommandDeviataZeta.Enabled = False
        
        CommandIstogramma.Enabled = False
        CommandCumulativa.Enabled = False
        CommandConfronto.Enabled = False
        CommandAppaiamento.Enabled = False
        CommandDispersione.Enabled = False
        CommandAssociazione.Enabled = False
        CommandAltman.Enabled = False
        CommandSovrapposizione.Enabled = False
        CommandLineaSpezzata.Enabled = False
        CommandBarreVerticali.Enabled = False
        CommandIstogrammiMultipli.Enabled = False
        CommandDistribuzione.Enabled = False
        CommandErroreStandard.Enabled = False
        CommandQuartiliERange.Enabled = False
        
        CommandStatisticheElementari.Enabled = False
        CommandConfrontoTraMedie.Enabled = False
        CommandAnalisiDEllaVarianza.Enabled = False
        CommandRegressioneLineare.Enabled = False
        CommandRegressionePolinomiale.Enabled = False
        CommandRegressioneMultipla.Enabled = False
        CommandStatisticaBayesiana.Enabled = False
        CommandAnalisiDellaSomiglianza.Enabled = False
        CommandTestChiQuadrato.Enabled = False
        CommandStatNon.Enabled = False
        CommandConfrontoTraMediane.Enabled = False
        CommandRegNon.Enabled = False
    

        CommandNumera.BevelWidth = 1
        CommandOrdinaCrescente.BevelWidth = 1
        CommandOrdinaDecrescente.BevelWidth = 1
        CommandRango.BevelWidth = 1
        CommandPercentili.BevelWidth = 1
        CommandDeviataZeta.BevelWidth = 1
        
        CommandIstogramma.BevelWidth = 1
        CommandCumulativa.BevelWidth = 1
        CommandConfronto.BevelWidth = 1
        CommandAppaiamento.BevelWidth = 1
        CommandDispersione.BevelWidth = 1
        CommandAltman.BevelWidth = 1
        CommandAssociazione.BevelWidth = 1
        CommandSovrapposizione.BevelWidth = 1
        CommandLineaSpezzata.BevelWidth = 1
        CommandBarreVerticali.BevelWidth = 1
        CommandIstogrammiMultipli.BevelWidth = 1
        CommandDistribuzione.BevelWidth = 1
        CommandErroreStandard.BevelWidth = 1
        CommandQuartiliERange.BevelWidth = 1
        
        CommandStatisticheElementari.BevelWidth = 1
        CommandConfrontoTraMedie.BevelWidth = 1
        CommandAnalisiDEllaVarianza.BevelWidth = 1
        CommandRegressioneLineare.BevelWidth = 1
        CommandRegressionePolinomiale.BevelWidth = 1
        CommandRegressioneMultipla.BevelWidth = 1
        CommandStatisticaBayesiana.BevelWidth = 1
        CommandAnalisiDellaSomiglianza.BevelWidth = 1
        CommandTestChiQuadrato.BevelWidth = 1
        CommandStatNon.BevelWidth = 1
        CommandConfrontoTraMediane.BevelWidth = 1
        CommandRegNon.BevelWidth = 1
        
        ColonnaSelezionataInPrecedenza = False
    End If ' if ColonnaSelezionataInPrecedenza = True !
    End If ' if BarraDegliStrumenti...
        Exit Sub
    End If

' selezionate una o più colonne
    If FormPrincipale!Grid1.SelStartRow = 1 And FormPrincipale!Grid1.SelEndRow = NUM_RIGHE - 1 Then
        ColonnaSelezionataInPrecedenza = True
    If BarraDegliStrumenti = True Then
        
        CommandApri.Enabled = False
        CommandSalva.Enabled = False
        CommandNuovo.Enabled = False
        CommandStampaColonne1.Enabled = False
        CommandStampaColonne6.Enabled = False
        
        CommandNumera.Enabled = False
        CommandOrdinaCrescente.Enabled = False
        CommandOrdinaDecrescente.Enabled = False
        CommandRango.Enabled = False
        CommandPercentili.Enabled = False
        CommandDeviataZeta.Enabled = False
        
        CommandIstogramma.Enabled = False
        CommandCumulativa.Enabled = False
        CommandConfronto.Enabled = False
        CommandAppaiamento.Enabled = False
        CommandDispersione.Enabled = False
        CommandAltman.Enabled = False
        CommandAssociazione.Enabled = False
        CommandSovrapposizione.Enabled = False
        CommandLineaSpezzata.Enabled = False
        CommandBarreVerticali.Enabled = False
        CommandIstogrammiMultipli.Enabled = False
        CommandDistribuzione.Enabled = False
        CommandErroreStandard.Enabled = False
        CommandQuartiliERange.Enabled = False
        
        CommandStatisticheElementari.Enabled = False
        CommandConfrontoTraMedie.Enabled = False
        CommandAnalisiDEllaVarianza.Enabled = False
        CommandRegressioneLineare.Enabled = False
        CommandRegressionePolinomiale.Enabled = False
        CommandRegressioneMultipla.Enabled = False
        CommandStatisticaBayesiana.Enabled = False
        CommandAnalisiDellaSomiglianza.Enabled = False
        CommandTestChiQuadrato.Enabled = False
        CommandStatNon.Enabled = False
        CommandConfrontoTraMediane.Enabled = False
        CommandRegNon.Enabled = False
    
        CommandEsploraPar.Enabled = False
        CommandTorta.Enabled = False
        CommandEsploraNonPar.Enabled = False

        CommandApri.BevelWidth = 1
        CommandSalva.BevelWidth = 1
        CommandNuovo.BevelWidth = 1
        CommandStampaColonne1.BevelWidth = 1
        CommandStampaColonne6.BevelWidth = 1
        
        CommandNumera.BevelWidth = 1
        CommandOrdinaCrescente.BevelWidth = 1
        CommandOrdinaDecrescente.BevelWidth = 1
        CommandRango.BevelWidth = 1
        CommandPercentili.BevelWidth = 1
        CommandDeviataZeta.BevelWidth = 1
        
        CommandIstogramma.BevelWidth = 1
        CommandCumulativa.BevelWidth = 1
        CommandConfronto.BevelWidth = 1
        CommandAppaiamento.BevelWidth = 1
        CommandDispersione.BevelWidth = 1
        CommandAltman.BevelWidth = 1
        CommandAssociazione.BevelWidth = 1
        CommandSovrapposizione.BevelWidth = 1
        CommandLineaSpezzata.BevelWidth = 1
        CommandBarreVerticali.BevelWidth = 1
        CommandIstogrammiMultipli.BevelWidth = 1
        CommandDistribuzione.BevelWidth = 1
        CommandErroreStandard.BevelWidth = 1
        CommandQuartiliERange.BevelWidth = 1
        
        CommandStatisticheElementari.BevelWidth = 1
        CommandConfrontoTraMedie.BevelWidth = 1
        CommandAnalisiDEllaVarianza.BevelWidth = 1
        CommandRegressioneLineare.BevelWidth = 1
        CommandRegressionePolinomiale.BevelWidth = 1
        CommandRegressioneMultipla.BevelWidth = 1
        CommandStatisticaBayesiana.BevelWidth = 1
        CommandAnalisiDellaSomiglianza.BevelWidth = 1
        CommandTestChiQuadrato.BevelWidth = 1
        CommandStatNon.BevelWidth = 1
        CommandConfrontoTraMediane.BevelWidth = 1
        CommandRegNon.BevelWidth = 1

        CommandEsploraPar.BevelWidth = 1
        CommandTorta.BevelWidth = 1
        CommandEsploraNonPar.BevelWidth = 1

' selezionata una colonna
        If FormPrincipale!Grid1.SelEndCol - FormPrincipale!Grid1.SelStartCol = 0 Then
        If QualcosaNelRitaglio = False Then
            
            CommandNumera.Enabled = True
            CommandOrdinaCrescente.Enabled = True
            CommandOrdinaDecrescente.Enabled = True
            CommandRango.Enabled = True
            CommandPercentili.Enabled = True
            CommandDeviataZeta.Enabled = True
            
            CommandIstogramma.Enabled = True
            CommandCumulativa.Enabled = True
            
            CommandStatisticheElementari.Enabled = True
            CommandAnalisiDellaSomiglianza.Enabled = True
            CommandStatNon.Enabled = True
        
            CommandNumera.BevelWidth = 3
            CommandOrdinaCrescente.BevelWidth = 3
            CommandOrdinaDecrescente.BevelWidth = 3
            CommandRango.BevelWidth = 3
            CommandPercentili.BevelWidth = 3
            CommandDeviataZeta.BevelWidth = 3
            
            CommandIstogramma.BevelWidth = 3
            CommandCumulativa.BevelWidth = 3
            
            CommandStatisticheElementari.BevelWidth = 3
            CommandAnalisiDellaSomiglianza.BevelWidth = 3
            CommandStatNon.BevelWidth = 3
        
        End If
        End If

' selezionate due colonne
        If FormPrincipale!Grid1.SelEndCol - FormPrincipale!Grid1.SelStartCol = 1 Then
            
            CommandConfronto.Enabled = True
            CommandAppaiamento.Enabled = True
            CommandDispersione.Enabled = True
            CommandAltman.Enabled = True
            CommandLineaSpezzata.Enabled = True
            CommandBarreVerticali.Enabled = True
            CommandIstogrammiMultipli.Enabled = True
            CommandDistribuzione.Enabled = True
            CommandErroreStandard.Enabled = True
            CommandQuartiliERange.Enabled = True
            
            CommandConfrontoTraMedie.Enabled = True
            CommandAnalisiDEllaVarianza.Enabled = True
            CommandRegressioneLineare.Enabled = True
            CommandRegressionePolinomiale.Enabled = True
            CommandRegressioneMultipla.Enabled = True
            CommandStatisticaBayesiana.Enabled = True
            CommandAnalisiDellaSomiglianza.Enabled = True
            CommandTestChiQuadrato.Enabled = True
            CommandConfrontoTraMediane.Enabled = True
            CommandRegNon.Enabled = True
        
            CommandConfronto.BevelWidth = 3
            CommandAppaiamento.BevelWidth = 3
            CommandDispersione.BevelWidth = 3
            CommandAltman.BevelWidth = 3
            CommandLineaSpezzata.BevelWidth = 3
            CommandBarreVerticali.BevelWidth = 3
            CommandIstogrammiMultipli.BevelWidth = 3
            CommandDistribuzione.BevelWidth = 3
            CommandErroreStandard.BevelWidth = 3
            CommandQuartiliERange.BevelWidth = 3
            
            CommandConfrontoTraMedie.BevelWidth = 3
            CommandAnalisiDEllaVarianza.BevelWidth = 3
            CommandRegressioneLineare.BevelWidth = 3
            CommandRegressionePolinomiale.BevelWidth = 3
            CommandRegressioneMultipla.BevelWidth = 3
            CommandStatisticaBayesiana.BevelWidth = 3
            CommandAnalisiDellaSomiglianza.BevelWidth = 3
            CommandTestChiQuadrato.BevelWidth = 3
            CommandConfrontoTraMediane.BevelWidth = 3
            CommandRegNon.BevelWidth = 3

        End If

' selezionate tre colonne
        If FormPrincipale!Grid1.SelEndCol - FormPrincipale!Grid1.SelStartCol = 2 Then
            
            CommandAssociazione.Enabled = True
            CommandSovrapposizione.Enabled = True
            CommandLineaSpezzata.Enabled = True
            CommandBarreVerticali.Enabled = True
            CommandIstogrammiMultipli.Enabled = True
            CommandDistribuzione.Enabled = True
            CommandErroreStandard.Enabled = True
            CommandQuartiliERange.Enabled = True
            
            CommandAnalisiDEllaVarianza.Enabled = True
            CommandRegressioneMultipla.Enabled = True
            CommandAnalisiDellaSomiglianza.Enabled = True
            CommandTestChiQuadrato.Enabled = True
        
            CommandAssociazione.BevelWidth = 3
            CommandSovrapposizione.BevelWidth = 3
            CommandLineaSpezzata.BevelWidth = 3
            CommandBarreVerticali.BevelWidth = 3
            CommandIstogrammiMultipli.BevelWidth = 3
            CommandDistribuzione.BevelWidth = 3
            CommandErroreStandard.BevelWidth = 3
            CommandQuartiliERange.BevelWidth = 3
            
            CommandAnalisiDEllaVarianza.BevelWidth = 3
            CommandRegressioneMultipla.BevelWidth = 3
            CommandAnalisiDellaSomiglianza.BevelWidth = 3
            CommandTestChiQuadrato.BevelWidth = 3

        End If

' selezionate quattro colonne
        If FormPrincipale!Grid1.SelEndCol - FormPrincipale!Grid1.SelStartCol = 3 Then
            
            CommandSovrapposizione.Enabled = True
            CommandLineaSpezzata.Enabled = True
            CommandBarreVerticali.Enabled = True
            CommandIstogrammiMultipli.Enabled = True
            CommandDistribuzione.Enabled = True
            CommandErroreStandard.Enabled = True
            CommandQuartiliERange.Enabled = True
            
            CommandAnalisiDEllaVarianza.Enabled = True
            CommandRegressioneMultipla.Enabled = True
            CommandAnalisiDellaSomiglianza.Enabled = True
            CommandTestChiQuadrato.Enabled = True
        
            CommandSovrapposizione.BevelWidth = 3
            CommandLineaSpezzata.BevelWidth = 3
            CommandBarreVerticali.BevelWidth = 3
            CommandIstogrammiMultipli.BevelWidth = 3
            CommandDistribuzione.BevelWidth = 3
            CommandErroreStandard.BevelWidth = 3
            CommandQuartiliERange.BevelWidth = 3
            
            CommandAnalisiDEllaVarianza.BevelWidth = 3
            CommandRegressioneMultipla.BevelWidth = 3
            CommandAnalisiDellaSomiglianza.BevelWidth = 3
            CommandTestChiQuadrato.BevelWidth = 3

        End If

' selezionate cinque colonne
        If FormPrincipale!Grid1.SelEndCol - FormPrincipale!Grid1.SelStartCol = 4 Then
            
            CommandSovrapposizione.Enabled = True
            CommandLineaSpezzata.Enabled = True
            CommandBarreVerticali.Enabled = True
            
            CommandAnalisiDEllaVarianza.Enabled = True
            CommandRegressioneMultipla.Enabled = True
            CommandAnalisiDellaSomiglianza.Enabled = True
            CommandTestChiQuadrato.Enabled = True
        
            CommandSovrapposizione.BevelWidth = 3
            CommandLineaSpezzata.BevelWidth = 3
            CommandBarreVerticali.BevelWidth = 3
            
            CommandAnalisiDEllaVarianza.BevelWidth = 3
            CommandRegressioneMultipla.BevelWidth = 3
            CommandAnalisiDellaSomiglianza.BevelWidth = 3
            CommandTestChiQuadrato.BevelWidth = 3

        End If

' selezionate oltre cinque colonne
        If FormPrincipale!Grid1.SelEndCol - FormPrincipale!Grid1.SelStartCol > 4 Then
            
            CommandAnalisiDEllaVarianza.Enabled = True
            CommandRegressioneMultipla.Enabled = True
            CommandAnalisiDellaSomiglianza.Enabled = True
            CommandTestChiQuadrato.Enabled = True
        
            CommandAnalisiDEllaVarianza.BevelWidth = 3
            CommandRegressioneMultipla.BevelWidth = 3
            CommandAnalisiDellaSomiglianza.BevelWidth = 3
            CommandTestChiQuadrato.BevelWidth = 3

        End If
    End If ' if BarraDegliStrumenti...
        Exit Sub
    End If

' selezionata una riga
    If (FormPrincipale!Grid1.SelStartCol = 1 And FormPrincipale!Grid1.SelEndCol = NUM_COLONNE - 1) And (FormPrincipale!Grid1.SelStartRow = FormPrincipale!Grid1.SelEndRow) Then
    If BarraDegliStrumenti = True Then

        CommandApri.Enabled = False
        CommandSalva.Enabled = False
        CommandNuovo.Enabled = False
        CommandStampaColonne1.Enabled = False
        CommandStampaColonne6.Enabled = False
        
        CommandNumera.Enabled = False
        CommandOrdinaCrescente.Enabled = False
        CommandOrdinaDecrescente.Enabled = False
        CommandRango.Enabled = False
        CommandPercentili.Enabled = False
        CommandDeviataZeta.Enabled = False

        CommandIstogramma.Enabled = False
        CommandCumulativa.Enabled = False
        CommandConfronto.Enabled = False
        CommandAppaiamento.Enabled = False
        CommandDispersione.Enabled = False
        CommandAltman.Enabled = False
        CommandAssociazione.Enabled = False
        CommandSovrapposizione.Enabled = False
        CommandLineaSpezzata.Enabled = False
        CommandBarreVerticali.Enabled = False
        CommandIstogrammiMultipli.Enabled = False
        CommandDistribuzione.Enabled = False
        CommandErroreStandard.Enabled = False
        CommandQuartiliERange.Enabled = False
        
        CommandStatisticheElementari.Enabled = False
        CommandConfrontoTraMedie.Enabled = False
        CommandAnalisiDEllaVarianza.Enabled = False
        CommandRegressioneLineare.Enabled = False
        CommandRegressionePolinomiale.Enabled = False
        CommandRegressioneMultipla.Enabled = False
        CommandStatisticaBayesiana.Enabled = False
        CommandAnalisiDellaSomiglianza.Enabled = False
        CommandTestChiQuadrato.Enabled = False
        CommandStatNon.Enabled = False
        CommandConfrontoTraMediane.Enabled = False
        CommandRegNon.Enabled = False
    
        CommandEsploraPar.Enabled = False
        CommandTorta.Enabled = True
        CommandEsploraNonPar.Enabled = False

        CommandApri.BevelWidth = 1
        CommandSalva.BevelWidth = 1
        CommandNuovo.BevelWidth = 1
        CommandStampaColonne1.BevelWidth = 1
        CommandStampaColonne6.BevelWidth = 1
        
        CommandNumera.BevelWidth = 1
        CommandOrdinaCrescente.BevelWidth = 1
        CommandOrdinaDecrescente.BevelWidth = 1
        CommandRango.BevelWidth = 1
        CommandPercentili.BevelWidth = 1
        CommandDeviataZeta.BevelWidth = 1

        CommandIstogramma.BevelWidth = 1
        CommandCumulativa.BevelWidth = 1
        CommandConfronto.BevelWidth = 1
        CommandAppaiamento.BevelWidth = 1
        CommandDispersione.BevelWidth = 1
        CommandAltman.BevelWidth = 1
        CommandAssociazione.BevelWidth = 1
        CommandSovrapposizione.BevelWidth = 1
        CommandLineaSpezzata.BevelWidth = 1
        CommandBarreVerticali.BevelWidth = 1
        CommandIstogrammiMultipli.BevelWidth = 1
        CommandDistribuzione.BevelWidth = 1
        CommandErroreStandard.BevelWidth = 1
        CommandQuartiliERange.BevelWidth = 1
        
        CommandStatisticheElementari.BevelWidth = 1
        CommandConfrontoTraMedie.BevelWidth = 1
        CommandAnalisiDEllaVarianza.BevelWidth = 1
        CommandRegressioneLineare.BevelWidth = 1
        CommandRegressionePolinomiale.BevelWidth = 1
        CommandRegressioneMultipla.BevelWidth = 1
        CommandStatisticaBayesiana.BevelWidth = 1
        CommandAnalisiDellaSomiglianza.BevelWidth = 1
        CommandTestChiQuadrato.BevelWidth = 1
        CommandStatNon.BevelWidth = 1
        CommandConfrontoTraMediane.BevelWidth = 1
        CommandRegNon.BevelWidth = 1

        CommandEsploraPar.BevelWidth = 1
        CommandTorta.BevelWidth = 3
        CommandEsploraNonPar.BevelWidth = 1

    End If ' if BarraDegliStrumenti...
        Exit Sub
    End If

' selezione non ammessa

    FormPrincipale!Grid1.HighLight = False
    FormPrincipale!Grid1.SelEndCol = FormPrincipale!Grid1.SelStartCol
    FormPrincipale!Grid1.SelEndRow = FormPrincipale!Grid1.SelStartRow
    FormPrincipale!Grid1.Col = FormPrincipale!Grid1.SelStartCol
    FormPrincipale!Grid1.Row = FormPrincipale!Grid1.SelStartRow

End Sub

Private Sub MenuCalcoli_Click()

' se non è stato selezionato nulla
    If (Grid1.SelStartCol = Grid1.SelEndCol) And (Grid1.SelStartRow = Grid1.SelEndRow) Then
    FormPrincipale!Grid1.HighLight = False
    End If

' se è stata selezionata una colonna
    If Grid1.SelStartCol = Grid1.SelEndCol And Grid1.SelStartRow = 1 And Grid1.SelEndRow = NUM_RIGHE - 1 Then
' se è stato tagliato qualcosa disabilita tutto
        If QualcosaNelRitaglio = True Then
            FormPrincipale!CalcoliSomma.Enabled = False
            FormPrincipale!CalcoliSottrai.Enabled = False
            FormPrincipale!CalcoliMoltiplica.Enabled = False
            FormPrincipale!CalcoliDividi.Enabled = False
            FormPrincipale!CalcoliQuadrato.Enabled = False
            FormPrincipale!CalcoliRadice.Enabled = False
            FormPrincipale!CalcoliLogaritmo.Enabled = False
            FormPrincipale!CalcoliCambiaSegno.Enabled = False
            FormPrincipale!CalcoliValoreAssoluto.Enabled = False
            FormPrincipale!CalcoliNumera.Enabled = False
            FormPrincipale!CalcoliRango.Enabled = False
            FormPrincipale!CalcoliPercentili.Enabled = False
            FormPrincipale!CalcoliDeviataZeta.Enabled = False
        Else
            FormPrincipale!CalcoliSomma.Enabled = True
            FormPrincipale!CalcoliSottrai.Enabled = True
            FormPrincipale!CalcoliMoltiplica.Enabled = True
            FormPrincipale!CalcoliDividi.Enabled = True
            FormPrincipale!CalcoliQuadrato.Enabled = True
            FormPrincipale!CalcoliRadice.Enabled = True
            FormPrincipale!CalcoliLogaritmo.Enabled = True
            FormPrincipale!CalcoliCambiaSegno.Enabled = True
            FormPrincipale!CalcoliValoreAssoluto.Enabled = True
            FormPrincipale!CalcoliNumera.Enabled = True
            FormPrincipale!CalcoliRango.Enabled = True
            FormPrincipale!CalcoliPercentili.Enabled = True
            FormPrincipale!CalcoliDeviataZeta.Enabled = True
        End If
' se non è stata selezionata una colonna disabilita tutto
    Else
        FormPrincipale!CalcoliSomma.Enabled = False
        FormPrincipale!CalcoliSottrai.Enabled = False
        FormPrincipale!CalcoliMoltiplica.Enabled = False
        FormPrincipale!CalcoliDividi.Enabled = False
        FormPrincipale!CalcoliQuadrato.Enabled = False
        FormPrincipale!CalcoliRadice.Enabled = False
        FormPrincipale!CalcoliLogaritmo.Enabled = False
        FormPrincipale!CalcoliCambiaSegno.Enabled = False
        FormPrincipale!CalcoliValoreAssoluto.Enabled = False
        FormPrincipale!CalcoliNumera.Enabled = False
        FormPrincipale!CalcoliRango.Enabled = False
        FormPrincipale!CalcoliPercentili.Enabled = False
        FormPrincipale!CalcoliDeviataZeta.Enabled = False
    End If
' se sono state selezionate due colonne
    If Grid1.SelStartCol = (Grid1.SelEndCol - 1) And Grid1.SelStartRow = 1 And Grid1.SelEndRow = NUM_RIGHE - 1 Then
        FormPrincipale!CalcoliDifferenza.Enabled = True
        FormPrincipale!CalcoliMedia.Enabled = True
        FormPrincipale!CalcoliRapporto.Enabled = True
    Else
        FormPrincipale!CalcoliDifferenza.Enabled = False
        FormPrincipale!CalcoliMedia.Enabled = False
        FormPrincipale!CalcoliRapporto.Enabled = False
    End If
End Sub

Private Sub MenuFile_Click()

' se non è stato selezionato nulla
    If (Grid1.SelStartCol = Grid1.SelEndCol) And (Grid1.SelStartRow = Grid1.SelEndRow) Then
    FormPrincipale!Grid1.HighLight = False
    End If

' se è stata selezionata una colonna o una riga disabilita tutto
    If (Grid1.SelStartRow = 1 And Grid1.SelEndRow = NUM_RIGHE - 1) Or (Grid1.SelStartCol = 1 And Grid1.SelEndCol = NUM_COLONNE - 1) Then
        FormPrincipale!FileNuovo.Enabled = False
        FormPrincipale!FileApri.Enabled = False
        FormPrincipale!FileImportaFile.Enabled = False
        FormPrincipale!FileEsporta.Enabled = False
        FormPrincipale!FileSalva.Enabled = False
        FormPrincipale!FileSalvaConNome.Enabled = False
        FormPrincipale!FileCancella.Enabled = False
        FormPrincipale!FileStampa.Enabled = False
        FormPrincipale!FileEsci.Enabled = False
' se non è stata selezionata una colonna abilita tutto...
    Else
        FormPrincipale!FileNuovo.Enabled = True
        FormPrincipale!FileApri.Enabled = True
        FormPrincipale!FileImportaFile.Enabled = True
        FormPrincipale!FileSalvaConNome.Enabled = True
        FormPrincipale!FileEsci.Enabled = True
' ...tranne Salva, Cancella ed Esporta
    If FileSalvatoInPrecedenza = True Then
        If LCase(Right(FormPrincipale.Caption, 5)) = ".mi1]" Then
            FileEsporta.Enabled = True
            FileSalva.Enabled = True
            FileCancella.Enabled = True
            FileStampa.Enabled = True
        Else
            FileEsporta.Enabled = False
            FileSalva.Enabled = False
            FileCancella.Enabled = False
            FileStampa.Enabled = False
        End If
    Else
        FileEsporta.Enabled = False
        FileSalva.Enabled = False
        FileCancella.Enabled = False
        FileStampa.Enabled = False
    End If

    End If

End Sub

Private Sub MenuGrafica_Click()

' se non è stato selezionato nulla
    If (Grid1.SelStartCol = Grid1.SelEndCol) And (Grid1.SelStartRow = Grid1.SelEndRow) Then
    FormPrincipale!Grid1.HighLight = False
    End If

' se non è stato selezionato nulla disabilita tutto
'    If (Grid1.SelStartCol = Grid1.SelEndCol) And (Grid1.SelStartRow = Grid1.SelEndRow) Then
        GraficaTorta.Enabled = False
        GraficaIstogramma.Enabled = False
        GraficaCumulativa.Enabled = False
        GraficaPrimaDopo.Enabled = False
        GraficaDispersione.Enabled = False
        GraficaAltman.Enabled = False
        GraficaAssociazione.Enabled = False
        GraficaSovrapposizione.Enabled = False
        GraficaErroreStandard.Enabled = False
        GraficaQuartili.Enabled = False
        GraficaDistribuzione.Enabled = False
        GraficaLineaSpezzata.Enabled = False
        GraficaBarreVerticali.Enabled = False
        GraficaConfronto.Enabled = False
        GraficaIstogrammiMultipli.Enabled = False
'    End If

' se è stata selezionata una riga
    If FormPrincipale!Grid1.SelStartCol <> FormPrincipale!Grid1.SelEndCol Then
    If FormPrincipale!Grid1.SelStartRow = FormPrincipale!Grid1.SelEndRow Then
        GraficaTorta.Enabled = True
    End If
    End If
    
' se è stata selezionata una colonna
    If (Grid1.SelStartCol = Grid1.SelEndCol) And (Grid1.SelStartRow = 1 And Grid1.SelEndRow = NUM_RIGHE - 1) Then
' se è stato tagliato qualcosa disabilita
        If QualcosaNelRitaglio = True Then
            GraficaIstogramma.Enabled = False
            GraficaCumulativa.Enabled = False
            GraficaPrimaDopo.Enabled = False
            GraficaDispersione.Enabled = False
            GraficaAltman.Enabled = False
            GraficaAssociazione.Enabled = False
            GraficaSovrapposizione.Enabled = False
            GraficaErroreStandard.Enabled = False
            GraficaQuartili.Enabled = False
            GraficaDistribuzione.Enabled = False
            GraficaLineaSpezzata.Enabled = False
            GraficaBarreVerticali.Enabled = False
            GraficaConfronto.Enabled = False
            GraficaIstogrammiMultipli.Enabled = False
        Else
            GraficaIstogramma.Enabled = True
            GraficaCumulativa.Enabled = True
            GraficaPrimaDopo.Enabled = False
            GraficaDispersione.Enabled = False
            GraficaAltman.Enabled = False
            GraficaAssociazione.Enabled = False
            GraficaSovrapposizione.Enabled = False
            GraficaErroreStandard.Enabled = False
            GraficaQuartili.Enabled = False
            GraficaDistribuzione.Enabled = False
            GraficaLineaSpezzata.Enabled = False
            GraficaBarreVerticali.Enabled = False
            GraficaConfronto.Enabled = False
            GraficaIstogrammiMultipli.Enabled = False
        End If
    End If

' se sono state selezionate due colonne
    If ((Grid1.SelEndCol - Grid1.SelStartCol) = 1) And (Grid1.SelStartRow = 1 And Grid1.SelEndRow = NUM_RIGHE - 1) Then
        GraficaIstogramma.Enabled = False
        GraficaCumulativa.Enabled = False
        GraficaPrimaDopo.Enabled = True
        GraficaDispersione.Enabled = True
        GraficaAltman.Enabled = True
        GraficaAssociazione.Enabled = False
        GraficaSovrapposizione.Enabled = False
        GraficaErroreStandard.Enabled = True
        GraficaQuartili.Enabled = True
        GraficaDistribuzione.Enabled = True
        GraficaLineaSpezzata.Enabled = True
        GraficaBarreVerticali.Enabled = True
        GraficaConfronto.Enabled = True
        GraficaIstogrammiMultipli.Enabled = True
    End If

' se sono state selezionate tre colonne
    If ((Grid1.SelEndCol - Grid1.SelStartCol) = 2) And (Grid1.SelStartRow = 1 And Grid1.SelEndRow = NUM_RIGHE - 1) Then
        GraficaIstogramma.Enabled = False
        GraficaCumulativa.Enabled = False
        GraficaPrimaDopo.Enabled = False
        GraficaDispersione.Enabled = False
        GraficaAltman.Enabled = False
        GraficaAssociazione.Enabled = True
        GraficaSovrapposizione.Enabled = True
        GraficaErroreStandard.Enabled = True
        GraficaQuartili.Enabled = True
        GraficaDistribuzione.Enabled = True
        GraficaLineaSpezzata.Enabled = True
        GraficaBarreVerticali.Enabled = True
        GraficaConfronto.Enabled = False
        GraficaIstogrammiMultipli.Enabled = True
    End If

' se sono state selezionate quattro colonne
    If ((Grid1.SelEndCol - Grid1.SelStartCol) = 3) And (Grid1.SelStartRow = 1 And Grid1.SelEndRow = NUM_RIGHE - 1) Then
        GraficaIstogramma.Enabled = False
        GraficaCumulativa.Enabled = False
        GraficaPrimaDopo.Enabled = False
        GraficaDispersione.Enabled = False
        GraficaAltman.Enabled = False
        GraficaAssociazione.Enabled = False
        GraficaSovrapposizione.Enabled = True
        GraficaErroreStandard.Enabled = True
        GraficaQuartili.Enabled = True
        GraficaDistribuzione.Enabled = True
        GraficaLineaSpezzata.Enabled = True
        GraficaBarreVerticali.Enabled = True
        GraficaConfronto.Enabled = False
        GraficaIstogrammiMultipli.Enabled = True
    End If

' se sono state selezionate cinque colonne
    If ((Grid1.SelEndCol - Grid1.SelStartCol) = 4) And (Grid1.SelStartRow = 1 And Grid1.SelEndRow = NUM_RIGHE - 1) Then
        GraficaIstogramma.Enabled = False
        GraficaCumulativa.Enabled = False
        GraficaPrimaDopo.Enabled = False
        GraficaDispersione.Enabled = False
        GraficaAltman.Enabled = False
        GraficaAssociazione.Enabled = False
        GraficaSovrapposizione.Enabled = True
        GraficaErroreStandard.Enabled = False
        GraficaQuartili.Enabled = False
        GraficaDistribuzione.Enabled = False
        GraficaLineaSpezzata.Enabled = True
        GraficaBarreVerticali.Enabled = True
        GraficaConfronto.Enabled = False
        GraficaIstogrammiMultipli.Enabled = False
    End If

' se sono state selezionate più di cinque colonne
    If ((Grid1.SelEndCol - Grid1.SelStartCol) = 5) And (Grid1.SelStartRow = 1 And Grid1.SelEndRow = NUM_RIGHE - 1) Then
        GraficaIstogramma.Enabled = False
        GraficaCumulativa.Enabled = False
        GraficaPrimaDopo.Enabled = False
        GraficaDispersione.Enabled = False
        GraficaAltman.Enabled = False
        GraficaAssociazione.Enabled = False
        GraficaSovrapposizione.Enabled = False
        GraficaErroreStandard.Enabled = False
        GraficaQuartili.Enabled = False
        GraficaDistribuzione.Enabled = False
        GraficaLineaSpezzata.Enabled = False
        GraficaBarreVerticali.Enabled = False
        GraficaConfronto.Enabled = False
        GraficaIstogrammiMultipli.Enabled = False
    End If

End Sub

Private Sub MenuModifica_Click()

' se non è stato selezionato nulla
    If (Grid1.SelStartCol = Grid1.SelEndCol) And (Grid1.SelStartRow = Grid1.SelEndRow) Then
    FormPrincipale!Grid1.HighLight = False
    End If

' se è stata selezionata una colonna
    If Grid1.SelStartCol = Grid1.SelEndCol And Grid1.SelStartRow = 1 And Grid1.SelEndRow = NUM_RIGHE - 1 Then
' se è stato tagliato qualcosa abilita solo Incolla
        If QualcosaNelRitaglio = True Then
            FormPrincipale!ModificaTagliaColonna = False
            FormPrincipale!ModificaCopiaColonna = False
            FormPrincipale!ModificaIncollaColonna.Enabled = True
            FormPrincipale!ModificaAnnullaIncolla.Enabled = True
            FormPrincipale!ModificaCella = False
            FormPrincipale!ModificaColonna = False
            FormPrincipale!ModificaRiga = False
            FormPrincipale!ModificaUnisciColonne = False
            FormPrincipale!ModificaScambiaColonne = False
            FormPrincipale!ModificaOrdina = False
        Else
            FormPrincipale!ModificaTagliaColonna = True
            FormPrincipale!ModificaCopiaColonna = True
            FormPrincipale!ModificaIncollaColonna.Enabled = False
            FormPrincipale!ModificaAnnullaIncolla.Enabled = False
            FormPrincipale!ModificaCella = False
            FormPrincipale!ModificaColonna = True
            FormPrincipale!ModificaRiga = False
            FormPrincipale!ModificaUnisciColonne = False
            FormPrincipale!ModificaScambiaColonne = False
            FormPrincipale!ModificaOrdina = True
        End If
' se non è stata selezionata una colonna disabilita tutto
    Else
            FormPrincipale!ModificaTagliaColonna = False
            FormPrincipale!ModificaCopiaColonna = False
            FormPrincipale!ModificaIncollaColonna.Enabled = False
            FormPrincipale!ModificaAnnullaIncolla.Enabled = False
            FormPrincipale!ModificaCella = False
            FormPrincipale!ModificaColonna = False
            FormPrincipale!ModificaRiga = False
            FormPrincipale!ModificaUnisciColonne = False
            FormPrincipale!ModificaScambiaColonne = False
            FormPrincipale!ModificaOrdina = False
' tranne che se non è state selezionato niente !
            If Grid1.SelEndRow = Grid1.SelStartRow And Grid1.SelEndCol = Grid1.SelStartCol Then
                FormPrincipale!ModificaCella = True
            Else
                FormPrincipale!ModificaCella = False
            End If
' e tranne che se è stata selezionata una riga !
            If Grid1.SelEndRow = Grid1.SelStartRow And Grid1.SelStartCol = 1 And Grid1.SelEndCol = (NUM_COLONNE - 1) Then
                FormPrincipale!ModificaRiga = True
            Else
                FormPrincipale!ModificaRiga = False
            End If
    End If
' e se sono state selezionate due colonne ?
    If Grid1.SelStartCol = (Grid1.SelEndCol - 1) And Grid1.SelStartRow = 1 And Grid1.SelEndRow = NUM_RIGHE - 1 Then
        FormPrincipale!ModificaUnisciColonne = True
        FormPrincipale!ModificaScambiaColonne = True
    Else
        FormPrincipale!ModificaUnisciColonne = False
        FormPrincipale!ModificaScambiaColonne = False
    End If

End Sub

Private Sub MenuOpzioni_Click()

' se non è stato selezionato nulla
    If (Grid1.SelStartCol = Grid1.SelEndCol) And (Grid1.SelStartRow = Grid1.SelEndRow) Then
    FormPrincipale!Grid1.HighLight = False
    End If

' se è stata selezionata una colonna
    If Grid1.SelStartCol = Grid1.SelEndCol And Grid1.SelStartRow = 1 And Grid1.SelEndRow = NUM_RIGHE - 1 Then
' blocca tutto per permettere solo di incollare
        If QualcosaNelRitaglio = True Then
            FormPrincipale!OpzioniNomeDellaVariabile.Enabled = False
            FormPrincipale!OpzioniNumeroDiDecimali.Enabled = False
            FormPrincipale!OpzioniGriglia.Enabled = False
            FormPrincipale!OpzioniBarraDegliStrumenti.Enabled = False
            FormPrincipale!OpzioniRipristinaAltezzaRighe.Enabled = False
            FormPrincipale!OpzioniRipristinaLarghezzaColonne.Enabled = False
        Else
            FormPrincipale!OpzioniNomeDellaVariabile.Enabled = True
            FormPrincipale!OpzioniNumeroDiDecimali.Enabled = True
            FormPrincipale!OpzioniGriglia.Enabled = False
            FormPrincipale!OpzioniBarraDegliStrumenti.Enabled = False
            FormPrincipale!OpzioniRipristinaAltezzaRighe.Enabled = False
            FormPrincipale!OpzioniRipristinaLarghezzaColonne.Enabled = False
        End If
' se non è stata selezionata una colonna...
    Else
        FormPrincipale!OpzioniNomeDellaVariabile.Enabled = False
        FormPrincipale!OpzioniNumeroDiDecimali.Enabled = False
        FormPrincipale!OpzioniGriglia.Enabled = True
        FormPrincipale!OpzioniBarraDegliStrumenti.Enabled = True
        FormPrincipale!OpzioniRipristinaAltezzaRighe.Enabled = True
        FormPrincipale!OpzioniRipristinaLarghezzaColonne.Enabled = True
' ...ma se sono state selezionate più colonne...
            If Grid1.SelStartRow = 1 And Grid1.SelEndRow = NUM_RIGHE - 1 Then
                FormPrincipale!OpzioniGriglia.Enabled = False
                FormPrincipale!OpzioniBarraDegliStrumenti.Enabled = False
                FormPrincipale!OpzioniRipristinaAltezzaRighe.Enabled = False
                FormPrincipale!OpzioniRipristinaLarghezzaColonne.Enabled = False
            End If
' ...e se è stata selezionata una riga !!!!
            If Grid1.SelStartCol = 1 And Grid1.SelEndCol = NUM_COLONNE - 1 Then
                FormPrincipale!OpzioniGriglia.Enabled = False
                FormPrincipale!OpzioniBarraDegliStrumenti.Enabled = False
                FormPrincipale!OpzioniRipristinaAltezzaRighe.Enabled = False
                FormPrincipale!OpzioniRipristinaLarghezzaColonne.Enabled = False
            End If
    End If

End Sub

Private Sub MenuStatistica_Click()

' se non è stato selezionato nulla
    If (Grid1.SelStartCol = Grid1.SelEndCol) And (Grid1.SelStartRow = Grid1.SelEndRow) Then
    FormPrincipale!Grid1.HighLight = False
    End If

' se non è stato selezionato nulla disabilita tutto
'    If (Grid1.SelStartCol = Grid1.SelEndCol) And (Grid1.SelStartRow = Grid1.SelEndRow) Then
        StatisticaEsploraPar.Enabled = True
        StatisticaStatisticheElementari.Enabled = False
        StatisticaT.Enabled = False
        StatisticaRegressioneLineare.Enabled = False
        StatisticaRegressionePolinomiale.Enabled = False
        StatisticaRegressioneMultipla.Enabled = False
        StatisticaAnova.Enabled = False
        StatisticaBayesiana = False
        StatisticaNonPArametrica.Enabled = True
            StatisticaEsploraNonPar.Enabled = True
            StatisticaNonParametricaCumulativa.Enabled = False
            StatisticaNonParametricaReglin.Enabled = False
            StatisticaNonParametricaW.Enabled = False
            StatisticaNonParametricaChi.Enabled = False
            StatisticaCluster = False
'    End If

' se è stata selezionata una colonna
    If (Grid1.SelStartCol = Grid1.SelEndCol) And (Grid1.SelStartRow = 1 And Grid1.SelEndRow = NUM_RIGHE - 1) Then
' se è stato tagliato qualcosa disabilita
        If QualcosaNelRitaglio = True Then
            StatisticaEsploraPar.Enabled = False
            StatisticaStatisticheElementari.Enabled = False
            StatisticaT.Enabled = False
            StatisticaRegressioneLineare.Enabled = False
            StatisticaRegressionePolinomiale.Enabled = False
            StatisticaRegressioneMultipla.Enabled = False
            StatisticaAnova.Enabled = False
            StatisticaBayesiana = False
            StatisticaCluster = False
            StatisticaNonPArametrica.Enabled = False
                StatisticaEsploraNonPar.Enabled = False
                StatisticaNonParametricaCumulativa.Enabled = False
                StatisticaNonParametricaReglin.Enabled = False
                StatisticaNonParametricaW.Enabled = False
                StatisticaNonParametricaChi.Enabled = False
                StatisticaCluster = False
        Else
            StatisticaEsploraPar.Enabled = False
            StatisticaStatisticheElementari.Enabled = True
            StatisticaT.Enabled = False
            StatisticaRegressioneLineare.Enabled = False
            StatisticaRegressionePolinomiale.Enabled = False
            StatisticaRegressioneMultipla.Enabled = False
            StatisticaAnova.Enabled = False
            StatisticaBayesiana = False
            StatisticaCluster = True
            StatisticaNonPArametrica.Enabled = True
                StatisticaEsploraNonPar.Enabled = False
                StatisticaNonParametricaCumulativa.Enabled = True
                StatisticaNonParametricaReglin.Enabled = False
                StatisticaNonParametricaW.Enabled = False
                StatisticaNonParametricaChi.Enabled = False
                StatisticaCluster = True
        End If
    End If

' se sono state selezionate due colonne
    If ((Grid1.SelEndCol - Grid1.SelStartCol) = 1) And (Grid1.SelStartRow = 1 And Grid1.SelEndRow = NUM_RIGHE - 1) Then
        StatisticaEsploraPar.Enabled = False
        StatisticaStatisticheElementari.Enabled = False
        StatisticaT.Enabled = True
        StatisticaRegressioneLineare.Enabled = True
        StatisticaRegressionePolinomiale.Enabled = True
        StatisticaRegressioneMultipla.Enabled = True
        StatisticaAnova.Enabled = True
        StatisticaBayesiana = True
        StatisticaCluster = True
        StatisticaNonPArametrica.Enabled = True
            StatisticaEsploraNonPar.Enabled = False
            StatisticaNonParametricaCumulativa.Enabled = False
            StatisticaNonParametricaReglin.Enabled = True
            StatisticaNonParametricaW.Enabled = True
            StatisticaNonParametricaChi.Enabled = True
            StatisticaCluster = True
    End If

' se sono state selezionate più di due colonne
    If ((Grid1.SelEndCol - Grid1.SelStartCol) > 1) And (Grid1.SelStartRow = 1 And Grid1.SelEndRow = NUM_RIGHE - 1) Then
        StatisticaEsploraPar.Enabled = False
        StatisticaStatisticheElementari.Enabled = False
        StatisticaT.Enabled = False
        StatisticaRegressioneLineare.Enabled = False
        StatisticaRegressionePolinomiale.Enabled = False
        StatisticaRegressioneMultipla.Enabled = True
        StatisticaAnova.Enabled = True
        StatisticaBayesiana = False
        StatisticaCluster = True
        StatisticaNonPArametrica.Enabled = True
            StatisticaEsploraNonPar.Enabled = False
            StatisticaNonParametricaCumulativa.Enabled = False
            StatisticaNonParametricaReglin.Enabled = False
            StatisticaNonParametricaW.Enabled = False
            StatisticaNonParametricaChi.Enabled = True
            StatisticaCluster = True
    End If

' selezionata una riga
    If (FormPrincipale!Grid1.SelStartCol = 1 And FormPrincipale!Grid1.SelEndCol = NUM_COLONNE - 1) And (FormPrincipale!Grid1.SelStartRow = FormPrincipale!Grid1.SelEndRow) Then
        StatisticaEsploraPar.Enabled = False
        StatisticaNonPArametrica.Enabled = False
    End If

End Sub

Private Sub ModificaAnnullaIncolla_Click()

    ReDim Ritaglio(0)
    QualcosaNelRitaglio = False
    RipristinaSchermo
    DeselezionaColonna
    FormPrincipale!Grid1.SetFocus

End Sub

Private Sub ModificaCopiaColonna_Click()
    Dim Riga
    
    FormPrincipale!Grid1.Row = 0
    NomeDellaVariabile = FormPrincipale!Grid1.Text
    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Copia colonna " & NomeDellaVariabile
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    ReDim Ritaglio(0 To NUM_RIGHE - 1)

    For Riga = 0 To NUM_RIGHE - 1

        FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        FormPrincipale!Grid1.Row = Riga
        Ritaglio(Riga) = FormPrincipale!Grid1.Text
    Next Riga

    QualcosaNelRitaglio = True
    Unload FormGauge
    RipristinaSchermo
    DeselezionaColonna
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

Private Sub ModificaEliminaCella_Click()
    Dim Riga, Risposta, Dummy
    Dim RigaDellaCella
    
    FormPrincipale!Grid1.SetFocus
    FormPrincipale!Grid1.HighLight = True

    If Trim(Grid1.Text) <> "" Then
        Risposta = MsgBox("Il dato della cella sarà perso." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Elimina cella (dato: " & Grid1.Text & ")")
    Else
        Risposta = MsgBox("Il dato della cella sarà perso." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Elimina cella")
    End If

        If Risposta = IDNO Then
            RipristinaSchermo
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If

    RigaDellaCella = FormPrincipale!Grid1.Row
    Screen.MousePointer = 11
    FormGauge.Show
    If Trim(Grid1.Text) <> "" Then
        FormGauge.Caption = "Elimina cella (dato: " & Grid1.Text & ")"
    Else
        FormGauge.Caption = "Elimina cella"
    End If
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    For Riga = RigaDellaCella To (NUM_RIGHE - 1) - 1

        FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        FormPrincipale!Grid1.Row = Riga + 1
        Dummy = FormPrincipale!Grid1.Text
        FormPrincipale!Grid1.Row = Riga
        FormPrincipale!Grid1.Text = Dummy

    Next Riga
        Riga = NUM_RIGHE - 1
        FormPrincipale!Grid1.Row = Riga
        FormPrincipale!Grid1.Text = ""
    
    FormGauge.Label1.Caption = "100%"
    FormGauge!Gauge1.Value = 100
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    Unload FormGauge
    RipristinaSchermo
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

Private Sub ModificaEliminaColonna_Click()
    Dim Riga, Colonna, Risposta, I, Dummy
    
    FormPrincipale!Grid1.Row = 0
    NomeDellaVariabile = FormPrincipale!Grid1.Text
    Risposta = MsgBox("I dati della colonna saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Elimina colonna " & NomeDellaVariabile)
        If Risposta = IDNO Then
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
    
    Screen.MousePointer = 11
    Colonna = FormPrincipale!Grid1.SelStartCol
    FormGauge.Show
    FormGauge.Caption = "Elimina colonna " & NomeDellaVariabile
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

' il caso generale delle prime 9 colonne
    If Colonna < NUM_COLONNE - 1 Then
    
    For Riga = 0 To NUM_RIGHE - 1

        FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    For I = Colonna + 1 To NUM_COLONNE - 1
        FormPrincipale!Grid1.Row = Riga
        FormPrincipale!Grid1.Col = I
        Dummy = FormPrincipale!Grid1.Text
        FormPrincipale!Grid1.Text = ""
        FormPrincipale!Grid1.Col = I - 1
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
    RipristinaSchermo
    DeselezionaColonna
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

Private Sub ModificaEliminaRiga_Click()
    Dim Riga, Colonna, Risposta, Dummy
    Dim RigaDellaCella
    
    Risposta = MsgBox("I dati della riga saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Elimina riga " & Grid1.Row)
        If Risposta = IDNO Then
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If

    RigaDellaCella = FormPrincipale!Grid1.Row
    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Elimina riga " & RigaDellaCella
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    For Riga = RigaDellaCella To (NUM_RIGHE - 1) - 1

        FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        For Colonna = 1 To NUM_COLONNE - 1
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Row = Riga + 1
            Dummy = FormPrincipale!Grid1.Text
            FormPrincipale!Grid1.Row = Riga
            FormPrincipale!Grid1.Text = Dummy
        Next Colonna

    Next Riga
        Riga = NUM_RIGHE - 1
        FormPrincipale!Grid1.Row = Riga
        For Colonna = 1 To NUM_COLONNE - 1
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = ""
        Next Colonna
    
    FormGauge.Label1.Caption = "100%"
    FormGauge!Gauge1.Value = 100
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    Unload FormGauge
    RipristinaSchermo
    DeselezionaColonna
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

Private Sub ModificaIncollaColonna_Click()
    Dim Riga
    
    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Incolla colonna " & NomeDellaVariabile
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    For Riga = 0 To NUM_RIGHE - 1

        FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        FormPrincipale!Grid1.Row = Riga
        FormPrincipale!Grid1.Text = Ritaglio(Riga)
    Next Riga

    ReDim Ritaglio(0)
    QualcosaNelRitaglio = False
    Unload FormGauge
    RipristinaSchermo
    DeselezionaColonna
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

Private Sub ModificaInserisciCella_Click()
    Dim Riga, Risposta, Dummy
    Dim RigaDellaCella
    
    FormPrincipale!Grid1.SetFocus
    FormPrincipale!Grid1.HighLight = True
    
    Risposta = MsgBox("Il dato della riga " & NUM_RIGHE - 1 & " sarà perso." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Inserisci cella in posizione c" & FormPrincipale!Grid1.Col & "r" & FormPrincipale!Grid1.Row)
        If Risposta = IDNO Then
            RipristinaSchermo
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If

    RigaDellaCella = FormPrincipale!Grid1.Row
    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Inserisci cella in posizione c" & FormPrincipale!Grid1.Col & "r" & FormPrincipale!Grid1.Row
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    For Riga = NUM_RIGHE - 1 To (RigaDellaCella + 1) Step -1

' gauge modificato per tornare all'indietro !!!!
        FormGauge.Label1.Caption = Int(100 * (RigaDellaCella + NUM_RIGHE - Riga + 1) / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * (RigaDellaCella + NUM_RIGHE - Riga + 1) / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        FormPrincipale!Grid1.Row = Riga - 1
        Dummy = FormPrincipale!Grid1.Text
        FormPrincipale!Grid1.Row = Riga
        FormPrincipale!Grid1.Text = Dummy

    Next Riga
        FormPrincipale!Grid1.Row = RigaDellaCella
        FormPrincipale!Grid1.Text = ""

    FormGauge.Label1.Caption = "100%"
    FormGauge!Gauge1.Value = 100
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    Unload FormGauge
    RipristinaSchermo
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

Private Sub ModificaInserisciColonna_Click()

    FormGauge!Label6.Caption = "[1/1]"
    InserisciColonna

    RipristinaSchermo
    DeselezionaColonna
    FormPrincipale!Grid1.SetFocus

End Sub

Private Sub ModificaInserisciRiga_Click()
    Dim Riga, Colonna, Risposta, Dummy
    Dim RigaDellaCella
    
    Risposta = MsgBox("I dati della riga " & NUM_RIGHE - 1 & " saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Inserisci riga in posizione " & Grid1.Row)
        If Risposta = IDNO Then
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
           Exit Sub
        End If

    RigaDellaCella = FormPrincipale!Grid1.Row
    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Inserisci riga in posizione " & RigaDellaCella
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    For Riga = NUM_RIGHE - 1 To (RigaDellaCella + 1) Step -1
' gauge modificato per tornare all'indietro !!!!

        FormGauge.Label1.Caption = Int(100 * (RigaDellaCella + NUM_RIGHE - Riga + 1) / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * (RigaDellaCella + NUM_RIGHE - Riga + 1) / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        For Colonna = 1 To NUM_COLONNE - 1
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Row = Riga - 1
            Dummy = FormPrincipale!Grid1.Text
            FormPrincipale!Grid1.Row = Riga
            FormPrincipale!Grid1.Text = Dummy
        Next Colonna

    Next Riga
        Riga = RigaDellaCella
        FormPrincipale!Grid1.Row = Riga
        For Colonna = 1 To NUM_COLONNE - 1
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = ""
        Next Colonna

    FormGauge.Label1.Caption = "100%"
    FormGauge!Gauge1.Value = 100
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
    Unload FormGauge
    RipristinaSchermo
    DeselezionaColonna
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

Private Sub ModificaOrdinaCrescente_Click()
    Dim Riga, Colonna, ColonnaDati
    Dim DatoRigaJ, DatoRigaL, Dummy
    Dim NumeroRigaJ As Double
    Dim NumeroRigaL As Double
    Dim Z1, D, I, J, L, X
    Dim NomeDellaVariabile, Risposta
    Dim ValoreInizialeD
    
    ColonnaDati = Grid1.SelStartCol
    
    FormPrincipale!Grid1.Row = 0
    NomeDellaVariabile = FormPrincipale!Grid1.Text
    Risposta = MsgBox("I dati verranno ordinati in ordine crescente." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONINFORMATION + MB_DEFBUTTON2, "Ordinamento in base a " & NomeDellaVariabile)
        If Risposta = IDNO Then
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
    
    Screen.MousePointer = 11
        FormGauge.Show
        FormGauge.Caption = "Ordinamento in base a " & NomeDellaVariabile
        FormGauge!Label6.Caption = "[1/1]"
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    Z1 = 0
    For Riga = 1 To NUM_RIGHE - 1
        Grid1.Row = Riga
        If Grid1.Text <> "" Then
            Z1 = Z1 + 1
        Else
            Exit For
        End If
    Next Riga

    ValoreInizialeD = 0

' inizio ordinamento dati
c330:
    D = 1
c340:
    D = 2 * D: If D <= Z1 Then GoTo c340
c350:
    D = Int((D - 1) / 2): If D = 0 Then GoTo c420
c360:
    For I = 1 To Z1 - D
c370:
        J = I
c380:
        L = J + D
        Grid1.Row = J
        NumeroRigaJ = CDbl(Grid1.Text) ' Format(Grid1.Text, "#####0.0####;-#####0.0####")
        Grid1.Row = L
        NumeroRigaL = CDbl(Grid1.Text) ' Format(Grid1.Text, "#####0.0####;-#####0.0####")
        If NumeroRigaJ <= NumeroRigaL Then GoTo c400
                    
c390:
        For Colonna = 1 To NUM_COLONNE - 1
            Grid1.Col = Colonna
            Grid1.Row = J
            DatoRigaJ = Grid1.Text
            Grid1.Row = L
            DatoRigaL = Grid1.Text
            X = DatoRigaJ: DatoRigaJ = DatoRigaL: DatoRigaL = X
            Grid1.Row = J
            Grid1.Text = DatoRigaJ
            Grid1.Row = L
            Grid1.Text = DatoRigaL
        Next Colonna

            If I > ValoreInizialeD And 100 * I / Z1 >= 1 Then
                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If
            Grid1.Col = ColonnaDati
            J = J - D: If J > 0 Then GoTo c380
c400:
    Next I
c410:
    ValoreInizialeD = Z1 - D
    GoTo c350
c420:
' fine ordinamento dati

    FormGauge.Label1.Caption = "100%"
    FormGauge!Gauge1.Value = 100
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    Unload FormGauge

    RipristinaSchermo
    DeselezionaColonna
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

Private Sub ModificaOrdinaDecrescente_Click()
    Dim Riga, Colonna, ColonnaDati
    Dim DatoRigaJ, DatoRigaL, Dummy
    Dim NumeroRigaJ As Double
    Dim NumeroRigaL As Double
    Dim Z1, D, I, J, L, X
    Dim NomeDellaVariabile, Risposta
    Dim ValoreInizialeD
    
    ColonnaDati = Grid1.SelStartCol
    
    FormPrincipale!Grid1.Row = 0
    NomeDellaVariabile = FormPrincipale!Grid1.Text
    Risposta = MsgBox("I dati verranno ordinati in ordine decrescente." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONINFORMATION + MB_DEFBUTTON2, "Ordinamento in base a " & NomeDellaVariabile)
        If Risposta = IDNO Then
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
    
    Screen.MousePointer = 11
        FormGauge.Show
        FormGauge.Caption = "Ordinamento in base a " & NomeDellaVariabile
        FormGauge!Label6.Caption = "[1/1]"
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    Z1 = 0
    For Riga = 1 To NUM_RIGHE - 1
        Grid1.Row = Riga
        If Grid1.Text <> "" Then
            Z1 = Z1 + 1
        Else
            Exit For
        End If
    Next Riga

    ValoreInizialeD = 0

' inizio ordinamento dati
d330:
    D = 1
d340:
    D = 2 * D: If D <= Z1 Then GoTo d340
d350:
    D = Int((D - 1) / 2): If D = 0 Then GoTo d420
d360:
    For I = 1 To Z1 - D
d370:
        J = I
d380:
        L = J + D
        Grid1.Row = J
        NumeroRigaJ = CDbl(Grid1.Text) ' Format(Grid1.Text, "#####0.0####;-#####0.0####")
        Grid1.Row = L
        NumeroRigaL = CDbl(Grid1.Text) ' Format(Grid1.Text, "#####0.0####;-#####0.0####")
        If NumeroRigaJ >= NumeroRigaL Then GoTo d400
                    
d390:
        For Colonna = 1 To NUM_COLONNE - 1
            Grid1.Col = Colonna
            Grid1.Row = J
            DatoRigaJ = Grid1.Text
            Grid1.Row = L
            DatoRigaL = Grid1.Text
            X = DatoRigaJ: DatoRigaJ = DatoRigaL: DatoRigaL = X
            Grid1.Row = J
            Grid1.Text = DatoRigaJ
            Grid1.Row = L
            Grid1.Text = DatoRigaL
        Next Colonna

            If I > ValoreInizialeD And 100 * I / Z1 >= 1 Then
                FormGauge.Label1.Caption = Int(100 * I / Z1) & "%"
                FormGauge!Gauge1.Value = Int(100 * I / Z1)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            End If
            Grid1.Col = ColonnaDati
            J = J - D: If J > 0 Then GoTo d380
d400:
    Next I
d410:
    ValoreInizialeD = Z1 - D
    GoTo d350
d420:
' fine ordinamento dati

    FormGauge.Label1.Caption = "100%"
    FormGauge!Gauge1.Value = 100
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    Unload FormGauge

    RipristinaSchermo
    DeselezionaColonna
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

Private Sub ModificaRipristinaAltezzaRighe_Click()
    Dim Riga

    For Riga = 0 To NUM_RIGHE - 1
        FormPrincipale!Grid1.RowHeight(Riga) = 255
    Next Riga

End Sub

Private Sub ModificaRipristinaLarghezzaColonne_Click()
    Dim Colonna
    
    FormPrincipale!Grid1.ColWidth(0) = 574
    FormPrincipale!Grid1.Row = 0
        For Colonna = 1 To NUM_COLONNE - 1
            FormPrincipale!Grid1.ColWidth(Colonna) = LARGHEZZA_COLONNE
        Next Colonna

End Sub

Private Sub ModificaScambiaColonne_Click()
    
    Dim Riga, Colonna, Dato1, Dato2, ColonnaDato1, ColonnaDato2
    Dim NomeDellaVariabile1, NomeDellaVariabile2
    Dim Risposta, Dummy, NumDati, NumDati1, NumDati2
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    ColonnaDato1 = Colonna
    ColonnaDato2 = Colonna + 1
    
    If Colonna > (NUM_COLONNE - 1) - 2 Then
        MsgBox "Lo scambio non può essere" + Chr(10) + "effettuato oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Scambia tra di loro due colonne"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    If Colonna <= (NUM_COLONNE - 1) - 2 Then
    
        FormPrincipale!Grid1.Row = 0
        FormPrincipale!Grid1.Col = ColonnaDato1
        NomeDellaVariabile1 = FormPrincipale!Grid1.Text
        FormPrincipale!Grid1.Col = ColonnaDato2
        NomeDellaVariabile2 = FormPrincipale!Grid1.Text
        Risposta = MsgBox("Verranno scambiate tra di loro" + Chr(10) + "le due colonne selezionate." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Scambia colonne " & NomeDellaVariabile1 & " e " & NomeDellaVariabile2)
            
        If Risposta <> IDNO Then
        
            Screen.MousePointer = 11
            FormGauge.Show
            FormGauge.Caption = "Scambia colonne " & NomeDellaVariabile1 & " e " & NomeDellaVariabile2
            FormGauge!Label6.Caption = "[1/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            NumDati = 0: NumDati1 = 0: NumDati2 = 0
            
            For Riga = 1 To NUM_RIGHE - 1
                    
                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Col = ColonnaDato1
                    If FormPrincipale!Grid1.Text <> "" Then
                        NumDati1 = NumDati1 + 1
                    End If
                FormPrincipale!Grid1.Col = ColonnaDato2
                    If FormPrincipale!Grid1.Text <> "" Then
                        NumDati2 = NumDati2 + 1
                    End If

            Next Riga
                
            NumDati = NumDati1
            If NumDati1 < NumDati2 Then NumDati = NumDati2

            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
            
            For Riga = 1 To NumDati
                    
                FormGauge.Label1.Caption = Int(100 * Riga / NumDati) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / NumDati)
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Col = ColonnaDato1
                Dato1 = FormPrincipale!Grid1.Text
                FormPrincipale!Grid1.Col = ColonnaDato2
                Dato2 = FormPrincipale!Grid1.Text

                Dummy = Dato1
                Dato1 = Dato2
                Dato2 = Dummy

                FormPrincipale!Grid1.Col = ColonnaDato1
                FormPrincipale!Grid1.Text = Dato1
                FormPrincipale!Grid1.Col = ColonnaDato2
                FormPrincipale!Grid1.Text = Dato2

            Next Riga

            Dummy = NomeDellaVariabile1
            NomeDellaVariabile1 = NomeDellaVariabile2
            NomeDellaVariabile2 = Dummy
            
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = ColonnaDato1
            FormPrincipale!Grid1.Text = NomeDellaVariabile1
            FormPrincipale!Grid1.Col = ColonnaDato2
            FormPrincipale!Grid1.Text = NomeDellaVariabile2
            FormPrincipale!Grid1.Row = 1

            Unload FormGauge
'            RipristinaSchermo
            DeselezionaColonna
            InseritaNuovaColonna = False
            FormPrincipale!Grid1.SetFocus
            Screen.MousePointer = 1
        Else
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
    End If

End Sub

Private Sub ModificaTagliaColonna_Click()
    Dim Riga
    
    FormPrincipale!Grid1.Row = 0
    NomeDellaVariabile = FormPrincipale!Grid1.Text
    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Taglia colonna " & NomeDellaVariabile
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    ReDim Ritaglio(0 To NUM_RIGHE - 1)

    For Riga = 0 To NUM_RIGHE - 1

        FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        FormPrincipale!Grid1.Row = Riga
        Ritaglio(Riga) = FormPrincipale!Grid1.Text
        FormPrincipale!Grid1.Text = ""
    Next Riga

    QualcosaNelRitaglio = True
    Unload FormGauge
    RipristinaSchermo
    DeselezionaColonna
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

Private Sub ModificaUnisciColonne_Click()
    Dim Riga, Colonna, Dato1, Dato2, ColonnaDato1, ColonnaDato2
    Dim NomeDellaVariabile1, NomeDellaVariabile2
    Dim RigaUltimoDato
    On Error Resume Next

    Colonna = FormPrincipale!Grid1.SelStartCol
    ColonnaDato1 = Colonna
    ColonnaDato2 = Colonna + 1
    
    If Colonna > (NUM_COLONNE - 1) - 2 Then
        MsgBox "L'unione non può essere" + Chr(10) + "effettuata oltre la colonna 9.", MB_OK + MB_ICONSTOP + MB_DEFBUTTON1, "Unisci i dati di due colonne"
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
        Exit Sub
    End If
    If Colonna <= (NUM_COLONNE - 1) - 2 Then
    
        FormPrincipale!Grid1.Row = 0
        FormPrincipale!Grid1.Col = ColonnaDato1
        NomeDellaVariabile1 = FormPrincipale!Grid1.Text
        FormPrincipale!Grid1.Col = ColonnaDato2
        NomeDellaVariabile2 = FormPrincipale!Grid1.Text
        MsgBox "Verranno uniti in una nuova colonna" + Chr(10) + "i dati delle due colonne selezionate.", MB_OK + MB_ICONINFORMATION + MB_DEFBUTTON1, "Unisci i dati di " & NomeDellaVariabile1 & " e di " & NomeDellaVariabile2
            
        FormGauge!Label6.Caption = "[1/2]"
        InserisciColonna
            
        If InseritaNuovaColonna = True Then
        
            FormPrincipale.Refresh
            FormPrincipale.Refresh
            Screen.MousePointer = 11
            ColonnaDato1 = ColonnaDato1 + 1
            ColonnaDato2 = ColonnaDato2 + 1
            FormGauge.Show
            FormGauge.Caption = "Unisci i dati di " & NomeDellaVariabile1 & " e di " & NomeDellaVariabile2
            FormGauge!Label6.Caption = "[2/2]"
            FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

            For Riga = 1 To NUM_RIGHE - 1
                    
                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Col = ColonnaDato1
                Dato1 = FormPrincipale!Grid1.Text
                    If Dato1 <> "" Then
                        FormPrincipale!Grid1.Col = Colonna
                        FormPrincipale!Grid1.Text = Dato1
                        RigaUltimoDato = Riga
                    Else
                        RigaUltimoDato = Riga - 1
                        Exit For
                    End If
            Next Riga
                
            If RigaUltimoDato = NUM_RIGHE - 1 Then GoTo Esci

            For Riga = RigaUltimoDato + 1 To NUM_RIGHE - 1
                    
                FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
                FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
                FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                
                FormPrincipale!Grid1.Row = Riga - RigaUltimoDato
                FormPrincipale!Grid1.Col = ColonnaDato2
                Dato2 = FormPrincipale!Grid1.Text
                    If Dato2 = "" Then
                        FormGauge.Label1.Caption = "100%"
                        FormGauge!Gauge1.Value = 100
                        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh
                        Exit For
                    End If
                FormPrincipale!Grid1.Row = Riga
                FormPrincipale!Grid1.Col = Colonna
                FormPrincipale!Grid1.Text = Dato2
            Next Riga
Esci:
            FormPrincipale!Grid1.Row = 0
            FormPrincipale!Grid1.Col = Colonna
            FormPrincipale!Grid1.Text = "Unione dati"
            FormPrincipale!Grid1.Row = 1
            Unload FormGauge
'            RipristinaSchermo
            DeselezionaColonna
            InseritaNuovaColonna = False
            FormPrincipale!Grid1.SetFocus
            Screen.MousePointer = 1
        Else
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If
    End If

End Sub

Private Sub ModificaVuotaCella_Click()
    Dim Risposta
    
    If Trim(Grid1.Text) <> "" Then
        Risposta = MsgBox("Il dato della cella sarà perso." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Vuota cella (dato: " & Grid1.Text & ")")
    Else
        Risposta = IDNO
    End If

        If Risposta = IDNO Then
            RipristinaSchermo
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If

    FormPrincipale!Grid1.Text = ""

    RipristinaSchermo
    FormPrincipale!Grid1.SetFocus

End Sub

Private Sub ModificaVuotaColonna_Click()
    Dim Riga, Risposta
    
    FormPrincipale!Grid1.Row = 0
    NomeDellaVariabile = FormPrincipale!Grid1.Text
    Risposta = MsgBox("I dati della colonna saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Vuota colonna " & NomeDellaVariabile)
        If Risposta = IDNO Then
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If

    Screen.MousePointer = 11
    FormGauge.Show
    FormGauge.Caption = "Vuota colonna " & NomeDellaVariabile
    FormGauge!Label6.Caption = "[1/1]"
    FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

    For Riga = 0 To NUM_RIGHE - 1

        FormGauge.Label1.Caption = Int(100 * Riga / (NUM_RIGHE - 1)) & "%"
        FormGauge!Gauge1.Value = Int(100 * Riga / (NUM_RIGHE - 1))
        FormGauge!Label1.Refresh: FormGauge!Gauge1.Refresh

        FormPrincipale!Grid1.Row = Riga
        FormPrincipale!Grid1.Text = ""
    Next Riga

    Unload FormGauge
    RipristinaSchermo
    DeselezionaColonna
    FormPrincipale!Grid1.SetFocus
    Screen.MousePointer = 1

End Sub

Private Sub ModificaVuotaRiga_Click()
    Dim Risposta, Colonna
    
    Risposta = MsgBox("I dati della riga saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Vuota riga " & Grid1.Row)
        If Risposta = IDNO Then
            RipristinaSchermo
            DeselezionaColonna
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If

    For Colonna = 1 To NUM_COLONNE - 1
        FormPrincipale!Grid1.Col = Colonna
        FormPrincipale!Grid1.Text = ""
    Next Colonna

    RipristinaSchermo
    DeselezionaColonna
    FormPrincipale!Grid1.SetFocus

End Sub

Private Sub NomeDellaVariabile_Click()
    FormNomeVariabile.Show 1
End Sub

Private Sub NonParametrica_Click()
'    Screen.MousePointer = 11
'    FormNonParametrica.Show
'    ' FormPrincipale.Hide
'    Screen.MousePointer = 1
End Sub

Private Sub OpzioniBarraDegliStrumenti_Click()
    
    OpzioniBarraDegliStrumenti.Checked = Not OpzioniBarraDegliStrumenti.Checked
    BarraDegliStrumenti = Not BarraDegliStrumenti
    If BarraDegliStrumenti = True Then
        
        Grid1.Move 0, 1800
        Grid1.Width = ScaleWidth
        Grid1.Height = ScaleHeight - 1800
        
        CommandApri.Visible = True
        CommandSalva.Visible = True
        CommandNuovo.Visible = True
        CommandStampaColonne1.Visible = True
        CommandStampaColonne6.Visible = True
        
        CommandGuida.Visible = True
        
        CommandNumera.Visible = True
        CommandOrdinaCrescente.Visible = True
        CommandOrdinaDecrescente.Visible = True
        CommandRango.Visible = True
        CommandPercentili.Visible = True
        CommandDeviataZeta.Visible = True
        
        CommandIstogramma.Visible = True
        CommandCumulativa.Visible = True
        CommandConfronto.Visible = True
        CommandAppaiamento.Visible = True
        CommandDispersione.Visible = True
        CommandAltman.Visible = True
        CommandAssociazione.Visible = True
        CommandSovrapposizione.Visible = True
        CommandLineaSpezzata.Visible = True
        CommandBarreVerticali.Visible = True
        CommandIstogrammiMultipli.Visible = True
        CommandDistribuzione.Visible = True
        CommandErroreStandard.Visible = True
        CommandQuartiliERange.Visible = True
        
        CommandEsploraPar.Visible = True
        CommandTorta.Visible = True
        CommandStatisticheElementari.Visible = True
        CommandConfrontoTraMedie.Visible = True
        CommandAnalisiDEllaVarianza.Visible = True
        CommandRegressioneLineare.Visible = True
        CommandRegressionePolinomiale.Visible = True
        CommandRegressioneMultipla.Visible = True
        CommandStatisticaBayesiana.Visible = True
        CommandEsploraNonPar.Visible = True
        CommandStatNon.Visible = True
        CommandConfrontoTraMediane.Visible = True
        CommandRegNon.Visible = True
        CommandTestChiQuadrato.Visible = True
        CommandAnalisiDellaSomiglianza.Visible = True
    
    Else
        
        Grid1.Move 0, 0
        Grid1.Width = ScaleWidth
        Grid1.Height = ScaleHeight
        
        CommandApri.Visible = False
        CommandSalva.Visible = False
        CommandNuovo.Visible = False
        CommandStampaColonne1.Visible = False
        CommandStampaColonne6.Visible = False
        
        CommandGuida.Visible = False
        
        CommandNumera.Visible = False
        CommandOrdinaCrescente.Visible = False
        CommandOrdinaDecrescente.Visible = False
        CommandRango.Visible = False
        CommandPercentili.Visible = False
        CommandDeviataZeta.Visible = False
        
        CommandIstogramma.Visible = False
        CommandCumulativa.Visible = False
        CommandConfronto.Visible = False
        CommandAppaiamento.Visible = False
        CommandDispersione.Visible = False
        CommandAltman.Visible = False
        CommandAssociazione.Visible = False
        CommandSovrapposizione.Visible = False
        CommandLineaSpezzata.Visible = False
        CommandBarreVerticali.Visible = False
        CommandIstogrammiMultipli.Visible = False
        CommandDistribuzione.Visible = False
        CommandErroreStandard.Visible = False
        CommandQuartiliERange.Visible = False
        
        CommandEsploraPar.Visible = False
        CommandTorta.Visible = False
        CommandStatisticheElementari.Visible = False
        CommandConfrontoTraMedie.Visible = False
        CommandAnalisiDEllaVarianza.Visible = False
        CommandRegressioneLineare.Visible = False
        CommandRegressionePolinomiale.Visible = False
        CommandRegressioneMultipla.Visible = False
        CommandStatisticaBayesiana.Visible = False
        CommandEsploraNonPar.Visible = False
        CommandStatNon.Visible = False
        CommandConfrontoTraMediane.Visible = False
        CommandRegNon.Visible = False
        CommandTestChiQuadrato.Visible = False
        CommandAnalisiDellaSomiglianza.Visible = False
    
    End If

    FormPrincipale!Grid1.SetFocus

End Sub

Private Sub OpzioniCinqueDecimali_Click()
    Dim StringaDiFormattazione As String, Risposta

    Risposta = MsgBox("In seguito all'operazione di arrotondamento" + Chr(10) + "i decimali in eccesso saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Arrotonda a cinque decimali")
        If Risposta = IDNO Then
            DeselezionaColonna
            RipristinaSchermo
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If

' limita a 12 i caratteri, inclusa la virgola
    StringaDiFormattazione = "#####0.00000"
    FormattaNumero (StringaDiFormattazione)
    FormPrincipale!Grid1.SetFocus

End Sub

Private Sub OpzioniDueDecimali_Click()
    Dim StringaDiFormattazione As String, Risposta
    
    Risposta = MsgBox("In seguito all'operazione di arrotondamento" + Chr(10) + "i decimali in eccesso saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Arrotonda a due decimali")
        If Risposta = IDNO Then
            DeselezionaColonna
            RipristinaSchermo
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If

' limita a 12 i caratteri, inclusa la virgola
    StringaDiFormattazione = "########0.00"
    FormattaNumero (StringaDiFormattazione)
    FormPrincipale!Grid1.SetFocus

End Sub

Private Sub OpzioniGriglia_Click()
    
    OpzioniGriglia.Checked = Not OpzioniGriglia.Checked
    Grid1.GridLines = Not Grid1.GridLines
    FormPrincipale!Grid1.SetFocus

End Sub

Private Sub OpzioniNessunDecimale_Click()
    Dim StringaDiFormattazione As String, Risposta

    Risposta = MsgBox("In seguito all'operazione di arrotondamento" + Chr(10) + "i decimali in eccesso saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Arrotonda a nessun decimale")
        If Risposta = IDNO Then
            DeselezionaColonna
            RipristinaSchermo
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If

' limita a 12 i caratteri, inclusa la virgola
    StringaDiFormattazione = "###########0"
    FormattaNumero (StringaDiFormattazione)
    FormPrincipale!Grid1.SetFocus

End Sub

Private Sub OpzioniNomeDellaVariabile_Click()
    
    FormNomeVariabile.Show 1
    FormPrincipale!Grid1.SetFocus

End Sub

Private Sub OpzioniQuattroDecimali_Click()
    Dim StringaDiFormattazione As String, Risposta

    Risposta = MsgBox("In seguito all'operazione di arrotondamento" + Chr(10) + "i decimali in eccesso saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Arrotonda a quattro decimali")
        If Risposta = IDNO Then
            DeselezionaColonna
            RipristinaSchermo
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If

' limita a 12 i caratteri, inclusa la virgola
    StringaDiFormattazione = "######0.0000"
    FormattaNumero (StringaDiFormattazione)
    FormPrincipale!Grid1.SetFocus

End Sub

Private Sub OpzioniRipristinaAltezzaRighe_Click()
    Dim Riga
    
    Screen.MousePointer = 11
    
    FormPrincipale!Grid1.HighLight = False

    For Riga = 0 To NUM_RIGHE - 1
        FormPrincipale!Grid1.RowHeight(Riga) = 255
    Next Riga

' ripristina schermo
    FormPrincipale!Grid1.Row = Rig
    FormPrincipale!Grid1.Col = Col
    FormPrincipale!Grid1.TopRow = RigSup
    FormPrincipale!Grid1.LeftCol = ColSin
    FormPrincipale!Grid1.SetFocus
    
    Screen.MousePointer = 1

End Sub

Private Sub OpzioniRipristinaLarghezzaColonne_Click()
    Dim Colonna
    
    Screen.MousePointer = 11
    
    FormPrincipale!Grid1.HighLight = False
    
    FormPrincipale!Grid1.ColWidth(0) = 574
    FormPrincipale!Grid1.Row = 0
        For Colonna = 1 To NUM_COLONNE - 1
            FormPrincipale!Grid1.ColWidth(Colonna) = LARGHEZZA_COLONNE
        Next Colonna

' ripristina schermo
    FormPrincipale!Grid1.Row = Rig
    FormPrincipale!Grid1.Col = Col
    FormPrincipale!Grid1.TopRow = RigSup
    FormPrincipale!Grid1.LeftCol = ColSin
    FormPrincipale!Grid1.SetFocus
    
    Screen.MousePointer = 1

End Sub

Private Sub OpzioniTreDecimali_Click()
    Dim StringaDiFormattazione As String, Risposta

    Risposta = MsgBox("In seguito all'operazione di arrotondamento" + Chr(10) + "i decimali in eccesso saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Arrotonda a tre decimali")
        If Risposta = IDNO Then
            DeselezionaColonna
            RipristinaSchermo
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If

' limita a 12 i caratteri, inclusa la virgola
    StringaDiFormattazione = "#######0.000"
    FormattaNumero (StringaDiFormattazione)
    FormPrincipale!Grid1.SetFocus

End Sub

Private Sub OpzioniUnDecimale_Click()
    Dim StringaDiFormattazione As String, Risposta
    
    Risposta = MsgBox("In seguito all'operazione di arrotondamento" + Chr(10) + "i decimali in eccesso saranno persi." + Chr(10) + "Conferma ?", MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2, "Arrotonda a un decimale")
        If Risposta = IDNO Then
            DeselezionaColonna
            RipristinaSchermo
            FormPrincipale!Grid1.SetFocus
            Exit Sub
        End If

' limita a 12 i caratteri, inclusa la virgola
    StringaDiFormattazione = "#########0.0"
    FormattaNumero (StringaDiFormattazione)
    FormPrincipale!Grid1.SetFocus

End Sub

Private Sub RipristinaLarghezzaColonne_Click()
    Dim Colonna
    
    FormPrincipale!Grid1.ColWidth(0) = 574
    FormPrincipale!Grid1.Row = 0
        For Colonna = 1 To NUM_COLONNE - 1
            FormPrincipale!Grid1.ColWidth(Colonna) = LARGHEZZA_COLONNE
        Next Colonna

End Sub

Private Sub StatisticaAnova_Click()

'    Screen.MousePointer = 11
    Load FormANOVA
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormANOVA
        End If

End Sub

Private Sub StatisticaBayesiana_Click()

'    Screen.MousePointer = 11
    Load FormBayes
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormBayes
        End If

End Sub

Private Sub StatisticaChiQuadrato_Click()

    Screen.MousePointer = 11
    Load FormChi
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormChi
        End If

End Sub

Private Sub StatisticaCluster_Click()
    
'    Screen.MousePointer = 11
    Load FormCluster
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormCluster
        End If

End Sub

Private Sub StatisticaEsploraNonPar_Click()

'    Screen.MousePointer = 11
    Load FormEsploraNonPar
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormEsploraNonPar
        End If

End Sub

Private Sub StatisticaEsploraPar_Click()

'    Screen.MousePointer = 11
    Load FormEsploraPar
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormEsploraPar
        End If

End Sub

Private Sub StatisticaNonParametricaChi_Click()

'    Screen.MousePointer = 11
    Load FormChi
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormChi
        End If

End Sub

Private Sub StatisticaNonParametricaCumulativa_Click()

'    Screen.MousePointer = 11
    Load FormCumulativa
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormCumulativa
        End If

End Sub

Private Sub StatisticaNonParametricaNumera_Click()

    CalcoliNumera_Click
    
End Sub

Private Sub StatisticaNonParametricaPercentili_Click()

    CalcoliPercentili_Click

End Sub

Private Sub StatisticaNonParametricaRango_Click()

    CalcoliRango_Click

End Sub

Private Sub StatisticaNonParametricaReglin_Click()

'    Screen.MousePointer = 11
    Load FormNonParametrica
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormNonParametrica
        End If

End Sub

Private Sub StatisticaNonParametricaW_Click()
    
'    Screen.MousePointer = 11
    Load FormW
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormW
        End If

End Sub

Private Sub StatisticaRegressioneLineare_Click()

'    Screen.MousePointer = 11
    Load FormParametrica
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormParametrica
        End If

End Sub

Private Sub StatisticaRegressioneMultipla_Click()
    
'    Screen.MousePointer = 11
    Load FormMultipla
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormMultipla
        End If

End Sub

Private Sub StatisticaRegressionePolinomiale_Click()
    
'    Screen.MousePointer = 11
    Load FormPolinomiale
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormPolinomiale
        End If

End Sub

Private Sub StatisticaStatisticheElementari_Click()

'    Screen.MousePointer = 11
    Load FormIstogramma
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormIstogramma
        End If

End Sub

Private Sub StatisticaT_Click()

'    Screen.MousePointer = 11
    Load FormT
        If AnnullatoCaricamentoForm = True Then
            AnnullatoCaricamentoForm = False
            Unload FormT
        End If

End Sub

