VERSION 5.00
Begin VB.Form HoofdstedenVenster 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3660
   ClipControls    =   0   'False
   Icon            =   "Hoofdsteden.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8.625
   ScaleMode       =   4  'Character
   ScaleWidth      =   30.5
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox RondePaneel 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1935
      ScaleWidth      =   3495
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton ControleerKnop 
         Caption         =   "&Controleer."
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox AntwoordVeld 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   255
         TabIndex        =   1
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label AantalVragenLabel 
         Caption         =   "Vraag: 0/0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label GoedFoutLabel 
         Caption         =   "Goed: 0 - Fout: 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label VraagLabel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label AntwoordLabel 
         Caption         =   "Antwoord:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Menu ProgrammaHoofdmenu 
      Caption         =   "&Programma"
      Begin VB.Menu InformatieMenu 
         Caption         =   "&Informatie"
         Shortcut        =   ^I
      End
      Begin VB.Menu SluitenMenu 
         Caption         =   "&Sluiten."
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu RegioHoofdmenu 
      Caption         =   "&Regio"
      Begin VB.Menu RegioMenu 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu RondeHoofdmenu 
      Caption         =   "&Ronde"
      Begin VB.Menu OpnieuwBeginnenMenu 
         Caption         =   "&Opnieuw beginnen."
         Shortcut        =   ^O
      End
      Begin VB.Menu ToonFoutenMenu 
         Caption         =   "&Toon fouten."
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "HoofdstedenVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het hoofdvenster.
Option Explicit






'Deze procedure geeft opdracht om het antwoord te controleren en daarna een nieuwe vraag te stellen.
Private Sub ControleerKnop_Click()
On Error GoTo Fout
   Controleer AntwoordVeld.Text
   ToonStatus AantalVragenLabel, GoedFoutLabel
   VraagLabel.Caption = StelVraag(RondePaneel, AntwoordVeld)
   
   If RondePaneel.Enabled Then AntwoordVeld.SetFocus
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub

'Deze procedure geeft de opdracht om de status te tonen.
Private Sub Form_Activate()
On Error GoTo Fout
   ToonStatus AantalVragenLabel, GoedFoutLabel
   VraagLabel.Caption = StelVraag(RondePaneel, AntwoordVeld)
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub
'Deze procedure stelt dit venster in.
Private Sub Form_Load()
On Error GoTo Fout
   Me.Caption = ProgrammaInformatie()
   
   MaakRegioLijst RegioMenu
   LaadRegioLijst RegioMenu, RegioMenu.LBound()
   
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub

'Deze procedure beëindigt dit programma na bevestiging van de gebruiker.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Fout
Dim Keuze As Long

   Keuze = MsgBox("Dit programma afsluiten?", vbQuestion Or vbYesNo Or vbDefaultButton2)
   Cancel = (Keuze = vbNo)

EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub


'Deze procedure geeft de opdracht om de informatie over dit programma te tonen.
Private Sub InformatieMenu_Click()
On Error GoTo Fout
   MsgBox ProgrammaInformatie(), vbInformation
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub

'Deze procedure begint opnieuw met de vragen na bevestiging van de gebruiker.
Private Sub OpNieuwBeginnenMenu_Click()
On Error GoTo Fout
   If MsgBox("Nog een keer oefenen?", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes Then
      NieuweRonde RondePaneel
      ToonStatus AantalVragenLabel, GoedFoutLabel
      VraagLabel.Caption = "Wat is de hoofdstad van """ & Ronde.Vragen(WillekeurigeVraag(NieuweVraag:=True)).Gebied & """?"
   End If
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub


'Deze procedure geeft de opdracht om de geselecteerde regio te laden.
Private Sub RegioMenu_Click(Index As Integer)
On Error GoTo Fout
   If Not RegioMenu(Index).Caption = vbNullString Then
      If MsgBox("Nieuwe regio laden?", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes Then
         LaadRegioLijst RegioMenu, CLng(Index)
         NieuweRonde
         ToonStatus AantalVragenLabel, GoedFoutLabel
         VraagLabel.Caption = StelVraag(RondePaneel, AntwoordVeld)
      End If
   End If
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub

'Deze procedure sluit het hoofdvenster.
Private Sub SluitenMenu_Click()
On Error GoTo Fout
   Unload Me
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub


'Deze procedure geeft de opdracht om het fouten venster te tonen.
Private Sub ToonFoutenMenu_Click()
On Error GoTo Fout
   With Ronde
      If .AantalFout + .AantalGoed < Abs(UBound(.Vragen()) - LBound(.Vragen())) Then
         MsgBox "Eerst moeten alle vragen beantwoord zijn.", vbExclamation
      Else
         FoutenVenster.Show vbModal
      End If
   End With
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub


