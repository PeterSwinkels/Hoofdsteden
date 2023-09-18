VERSION 5.00
Begin VB.Form FoutenVenster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fout beantwoorde vragen:"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   ClipControls    =   0   'False
   Icon            =   "Fouten.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12.625
   ScaleMode       =   4  'Character
   ScaleWidth      =   39.375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox FoutenLijst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "FoutenVenster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deze module bevat het Fouten venster.
Option Explicit

'Deze procedure geeft de opdracht om de lijst van fouten te tonen wanneer dit venster actief wordt.
Private Sub Form_Activate()
On Error GoTo Fout
   ToonFouten FoutenLijst
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub

'Deze procedure stelt dit venster in wanneer het geopend wordt.
Private Sub Form_Load()
On Error Resume Next

   Me.Width = Screen.Width / 2
   Me.Height = Screen.Height / 2
   
   FoutenLijst.Width = Me.ScaleWidth - 2
   FoutenLijst.Height = Me.ScaleHeight - 1
End Sub



