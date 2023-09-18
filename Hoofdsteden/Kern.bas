Attribute VB_Name = "KernModule"
'Deze module bevat de kern procedures van dit programma.
Option Explicit

'Deze structuur definieert een vraag.
Public Type VraagStr
   AntwoordGoed As Boolean   'Geeft aan of de vraag goed is beantwoord.
   Gebied As String          'Definieert de naam van een gebied met een hoofdstad.
   Gehad As Boolean          'Geeft aan of de hoofdstad voor een gebied al eerder is gevraagd.
   HoofdStad() As String     'Definieert de hoofdstad van een gebied.
End Type

'Deze structuur definieert een ronde van vragen.
Public Type RondeStr
   AantalFout As Long      'Definieert  het aantal fout beantwoorde vragen.
   AantalGoed As Long      'Definieert het aantal goed beantwoorde vragen.
   Vragen() As VraagStr    'Definieert de lijst van vragen.
End Type

Public Ronde As RondeStr   'Bevat een ronde van vragen.

'Deze procedure controleert het door de gebruiker gegeven antwoord.
Public Sub Controleer(Antwoord As String)
On Error GoTo Fout
Dim Index As Long

   With Ronde
      If .AantalFout + .AantalGoed <= Abs(UBound(.Vragen()) - LBound(.Vragen())) Then
         Antwoord = LCase$(Trim$(Antwoord))
         
         With .Vragen(WillekeurigeVraag())
            .Gehad = True
            
            For Index = LBound(.HoofdStad()) To UBound(.HoofdStad()) - 1
               If Antwoord = LCase$(Trim$(.HoofdStad(Index))) Then
                  Ronde.AantalGoed = Ronde.AantalGoed + 1
                  .AntwoordGoed = True
                  Exit For
               End If
            Next Index
         End With
         
         If Not .Vragen(WillekeurigeVraag()).AntwoordGoed Then HandelFoutAntwoordAf
      End If
   End With
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub

'Deze procedure handelt eventuele fouten af.
Public Sub HandelFoutAf()
Dim FoutCode As Long
Dim Omschrijving As String

   FoutCode = Err.Number
   Omschrijving = Err.Description
   
   On Error Resume Next
   MsgBox "Foutcode: " & FoutCode & vbCr & Omschrijving, vbExclamation
End Sub

'Deze procedure handelt eventuele foute antwoorden af.
Public Sub HandelFoutAntwoordAf()
On Error GoTo Fout
Dim GoedeAntwoorden As String
Dim Index As Long

   Ronde.AantalFout = Ronde.AantalFout + 1
   
   With Ronde.Vragen(WillekeurigeVraag())
      If Abs((UBound(.HoofdStad()) - 1) - LBound(.HoofdStad())) > 1 Then
         GoedeAntwoorden = "De mogelijke goede antwoorden zijn: "
      Else
         GoedeAntwoorden = "Het goede antwoord is: "
      End If
      
      For Index = LBound(.HoofdStad()) To UBound(.HoofdStad()) - 1
         GoedeAntwoorden = GoedeAntwoorden & """" & .HoofdStad(Index) & """"
         If Index = UBound(.HoofdStad()) - 2 Then
            GoedeAntwoorden = GoedeAntwoorden & " of "
         ElseIf Index < UBound(.HoofdStad()) - 2 Then
            GoedeAntwoorden = GoedeAntwoorden & ", "
         End If
      Next Index
   End With
   
   MsgBox "Fout. " & GoedeAntwoorden & ".", vbExclamation
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub

'Deze procedure laadt de regio lijst met de opgegeven index.
Public Sub LaadRegioLijst(RegioMenu As Object, Index As Long)
On Error GoTo Fout
Static VorigeIndex As Long

   If Not RegioMenu(Index) = vbNullString Then
      RegioMenu(VorigeIndex).Checked = False
      VorigeIndex = Index
      LeesHoofdsteden Mid$(RegioMenu(Index).Caption, 2)
      NieuweRonde
      RegioMenu(Index).Checked = True
   End If
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub

'Deze procedure leest de lijst van gebieden en bijbehorende hoofdsteden in.
Public Sub LeesHoofdsteden(Lijst As String)
On Error GoTo Fout
Dim BestandH As Long
ReDim Ronde.Vragen(0 To 0) As VraagStr

   BestandH = FreeFile()
   Open ".\Data\" & Lijst & ".txt" For Input As BestandH
      Do Until EOF(BestandH)
         With Ronde.Vragen(UBound(Ronde.Vragen()))
            Input #BestandH, .Gebied
            
            ReDim .HoofdStad(0 To 0) As String
            Do
               Input #BestandH, .HoofdStad(UBound(.HoofdStad()))
               If .HoofdStad(UBound(.HoofdStad())) = "*" Then Exit Do
               ReDim Preserve .HoofdStad(LBound(.HoofdStad()) To UBound(.HoofdStad()) + 1) As String
            Loop
         End With
      
         If Not EOF(BestandH) Then ReDim Preserve Ronde.Vragen(LBound(Ronde.Vragen()) To UBound(Ronde.Vragen()) + 1) As VraagStr
      Loop
EindeProcedure:
   Close BestandH
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub

'Deze procedure maakt de lijst van beschikbare regio's.
Public Sub MaakRegioLijst(RegioMenu As Object)
On Error GoTo Fout
Dim Lijst As String

   Do While RegioMenu.Count > 1 And DoEvents() > 0
      Unload RegioMenu(RegioMenu(RegioMenu.Count() - 1))
   Loop
   RegioMenu(RegioMenu.Count() - 1).Caption = vbNullString
   
   Lijst = Dir$(".\Data\*.txt", vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
   Do Until Lijst = vbNullString
      Lijst = Left$(Lijst, InStrRev(Lijst, ".") - 1)
      RegioMenu(RegioMenu.UBound()).Caption = "&" & Lijst
      Lijst = Dir$()
      If Not Lijst = vbNullString Then Load RegioMenu(RegioMenu.Count)
   Loop
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub

'Deze procedure wordt uitgevoerd wanneer dit programma wordt gestart.
Public Sub Main()
On Error GoTo Fout
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   Randomize Timer
   
   HoofdstedenVenster.Show
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub


'Deze procedure start een nieuwe ronde.
Public Sub NieuweRonde(Optional RondePaneel As PictureBox = Nothing)
On Error GoTo Fout
Dim Index As Long

   With Ronde
      .AantalFout = 0
      .AantalGoed = 0
      For Index = LBound(.Vragen()) To UBound(.Vragen())
         With .Vragen(Index)
            .AntwoordGoed = False
            .Gehad = False
         End With
      Next Index
      
      If Not RondePaneel Is Nothing Then RondePaneel.Enabled = True
   End With
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub


'Deze procedure stuurt de informatie over dit programma terug.
Public Function ProgrammaInformatie() As String
On Error GoTo Fout
Dim Informatie As String

   With App
      Informatie = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & ", by: " & .CompanyName
   End With
   Informatie = Informatie & " - ***2014***"
   
EindeProcedure:
   ProgrammaInformatie = Informatie
   Exit Function
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Function


'Deze procedure stuurt een vraag terug.
Public Function StelVraag(RondePaneel As PictureBox, AntwoordVeld As TextBox) As String
On Error GoTo Fout
Dim Vraag As String

   With Ronde
      If .AantalGoed + .AantalFout > Abs(UBound(.Vragen()) - LBound(.Vragen())) Then
         RondePaneel.Enabled = False
         Vraag = vbNullString
         MsgBox "Alle vragen zijn beantwoord.", vbInformation
      Else
         Vraag = "Wat is de hoofdstad van """ & .Vragen(WillekeurigeVraag(NieuweVraag:=True)).Gebied & """?"
         AntwoordVeld.Text = vbNullString
         AntwoordVeld.SetFocus
      End If
   End With
EindeProcedure:
   StelVraag = Vraag
   Exit Function
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Function

'Deze procedure toont de lijst van fout beantwoorde vragen.
Public Sub ToonFouten(FoutenLijst As ListBox)
On Error GoTo Fout
Dim Index As Long
 
   With Ronde
      FoutenLijst.Clear
      For Index = LBound(.Vragen()) To UBound(.Vragen())
         If Not .Vragen(Index).AntwoordGoed Then
            FoutenLijst.AddItem .Vragen(Index).Gebied & " = " & .Vragen(Index).HoofdStad(LBound(.Vragen(Index).HoofdStad()))
         End If
      Next Index
   End With
   
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub
'Deze procedure toont de status van dit programma.
Public Sub ToonStatus(AantalVragenLabel As Label, GoedFoutLabel As Label)
On Error GoTo Fout
   With Ronde
      If .AantalGoed + .AantalFout <= Abs(UBound(.Vragen()) - LBound(.Vragen())) Then
         AantalVragenLabel.Caption = "Vraag: " & CStr((.AantalGoed + .AantalFout) + 1) & "/" & CStr(Abs(UBound(.Vragen()) - LBound(.Vragen())) + 1)
      End If
      GoedFoutLabel.Caption = "Goed: " & CStr(.AantalGoed) & " - Fout: " & CStr(.AantalFout)
   End With
EindeProcedure:
   Exit Sub
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Sub

'Deze procedure stuurt een willekeurige vraag terug.
Public Function WillekeurigeVraag(Optional NieuweVraag As Boolean = False) As Long
Static Vraag As Long

   If NieuweVraag Then
      With Ronde
         If .AantalFout + .AantalGoed <= Abs(UBound(.Vragen()) - LBound(.Vragen())) Then
            Do
               Vraag = CInt(Rnd * Abs(UBound(.Vragen()) - LBound(.Vragen())))
            Loop While .Vragen(Vraag).Gehad And DoEvents() > 0
         End If
      End With
   End If
   
EindeProcedure:
   WillekeurigeVraag = Vraag
   Exit Function
   
Fout:
   HandelFoutAf
   Resume EindeProcedure
End Function




