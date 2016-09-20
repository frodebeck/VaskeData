# VaskeData
Testing
Skal se hvordan dette fungerer
Public Sub FormatereLønnsdataFraHL()

'Makroen går gjennom uttrekket fra Huldt og Lillevik. Flytter navnet til kolonne 11 på cellen under og sletter så navneraden. For hver rad hvor kol 7 ikke er tom eller lik 'Sum' henter han navnet og setter i kol 11. 'Sum'-rader slettes av makroen. Makroen er veldig avhengig av at kol G er stabil.

Dim MakroNavn As String, StatusFortsett As Integer
'Dim KodeModulnavn As CodeModule
'Set KodeModulnavn = Application.VBE.ActiveCodePane.CodeModule

'Sjekker tilgang
MakroNavn = "FormatereLønnsdataFraHL"
If BrukerHarTilgangTilMakro(MakroNavn) = False Then Exit Sub

'Dekker informasjonsbehovet for brukeren
If SjekkInformasjon = True Then StatusFortsett = MsgBox(SjekkMakro(MakroNavn, "Informasjon") & vbLf & "Trykk 'Ok' om dette virker greit og 'Avbryt' om du ikke ønsker å gå videre.", vbOKCancel, "Informasjon")

'Finn antall rader
Dim AntallRader As Long, i As Long
AntallRader = AntallRaderKolonne(1)

Application.ScreenUpdating = False

'Velger område
'Range(Cells(1, 1), Cells(AntallRader, 11)).Select

'Fyller på med data eller sletter rader
    For i = 1 To AntallRader
    
        'tar høyde for at ingenting skal skje i rad 1
        If i = 1 Then
        'Flytter navnet
        ElseIf Cells(i, 7).Value = "" Then
            Cells(i + 1, 11).Value = Cells(i, 1).Value
            
        'Sletter SUM-rader
        ElseIf Cells(i, 7).Value = "Sum" Then
            Rows(i).Delete
            i = i - 1
        
        'Setter inn ref til navn
        ElseIf Len(Cells(i, 7).Value) > 0 Then
            If Len(Cells(i, 11).Value) = 0 Then Cells(i, 11).Formula = Cells(i - 1, 11).Value

        End If

    Next

    For i = AntallRader To 2 Step -1
        If Len(Cells(i, 7).Value) = 0 Then Rows(i).Delete
    Next
Cells(i, 11).Value = "Navn"
Application.ScreenUpdating = True

Call TilpassKolonnerKode(True)

Call SkrivMakroLogg(ModulNavn, MakroNavn)

End Sub
