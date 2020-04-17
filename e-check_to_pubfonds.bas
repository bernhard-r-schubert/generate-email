Attribute VB_Name = "Modul1"
Sub Metadaten_EligibilityCheck_nach_PubFonds()
Attribute Metadaten_EligibilityCheck_nach_PubFonds.VB_ProcData.VB_Invoke_Func = " \n14"

' Metadaten_EligibilityCheck_nach_PubFonds Makro, Variablen definieren

    Dim E_Check, PubFonds As Integer
    Dim Eingangsdatum, Checkdatum As Date
    Dim Typ, Verlag, Corr_Author, Titel, Journal, DOI As String 'Typ ist für IOP-Prüfung notwendig
                
Quelleneingabe:
                
   E_Check = InputBox("Bitte zu verwendende Quellzeile aus dem Eligility-Check-Masterfile eingeben:")

    If IsNumeric(E_Check) = False Then 'Check, dass Wert eine Zahl ist
       MsgBox "Zahlenwert erwartet!", vbOKOnly
       GoTo Quelleneingabe
    End If

   Windows("01 Eligibility-Check-Masterfile.xlsm").Activate

   If Cells(E_Check, 2).Value = "" Then 'Check, dass nur ausgef_llte Zeile ausgew_hlt wird
       MsgBox "Quellzeile ist leer!", vbOKOnly
       GoTo Quelleneingabe
   End If
    
'   Variable Werte aus EligibilityCheck lesen
  
    Typ = Cells(E_Check, 1)
    Eingangsdatum = Cells(E_Check, 2)
    Checkdatum = Cells(E_Check, 3)
    Verlag = Cells(E_Check, 4)
    Corr_Author = Cells(E_Check, 6)
    Titel = Cells(E_Check, 7)
    Journal = Cells(E_Check, 8)
    DOI = Cells(E_Check, 17)

'   Fixe und variable Werte in Publikationsfonds schreiben

    Windows("Publikationsfonds Kontostand SAP.xlsx").Activate
    
    Sheets("Publikationsfonds APCs").Select 'Richtiges Blatt auswählen
    Range("A16").Select 'Beginn der eigentlichen Tabelle auswählen

    PubFonds = Selection.End(xlDown).Row + 1 'Erste leere Zeile auswählen
    
'   Fixe Werte einf_gen

    Cells(PubFonds, 1).Value = "Zusage"
    Cells(PubFonds, 3).Value = "APC"
    If Verlag = "de Gruyter" Then
        Cells(PubFonds, 4).Value = "ja"
    ElseIf Verlag = "SAGE" Then
        Cells(PubFonds, 4).Value = "ja"
        Cells(PubFonds, 20).Value = "GBP 200"
    ElseIf Verlag = "IOP" And Typ = "Deal" Then
        Cells(PubFonds, 4).Value = "ja"
    Else
        Cells(PubFonds, 4).Value = "nein"
    End If
    Cells(PubFonds, 7).Value = "Wien U"

'   Variable Werte einf_gen

    Cells(PubFonds, 5).Value = Corr_Author
    Cells(PubFonds, 8).Value = Titel
    Cells(PubFonds, 9).Value = Journal
    Cells(PubFonds, 10).Value = Verlag
    Cells(PubFonds, 11).Value = DOI
    Cells(PubFonds, 16).Value = Eingangsdatum
    Cells(PubFonds, 17).Value = Checkdatum
    
End Sub
