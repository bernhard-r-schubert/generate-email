Sub EMail_Erzeugen()

''''''''''''''''''''''
' Inhaltsverzeichnis '
''''''''''''''''''''''

'0 Definition der Variablen

'0.1 Allgemeine Variablen
'0.2 Variablen für Metadaten aus der Liste
'0.3 Variablen für E-Mail-Textteile
'0.4 Variable Werte aus PubFonds

'1 Ablehnung an Autor*in

'1.1 BPC
'1.2 Hybrid außerhalb von Abkommen
'1.3 Limit erreicht (BMC, Frontiers, MDPI, Publikationsfonds)
'1.4 FWF-Ablehnung und -Zuweisung
'1.5 EU-Ablehnung
'1.6 Nicht affiliiert zum Stichtag

'2 Rechnungslegung an Quästur und Zahlungsbestätigung an Autor*in (de Gruyter, Frontiers, MDPI, SAGE, Publikationsfonds)

'2.1 Sofort zu zahlen
'2.2 Nicht sofort zu zahlen

'3 FWF-Klärung und -Zuweisung an FWF bzw. Verlag

'3.1 FWF-Nachfrage
'3.1.1 Elsevier, SAGE

'3.2 FWF-Zuweisung
'3.2.1 BMC, Frontiers, Elsevier, IOP, SAGE, Wiley
'3.2.2 MDPI

'4 Abkommen-spezifische Nachricht an Verlag

'4.1 BMC
'4.1.1 Anforderung Information
'4.1.2 Ablehnung wegen Nichtaffiliation
'4.1.3 Ablehnung wegen EU funding
'4.1.4 Bestätigung

'4.2 IOP
'4.2.1 Ablehnung wegen Nichtaffiliation
'4.2.2 Bestätigung

'4.3 T&F
'4.3.1 Approval disabled

'4.4 de Gruyter, SAGE
'4.4.1 Reklamation: Artikel nicht OA

'5 Hinweis auf Abkommen/Retro-OA an Autor*in

'5.1 ACS, de Gruyter
'5.2 CUP
'5.3 Elsevier
'5.4 IEEE

'6 Bestätigung der Kostenübernahme an Autor*in

'6.1 Bestätigung Publikationsfonds
'6.1.1 Über Kostengrenze
'6.1.2 Regulär

'6.2 Bestätigung Verlagsabkommen
'6.2.1 Frontiers

''''''''''''''''''''''''''''''''
' 0 Definition der Variablen   '
''''''''''''''''''''''''''''''''

' 0.1 Allgemeine Variablen
''''''''''''''''''''''''''

    Dim masterlist As Integer
    Dim url As String
    
'0.2 Variablen für Metadaten aus der Liste
''''''''''''''''''''''''''''''''''''''''''

    Dim open_access_deal, type_of_charge, publisher, title, source_full_title, article_id, doi, funder, echeck_status, due_date, invoice_status, InvoiceNr, id, license_ref As String
    Dim corresponding_author As String 'weil Listen von Variablen eigentlich als Variant eingerichtet werden muss das hier extra sein
    Dim reject_reason, oa_status As String
    Dim doaj, affiliated As Boolean
    Dim euro As Integer
    Dim echeck_date As Date
        

'0.3 Variablen für E-Mail-Textteile
'''''''''''''''''''''''''''''''''''

Dim PublisherContact, PriceReductionGer, PriceReductionEng, RejectionReasonGer, RejectionReasonEng, RejectionHeaderGer, RejectionHeaderEng, PaymentOrApprovalGer, PaymentOrApprovalEng, CCCNotification, CCCNotificationEng, FWFDashboardAccount As String

Zeilenauswahl:

    Select Case MsgBox("Zeile " & ActiveCell.Row & " ist ausgewählt. Übernehmen?", vbYesNoCancel) '3 Wege: Auswahl übernehmen, Zeile eingeben oder abbrechen
        
        Case vbYes 'Übernehmen
            masterlist = ActiveCell.Row
        
        Case vbNo 'Auswahl eingeben mit Checks
            
Quelleneingabe:

            masterlist = Application.InputBox("Bitte zu verwendende Quellzeile eingeben:")
   
            If masterlist = False Then 'Beim Abbrechen beenden
                GoTo Ende
            ElseIf IsNumeric(masterlist) = False Then 'Check, dass Wert eine Zahl ist
                MsgBox "Zahlenwert erwartet!", vbOKOnly
            GoTo Quelleneingabe
            End If
            
        Case vbCancel 'Abbrechen
            GoTo Ende
    
    End Select
                            

   Windows("OAO Funding Masterlist.xlsm").Activate
      
   If (Cells(masterlist, 3).Value = "" And type_of_charge <> "OA support") Then 'Check, dass nur ausgefüllte Zeile ausgewählt wird - Ausnahme: OA-Infrastrukturkosten
       MsgBox "Quellzeile ist leer!", vbOKOnly
       GoTo Zeilenauswahl
   End If
    

'0.4 Variable Werte aus PubFonds
''''''''''''''''''''''''''''''''
  
    id = Cells(masterlist, 1)
    type_of_charge = Cells(masterlist, 2)
    limit_amount = Cells(masterlist, 34)
    open_access_deal = Cells(masterlist, 24)
    funder = Cells(masterlist, 7)
    article_id = Cells(masterlist, 14)
    due_date = Cells(masterlist, 45)
    publisher = Cells(masterlist, 5)
    corresponding_author = Cells(masterlist, 3)
    title = Cells(masterlist, 4)
    title = Trim(Replace(Replace(title, Chr(10), ""), Chr(13), "")) 'Whitespace entfernen
    source_full_title = Cells(masterlist, 6)
    source_full_title = Trim(Replace(Replace(source_full_title, Chr(10), ""), Chr(13), "")) 'Whitespace entfernen
    doi = Cells(masterlist, 8)
    license_ref = Cells(masterlist, 9)
    echeck_date = Cells(masterlist, 23)
    echeck_status = Cells(masterlist, 30)
    reject_reason = Cells(masterlist, 31)
    invoice_status = Cells(masterlist, 42)
    InvoiceNr = Cells(masterlist, 51)
    euro = -1 * (Val(Cells(masterlist, 50)))
    If Cells(masterlist, 11) = "YES" Then
        doaj = True
        Else: doaj = False
    End If
    If Cells(masterlist, 12) = "TRUE" Then
        is_hybrid = True
        Else: is_hybrid = False
    End If
    oa_status = Cells(masterlist, 13)
    If Cells(masterlist, 25) = "ja" Then
        affiliated = True
        Else: affiliated = False
    End If
    
'''''''''''''''''''''''''''''
' 1 Ablehnungen an Autor*in '
'''''''''''''''''''''''''''''

' 1.1 BPC
'''''''''

'endif siehe 1.2

    If open_access_deal = "no agreement" Then 'PubFonds-Rejections
    
        If type_of_charge = "Book (BPC)" Then 'BPC-Rejection
            
            'Deutsch
            
            EMailGenerate "S.g. NNNNN," & vbCrLf & vbCrLf & _
            "leider können Monografien nicht aus zentralen Mitteln für das Open-Access-Publizieren gefördert werden. Wir möchten in diesem Zusammenhang aber auf ein Förderprogramm des FWF hinweisen, das auch eine Open-Access-Publikation ermöglicht: https://www.fwf.ac.at/de/forschungsfoerderung/fwf-programme/selbststaendige-publikationen/" & vbCrLf & vbCrLf & _
            "Sollten Sie noch offene Fragen haben, stehen wir gerne zur Verfügung." & vbCrLf & vbCrLf & _
            "Mit besten Grüßen" & vbCrLf & vbCrLf & _
            "Guido Blechl / Bernhard Schubert / Klara Schellander"
            
            'Englisch
            
            EMailGenerate "Dear NNNN," & vbCrLf & vbCrLf & _
            "unfortunately monographs cannot be covered by the central Open Access Publishing Fund. We would like to mention an FWF funding programme that does allow OA publication for monographs: https://www.fwf.ac.at/en/research-funding/fwf-programmes/stand-alone-publications/" & vbCrLf & vbCrLf & _
            "Please do not hesitate to ask should you have any further questions." & vbCrLf & vbCrLf & _
            "Kind regards" & vbCrLf & vbCrLf & _
            "Guido Blechl / Bernhard Schubert / Klara Schellander"
        
        End If
        
' 1.2 Hybrid außerhalb von Abkommen
'''''''''''''''''''''''''''''''''''

'nested if, siehe 1.1
        
        If doaj = False And is_hybrid = True Then 'Hybrid-Rejection
        
            UFind corresponding_author 'Suche nach corresponding_author in u:find
        
            'Deutsch
        
            EMailGenerate "S.g. NNNNN," & vbCrLf & vbCrLf & _
            "vielen Dank für Ihren Antrag auf Förderung des Artikels """ & title & """." & vbCrLf & vbCrLf & _
            "Leider können wir Ihre Publikation in der Zeitschrift """ & source_full_title & """ nicht fördern, da sie in einem sogenannten ""Hybrid-Journal"" (= Subskriptionsjournal, das den Freikauf einzelner Artikel anbietet) erscheinen soll und dies gemäß Open Access Policy der Universität Wien und gemäß Förderkriterium 2a (http://openaccess.univie.ac.at/foerderkriterien) aus dem zentralen Publikationsfonds prinzipiell nicht unterstützt wird. Bitte haben Sie Verständnis, dass die Universität Wien Hybrid-Modelle nur im Rahmen von Spezialabkommen mit Verlagen fördert (siehe auch: https://openaccess.univie.ac.at/foerderungen/oa-verlagsabkommen/)." & vbCrLf & vbCrLf & _
            "Dem Verzeichnis SHERPA-RoMEO (https://v2.sherpa.ac.uk/cgi/search/publication/basic?publication_title-auto=" & source_full_title & ") entnehmen wir, dass die Policy von """ & source_full_title & """ es erlaubt, die NNNNNNNN--Version--NNNNNN Ihres Artikels NNNNNNNNN--nach x Monaten--NNNNNNNN über NNNNNNNNN--das institutionelle Repositorium u:scholar (https://uscholar.univie.ac.at/) oder ein Fachrepositorium--NNNNNNNN frei zugänglich zu machen, sofern dies eine Option für Sie darstellt." & vbCrLf & vbCrLf & _
            "Sollten Sie dazu oder zu anderen Open-Access-Themen noch Fragen haben, so helfen wir Ihnen gerne weiter!" & vbCrLf & vbCrLf & _
            "Mit freundlichen Grüßen" & vbCrLf & vbCrLf & _
            "Guido Blechl / Bernhard Schubert / Klara Schellander"
            
            'Englisch
            
            EMailGenerate "Dear NNNNN," & vbCrLf & vbCrLf & _
            "thank you for your request to fund the article """ & title & """." & vbCrLf & vbCrLf & _
            "We regret to inform you that we cannot fund your publication in """ & source_full_title & """ since it is to appear in a so-called ""hybrid journal"" (= subscription journal that makes individual articles Open Access for a fee), which is generally not supported according to the Open Access Policy of the University of Vienna and according to funding criterion 2a (http://openaccess.univie.ac.at/en/funding/oa-publishing-fund/) of the Central Open Access Publishing Fund. Please understand that the University of Vienna supports hybrid publication models only if they are part of special agreements with publishers (see also: https://openaccess.univie.ac.at/en/funding/oa-publishing-agreements/)." & vbCrLf & vbCrLf & _
            "According to the SHERPA-RoMEO (https://v2.sherpa.ac.uk/cgi/search/publication/basic?publication_title-auto=" & source_full_title & ") directory the policy of """ & source_full_title & """ allows making the NNNNNNNN--Version--NNNNNN of your article freely avalaible NNNNNNNNN--after x months--NNNNNNNN via NNNNNNNN--the institutional repository u:scholar (https://uscholar.univie.ac.at/) or a subject repository--NNNNNNNN." & vbCrLf & vbCrLf & _
            "Should you have any futher questions on this or other topics related to Open Access please do not hesitate to contact us!" & vbCrLf & vbCrLf & _
            "Kind regards" & vbCrLf & vbCrLf & _
            "Guido Blechl / Bernhard Schubert / Klara Schellander"
        
        End If
    
    End If

' 1.3 Limit erreicht (BMC, Frontiers, MDPI, Publikationsfonds)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If (publisher = "Frontiers" Or publisher = "MDPI" Or publisher = "BMC" Or open_access_deal = "no agreement") And (reject_reason = "limit reached") Then 'Limit erreicht
        
        UFind corresponding_author 'Suche nach corresponding_author in u:find
        
        'Deutsch
                               
        If open_access_deal = "no agreement" Then
            url = "https://openaccess.univie.ac.at/publikationsfonds/"
        Else
            url = "https://openaccess.univie.ac.at/" & LCase(publisher)
        End If
        
        EMailGenerate "Ihre neueste Einreichung bei " & publisher & ": Fördergrenze erreicht" & vbCrLf & vbCrLf & _
        "S.g. NNNN NNNNNN," & vbCrLf & vbCrLf & _
        "unglücklicherweise müssen wir Ihnen mitteilen, dass wir die Kosten für Ihre neueste " & publisher & "-Einreichung """ & title & """ (und ggf. weitere Einreichungen im aktuellen Jahr) nicht übernehmen können. Die finanziellen Mittel in unserem Publikationsfonds sind begrenzt, weshalb es eine Obergrenze von drei Artikeln pro Jahr gibt (siehe dazu " & url & ")." & vbCrLf & vbCrLf & _
        "Selbstverständlich übernehmen wir die Kosten für die übrigen Artikel, die wir in diesem Jahr bestätigt haben. Teilen Sie uns bitte mit, falls aktuelle Einreichungen nicht angenommen werden - in diesem Fall zählen diese nicht zu Ihrem Artikellimit." & vbCrLf & vbCrLf & _
        "Sollten Sie Fragen haben, melden Sie sich bitte." & vbCrLf & vbCrLf & _
        "Mit freundlichen Grüßen," & vbCrLf & vbCrLf & _
        "Guido Blechl / Bernhard Schubert / Klara Schellander"
        
        'Englisch
                               
        If open_access_deal = "no agreement" Then
            url = "https://openaccess.univie.ac.at/en/publikationsfonds/"
        Else
            url = "https://openaccess.univie.ac.at/en/" & LCase(publisher)
        End If
        
        EMailGenerate "Your latest " & publisher & " submission: Funding limit reached" & vbCrLf & vbCrLf & _
        "Dear NNNN NNNNNN," & vbCrLf & vbCrLf & _
        "unfortunately we have to inform you that we are unable to cover the costs for your latest " & publisher & " submission """ & title & """ (and possibly other submissions in the current year). The financial resources of our OA publishing fund are limited, which is why there is a funding cap of three publications per author per year (see " & url & ")." & vbCrLf & vbCrLf & _
        "Please note that we will of course cover the costs for the articles that have already been confirmed. In case the publisher does not accept your contributions please let us know so we can reallocate the funds set aside and the article does not count towards your funding limit." & vbCrLf & vbCrLf & _
        "If you have any questions please do not hesitate to contact us." & vbCrLf & vbCrLf & _
        "Kind regards," & vbCrLf & vbCrLf & _
        "Guido Blechl / Bernhard Schubert / Klara Schellander"
                
    End If

' 1.4 FWF-Ablehnung und -Zuweisungen
''''''''''''''''''''''''''''''''''''

'endif siehe 1.6

    If (publisher = "Frontiers" Or publisher = "ACS" Or publisher = "IOP" Or publisher = "MDPI" Or publisher = "T&F" Or publisher = "Elsevier" Or publisher = "OUP" Or publisher = "CUP" Or publisher = "Wiley" Or publisher = "Springer" Or publisher = "BMC") And (reject_reason = "FWF funded" Or reject_reason = "EU funded" Or reject_reason = "author not affiliated at relevant date") Then 'Frontiers-/ACS-/IOP-/MDPI-/T&F-/Elsevier-/OUP-Ablehungsemails
       
       RejectionHeaderGer = "Förderabsage"
       RejectionHeaderEng = "Funding declined for"
       PaymentOrApprovalGer = "Übernahme der Publikationskosten"
       PaymentOrApprovalEng = "cover the publishing charges for"
       
        If reject_reason = "FWF funded" Then
            PaymentOrApprovalGer = "Bestätigung im Rahmen unseres Verlagsabkommens"
            PaymentOrApprovalEng = "approve under our publishing agreement"
            If publisher = "BMC" Then
                RejectionHeaderGer = "Förderabsage"
                RejectionHeaderEng = "Funding declined for"
                RejectionReasonGer = "Wir können Ihre Publikation nicht im Rahmen unseres Abkommens bestätigen, da gemäß der Artikelmetadaten ein FWF Funding vorliegt (" & funder & ") und seitens der Universität Wien deshalb keine Abwicklung möglich ist (siehe https://openaccess.univie.ac.at/" & LCase(publisher) & "). Bitte kontaktieren Sie zur Übernahme der Publikationskosten den FWF direkt: publikationskosten@fwf.ac.at. Ein Ansuchen auf Refundierung ist in diesem Zusammenhang nicht erforderlich, da der Artikel über ein Abkommen zwischen dem Verlag und dem FWF gefördert wird (siehe https://www.fwf.ac.at/de/forschungsfoerderung/fwf-programme/referierte-publikationen)."
                RejectionReasonEng = "We cannot approve your publication as part of our agreement since the article metadata indicate FWF funding (" & funder & ") and the University of Vienna cannot process articles in this case (see https://openaccess.univie.ac.at/en/" & LCase(publisher) & "). Please contact the FWF directly in order to have the charges covered: publikationskosten@fwf.ac.at. A refund request is not necessary in this case since the article is covered by an agreement between the publisher and the FWF (see https://www.fwf.ac.at/en/forschungsfoerderung/fwf-programme/referierte-publikationen)."
            Else
                RejectionHeaderGer = "Zuweisung an den FWF"
                RejectionHeaderEng = "Referral to the FWF of"
                If publisher = "T&F" Then
                    RejectionReasonGer = "Wir können Ihre Publikation nicht im Rahmen unseres Abkommens bestätigen, da gemäß der Artikelmetadaten ein FWF Funding vorliegt (" & funder & ") und seitens der Universität Wien deshalb keine Abwicklung möglich ist (siehe https://openaccess.univie.ac.at/taylor-francis). Wir haben deshalb veranlasst, dass " & publisher & " den Artikel dem FWF zur Bestägigung zuweist. Ein Ansuchen auf Refundierung ist in diesem Zusammenhang nicht erforderlich, da der Artikel über ein Abkommen zwischen dem Verlag und dem FWF gefördert wird (siehe https://www.fwf.ac.at/de/forschungsfoerderung/fwf-programme/referierte-publikationen)."
                    RejectionReasonEng = "We cannot approve your publication as part of our agreement since the article metadata indicate FWF funding (" & funder & ") and the University of Vienna cannot process articles in this case (see https://openaccess.univie.ac.at/en/taylor-francis). We have therefore asked " & publisher & " to forward the article to the FWF for approval. A refund request is not necessary in this case since the article is covered by an agreement between the publisher and the FWF (see https://www.fwf.ac.at/en/forschungsfoerderung/fwf-programme/referierte-publikationen)."
                Else
                    RejectionReasonGer = "Wir können Ihre Publikation nicht im Rahmen unseres Abkommens bestätigen, da gemäß der Artikelmetadaten ein FWF Funding vorliegt (" & funder & ") und seitens der Universität Wien deshalb keine Abwicklung möglich ist (siehe https://openaccess.univie.ac.at/" & LCase(publisher) & "). Wir haben deshalb veranlasst, dass " & publisher & " den Artikel dem FWF zur Bestägigung zuweist. Ein Ansuchen auf Refundierung ist in diesem Zusammenhang nicht erforderlich, da der Artikel über ein Abkommen zwischen dem Verlag und dem FWF gefördert wird (siehe https://www.fwf.ac.at/de/forschungsfoerderung/fwf-programme/referierte-publikationen)."
                    RejectionReasonEng = "We cannot approve your publication as part of our agreement since the article metadata indicate FWF funding (" & funder & ") and the University of Vienna cannot process articles in this case (see https://openaccess.univie.ac.at/en/" & LCase(publisher) & "). We have therefore asked " & publisher & " to forward the article to the FWF for approval. A refund request is not necessary in this case since the article is covered by an agreement between the publisher and the FWF (see https://www.fwf.ac.at/en/forschungsfoerderung/fwf-programme/referierte-publikationen)."
                End If
            End If
            
' 1.5 EU-Ablehnungen
''''''''''''''''''''

'nested if, siehe 1.4
            
        ElseIf reject_reason = "EU funded" Then
            RejectionReasonGer = "Leider können wir Ihre Publikation nicht fördern, da gemäß der Artikelmetadaten ein EU Funding vorliegt (" & funder & ") und seitens der Universität Wien deshalb keine Förderung möglich ist (siehe https://openaccess.univie.ac.at/" & LCase(publisher) & ")."
            RejectionReasonEng = "Unfortunately we cannot cover the charges since the article metadata indicate EU funding (" & funder & ") and the University of Vienna cannot provide funding in this case (see https://openaccess.univie.ac.at/en/" & LCase(publisher) & ")."
        
' 1.6 Nicht affiliiert zum Stichtag
'''''''''''''''''''''''''''''''''''

'nested if, siehe 1.4
        
        ElseIf affiliated = False Then
            PaymentOrApprovalGer = "Bestätigung im Rahmen unseres Verlagsabkommens"
            PaymentOrApprovalEng = "approve under our publishing agreement"
            If (publisher = "ACS" Or publisher = "IOP" Or publisher = "T&F" Or publisher = "OUP" Or publisher = "Wiley" Or publisher = "Springer") Then 'Relevanter Zeitpunkt ist acceptance
                If publisher = "T&F" Then
                    'RejectionReasonGer = "Leider können wir Ihre Publikation nicht fördern, da Sie zum Zeitpunkt der Acceptance nicht Angehörige*r der Universität Wien waren und seitens der Universität deshalb keine Förderung möglich ist (siehe https://openaccess.univie.ac.at/taylor-francis)."
                    'RejectionReasonEng = "Unfortunately we cannot cover the charges since you were not affiliated with the University of Vienna at the date of acceptance and the University cannot provide funding in this case (see https://openaccess.univie.ac.at/en/taylor-francis)."
                    RejectionReasonGer = "Leider können wir Ihre Publikation nicht im Rahmen unseres Abkommens bestätigen, da Sie zum Zeitpunkt der Acceptance nicht Angehörige*r der Universität Wien waren und seitens der Universität deshalb keine Abwicklung möglich ist (siehe https://openaccess.univie.ac.at/taylor-francis)."
                    RejectionReasonEng = "Unfortunately we cannot approve your publication as part of our agreement since you were not affiliated with the University of Vienna at the date of acceptance and the University cannot process articles in this case (see https://openaccess.univie.ac.at/en/taylor-francis)."
                Else
                    RejectionReasonGer = "Leider können wir Ihre Publikation nicht im Rahmen unseres Abkommens bestätigen, da Sie zum Zeitpunkt der Acceptance nicht Angehörige*r der Universität Wien waren und seitens der Universität deshalb keine Abwicklung möglich ist (siehe https://openaccess.univie.ac.at/" & LCase(publisher) & ")."
                    RejectionReasonEng = "Unfortunately we cannot approve your publication as part of our agreement since you were not affiliated with the University of Vienna at the date of acceptance and the University cannot process articles in this case (see https://openaccess.univie.ac.at/en/" & LCase(publisher) & ")."
                End If
            ElseIf publisher = "CUP" Then
                RejectionReasonGer = "Leider können wir Ihre Publikation nicht fördern, da Sie zum Zeitpunkt der Acceptance nicht Angehörige*r der Universität Wien waren und seitens der Universität deshalb keine Förderung möglich ist (siehe https://openaccess.univie.ac.at/" & LCase(publisher) & "). Sollten Sie keine Mittel für Open Access zur Verfügung haben, schreiben Sie bitte an oaqueries@cambridge.org, damit Ihr Auftrag storniert wird (der Artikel wird dann Closed Access veröffentlicht)."
                RejectionReasonEng = "Unfortunately we cannot cover the charges since you were not affiliated with the University of Vienna at the date of acceptance and the University cannot provide funding in this case (see https://openaccess.univie.ac.at/en/" & LCase(publisher) & "). If you do not have funds available for Open Access please write to oaqueries@cambridge.org to cancel your request (the article will then be published Closed Access)."
            ElseIf (publisher = "MDPI" Or publisher = "Frontiers" Or publisher = "BMC") Then 'Relevanter Zeitpunkt ist Submission
                RejectionReasonGer = "Leider können wir Ihre Publikation nicht fördern, da Sie zum Zeitpunkt der Einreichung nicht Angehörige*r der Universität Wien waren und seitens der Universität deshalb keine Förderung möglich ist (siehe https://openaccess.univie.ac.at/" & LCase(publisher) & ")."
                RejectionReasonEng = "Unfortunately we cannot cover the charges since you were not affiliated with the University of Vienna at the date of submission and the University cannot provide funding in this case (see https://openaccess.univie.ac.at/en/" & LCase(publisher) & ")."
            ElseIf publisher = "Elsevier" Then
                RejectionReasonGer = "Leider können wir Ihre Publikation nicht fördern, da Sie zum Zeitpunkt der Einreichung nicht Angehörige*r der Universität Wien waren und seitens der Universität deshalb keine Förderung möglich ist (siehe https://openaccess.univie.ac.at/" & LCase(publisher) & "). Sollten Sie keine Mittel für Open Access zur Verfügung haben, können Sie die Rechnung, die Ihnen ausgestellt wird, binnen zwei Wochen ab Rechnungsdatum stornieren (der Artikel wird dann Closed Access veröffentlicht)."
                RejectionReasonEng = "Unfortunately we cannot cover the charges since you were not affiliated with the University of Vienna at the date of submission and the University cannot provide funding in this case (see https://openaccess.univie.ac.at/en/" & LCase(publisher) & "). If you do not have funds available for Open Access you can cancel the invoice you will receive up until two weeks after the invoice date (the article will then be published Closed Access)."
            End If
        End If
        
        'corresponding_author = corresponding_author & " %2Binaktiv" 'Suche nach inaktivem corresponding_author in u:find
        UFind corresponding_author
        
        'Deutsch
        
        EMailGenerate RejectionHeaderGer & " für Ihren " & publisher & "-Artikel """ & title & """" & vbCrLf & vbCrLf & _
        "S.g. NNNNNNN," & vbCrLf & vbCrLf & _
        publisher & " hat uns Ihren Artikel """ & title & """ in der Zeitschrift """ & source_full_title & """ zur " & PaymentOrApprovalGer & " übermittelt." & vbCrLf & vbCrLf & _
        RejectionReasonGer & vbCrLf & vbCrLf & _
        "Sollten Sie noch offene Fragen haben, melden Sie sich bitte." & vbCrLf & vbCrLf & _
        "Mit freundlichen Grüßen" & vbCrLf & vbCrLf & _
        "Guido Blechl / Bernhard Schubert / Klara Schellander"
        
        'Englisch
        
        EMailGenerate RejectionHeaderEng & " your " & publisher & " article """ & title & """" & vbCrLf & vbCrLf & _
        "Dear NNNNNNN," & vbCrLf & vbCrLf & _
        publisher & " has asked us to " & PaymentOrApprovalEng & " your article """ & title & """ in """ & source_full_title & """." & vbCrLf & vbCrLf & _
        RejectionReasonEng & vbCrLf & vbCrLf & _
        "Please do not hesitate to ask should you have any questions." & vbCrLf & vbCrLf & _
        "Kind regards" & vbCrLf & vbCrLf & _
        "Guido Blechl / Bernhard Schubert / Klara Schellander"
    
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2 Rechnungslegung an Quästur und Zahlungsbestätigung an Autor*in (de Gruyter, Frontiers, MDPI, SAGE, Publikationsfonds) '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'endif siehe 2.2

        If invoice_status = "Invoiced" And (oa_status = "") Then 'Zahlungsanweisung & Autor*inneninfo
        
            If InvoiceNr = "" Or InvoiceNr = False Then
                MsgBox "Rechnungsnummer fehlt!", vbOKOnly
                GoTo Ende
            ElseIf due_date = "" Or due_date = False Then
                MsgBox "Zahlungsziel fehlt!", vbOKOnly
                GoTo Ende
            End If
            
            If (publisher = "SAGE") Or (publisher = "de Gruyter") Then 'Ergänzung zur reduzierten Gebühren für SAGE/de Gruyter
                PriceReductionGer = " (aufgrund des Verlagsabkommens mit stark reduzierten Publikationsgebühren)"
                PriceReductionEng = " (priced at a greatly reduced rate as part of our publishing agreement)"
            Else
                PriceReductionGer = ""
                PriceReductionEng = ""
            End If
            
' 2.1 Sofort zu zahlen
''''''''''''''''''''''

'nested if, siehe 2

            If due_date = "sofort" Then
                
                
                'Zahlungsanweisung, sofort zu zahlen
                
                EMailGenerate "IP150063: Open Access Rechnung (" & id & "/" & publisher & "/" & corresponding_author & "/" & InvoiceNr & "): Bitte einzahlen" & vbCrLf & vbCrLf & _
                "Liebe Kolleg*innen," & vbCrLf & vbCrLf & "anbei übermittle ich eine neue Rechnung, die wir via Kostenstelle IP150063 (OA Publikationsfonds) bezahlen:" & vbCrLf & vbCrLf & _
                ".) Ich bestätige hiermit die sachliche Richtigkeit." & vbCrLf & vbCrLf & _
                ".) Buchen unter IP150063" & vbCrLf & vbCrLf & _
                ".) *Bitte möglichst schnell und ohne Verzögerung einzahlen!* Ist sehr wichtig, damit der Artikel ohne Verzögerung freigeschaltet wird!" & vbCrLf & vbCrLf & _
                ".) Rückfragen: Guido Blechl (27607), Bernhard Schubert (27608), Klara Schellander (16660), Brigitte Kromp (27603)" & vbCrLf & vbCrLf & _
                "Vielen Dank und beste Grüße," & vbCrLf & _
                "Guido Blechl / Bernhard Schubert / Klara Schellander"
                
                UFind corresponding_author 'Suche nach corresponding_author in u:find
                
                'Zahlungsbestätigung, sofort zu zahlen, deutsch
                
                EMailGenerate "Rechnung in Zahlung (" & title & ")" & vbCrLf & vbCrLf & _
                "S.g. NNNN NNNNNN," & vbCrLf & vbCrLf & _
                "die Rechnung für den Artikel """ & title & """ in der Zeitschrift """ & source_full_title & """ wurde in den Rechnungslauf der Universität Wien eingebracht. Der Betrag sollte innerhalb der nächsten 5-10 Tage am Konto des Empfängers eingelangt sein. Die Bezahlung erfolgt aus Mitteln des zentralen Open-Access-Publikationsfonds" & PriceReductionGer & ". Von Ihrer Seite ist nichts weiter zu tun." & vbCrLf & vbCrLf & _
                "Mit freundlichen Grüßen," & vbCrLf & vbCrLf & _
                "Guido Blechl / Bernhard Schubert / Klara Schellander"
                
                'Zahlungsbestätigung, sofort zu zahlen, englisch
                
                EMailGenerate "Invoice Payment (" & title & ")" & vbCrLf & vbCrLf & _
                "Dear NNNN NNNNNN," & vbCrLf & vbCrLf & _
                "the invoice for the article """ & title & """ in """ & source_full_title & """ has been forwarded to the university's accounting office and will be authorised for payment within the next 5-10 days. Payment is provided by the central Open Access publishing fund" & PriceReductionEng & ". No further action is required on your part." & vbCrLf & vbCrLf & _
                "Kind regards," & vbCrLf & vbCrLf & _
                "Guido Blechl / Bernhard Schubert / Klara Schellander"
                            
' 2.2 Nicht sofort zu zahlen
''''''''''''''''''''''''''''

'nested if, siehe 2
                            
            Else 'nicht sofort zu zahlen
            
                EMailGenerate "IP150063: Open Access Rechnung (" & id & "/" & publisher & "/" & corresponding_author & "/" & InvoiceNr & "): Bitte einzahlen" & vbCrLf & vbCrLf & _
                "Liebe Kolleg*innen," & vbCrLf & vbCrLf & "anbei übermittle ich eine neue Rechnung, die wir via Kostenstelle IP150063 (OA Publikationsfonds) bezahlen:" & vbCrLf & vbCrLf & _
                ".) Ich bestätige hiermit die sachliche Richtigkeit." & vbCrLf & vbCrLf & _
                ".) Buchen unter IP150063" & vbCrLf & vbCrLf & _
                ".) Rückfragen: Guido Blechl (27607), Bernhard Schubert (27608), Klara Schellander (16660), Brigitte Kromp (27603)" & vbCrLf & vbCrLf & _
                "Vielen Dank und beste Grüße," & vbCrLf & _
                "Guido Blechl / Bernhard Schubert / Klara Schellander"
                
                'Zahlungsbestätigung, nicht sofort zu zahlen, deutsch
                
                UFind corresponding_author 'Suche nach corresponding_author in u:find
                
                EMailGenerate "Rechnung in Zahlung (" & title & ")" & vbCrLf & vbCrLf & _
                "S.g. NNNN NNNNNN," & vbCrLf & vbCrLf & _
                "die Rechnung für den Artikel """ & title & """ in der Zeitschrift """ & source_full_title & """ wurde in den Rechnungslauf der Universität Wien eingebracht und wird mit dem auf der Rechnung angeführten Zahlungsziel (" & due_date & ") angewiesen. Die Bezahlung erfolgt aus Mitteln des zentralen Open-Access-Publikationsfonds" & PriceReductionGer & ". Von Ihrer Seite ist nichts weiter zu tun." & vbCrLf & vbCrLf & _
                "Mit freundlichen Grüßen," & vbCrLf & vbCrLf & _
                "Guido Blechl / Bernhard Schubert / Klara Schellander"
                
                'Zahlungsbestätigung, nicht sofort zu zahlen, englisch
                
                EMailGenerate "Invoice Payment (" & title & ")" & vbCrLf & vbCrLf & _
                "Dear NNNN NNNNNN," & vbCrLf & vbCrLf & _
                "the invoice for the article """ & title & """ in """ & source_full_title & """ has been forwarded to the university's accounting office and will be authorised for payment on the due date specified (" & due_date & "). Payment is provided by the central Open Access publishing fund" & PriceReductionEng & ". No further action is required on your part." & vbCrLf & vbCrLf & _
                "Kind regards," & vbCrLf & vbCrLf & _
                "Guido Blechl / Bernhard Schubert / Klara Schellander"
                
            End If
        
        End If

'''''''''''''''''''''''''''''''''''''''''''''''''''
' 3 FWF-Klärung und -Zuweisung an FWF bzw. Verlag '
'''''''''''''''''''''''''''''''''''''''''''''''''''

' 3.1 FWF-Nachfrage
'''''''''''''''''''

'3.1.1 Elsevier, SAGE
'''''''''''''''''''''

    If (publisher = "Elsevier" Or publisher = "SAGE") And echeck_status = "pending" And funder <> "" Then 'FWF-Nachfrage für Elsevier-/SAGE-Artikel
    
        If license_ref = "" Then
            license_ref = "nicht angegeben"
        End If
    
        EMailGenerate "publikationskosten@fwf.ac.at" & vbCrLf & vbCrLf & _
        "OA Förderung " & publisher & "/" & corresponding_author & "/" & funder & vbCrLf & vbCrLf & _
        "Liebe Kolleg*innen," & vbCrLf & vbCrLf & _
        publisher & " hat uns den folgenden Artikel zur Prüfung für eine Open-Access-Förderung übermittelt:" & vbCrLf & vbCrLf & _
        "> """ & title & """" & vbCrLf & _
        "> Corresponding author: " & corresponding_author & vbCrLf & _
        "> Article Submitted: NNNN-Datum-NNNN" & vbCrLf & _
        "> DOI: " & doi & vbCrLf & _
        "> Lizenz: " & license_ref & vbCrLf & _
        "> FWF Projekt: " & funder & vbCrLf & vbCrLf & _
        "Wurde dieser Artikel bereits durch den FWF abgelehnt oder hat noch keine Prüfung stattgefunden? Im zweiten Fall würden wir an " & publisher & " schreiben, damit der Beitrag dem FWF zugeordnet wird." & vbCrLf & vbCrLf & _
        "Vielen Dank und beste Grüße" & vbCrLf & vbCrLf & _
        "Guido Blechl / Bernhard Schubert / Klara Schellander"
        
    End If
    
' 3.2 FWF-Zuweisung
'''''''''''''''''''

' 3.2.1 BMC, Frontiers, Elsevier, IOP, SAGE, Wiley
''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If (publisher = "Wiley" Or publisher = "Elsevier" Or publisher = "Frontiers" Or publisher = "IOP" Or publisher = "SAGE" Or publisher = "BMC") And reject_reason = "FWF funded" Then 'Wiley/Elsevier/Frontiers/IOP/SAGE/BMC Artikel an FWF
    
        FWFDashboardAccount = "dashboard for the eligibility check"
    
        If publisher = "Elsevier" Then 'Variable für E-Mail-Adresse von Elsevier, bei Wiley kann auf das Benachrichtigungs-E-Mail geantwortet werden
            PublisherContact = "agreementactivation@elsevier.com"
        ElseIf publisher = "Frontiers" Then
            PublisherContact = "institutions@frontiersin.org"
            doi = article_id 'Kein DOI bei Frontiers-Artikeln vorhanden
        ElseIf publisher = "Wiley" Then
            PublisherContact = ""
        ElseIf publisher = "IOP" Then
            PublisherContact = ""
            doi = article_id 'Kein DOI bei IOP-Artikeln vorhanden
        ElseIf publisher = "SAGE" Then
            PublisherContact = ""
        ElseIf publisher = "BMC" Then
            PublisherContact = ""
            doi = article_id 'Kein DOI bei BMC-Artikeln vorhanden
            FWFDashboardAccount = "account"
        End If
            
        If publisher = "BMC" Then
            EMailGenerate "Dear " & publisher & " Support," & vbCrLf & vbCrLf & _
            "Thank you for checking. Please remove the Manuscript ID " & doi & " (see below) from our account. Reason: FWF funded (" & funder & "; see our funding requirements: https://openaccess.univie.ac.at/en/bmc/). Author can apply for funding at FWF: publikationskosten@fwf.ac.at" & vbCrLf & vbCrLf & _
            "If you or the author have any further questions, please do not hesitate to contact us." & vbCrLf & vbCrLf & _
            "Kind regards" & vbCrLf & vbCrLf & _
            "Guido Blechl / Bernhard Schubert / Klara Schellander"
        Else
            EMailGenerate PublisherContact & vbCrLf & "publikationskosten@fwf.ac.at" & vbCrLf & vbCrLf & _
            "Austrian OA Agreement: Please assign " & doi & " to FWF dashboard" & vbCrLf & vbCrLf & _
            "Dear " & publisher & " Support," & vbCrLf & vbCrLf & _
            "We cannot approve the article " & doi & " (" & title & ") due to FWF funding (" & funder & ")." & vbCrLf & vbCrLf & _
            "Could you please reassign the article to the FWF " & FWFDashboardAccount & ". Please confirm." & vbCrLf & vbCrLf & _
            "Kind regards," & vbCrLf & vbCrLf & _
            "Guido Blechl / Bernhard Schubert / Klara Schellander"
        End If
        
    End If

' 3.2.2 MDPI
''''''''''''
    
    If publisher = "MDPI" And reject_reason = "FWF funded" Then 'MDPI Artikel an FWF
    
        EMailGenerate "publikationskosten@fwf.ac.at" & vbCrLf & vbCrLf & _
        "Dear MDPI Support and NNNNNNNNNNNNNN-Author-NNNNNNNNNNNNNNNN," & vbCrLf & vbCrLf & _
        "We cannot approve the manuscript """ & article_id & """ (" & title & ") due to FWF funding (Austrian Science Fund " & funder & ")." & vbCrLf & vbCrLf & _
        "@MDPI: Could you please reassign the article to the FWF dashboard for the eligibility check. Please confirm." & vbCrLf & vbCrLf & _
        "Kind regards," & vbCrLf & vbCrLf & _
        "Guido Blechl / Bernhard Schubert / Klara Schellander"
        
    End If

''''''''''''''''''''''''''''''''''''''''''''''''
' 4 Abkommen-spezifische Nachrichten an Verlag '
''''''''''''''''''''''''''''''''''''''''''''''''

' 4.1 BMC
'''''''''

' 4.1.1 Anforderung Information
'''''''''''''''''''''''''''''''

'endif siehe 4.1.4
    
    If publisher = "BMC" Then
        If echeck_status = "pending" Then 'BMC funding Nachfrage
            EMailGenerate "Dear BMC Team," & vbCrLf & vbCrLf & _
            "Please provide us with funding information from the article metadata and the acknowledgements section from the article manuscript. As per our funding criteria we can only cover charges for articles that have not resulted from external funding." & vbCrLf & vbCrLf & _
            "Kind regards," & vbCrLf & vbCrLf & _
            "Guido Blechl / Bernhard Schubert / Klara Schellander"

' 4.1.2 Ablehnung wegen Nichtaffiliation
''''''''''''''''''''''''''''''''''''''''

'nested if, siehe 4.1.1
            
        ElseIf (reject_reason = "external affiliation") Or (reject_reason = "author not affiliated at relevant date") Then 'BMC Reject bei Nichtaffiliation
            EMailGenerate "Dear NNNN," & vbCrLf & vbCrLf & _
            "Please remove the Manuscript ID " & article_id & " (see below) from our account, because the submitting author (" & corresponding_author & ") is not affiliated with the University of Vienna." & vbCrLf & vbCrLf & _
            "If you or the author have any further questions, please do not hesitate to contact us." & vbCrLf & vbCrLf & _
            "Kind regards" & vbCrLf & vbCrLf & _
            "Guido Blechl / Bernhard Schubert / Klara Schellander"
        
' 4.1.3 Ablehnung wegen EU funding
''''''''''''''''''''''''''''''''''

'nested if, siehe 4.1.1
        
        ElseIf reject_reason = "EU funded" Then 'BMC EU forwarding
            EMailGenerate "Dear NNNN," & vbCrLf & vbCrLf & _
            "Thank you for checking. Please remove the Manuscript ID " & article_id & " (see below) from our account. Reason: EU funded (see our funding requirements: https://openaccess.univie.ac.at/en/bmc/)." & vbCrLf & vbCrLf & _
            "If you or the author have any further questions, please do not hesitate to contact us." & vbCrLf & vbCrLf & _
            "Kind regards" & vbCrLf & vbCrLf & _
            "Guido Blechl / Bernhard Schubert / Klara Schellander"
        
' 4.1.4 Bestätigung
'''''''''''''''''''

'nested if, siehe 4.1.1
        
        ElseIf echeck_status = "approved" Then 'BMC funding Bestätigung
            EMailGenerate "Dear NNNN," & vbCrLf & vbCrLf & _
            "thanks for letting us know, we will cover the charges." & vbCrLf & vbCrLf & _
            "Kind regards" & vbCrLf & vbCrLf & _
            "Guido Blechl / Bernhard Schubert / Klara Schellander"
        End If
    End If
        
' 4.2 IOP
'''''''''
            
' 4.2.1 Ablehnung wegen Nichtaffiliation
''''''''''''''''''''''''''''''''''''''''

'endif siehe 4.2.2
            
    If publisher = "IOP" And open_access_deal = "transformative agreement" Then 'IOP Hybrid approve und reject
            
        If affiliated = False Then 'Reject weil nicht affiliated
            EMailGenerate "Dear NNNNNNNN," & vbCrLf & vbCrLf & _
            "Thank you for your notification. This article is not eligible. Reason: Corresponding author is not affiliated with the University of Vienna." & vbCrLf & vbCrLf & _
            "If you or the author have any further questions, please do not hesitate to contact us." & vbCrLf & vbCrLf & _
            "Kind regards" & vbCrLf & vbCrLf & _
            "Guido Blechl / BErnhard Schubert / Klara Schellander"
            
' 4.2.2 Bestätigung
'''''''''''''''''''

'nested if, siehe 4.2.1
        
        ElseIf echeck_status = "approved" And invoice_status = "Zusage" Then 'Bestätigung
            EMailGenerate "Dear NNNNNNNN," & vbCrLf & vbCrLf & _
            "the article """ & title & """ qualifies for inclusion in our Open Access agreement." & vbCrLf & vbCrLf & _
            "Kind regards" & vbCrLf & vbCrLf & _
            "Guido Blechl / Bernhard Schubert / Klara Schellander"
            
            
        End If
    
    End If
        
        
' 4.3 T&F
'''''''''

' 4.3.1 Approval disabled
'''''''''''''''''''''''''
        
    If publisher = "T&F" And echeck_status = "pending" Then 'T&F, approval disabled
        EMailGenerate "APC@tandf.co.uk" _
        & vbCrLf & _
        "Approval disabled for " & corresponding_author & " - DOI: " & doi _
        & vbCrLf & vbCrLf & _
        "Dear APC Team," _
        & vbCrLf & vbCrLf & _
        "Our dashboard states ""Approval has been disabled for this article, please contact apc@tandf.co.uk"". This article is eligible, could you please enable approval or approve the article manually to be included in our agreement?" _
        & vbCrLf & vbCrLf & _
        "Kind regards," _
        & vbCrLf & vbCrLf & _
        "Guido Blechl / Bernhard Schubert / Klara Schellander"
    End If
        
    
' 4.4 de Gruyter, SAGE
''''''''''''''''''''''

' 4.4.1 Reklamation: Artikel nicht OA
'''''''''''''''''''''''''''''''''''''
    
    If (publisher = "SAGE" Or publisher = "de Gruyter") And oa_status = "NOT OA" Then 'SAGE/de Gruyter Reklamation, wenn Artikel nicht OA
    
        If publisher = "SAGE" Then 'Variable für E-Mail-Adresse von SAGE
            PublisherContact = "openaccess@sagepub.com"
        End If
            
        EMailGenerate PublisherContact & vbCrLf & vbCrLf & _
        "Uni Vienna agreement: Approved article " & doi & " not available Open Access" & vbCrLf & vbCrLf & _
        "Dear " & publisher & " OA team," & vbCrLf & vbCrLf & _
        "we approved the following article for Open Access funding on " & echeck_date & ":" & vbCrLf & vbCrLf & _
        "> title: " & title & vbCrLf & _
        "> corresponding author: " & corresponding_author & vbCrLf & _
        "> DOI: " & doi & vbCrLf & _
        "> article ID: " & article_id & vbCrLf & vbCrLf & _
        "Can you please let us know why this article is not available Open Access yet?" & vbCrLf & vbCrLf & _
        "Kind regards," & vbCrLf & vbCrLf & _
        "Guido Blechl / Bernhard Schubert / Klara Schellander"
        
    End If
    
    
'''''''''''''''''''''''''''''''''''''''''''''''
' 5 Hinweis auf Abkommen/Retro-OA an Autor*in '
'''''''''''''''''''''''''''''''''''''''''''''''

' 5.1 ACS, de Gruyter
'''''''''''''''''''''

    If ((publisher = "ACS" And doaj <> True) Or publisher = "de Gruyter") And echeck_status = "pending" Then 'ACS/de Gruyter Nachfrage
                   
        UFind corresponding_author 'Suche nach corresponding_author in u:find
        
               
        If publisher = "ACS" Then
            CCCNotification = "es entstehen Ihnen keine Kosten, da die Open-Access-Gebühren pauschal in unserem Verlagsvertrag inkludiert sind"
            CCCNotificationEng = "you will not incur any costs since any Open Access charges are already included in our contract sum"
        ElseIf publisher = "de Gruyter" Then
            CCCNotification = "die stark reduzierten Publikationskosten werden auf Basis des Vertrags mit de Gruyter von uns übernommen"
            CCCNotificationEng = "the greatly reduced publishing charges will be covered by us on the basis of our agreement with de Gruyter"
        End If
        
        'Deutsch
        
        url = "https://openaccess.univie.ac.at/" & LCase(Application.WorksheetFunction.Trim(publisher))
        url = Replace(url, " ", "")
        
        EMailGenerate "Open Access für Ihren " & publisher & "-Artikel """ & title & """" & vbCrLf & vbCrLf & _
        "S.g. NNN," & vbCrLf & vbCrLf & _
        "wir wurden von " & publisher & " darüber informiert, dass folgende Publikation über das Open-Access-Verlagsabkommen der Universität Wien gefördert werden könnte:" & vbCrLf & vbCrLf & _
        "> Manuscript Details" & vbCrLf & _
        "> " & doi & vbCrLf & _
        "> " & source_full_title & vbCrLf & _
        "> " & title & vbCrLf & vbCrLf & _
        "Wir würden uns freuen, wenn Sie dieses Angebot wahrnehmen würden -- " & CCCNotification & ". Falls Sie sich für Open Access entscheiden, folgen Sie bitte dem Link ""Click here"" im Acceptance-E-Mail von " & publisher & " und wählen Sie in Folge ""Seek funding from Universitat Wien"", damit die weitere Abwicklung von uns übernommen werden kann. (Da das Abkommen über das Copyright Clearance Center abgewickelt wird, ist eine einmalige Registrierung notwendig, um den Prozess abzuschließen.) Unsere Informationen zum Förderabkommen finden Sie unter " & url & "." & vbCrLf & vbCrLf & _
        "Sollten Sie noch Fragen haben, melden Sie sich bitte." & vbCrLf & vbCrLf & _
        "Mit freundlichen Grüßen," & vbCrLf & vbCrLf & _
        "Guido Blechl / Bernhard Schubert / Klara Schellander"
        
        'Englisch
                                
        url = "https://openaccess.univie.ac.at/en/" & LCase(Application.WorksheetFunction.Trim(publisher))
        url = Replace(url, " ", "")
                                
        EMailGenerate "Open Access for your " & publisher & " article """ & title & """" & vbCrLf & vbCrLf & _
        "Dear NNN," & vbCrLf & vbCrLf & _
        "we were informed by " & publisher & " that the publication below is eligible for Open Access funding as part of a publishing agreement with the University of Vienna:" & vbCrLf & vbCrLf & _
        "> Manuscript Details" & vbCrLf & _
        "> " & doi & vbCrLf & _
        "> " & source_full_title & vbCrLf & _
        "> " & title & vbCrLf & vbCrLf & _
        "We would be delighted if you would accept this offer -- " & CCCNotificationEng & ". If you opt for Open Access please follow the ""Click here"" link in " & publisher & "'s acceptance e-mail and choose to ""Seek funding from Universitat Wien"" so we can administer the remainder of the process. (Since the agreement is administrated by the Copyright Clearance Center a one-time registration is necessary in order to finish the process.) You can find additional information on the agreement under " & url & "." & vbCrLf & vbCrLf & _
        "Please do not hesitate to ask if you have any questions." & vbCrLf & vbCrLf & _
        "Kind regards," & vbCrLf & vbCrLf & _
        "Guido Blechl / Bernhard Schubert / Klara Schellander"
    
    End If

' 5.2 CUP
'''''''''

    If (publisher = "CUP") And echeck_status = "pending" Then 'CUP Nachfrage
                   
        UFind corresponding_author 'Suche nach corresponding_author in u:find
        
        'Deutsch
        
        url = "https://openaccess.univie.ac.at/" & LCase(Application.WorksheetFunction.Trim(publisher))
        url = Replace(url, " ", "")
        
        EMailGenerate "Open Access für Ihren " & publisher & "-Artikel """ & title & """" & vbCrLf & vbCrLf & _
        "S.g. NNN," & vbCrLf & vbCrLf & _
        "wir wurden von " & publisher & " darüber informiert, dass folgende Publikation über das Open-Access-Verlagsabkommen der Universität Wien gefördert werden könnte:" & vbCrLf & vbCrLf & _
        "> Manuscript Details" & vbCrLf & _
        "> " & doi & vbCrLf & _
        "> " & source_full_title & vbCrLf & _
        "> " & title & vbCrLf & vbCrLf & _
        "Wir würden uns freuen, wenn Sie dieses Angebot wahrnehmen würden. Es entstehen Ihnen keine Kosten. Falls Sie sich für Open Access entscheiden, öffnen Sie bitte den Link https://www.cambridge.org/core/services/open-access-policies/read-and-publish-agreements/convert-your-article-to-open-access und folgen Sie den Anweisungen. Als Open-Access-Lizenz empfehlen wir CC BY (siehe https://creativecommons.org/licenses/by/4.0/ bzw. auch https://openaccess.univie.ac.at/creativecommons/), welche eine bestmögliche Verbreitung Ihres Artikels ermöglicht. Bitte beachten Sie auch, dass einige Fördergeber - wie z.B. der FWF - diese Lizenz verpflichtend vorsehen." & vbCrLf & vbCrLf & _
        "Unsere Informationen zum Förderabkommen finden Sie unter " & url & "." & vbCrLf & vbCrLf & _
        "Sollten Sie noch Fragen haben, melden Sie sich bitte." & vbCrLf & vbCrLf & _
        "Mit freundlichen Grüßen," & vbCrLf & vbCrLf & _
        "Guido Blechl / Bernhard Schubert / Klara Schellander"
        
        'Englisch
                                
        url = "https://openaccess.univie.ac.at/en/" & LCase(Application.WorksheetFunction.Trim(publisher))
        url = Replace(url, " ", "")
                                
        EMailGenerate "Open Access for your " & publisher & " article """ & title & """" & vbCrLf & vbCrLf & _
        "Dear NNN," & vbCrLf & vbCrLf & _
        "we were informed by " & publisher & " that the publication below is eligible for Open Access funding as part of a publishing agreement with the University of Vienna:" & vbCrLf & vbCrLf & _
        "> Manuscript Details" & vbCrLf & _
        "> " & doi & vbCrLf & _
        "> " & source_full_title & vbCrLf & _
        "> " & title & vbCrLf & vbCrLf & _
        "We would be delighted if you would accept this offer. You will not incur any costs since any Open Access charges are already included in our contract sum. If you opt for Open Access please open the link https://www.cambridge.org/core/services/open-access-policies/read-and-publish-agreements/convert-your-article-to-open-access and follow the instructions. We recommend the CC BY license (see https://creativecommons.org/licenses/by/4.0/ and https://openaccess.univie.ac.at/en/creativecommons/, respectively), which ensures that your article can be disseminated as widely and as easily as possible. Please note that some funders (such as the FWF) mandate the CC BY license." & vbCrLf & vbCrLf & _
        "You can find additional information on the agreement under " & url & "." & vbCrLf & vbCrLf & _
        "Please do not hesitate to ask if you have any questions." & vbCrLf & vbCrLf & _
        "Kind regards," & vbCrLf & vbCrLf & _
        "Guido Blechl / Bernhard Schubert / Klara Schellander"
    
    End If

' 5.3 Elsevier
''''''''''''''

    If (publisher = "Elsevier") And echeck_status = "pending" And funder = "" Then 'Elsevier Nachfrage
                   
        UFind corresponding_author 'Suche nach corresponding_author in u:find
        
        'Deutsch
        
        url = "https://openaccess.univie.ac.at/" & LCase(Application.WorksheetFunction.Trim(publisher))
        url = Replace(url, " ", "")
        
        EMailGenerate "Open Access für Ihren " & publisher & "-Artikel """ & title & """" & vbCrLf & vbCrLf & _
        "S.g. NNN," & vbCrLf & vbCrLf & _
        "wir wurden von " & publisher & " darüber informiert, dass folgende Publikation über das Open-Access-Verlagsabkommen der Universität Wien gefördert werden könnte:" & vbCrLf & vbCrLf & _
        "> Manuscript Details" & vbCrLf & _
        "> " & doi & vbCrLf & _
        "> " & source_full_title & vbCrLf & _
        "> " & title & vbCrLf & vbCrLf & _
        "Wir würden uns freuen, wenn Sie dieses Angebot wahrnehmen würden. Es entstehen Ihnen keine Kosten. Falls Sie sich für Open Access entscheiden, folgen Sie bitte dem Link NNNNNNNNNNNNNNNN und klicken Sie im Bereich ""Rights and Access"" auf ""Make changes and re-submit"". Sollten Sie nochmals nach Ihrer Affiliation gefragt werden, wählen Sie bitte unbedingt ""University of Vienna"" aus, damit die Zuordnung zum Verlagsabkommen funktioniert. In Schritt 4 sollten Sie dann ""Publish Open Access"" auswählen können (mit dem Hinweis: ""As an author affiliated with an Austrian institution, upon validation, agreement between the Austrian institutions and Elsevier will cover the APC""). Unsere Informationen zum Förderabkommen finden Sie unter " & url & "." & vbCrLf & vbCrLf & _
        "Sollten Sie noch Fragen haben, melden Sie sich bitte." & vbCrLf & vbCrLf & _
        "Mit freundlichen Grüßen," & vbCrLf & vbCrLf & _
        "Guido Blechl / Bernhard Schubert / Klara Schellander"
        
        'Englisch
                                
        url = "https://openaccess.univie.ac.at/en/" & LCase(Application.WorksheetFunction.Trim(publisher))
        url = Replace(url, " ", "")
                                
        EMailGenerate "Open Access for your " & publisher & " article """ & title & """" & vbCrLf & vbCrLf & _
        "Dear NNN," & vbCrLf & vbCrLf & _
        "we were informed by " & publisher & " that the publication below is eligible for Open Access funding as part of a publishing agreement with the University of Vienna:" & vbCrLf & vbCrLf & _
        "> Manuscript Details" & vbCrLf & _
        "> " & doi & vbCrLf & _
        "> " & source_full_title & vbCrLf & _
        "> " & title & vbCrLf & vbCrLf & _
        "We would be delighted if you would accept this offer. You will not incur any costs since any Open Access charges are already included in our contract sum. If you opt for Open Access please follow the the link NNNNNNNNNNNN and click on ""Make changes and re-submit"" in the ""Rights and Access"" section. Should you be asked for your affiliation please make sure to choose ""University of Vienna"" so the article can be processed via our publishing agreement. In step 4 you should then be able to select ""Publish Open Access"" (with the note: ""As an author affiliated with an Austrian institution, upon validation, agreement between the Austrian institutions and Elsevier will cover the APC""). You can find additional information on the agreement under " & url & "." & vbCrLf & vbCrLf & _
        "Please do not hesitate to ask if you have any questions." & vbCrLf & vbCrLf & _
        "Kind regards," & vbCrLf & vbCrLf & _
        "Guido Blechl / Bernhard Schubert / Klara Schellander"
    
    End If

' 5.4 IEEE
''''''''''
    
    If (publisher = "IEEE") And echeck_status = "pending" Then 'IEEE Nachfrage
                   
        UFind corresponding_author 'Suche nach corresponding_author in u:find
        
        'Deutsch
        
        url = "https://openaccess.univie.ac.at/" & LCase(Application.WorksheetFunction.Trim(publisher))
        url = Replace(url, " ", "")
        
        EMailGenerate "Open Access für Ihren " & publisher & "-Artikel """ & title & """" & vbCrLf & vbCrLf & _
        "S.g. NNN," & vbCrLf & vbCrLf & _
        "wir wurden von " & publisher & " darüber informiert, dass folgende Publikation über das Open-Access-Verlagsabkommen der Universität Wien gefördert werden könnte:" & vbCrLf & vbCrLf & _
        "> Manuscript Details" & vbCrLf & _
        "> " & doi & vbCrLf & _
        "> " & source_full_title & vbCrLf & _
        "> " & title & vbCrLf & vbCrLf & _
        "Wir würden uns freuen, wenn Sie dieses Angebot wahrnehmen würden. Es entstehen Ihnen keine Kosten. Falls Sie sich für Open Access entscheiden, bitten wir Sie, uns Bescheid zu geben. Wir würden dies dann IEEE mitteilen." & vbCrLf & vbCrLf & _
        "Unsere Informationen zum Förderabkommen finden Sie unter " & url & "." & vbCrLf & vbCrLf & _
        "Sollten Sie noch Fragen haben, melden Sie sich bitte." & vbCrLf & vbCrLf & _
        "Mit freundlichen Grüßen," & vbCrLf & vbCrLf & _
        "Guido Blechl / Bernhard Schubert / Klara Schellander"
        
        'Englisch
                                
        url = "https://openaccess.univie.ac.at/en/" & LCase(Application.WorksheetFunction.Trim(publisher))
        url = Replace(url, " ", "")
                                
        EMailGenerate "Open Access for your " & publisher & " article """ & title & """" & vbCrLf & vbCrLf & _
        "Dear NNN," & vbCrLf & vbCrLf & _
        "we were informed by " & publisher & " that the publication below is eligible for Open Access funding as part of a publishing agreement with the University of Vienna:" & vbCrLf & vbCrLf & _
        "> Manuscript Details" & vbCrLf & _
        "> " & doi & vbCrLf & _
        "> " & source_full_title & vbCrLf & _
        "> " & title & vbCrLf & vbCrLf & _
        "We would be delighted if you would accept this offer. You will not incur any costs. If you decide to opt for Open Access please let us know. We will then notify IEEE accordingly." & vbCrLf & vbCrLf & _
        "You can find  information on the agreement under " & url & "." & vbCrLf & vbCrLf & _
        "Please do not hesitate to ask if you have any questions." & vbCrLf & vbCrLf & _
        "Kind regards," & vbCrLf & vbCrLf & _
        "Guido Blechl / Bernhard Schubert / Klara Schellander"
    
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''
' 6 Bestätigung der Kostenübernahme an Autor*in '
'''''''''''''''''''''''''''''''''''''''''''''''''

' 6.1 Bestätigung Publikationsfonds
'''''''''''''''''''''''''''''''''''

' 6.1.1 Über Kostengrenze
'''''''''''''''''''''''''
       
'endif siehe 6.1.2

        If invoice_status = "Zusage" And open_access_deal = "no agreement" Then  'Publikationsfonds Zusage
            If euro >= 2400 Then 'APC über Kostengrenze (netto)
            
            'Deutsch
            
                EMailGenerate "S.g. NNNNN," & vbCrLf & vbCrLf & _
                "vielen Dank für Ihren Antrag auf Open-Access-Förderung des Artikels """ & title & """ in der Zeitschrift """ & source_full_title & """. Obwohl die maximale Fördersumme überschritten wird, können von uns die vollen Publikationskosten übernommen werden." & vbCrLf & vbCrLf & _
                "Bitte teilen Sie dem Verlag (nach Acceptance) Folgendes mit:" & vbCrLf & vbCrLf & _
                "1. Article Acknowledgement" & vbCrLf & "--------------------------" & vbCrLf & "Open access funding provided by University of Vienna." & vbCrLf & vbCrLf & vbCrLf & _
                "2. Rechnungsadresse für die Publikationsgebühr (invoice address)" & vbCrLf & "----------------------------------------------------------------" & vbCrLf & "Postanschrift:" & vbCrLf & " Universität Wien" & vbCrLf & " Bibliotheks- und Archivwesen" & vbCrLf & " Open Access Office" & vbCrLf & " Boltzmanngasse 5" & vbCrLf & " A-1090 Wien" & vbCrLf & vbCrLf & "E-Mail:" & vbCrLf & vbCrLf & " openaccess@univie.ac.at" & vbCrLf & vbCrLf & "VAT identification number of the University of Vienna:" & vbCrLf & " ATU 37586901" & vbCrLf & vbCrLf & vbCrLf & _
                "3. Zahlungsziel" & vbCrLf & "---------------" & vbCrLf & "Um eine möglichst rasche Freischaltung Ihres Artikels zu gewährleisten, ist es notwendig, dass als Zahlungsziel auf der Rechnung ""nach Erhalt der Rechnung"" (""due on receipt"") angegeben wird. Dies ist erforderlich, da Zahlungen seitens der Quästur der Universität Wien immer mit dem auf der Rechnung angeführten Zahlungsziel erfolgen." & vbCrLf & vbCrLf & vbCrLf & _
                "Hinweise:" & vbCrLf & vbCrLf & ".) Sollte der Verlag die Rechnung nur direkt an Sie schicken können, so übermitteln Sie uns bitte diese Rechnung, damit wir sie bezahlen können. Zahlen Sie die Rechnung bitte nicht eigenständig ein!" & vbCrLf & vbCrLf & ".) Eine Rückerstattung von bereits bezahlten Rechnungen für Publikationsgebühren (APCs) ist nicht möglich." & vbCrLf & vbCrLf & ".) Sollte Ihr Beitrag vom Verlag nicht akzeptiert werden, bitten wir Sie, uns kurz zu informieren, damit wir die reservierten Mittel wieder freigeben können und der Artikel nicht zu Ihrem Publikationsfonds-Förderlimit zählt (zurzeit drei Artikel pro Jahr pro corresponding author). Selbstverständlich können Sie für eine Neueinreichung bei einer anderen Zeitschrift einen Neuantrag bei uns stellen. Angeforderte Mittel zur Publikationsförderung verfallen automatisch nach einem Jahr. Geben Sie uns deshalb bitte Bescheid, falls der Veröffentlichungsprozess länger dauern sollte." & vbCrLf & vbCrLf & _
                "Sollten Sie dazu oder zu anderen Open-Access-Themen noch Fragen haben, so helfen wir Ihnen gerne weiter!" & vbCrLf & vbCrLf & "Mit freundlichen Grüßen" & vbCrLf & vbCrLf & "Guido Blechl / Bernhard Schubert / Klara Schellander"
                
            'Englisch
            
                EMailGenerate "Dear NNNNN," & vbCrLf & vbCrLf & _
                "thank you for your application to fund the article """ & title & """ in """ & source_full_title & """. Despite the fact that the APCs exceed the maximum expected amount we can cover the charges in full." & vbCrLf & vbCrLf & _
                "Please inform the publisher of the following:" & vbCrLf & vbCrLf & _
                "1. Article Acknowledgement" & vbCrLf & "--------------------------" & vbCrLf & "Open access funding provided by University of Vienna." & vbCrLf & vbCrLf & vbCrLf & _
                "2. Invoice address for publication charges" & vbCrLf & "------------------------------------------" & vbCrLf & "Postal address:" & vbCrLf & " Universität Wien" & vbCrLf & " Bibliotheks- und Archivwesen" & vbCrLf & " Open Access Office" & vbCrLf & " Boltzmanngasse 5" & vbCrLf & " A-1090 Wien" & vbCrLf & vbCrLf & "E-Mail:" & vbCrLf & vbCrLf & " openaccess@univie.ac.at" & vbCrLf & vbCrLf & "VAT identification number of the University of Vienna:" & vbCrLf & " ATU 37586901" & vbCrLf & vbCrLf & vbCrLf & _
                "3. Due date" & vbCrLf & "-----------" & vbCrLf & "To ensure your article is published as soon as possible, the due date on the invoice has to be ""on receipt"". This is necessary because the University's accounting office only settles invoices on their due date." & vbCrLf & vbCrLf & vbCrLf & _
                "Notes:" & vbCrLf & vbCrLf & ".) In case the publisher can only send the invoice directly to you please forward it to us so we can pay it. Please do not pay it yourself!" & vbCrLf & vbCrLf & ".) Reimbursement of APC invoices already paid is not possible." & vbCrLf & vbCrLf & ".) In case the publisher does not accept your contribution please let us know so we can reallocate the funds set aside and the article does not count towards your funding limit (currently three articles per year per corresponding author). You may of course reapply for funding in order to publish in a different journal. Requested funds expire automatically after one year. For this reason please let us know in case the publication process takes longer than that." & vbCrLf & vbCrLf & _
                "If you have any questions regarding the process or other topics related to Open Access please do not hesitate to contact us!" & vbCrLf & vbCrLf & "Kind regards" & vbCrLf & vbCrLf & "Guido Blechl / Bernhard Schubert / Klara Schellander"
                
' 6.1.2 Regulär
'''''''''''''''

'nested if, siehe 6.1.1
                
            Else 'APC unter Kostengrenze
            
            'Deutsch
    
                EMailGenerate "S.g. NNNNN," & vbCrLf & vbCrLf & _
                "vielen Dank für Ihren Antrag auf Open-Access-Förderung des Artikels """ & title & """ in der Zeitschrift """ & source_full_title & """. Da die Förderkriterien erfüllt sind, wird Ihr Antrag bewilligt." & vbCrLf & vbCrLf & _
                "Bitte teilen Sie dem Verlag (nach Acceptance) Folgendes mit:" & vbCrLf & vbCrLf & _
                "1. Article Acknowledgement" & vbCrLf & "--------------------------" & vbCrLf & "Open access funding provided by University of Vienna." & vbCrLf & vbCrLf & vbCrLf & _
                "2. Rechnungsadresse für die Publikationsgebühr (invoice address)" & vbCrLf & "----------------------------------------------------------------" & vbCrLf & "Postanschrift:" & vbCrLf & " Universität Wien" & vbCrLf & " Bibliotheks- und Archivwesen" & vbCrLf & " Open Access Office" & vbCrLf & " Boltzmanngasse 5" & vbCrLf & " A-1090 Wien" & vbCrLf & vbCrLf & "E-Mail:" & vbCrLf & vbCrLf & " openaccess@univie.ac.at" & vbCrLf & vbCrLf & "VAT identification number of the University of Vienna:" & vbCrLf & " ATU 37586901" & vbCrLf & vbCrLf & vbCrLf & _
                "3. Zahlungsziel" & vbCrLf & "---------------" & vbCrLf & "Um eine möglichst rasche Freischaltung Ihres Artikels zu gewährleisten, ist es notwendig, dass als Zahlungsziel auf der Rechnung ""nach Erhalt der Rechnung"" (""due on receipt"") angegeben wird. Dies ist erforderlich, da Zahlungen seitens der Quästur der Universität Wien immer mit dem auf der Rechnung angeführten Zahlungsziel erfolgen." & vbCrLf & vbCrLf & vbCrLf & _
                "Hinweise:" & vbCrLf & vbCrLf & ".) Sollte der Verlag die Rechnung nur direkt an Sie schicken können, so übermitteln Sie uns bitte diese Rechnung, damit wir sie bezahlen können. Zahlen Sie die Rechnung bitte nicht eigenständig ein!" & vbCrLf & vbCrLf & ".) Eine Rückerstattung von bereits bezahlten Rechnungen für Publikationsgebühren (APCs) ist nicht möglich." & vbCrLf & vbCrLf & ".) Sollte Ihr Beitrag vom Verlag nicht akzeptiert werden, bitten wir Sie, uns kurz zu informieren, damit wir die reservierten Mittel wieder freigeben können und der Artikel nicht zu Ihrem Publikationsfonds-Förderlimit zählt (zurzeit drei Artikel pro Jahr pro corresponding author). Selbstverständlich können Sie für eine Neueinreichung bei einer anderen Zeitschrift einen Neuantrag bei uns stellen. Angeforderte Mittel zur Publikationsförderung verfallen automatisch nach einem Jahr. Geben Sie uns deshalb bitte Bescheid, falls der Veröffentlichungsprozess länger dauern sollte." & vbCrLf & vbCrLf & _
                "Sollten Sie dazu oder zu anderen Open-Access-Themen noch Fragen haben, so helfen wir Ihnen gerne weiter!" & vbCrLf & vbCrLf & "Mit freundlichen Grüßen" & vbCrLf & vbCrLf & "Guido Blechl / Bernhard Schubert / Klara Schellander"
        
            'Englisch
        
                EMailGenerate "Dear NNNNN," & vbCrLf & vbCrLf & _
                "thank you for your application to fund the article """ & title & """ in """ & source_full_title & """. Since the requirements for funding are met your request can be granted." & vbCrLf & vbCrLf & _
                "Please inform the publisher of the following:" & vbCrLf & vbCrLf & _
                "1. Article Acknowledgement" & vbCrLf & "--------------------------" & vbCrLf & "Open access funding provided by University of Vienna." & vbCrLf & vbCrLf & vbCrLf & _
                "2. Invoice address for publication charges" & vbCrLf & "------------------------------------------" & vbCrLf & "Postal address:" & vbCrLf & " Universität Wien" & vbCrLf & " Bibliotheks- und Archivwesen" & vbCrLf & " Open Access Office" & vbCrLf & " Boltzmanngasse 5" & vbCrLf & " A-1090 Wien" & vbCrLf & vbCrLf & "E-Mail:" & vbCrLf & vbCrLf & " openaccess@univie.ac.at" & vbCrLf & vbCrLf & "VAT identification number of the University of Vienna:" & vbCrLf & " ATU 37586901" & vbCrLf & vbCrLf & vbCrLf & _
                "3. Due date" & vbCrLf & "-----------" & vbCrLf & "To ensure your article is published as soon as possible, the due date on the invoice has to be ""on receipt"". This is necessary because the University's accounting office only settles invoices on their due date." & vbCrLf & vbCrLf & vbCrLf & _
                "Notes:" & vbCrLf & vbCrLf & ".) In case the publisher can only send the invoice directly to you please forward it to us so we can pay it. Please do not pay it yourself!" & vbCrLf & vbCrLf & ".) Reimbursement of APC invoices already paid is not possible." & vbCrLf & vbCrLf & ".) In case the publisher does not accept your contribution please let us know so we can reallocate the funds set aside and the article does not count towards your funding limit (currently three articles per year per corresponding author). You may of course reapply for funding in order to publish in a different journal. Requested funds expire automatically after one year. For this reason please let us know in case the publication process takes longer than that." & vbCrLf & vbCrLf & _
                "If you have any questions regarding the process or other topics related to Open Access please do not hesitate to contact us!" & vbCrLf & vbCrLf & "Kind regards" & vbCrLf & vbCrLf & "Guido Blechl / Bernhard Schubert / Klara Schellander"

            End If
        End If
        
' 6.2 Bestätigung Verlagsabkommen
'''''''''''''''''''''''''''''''''
        
        If invoice_status = "Zusage" And publisher = "Frontiers" Then 'Frontiers-Autor*inneninfo
        
            UFind corresponding_author 'Suche nach corresponding_author in u:find
        
            'Deutsch
            
            EMailGenerate "Open-Access-Förderung für Ihren Frontiers-Artikel """ & title & """" & vbCrLf & vbCrLf & _
            "S.g. NNNN," & vbCrLf & vbCrLf & _
            "wir wurden von Frontiers über folgende Einreichung informiert:" & vbCrLf & vbCrLf & _
            "> Manuscript Details" & vbCrLf & _
            "> Title: " & title & vbCrLf & _
            "> Journal: " & source_full_title & vbCrLf & _
            "> Corresponding author: " & corresponding_author & vbCrLf & vbCrLf & _
            "Wir haben die Übernahme der Open-Access-Publikationsgebühren im Rahmen unseres Abkommens gegenüber dem Verlag bereits bestätigt (Details siehe https://openaccess.univie.ac.at/frontiers/), sodass die Rechnung nach Acceptance zentral über das Open Access Office bezahlt wird, ohne dass Sie hier tätig werden müssen." & vbCrLf & vbCrLf & _
            "Sollten Sie noch Fragen haben, melden Sie sich bitte." & vbCrLf & vbCrLf & _
            "Mit freundlichen Grüßen," & vbCrLf & vbCrLf & _
            "Guido Blechl / Bernhard Schubert / Klara Schellander"
            
            'Englisch
            
            EMailGenerate "Open Access for your Frontiers article """ & title & """" & vbCrLf & vbCrLf & _
            "Dear NNNN," & vbCrLf & vbCrLf & _
            "we were notified of the submission below by Frontiers:" & vbCrLf & vbCrLf & _
            "> Manuscript Details" & vbCrLf & _
            "> Title: " & title & vbCrLf & _
            "> Journal: " & source_full_title & vbCrLf & _
            "> Corresponding author: " & corresponding_author & vbCrLf & vbCrLf & _
            "We have already informed the publisher that we will cover the charges under our agreement (see https://openaccess.univie.ac.at/en/frontiers/), which means that after acceptance the invoice will be paid centrally by the Open Access Office without any need for you to become involved." & vbCrLf & vbCrLf & _
            "Please do not hesitate to ask if you have any questions." & vbCrLf & vbCrLf & _
            "Kind regards," & vbCrLf & vbCrLf & _
            "Guido Blechl / Bernhard Schubert / Klara Schellander"
        
        End If

Ende:

End Sub