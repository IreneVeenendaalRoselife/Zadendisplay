'==============================================================================
' BAZA intekenmailing
'
' Maakt per winkel uit het winkeloverzicht (Power Query) een Outlook-mail aan
' met de persoonlijke formulierlink en de vier productafbeeldingen in de tekst.
' De mails worden alleen GEOPEND, niet verstuurd: je klikt zelf op Verzenden.
' Er worden steeds 25 mails tegelijk klaargezet. Daarna vraagt de macro of de
' volgende 25 klaargezet mogen worden.
'
' Elke klaargezette winkel komt in het tabblad "Verzendlog". Draai je de macro
' opnieuw, dan worden winkels uit de log overgeslagen. Wil je een winkel
' opnieuw laten klaarzetten, verwijder dan die regel uit de log.
'==============================================================================
Option Explicit

' ---- Instellingen ------------------------------------------------------------
Private Const AFZENDER As String = "plus@roselife.nl"
' Onderwerp: "<jaar> Intekenen Zaden (Displays & Mixdozen) | Winkelnummer/SAPnr Winkelnaam"
' Het jaar en het soort mail (gewoon, correctie of reminder) vraagt de macro bij het starten.
Private Const ONDERWERP As String = "Intekenen Zaden (Displays & Mixdozen)"
Private Const PER_KEER As Long = 25
Private Const CONTACTPERSONEN_IN_CC As Boolean = True   ' False = alle adressen in Aan
' Map met de afbeeldingen (Zadendisplay.jpg, Bloemenweide.jpg, AardbeienTuin.jpg, TomatenTuin.jpg)
' en de knop (KnopFormulier.png).
Private Const AFBEELDINGEN_MAP As String = "S:\Roselife\2027\Zadendisplays\Afbeeldingen displays dozen"
' -----------------------------------------------------------------------------

Private Const PR_ATTACH_CONTENT_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
Private Const PR_ATTACHMENT_HIDDEN As String = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

' Keuzes van deze run (worden bij het starten gevraagd)
Private mJaar As Long          ' bijvoorbeeld 2027
Private mSoort As String       ' "", "Correctie" of "Reminder"

Public Sub BazaMailingKlaarzetten()
    Dim lo As Object, ol As Object, account As Object
    Dim logWs As Worksheet, gehad As Object
    Dim map As String, afb As Variant, i As Long
    Dim cWnk As Long, cNaam As Long, cMail As Long, cC2 As Long, cC3 As Long, cC4 As Long, cLink As Long, cSap As Long
    Dim rij As Range, wnk As String, mailAdres As String
    Dim teDoen As Collection, item As Variant
    Dim aantal As Long, totaal As Long, mail As Object

    ' Jaar en soort mail vragen
    If Not VraagInstellingen() Then Exit Sub

    ' Tabel met het winkeloverzicht zoeken (herkenbaar aan de kolom Formulierlink)
    Set lo = ZoekTabel()
    If lo Is Nothing Then
        MsgBox "Geen tabel gevonden met de kolommen Winkelnummer, Emailadres en Formulierlink." & vbCrLf & _
               "Laad de Power Query eerst als tabel in deze werkmap.", vbExclamation
        Exit Sub
    End If
    cWnk = lo.ListColumns("Winkelnummer").Index
    cNaam = lo.ListColumns("Winkelnaam").Index
    cMail = lo.ListColumns("Emailadres").Index
    cLink = lo.ListColumns("Formulierlink").Index
    cSap = KolomIndex(lo, "SAPnr")
    cC2 = KolomIndex(lo, "Contactpersoon2")
    cC3 = KolomIndex(lo, "Contactpersoon3")
    cC4 = KolomIndex(lo, "Contactpersoon4")

    ' Afbeeldingen controleren
    afb = Array("Zadendisplay.jpg", "Bloemenweide.jpg", "AardbeienTuin.jpg", "TomatenTuin.jpg", "KnopFormulier.png")
    map = ZoekAfbeeldingenMap(afb)
    If map = "" Then Exit Sub

    ' Log ophalen en lijst met nog te doen winkels maken
    Set logWs = LogBlad()
    Set gehad = CreateObject("Scripting.Dictionary")
    For i = 2 To logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row
        gehad(CStr(logWs.Cells(i, 1).Value)) = True
    Next i

    Set teDoen = New Collection
    If Not lo.DataBodyRange Is Nothing Then
        For Each rij In lo.DataBodyRange.Rows
            wnk = CStr(rij.Cells(1, cWnk).Value)
            mailAdres = Trim$(CStr(rij.Cells(1, cMail).Value))
            If wnk <> "" And mailAdres <> "" And Not gehad.Exists(wnk) Then teDoen.Add rij
        Next rij
    End If

    totaal = teDoen.Count
    If totaal = 0 Then
        MsgBox "Er zijn geen winkels meer om klaar te zetten." & vbCrLf & _
               "(Alle winkels staan al in het tabblad " & LogNaam() & ".)", vbInformation
        Exit Sub
    End If
    If MsgBox("Jaar: " & mJaar & vbCrLf & _
              "Soort mail: " & IIf(mSoort = "", "Gewone mail", mSoort) & vbCrLf & _
              "Log: tabblad " & LogNaam() & vbCrLf & vbCrLf & _
              totaal & " winkels nog te mailen." & vbCrLf & vbCrLf & _
              "De eerste " & Application.Min(PER_KEER, totaal) & " mails klaarzetten in Outlook?", _
              vbQuestion + vbYesNo) = vbNo Then Exit Sub

    ' Outlook en afzenderaccount
    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    For Each account In ol.Session.Accounts
        If LCase$(account.SmtpAddress) = LCase$(AFZENDER) Then Exit For
    Next account

    For Each item In teDoen
        Set rij = item
        Set mail = ol.CreateItem(0)   ' olMailItem

        If Not account Is Nothing Then
            Set mail.SendUsingAccount = account
        Else
            mail.SentOnBehalfOfName = AFZENDER
        End If

        Dim extra As String
        extra = Adressen(rij, cC2, cC3, cC4)
        If CONTACTPERSONEN_IN_CC Then
            mail.To = Trim$(CStr(rij.Cells(1, cMail).Value))
            mail.CC = extra
        Else
            mail.To = Trim$(CStr(rij.Cells(1, cMail).Value)) & IIf(extra = "", "", ";" & extra)
        End If
        Dim sapWaarde As Variant
        sapWaarde = ""
        If cSap > 0 Then sapWaarde = rij.Cells(1, cSap).Value
        mail.Subject = Onderwerpregel(rij.Cells(1, cWnk).Value, rij.Cells(1, cNaam).Value, sapWaarde)

        ' Afbeeldingen als verborgen bijlage met Content-ID, zodat ze in de tekst staan
        For i = LBound(afb) To UBound(afb)
            Dim bijlage As Object
            Set bijlage = mail.Attachments.Add(map & afb(i), 1, 0)
            bijlage.PropertyAccessor.SetProperty PR_ATTACH_CONTENT_ID, "baza" & (i + 1)
            bijlage.PropertyAccessor.SetProperty PR_ATTACHMENT_HIDDEN, True
        Next i

        mail.BodyFormat = 2           ' olFormatHTML
        mail.HTMLBody = MailTekst(CStr(rij.Cells(1, cLink).Value))
        mail.Display False

        SchrijfLog logWs, rij.Cells(1, cWnk).Value, rij.Cells(1, cNaam).Value, mail.To, mail.CC
        aantal = aantal + 1

        If aantal Mod PER_KEER = 0 And aantal < totaal Then
            If MsgBox(aantal & " van " & totaal & " mails staan klaar in Outlook." & vbCrLf & _
                      "Controleer en verstuur ze zelf met Verzenden." & vbCrLf & vbCrLf & _
                      "De volgende " & Application.Min(PER_KEER, totaal - aantal) & " klaarzetten?", _
                      vbQuestion + vbYesNo) = vbNo Then Exit For
        End If
    Next item

    MsgBox aantal & " mails klaargezet in Outlook." & vbCrLf & _
           "Nog " & (totaal - aantal) & " winkels te gaan." & vbCrLf & vbCrLf & _
           "Let op: de mails worden pas verstuurd als je zelf op Verzenden klikt.", vbInformation
End Sub

' HTML-tekst van de mail
Private Function MailTekst(ByVal link As String) As String
    Dim h As String
    link = Replace(link, "&", "&amp;")

    h = "<html><body style=""font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#222222;"">"
    h = h & "<p>Goedendag,</p>"
    If mSoort = "Correctie" Then
        h = h & "<p>Onlangs heb je van ons een mail ontvangen over het intekenen op het BAZA-assortiment. "
        h = h & "Helaas werkte de knop naar het formulier in die mail niet. Onze excuses daarvoor. "
        h = h & "Hieronder vind je de informatie opnieuw, nu met een werkende link.</p>"
    ElseIf mSoort = "Reminder" Then
        h = h & "<p>Onlangs heb je van ons een mail ontvangen over het intekenen op het BAZA-assortiment. "
        h = h & "Wij hebben je reactie nog niet ontvangen. Wil je het formulier alsnog invullen? "
        h = h & "Ook als je dit jaar niet wilt intekenen, horen we dat graag.</p>"
        h = h & "<p>Heb je het formulier inmiddels al ingevuld? Dan kun je deze mail als niet verzonden beschouwen.</p>"
    End If
    h = h & "<p>Ook dit jaar is het weer mogelijk om in te tekenen op het assortiment van BAZA!<br>"
    h = h & "Naast de vertrouwde jaarlijkse zadendisplay kun je dit jaar ook kiezen voor diverse andere displaydozen.</p>"
    h = h & "<p>Het aanbod bestaat uit:</p>"

    ' Tabel met het aanbod: vier displays naast elkaar
    h = h & AanbodTabel()
    h = h & "<p>Voordat je gaat intekenen, is het belangrijk om rekening te houden met de volgende informatie:</p><ul>"
    h = h & "<li>Alle displays worden kant-en-klaar geleverd.</li>"
    h = h & "<li>Besteladvies: minimaal 1 zadendisplay (850 zakjes).</li>"
    h = h & "<li>Voor de zadendisplay geldt 100% recht van retour. Overgebleven zaden kunnen in week 22 van " & mJaar & " per post worden geretourneerd. Verdere instructies hierover volgen op een later moment.</li>"
    h = h & "</ul>"

    ' Oproep om ook bij geen interesse het formulier in te vullen
    h = h & "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""680"" style=""margin:8px 0 4px 0;""><tr>"
    h = h & "<td style=""background:#F3F8EA;border-left:4px solid #8DB63C;padding:12px 16px;font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#222222;"">"
    h = h & "<b>Geen interesse dit jaar?</b><br>"
    h = h & "Laat het ons dan ook weten via het formulier, "
    h = h & "dan ontvang je over deze actie geen herinneringen meer."
    h = h & "</td></tr></table>"

    ' Knop met de persoonlijke formulierlink. De knop is een afbeelding met een
    ' gewone link eromheen: dat werkt in Outlook en in alle andere mailprogramma's.
    h = h & "<table cellpadding=""0"" cellspacing=""0"" border=""0"" style=""margin:12px 0 8px 0;""><tr><td>"
    h = h & "<a href=""" & link & """ target=""_blank"">"
    h = h & "<img src=""cid:baza5"" width=""360"" height=""58"" border=""0"" alt=""Naar het formulier"" style=""display:block;border:0;"">"
    h = h & "</a></td></tr></table>"

    ' Dezelfde link voluit, voor als de knop niet werkt
    h = h & "<p style=""font-size:10pt;color:#555555;margin:0 0 24px 0;"">Werkt de knop niet? "
    h = h & "Klik dan op deze link of kopieer hem naar je browser:<br>"
    h = h & "<a href=""" & link & """ target=""_blank"" style=""color:#5A8A1F;word-break:break-all;"">" & link & "</a></p>"

    h = h & "<p>Met vriendelijke groet,</p>"
    h = h & "<p style=""color:#555555;"">Roselife &nbsp;|&nbsp; 2675 LN Honselersdijk &nbsp;|&nbsp; the Netherlands<br>"
    h = h & "+31 (0)174 - 671 414 &nbsp;|&nbsp; <a href=""https://www.roselife.nl"">www.roselife.nl</a></p>"
    h = h & "</body></html>"
    MailTekst = h
End Function

' Zoekt de map met de afbeeldingen. Volgorde: instelling AFBEELDINGEN_MAP,
' de eerder gekozen map, de submap "Afbeeldingen" naast de werkmap. Lukt dat
' niet (bijvoorbeeld omdat de werkmap op OneDrive/SharePoint staat en het pad
' dan een webadres is), dan mag je zelf de map aanwijzen. Die keuze wordt in de
' werkmap onthouden.
Private Function ZoekAfbeeldingenMap(ByVal afb As Variant) As String
    Dim kandidaten As Collection, k As Variant, gekozen As String, mist As String

    Set kandidaten = New Collection
    If AFBEELDINGEN_MAP <> "" Then kandidaten.Add AFBEELDINGEN_MAP
    gekozen = OnthoudenMap()
    If gekozen <> "" Then kandidaten.Add gekozen
    If LCase$(Left$(ThisWorkbook.Path, 4)) <> "http" And ThisWorkbook.Path <> "" Then
        kandidaten.Add ThisWorkbook.Path & Application.PathSeparator & "Afbeeldingen"
    End If

    For Each k In kandidaten
        If Ontbrekend(MetScheiding(CStr(k)), afb) = "" Then
            ZoekAfbeeldingenMap = MetScheiding(CStr(k))
            Exit Function
        End If
    Next k

    ' Niet gevonden: precies melden welke bestanden in de ingestelde map ontbreken
    If AFBEELDINGEN_MAP <> "" Then
        mist = Ontbrekend(MetScheiding(AFBEELDINGEN_MAP), afb)
        MsgBox "In deze map ontbreken bestanden:" & vbCrLf & AFBEELDINGEN_MAP & vbCrLf & vbCrLf & _
               "Ontbreekt:" & vbCrLf & mist & vbCrLf & vbCrLf & _
               "Let op de exacte bestandsnaam en extensie (bijvoorbeeld .png of .jpg)." & vbCrLf & _
               "Je kunt in het volgende venster ook een andere map kiezen.", vbExclamation
    Else
        MsgBox "De map met de afbeeldingen is niet automatisch gevonden." & vbCrLf & _
               "Kies in het volgende venster de map waarin deze bestanden staan:" & vbCrLf & Join(afb, vbCrLf), vbInformation
    End If

    Do
        With Application.FileDialog(4)   ' msoFileDialogFolderPicker
            .Title = "Kies de map met de BAZA-afbeeldingen"
            If .Show <> -1 Then Exit Function
            gekozen = MetScheiding(.SelectedItems(1))
        End With
        mist = Ontbrekend(gekozen, afb)
        If mist = "" Then
            OnthoudMap gekozen
            ZoekAfbeeldingenMap = gekozen
            Exit Function
        End If
        If MsgBox("In deze map ontbreekt:" & vbCrLf & mist & vbCrLf & vbCrLf & _
                  "Een andere map kiezen?", vbExclamation + vbYesNo) = vbNo Then Exit Function
    Loop
End Function

' Geeft de bestandsnamen terug die niet in de map staan (leeg = alles aanwezig)
Private Function Ontbrekend(ByVal map As String, ByVal afb As Variant) As String
    Dim i As Long, gevonden As String, r As String
    For i = LBound(afb) To UBound(afb)
        gevonden = ""
        On Error Resume Next
        gevonden = Dir(map & afb(i))
        On Error GoTo 0
        If gevonden = "" Then r = r & IIf(r = "", "", vbCrLf) & "- " & afb(i)
    Next i
    Ontbrekend = r
End Function

Private Function MetScheiding(ByVal map As String) As String
    If Right$(map, 1) <> Application.PathSeparator Then map = map & Application.PathSeparator
    MetScheiding = map
End Function

Private Function OnthoudenMap() As String
    On Error Resume Next
    OnthoudenMap = Evaluate(ThisWorkbook.Names("BazaAfbeeldingenMap").RefersTo)
    On Error GoTo 0
End Function

Private Sub OnthoudMap(ByVal map As String)
    On Error Resume Next
    ThisWorkbook.Names("BazaAfbeeldingenMap").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="BazaAfbeeldingenMap", RefersTo:="=""" & map & """", Visible:=False
End Sub

' Vier displays naast elkaar. Per kolom: omschrijving, foto, uitleverweek, inhoud, consumentenprijs en recht van retour.
' De nummers 1 t/m 4 verwijzen naar de afbeeldingen in de volgorde van afb.
Private Function AanbodTabel() As String
    Dim namen As Variant, weken As Variant, inhoud As Variant, prijs As Variant, retour As Variant
    Dim t As String, i As Long, cel As String

    namen = Array("Zadendisplay", "Bloemenweide mix doos", "Hangtuin Aardbeien mix Displaydoos", "Keuken Tuin Tomaten Displaydoos")
    weken = Array("Week 3", "Week 3", "Week 12", "Week 12")
    inhoud = Array("850 zakjes", "30 CE", "30 CE", "30 CE")
    ' Consumentenprijs per stuk (&euro; = euroteken in de mail)
    prijs = Array("&euro; 1,79", "&euro; 6,99", "&euro; 6,99", "&euro; 6,99")
    retour = Array(True, False, False, False)
    cel = "align=""center"" width=""25%"" style=""padding:8px 10px;font-family:Calibri,Arial,sans-serif;font-size:11pt;border-left:1px solid #DDDDDD;border-right:1px solid #DDDDDD;"

    t = "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""680"" style=""border-collapse:collapse;margin:4px 0 16px 0;"">"

    ' Omschrijving (groene kopregel)
    t = t & "<tr>"
    For i = 0 To 3
        t = t & "<td " & Replace(cel, "padding:8px 10px;", "padding:10px;background:#8DB63C;color:#ffffff;font-weight:bold;") & _
            "border-left:1px solid #FFFFFF;"" valign=""middle"" height=""48"">" & namen(i) & "</td>"
    Next i
    t = t & "</tr>"

    ' Foto
    t = t & "<tr>"
    For i = 0 To 3
        t = t & "<td " & cel & """ valign=""middle""><img src=""cid:baza" & (i + 1) & """ width=""150"" alt=""" & namen(i) & """ style=""display:block;margin:0 auto;""></td>"
    Next i
    t = t & "</tr>"

    ' Uitleverweek
    t = t & "<tr>"
    For i = 0 To 3
        t = t & "<td " & cel & "border-top:1px solid #DDDDDD;"">Uitlevering:<br><b>" & weken(i) & "</b></td>"
    Next i
    t = t & "</tr>"

    ' Inhoud
    t = t & "<tr>"
    For i = 0 To 3
        t = t & "<td " & cel & "border-top:1px solid #DDDDDD;"">Inhoud:<br><b>" & inhoud(i) & "</b></td>"
    Next i
    t = t & "</tr>"

    ' Consumentenprijs per stuk
    t = t & "<tr>"
    For i = 0 To 3
        t = t & "<td " & cel & "border-top:1px solid #DDDDDD;"">Consumentenprijs:<br><b>" & prijs(i) & "</b> per stuk</td>"
    Next i
    t = t & "</tr>"

    ' Recht van retour: Ja in groen, Nee in rood
    t = t & "<tr>"
    For i = 0 To 3
        t = t & "<td " & cel & "border-top:1px solid #DDDDDD;border-bottom:1px solid #DDDDDD;"">Recht van retour:<br>" & _
            IIf(retour(i), "<b style=""color:#8DB63C;"">JA</b>", "<b style=""color:#D52B1E;"">NEE</b>") & "</td>"
    Next i
    t = t & "</tr></table>"

    AanbodTabel = t
End Function

' Vraagt het jaar (standaard volgend jaar) en het soort mail.
Private Function VraagInstellingen() As Boolean
    Dim antw As String

    antw = InputBox("Voor welk jaar is deze mailing?", "BAZA mailing: jaar", CStr(Year(Date) + 1))
    If StrPtr(antw) = 0 Then Exit Function          ' Annuleren
    antw = Trim$(antw)
    If Not IsNumeric(antw) Or Len(antw) <> 4 Then
        MsgBox "Vul een jaartal in, bijvoorbeeld " & (Year(Date) + 1) & ".", vbExclamation
        Exit Function
    End If
    mJaar = CLng(antw)

    antw = InputBox("Welk soort mail?" & vbCrLf & vbCrLf & _
                    "1 = Gewone mail (niets voor het onderwerp)" & vbCrLf & _
                    "2 = Correctie" & vbCrLf & _
                    "3 = Reminder", "BAZA mailing: soort mail", "1")
    If StrPtr(antw) = 0 Then Exit Function
    Select Case Trim$(antw)
        Case "1", "": mSoort = ""
        Case "2": mSoort = "Correctie"
        Case "3": mSoort = "Reminder"
        Case Else
            MsgBox "Kies 1, 2 of 3.", vbExclamation
            Exit Function
    End Select

    VraagInstellingen = True
End Function

' Elk jaar en elk soort mail heeft een eigen log, bijvoorbeeld
' "Verzendlog 2027" of "Verzendlog 2027 Reminder".
Private Function LogNaam() As String
    LogNaam = "Verzendlog " & mJaar & IIf(mSoort = "", "", " " & mSoort)
End Function

' Onderwerp met winkelgegevens, bijvoorbeeld:
' "2027 Intekenen Zaden (Displays & Mixdozen) | 1234/6123 PLUS Jansen"
' Bij een correctie of reminder komt "Correctie: " of "Reminder: " ervoor.
' Het SAPnr komt er alleen bij als het bekend is. Staat het SAPnr al tussen
' haakjes in de winkelnaam, dan wordt het daar weggehaald zodat het niet dubbel staat.
Private Function Onderwerpregel(ByVal wnk As Variant, ByVal naam As Variant, ByVal sap As Variant) As String
    Dim n As String, s As String, deel As String
    n = Trim$(CStr(naam))
    s = Trim$(CStr(sap))
    If s <> "" Then
        n = Trim$(Replace(n, "(" & s & ")", ""))
        n = Replace(n, "  ", " ")
    End If
    deel = Trim$(CStr(wnk))
    If s <> "" Then deel = deel & "/" & s
    Onderwerpregel = IIf(mSoort = "", "", mSoort & ": ") & mJaar & " " & ONDERWERP & " | " & Trim$(deel & " " & n)
End Function

' Contactpersoon 2 t/m 4 als ;-gescheiden lijst
Private Function Adressen(ByVal rij As Range, ParamArray kolommen() As Variant) As String
    Dim k As Variant, a As String, r As String
    For Each k In kolommen
        If k > 0 Then
            a = Trim$(CStr(rij.Cells(1, k).Value))
            If a <> "" Then r = r & IIf(r = "", "", ";") & a
        End If
    Next k
    Adressen = r
End Function

Private Function ZoekTabel() As Object
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If KolomIndex(lo, "Formulierlink") > 0 And KolomIndex(lo, "Emailadres") > 0 _
               And KolomIndex(lo, "Winkelnummer") > 0 Then
                Set ZoekTabel = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function KolomIndex(ByVal lo As ListObject, ByVal naam As String) As Long
    On Error Resume Next
    KolomIndex = lo.ListColumns(naam).Index
    On Error GoTo 0
End Function

Private Function LogBlad() As Worksheet
    On Error Resume Next
    Set LogBlad = ThisWorkbook.Worksheets(LogNaam())
    On Error GoTo 0
    If LogBlad Is Nothing Then
        Set LogBlad = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        LogBlad.Name = LogNaam()
        LogBlad.Range("A1:E1").Value = Array("Winkelnummer", "Winkelnaam", "Aan", "CC", "Klaargezet op")
        LogBlad.Range("A1:E1").Font.Bold = True
    End If
End Function

Private Sub SchrijfLog(ByVal ws As Worksheet, ByVal wnk As Variant, ByVal naam As Variant, _
                       ByVal aan As String, ByVal cc As String)
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = wnk
    ws.Cells(r, 2).Value = naam
    ws.Cells(r, 3).Value = aan
    ws.Cells(r, 4).Value = cc
    ws.Cells(r, 5).Value = Now
End Sub
