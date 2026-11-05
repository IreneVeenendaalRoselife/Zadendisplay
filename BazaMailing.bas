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
Private Const ONDERWERP As String = "Intekenen BAZA zaden en displaydozen"
Private Const PER_KEER As Long = 25
Private Const LOG_BLAD As String = "Verzendlog"
Private Const CONTACTPERSONEN_IN_CC As Boolean = True   ' False = alle adressen in Aan
' Map met de afbeeldingen (Zadendisplay.jpg, Bloemenweide.jpg, AardbeienTuin.jpg, TomatenTuin.jpg).
Private Const AFBEELDINGEN_MAP As String = "S:\Roselife\2027\Zadendisplays\Afbeeldingen displays dozen"
' -----------------------------------------------------------------------------

Private Const PR_ATTACH_CONTENT_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
Private Const PR_ATTACHMENT_HIDDEN As String = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

Public Sub BazaMailingKlaarzetten()
    Dim lo As Object, ol As Object, account As Object
    Dim logWs As Worksheet, gehad As Object
    Dim map As String, afb As Variant, i As Long
    Dim cWnk As Long, cNaam As Long, cMail As Long, cC2 As Long, cC3 As Long, cC4 As Long, cLink As Long
    Dim rij As Range, wnk As String, mailAdres As String
    Dim teDoen As Collection, item As Variant
    Dim aantal As Long, totaal As Long, mail As Object

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
    cC2 = KolomIndex(lo, "Contactpersoon2")
    cC3 = KolomIndex(lo, "Contactpersoon3")
    cC4 = KolomIndex(lo, "Contactpersoon4")

    ' Afbeeldingen controleren
    afb = Array("Zadendisplay.jpg", "Bloemenweide.jpg", "AardbeienTuin.jpg", "TomatenTuin.jpg")
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
               "(Alle winkels staan al in het tabblad " & LOG_BLAD & ".)", vbInformation
        Exit Sub
    End If
    If MsgBox(totaal & " winkels nog te mailen." & vbCrLf & vbCrLf & _
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
        mail.Subject = ONDERWERP

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
    h = h & "<p>Ook dit jaar is het weer mogelijk om in te tekenen op het assortiment van BAZA!<br>"
    h = h & "Naast de vertrouwde jaarlijkse zadendisplay kun je dit jaar ook kiezen voor diverse andere displaydozen.</p>"
    h = h & "<p>Het aanbod bestaat uit:</p>"

    ' Tabel met het aanbod: afbeelding, omschrijving en inhoud
    h = h & "<table cellpadding=""8"" cellspacing=""0"" border=""0"" style=""border-collapse:collapse;margin:4px 0 16px 0;"">"
    h = h & "<tr style=""background:#8DB63C;color:#ffffff;"">"
    h = h & "<th align=""left"" style=""padding:8px 12px;font-family:Calibri,Arial,sans-serif;"">Afbeelding</th>"
    h = h & "<th align=""left"" style=""padding:8px 12px;font-family:Calibri,Arial,sans-serif;"">Omschrijving</th>"
    h = h & "<th align=""left"" style=""padding:8px 12px;font-family:Calibri,Arial,sans-serif;"">Inhoud</th></tr>"
    h = h & AanbodRij(1, "Zadendisplay", "850 zakjes")
    h = h & AanbodRij(2, "Bloemenweide mix doos", "30 CE")
    h = h & AanbodRij(3, "Hangtuin Aardbeien mix Displaydoos", "30 CE")
    h = h & AanbodRij(4, "Keuken Tuin Tomaten Displaydoos", "30 CE")
    h = h & "</table>"
    h = h & "<p>Voordat je gaat intekenen, is het belangrijk om rekening te houden met de volgende informatie:</p><ul>"
    h = h & "<li>De Zadendisplay en de Bloemenweide doos worden uitgeleverd in week 3.</li>"
    h = h & "<li>De Hangtuin Aardbeien Mix Box en de Tomaten Keukentuintjes Box worden uitgeleverd in week 12.</li>"
    h = h & "<li>Alle displays worden kant-en-klaar geleverd.</li></ul><ul>"
    h = h & "<li>Besteladvies: minimaal 1 zadendisplay (850 zakjes).</li>"
    h = h & "<li>Voor de zadendisplay geldt 100% recht van retour. Overgebleven zaden kunnen in week 22 van 2027 per post worden geretourneerd. Verdere instructies hierover volgen op een later moment.</li>"
    h = h & "<li>Voor de overige boxen geldt geen recht van retour.</li></ul>"

    ' Afgeronde knop met de persoonlijke formulierlink.
    ' Outlook (Windows) tekent afgeronde hoeken alleen via VML; andere
    ' mailprogramma's gebruiken de gewone link met border-radius.
    h = h & "<table cellpadding=""0"" cellspacing=""0"" border=""0"" style=""margin:8px 0 24px 0;""><tr><td>"
    h = h & "<!--[if mso]>"
    h = h & "<v:roundrect xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:w=""urn:schemas-microsoft-com:office:word"" "
    h = h & "href=""" & link & """ style=""height:46px;v-text-anchor:middle;width:260px;"" arcsize=""50%"" "
    h = h & "strokecolor=""#7AA22F"" fillcolor=""#8DB63C"">"
    h = h & "<w:anchorlock/><center style=""color:#ffffff;font-family:Calibri,Arial,sans-serif;font-size:13pt;font-weight:bold;"">"
    h = h & "Klik hier om in te tekenen</center></v:roundrect>"
    h = h & "<![endif]-->"
    h = h & "<!--[if !mso]><!-->"
    h = h & "<a href=""" & link & """ style=""display:inline-block;background:#8DB63C;color:#ffffff;"
    h = h & "font-family:Calibri,Arial,sans-serif;font-size:13pt;font-weight:bold;text-decoration:none;"
    h = h & "padding:12px 32px;border-radius:23px;border:1px solid #7AA22F;"">Klik hier om in te tekenen</a>"
    h = h & "<!--<![endif]-->"
    h = h & "</td></tr></table>"

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
    Dim kandidaten As Collection, k As Variant, gekozen As String

    Set kandidaten = New Collection
    If AFBEELDINGEN_MAP <> "" Then kandidaten.Add AFBEELDINGEN_MAP
    gekozen = OnthoudenMap()
    If gekozen <> "" Then kandidaten.Add gekozen
    If LCase$(Left$(ThisWorkbook.Path, 4)) <> "http" And ThisWorkbook.Path <> "" Then
        kandidaten.Add ThisWorkbook.Path & Application.PathSeparator & "Afbeeldingen"
    End If

    For Each k In kandidaten
        If AlleAfbeeldingen(MetScheiding(CStr(k)), afb) Then
            ZoekAfbeeldingenMap = MetScheiding(CStr(k))
            Exit Function
        End If
    Next k

    MsgBox "De map met de afbeeldingen (" & Join(afb, ", ") & ") is niet automatisch gevonden." & vbCrLf & _
           "Kies in het volgende venster de map waarin de afbeeldingen staan.", vbInformation
    Do
        With Application.FileDialog(4)   ' msoFileDialogFolderPicker
            .Title = "Kies de map met de BAZA-afbeeldingen"
            If .Show <> -1 Then Exit Function
            gekozen = MetScheiding(.SelectedItems(1))
        End With
        If AlleAfbeeldingen(gekozen, afb) Then
            OnthoudMap gekozen
            ZoekAfbeeldingenMap = gekozen
            Exit Function
        End If
        If MsgBox("In deze map staan niet alle afbeeldingen:" & vbCrLf & Join(afb, vbCrLf) & vbCrLf & vbCrLf & _
                  "Een andere map kiezen?", vbExclamation + vbYesNo) = vbNo Then Exit Function
    Loop
End Function

Private Function AlleAfbeeldingen(ByVal map As String, ByVal afb As Variant) As Boolean
    Dim i As Long, gevonden As String
    For i = LBound(afb) To UBound(afb)
        gevonden = ""
        On Error Resume Next
        gevonden = Dir(map & afb(i))
        On Error GoTo 0
        If gevonden = "" Then Exit Function
    Next i
    AlleAfbeeldingen = True
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

Private Function AanbodRij(ByVal nr As Long, ByVal omschrijving As String, ByVal inhoud As String) As String
    Dim td As String
    td = "style=""padding:8px 12px;border-bottom:1px solid #DDDDDD;font-family:Calibri,Arial,sans-serif;font-size:11pt;"""
    AanbodRij = "<tr>" & _
        "<td valign=""middle"" " & td & "><img src=""cid:baza" & nr & """ width=""180"" alt=""" & omschrijving & """ style=""display:block;""></td>" & _
        "<td valign=""middle"" " & td & "><b>" & omschrijving & "</b></td>" & _
        "<td valign=""middle"" " & td & ">" & inhoud & "</td></tr>"
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
    Set LogBlad = ThisWorkbook.Worksheets(LOG_BLAD)
    On Error GoTo 0
    If LogBlad Is Nothing Then
        Set LogBlad = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        LogBlad.Name = LOG_BLAD
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
