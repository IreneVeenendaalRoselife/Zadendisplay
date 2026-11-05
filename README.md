# Zadendisplay: BAZA intekenmailing

- `Winkeloverzicht.pq`: Power Query (M) voor het winkeloverzicht met contactpersonen en vooringevulde formulierlink.
- `BazaMailing.bas`: Excel-macro die per winkel een Outlook-mail klaarzet (nooit automatisch versturen), 25 per keer.
- `Afbeeldingen/`: de vier productafbeeldingen die in de mail worden geplaatst.

## Gebruik

1. Laad `Winkeloverzicht.pq` als tabel in een Excel-werkmap (.xlsm).
2. Zet de map `Afbeeldingen` naast die werkmap.
3. Alt+F11, Bestand > Bestand importeren, kies `BazaMailing.bas`.
4. Start de macro `BazaMailingKlaarzetten` (Alt+F8).
