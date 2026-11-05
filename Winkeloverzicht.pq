// Winkeloverzicht: één regel per winkel met Winkelnummer, SAPnr, Winkelnaam, Emailadres,
// Contactpersoon2 t/m 4, Adres, Postcode, Plaats en de vooringevulde Formulierlink.
// Bron: afleveradressen in Navision (SQL05, NAV-NOVIFLORA), zelfde logica als qpt_Winkels.
let
    Bron        = Sql.Database("SQL05", "NAV-NOVIFLORA"),
    Afleveradres = Bron{[Schema = "dbo", Item = "Roselife BV$Afleveradres"]}[Data],
    Klant       = Bron{[Schema = "dbo", Item = "Roselife BV$Klant"]}[Data],
    Contactpersonen = Bron{[Schema = "dbo", Item = "Roselife BV$Contactpersonen"]}[Data],

    KlantActief = Table.SelectColumns(Table.SelectRows(Klant, each [Geblokkeerd] = 0), {"Nr_"}),
    Adressen    = Table.SelectRows(Afleveradres, each [#"Distribution Centre (DC)"] = 0),
    MetKlant    = Table.Join(Adressen, "Klantnr_", KlantActief, "Nr_", JoinKind.Inner),
    Kolommen    = Table.SelectColumns(MetKlant, {"Code", "Fax", "Naam", "Naam 2", "Adres", "Adres 2", "Postcode", "Plaats", "E-mail", "Blocked"}),

    Getallen = Table.AddColumn(
        Table.AddColumn(Kolommen, "CodeNr", each try Int64.From(Text.Trim([Code])) otherwise null, Int64.Type),
        "FaxNr", each try Int64.From(Text.Trim([Fax])) otherwise null, Int64.Type),
    Winkels  = Table.SelectRows(Getallen, each [CodeNr] <> null),

    Namen = Table.AddColumn(Winkels, "NaamInfo", each
        let
            Volledig = Text.Trim(Text.Combine({[Naam], [#"Naam 2"]}, " ")),
            Pos      = Text.PositionOf(Volledig, "(", Occurrence.Last),
            Tussen   = if Pos >= 0 then Text.BeforeDelimiter(Text.Range(Volledig, Pos + 1), ")") else "",
            HeeftSAP = Pos >= 0 and List.Contains({"6", "7"}, Text.Start(Tussen, 1))
        in
            [ Winkelnaam = Volledig,
              Adresregel = Text.Trim(Text.Combine({[Adres], [#"Adres 2"]}, " ")),
              SAPNaam    = if HeeftSAP then (try Int64.From(Tussen) otherwise null) else null ]),
    NaamUit = Table.ExpandRecordColumn(Namen, "NaamInfo", {"Winkelnaam", "Adresregel", "SAPNaam"}),

    OudNaarNieuw = Table.Distinct(
        Table.RenameColumns(
            Table.SelectColumns(
                Table.SelectRows(NaamUit, each [CodeNr] < 2000 and [FaxNr] <> null and [FaxNr] >= 2000),
                {"FaxNr", "CodeNr"}),
            {{"FaxNr", "NieuwNr"}, {"CodeNr", "OudNr"}}),
        {"NieuwNr"}),
    MetOud = Table.ExpandTableColumn(
        Table.NestedJoin(NaamUit, {"CodeNr"}, OudNaarNieuw, {"NieuwNr"}, "Oud", JoinKind.LeftOuter),
        "Oud", {"OudNr"}),

    MetNummers = Table.AddColumn(
        Table.AddColumn(MetOud, "Winkelnummer", each
            if [CodeNr] < 2000 then [CodeNr]
            else if [FaxNr] <> null and [FaxNr] > 0 and [FaxNr] < 2000 then [FaxNr]
            else if [OudNr] <> null then [OudNr]
            else [CodeNr], Int64.Type),
        "SAPnr", each
            if [CodeNr] >= 6000 and [CodeNr] <= 7999 then [CodeNr]
            else if [SAPNaam] <> null then [SAPNaam]
            else if [FaxNr] <> null and [FaxNr] >= 6000 and [FaxNr] <= 7999 then [FaxNr]
            else null, Int64.Type),

    NietGeblokkeerd = Table.SelectRows(MetNummers, each [Blocked] = 0),

    Gegroepeerd = Table.Group(NietGeblokkeerd, {"Winkelnummer"}, {
        {"Regel", each Table.First(Table.Sort(_, {{"CodeNr", Order.Descending}})), type record},
        {"SAPnr", each List.Max([SAPnr]), Int64.Type},
        {"Emailadres", each List.First(List.Select(
            Table.Sort(_, {{"CodeNr", Order.Descending}})[#"E-mail"],
            each _ <> null and Text.Trim(_) <> "")), type text}}),
    Uitgeklapt = Table.ExpandRecordColumn(Gegroepeerd, "Regel", {"Winkelnaam", "Adresregel", "Postcode", "Plaats"}, {"Winkelnaam", "Adres", "Postcode", "Plaats"}),

    MetEmail = Table.SelectRows(Uitgeklapt, each [Emailadres] <> null),

    // Extra contactpersonen
    CodeNaarWinkel = Table.RenameColumns(
        Table.Distinct(Table.SelectColumns(MetNummers, {"CodeNr", "Winkelnummer"}), {"CodeNr"}),
        {{"CodeNr", "KoppelCode"}}),
    Contacten = Table.SelectRows(
        Table.SelectColumns(
            Table.SelectRows(Contactpersonen, each [Relatienr_] = "RL0002"),
            {"Ship-to Code", "E-mail"}),
        each [#"E-mail"] <> null and Text.Trim([#"E-mail"]) <> ""),
    ContactNr = Table.AddColumn(Contacten, "ContactCode",
        each try Int64.From(Text.Trim([#"Ship-to Code"])) otherwise null, Int64.Type),
    ContactWinkel = Table.Join(
        Table.SelectRows(ContactNr, each [ContactCode] <> null),
        "ContactCode", CodeNaarWinkel, "KoppelCode", JoinKind.Inner),
    ContactLijst = Table.Group(ContactWinkel, {"Winkelnummer"}, {
        {"ContactMails", each List.Distinct(List.Transform([#"E-mail"], Text.Trim), Comparer.OrdinalIgnoreCase), type list}}),
    ContactLijstHernoemd = Table.RenameColumns(ContactLijst, {{"Winkelnummer", "ContactWinkelnr"}}),
    MetContacten = Table.ExpandTableColumn(
        Table.NestedJoin(MetEmail, {"Winkelnummer"}, ContactLijstHernoemd, {"ContactWinkelnr"}, "C", JoinKind.LeftOuter),
        "C", {"ContactMails"}),
    MetExtra = Table.AddColumn(MetContacten, "Extra", each
        let
            Hoofd  = Text.Lower(Text.Trim([Emailadres])),
            Lijst  = if [ContactMails] = null then {} else [ContactMails],
            Zonder = List.Select(Lijst, (m) => Text.Lower(m) <> Hoofd)
        in
            List.FirstN(Zonder, 3)),
    MetKolommen = Table.AddColumn(
        Table.AddColumn(
            Table.AddColumn(MetExtra, "Contactpersoon2", each [Extra]{0}?, type text),
            "Contactpersoon3", each [Extra]{1}?, type text),
        "Contactpersoon4", each [Extra]{2}?, type text),

    // Vooringevulde link naar het Forms-formulier per winkel
    MetLink = Table.AddColumn(MetKolommen, "Formulierlink", each
        "https://forms.cloud.microsoft/Pages/ResponsePage.aspx?id=6ruqOTgp-Ey5TMeAk6ZcQ6RuejvDkKJHqyhO7o3ABYJUM1A3M0REOVg5QjNRWDlOOFhZUEVEWTlTSC4u"
        & "&rae0f80d9a3dc4cb3b6be82e0dd6715e3=" & Uri.EscapeDataString(Text.Trim([Emailadres]))
        & "&r75c04dd08fd342388a762338aa5a5a89=" & Text.From([Winkelnummer])
        & "&ra8cca77592514b5290732cf6569a67a7=" & (if [SAPnr] = null then "" else Text.From([SAPnr]))
        & "&r717aee5eb4e14518a9b14d601147f607=" & Uri.EscapeDataString(if [Winkelnaam] = null then "" else Text.Trim([Winkelnaam]))
        & "&r114949355c7840b9a2118856da1ea010=%22Intekenen%22"
        & "&r59b6a05a77854bbe9142b2a38d9f8ece=0"
        & "&rab46988234ee413bbc826d4caa57d40b=0"
        & "&r3845af3a89f04c14bdbf0cb3d9596bae=0"
        & "&r8480c3934953422e9e0ad50c304e7d23=0",
        type text),

    Resultaat = Table.Sort(
        Table.TransformColumnTypes(
            Table.SelectColumns(MetLink, {"Winkelnummer", "SAPnr", "Winkelnaam", "Emailadres",
                "Contactpersoon2", "Contactpersoon3", "Contactpersoon4", "Adres", "Postcode", "Plaats", "Formulierlink"}),
            {{"Winkelnaam", type text}, {"Emailadres", type text}, {"Adres", type text}, {"Postcode", type text}, {"Plaats", type text}}),
        {{"Winkelnummer", Order.Ascending}})
in
    Resultaat
