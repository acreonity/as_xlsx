# Style-Konzept für as_xlsx (Entwurf)

Ziel: Excel-Exporte mit einem eigenen, wiederverwendbaren Style erzeugen
(Schrift, Farben, Rahmen, Zahlformate, Formeln), ohne für jede Zelle
`get_font`/`get_fill`/`get_border`/`get_xfid` von Hand zu kombinieren.

Muster zum Anschauen: [`muster-template.xlsx`](muster-template.xlsx)
(Blatt "Muster-Export" und "Style-Katalog").

## 1. Standard-Template

| Element | Wert |
|---|---|
| Standardschrift | Hack 8, immer monospaced (gilt für das ganze Workbook, auch leere Zellen). Fehlt Hack auf einem PC, ersetzt Excel sie durch eine andere Schrift; Fallback-Empfehlung Consolas. |
| Akzentfarbe | `#2F4F6F` (Titel, Kopfzeile, Summenlinien, Registerfarbe) |
| Zebra-Füllung | `#F2F4F7` jede zweite Datenzeile |
| Linien | `#D0D5DD` fein, nur unten |
| Gitternetz | ausgeblendet |

Benannte Styles:

| Name | Schrift | Füllung / Rahmen | Zahlformat |
|---|---|---|---|
| `titel` | 12 fett, Akzent | – | – |
| `untertitel` | 8 kursiv, grau | – | – |
| `header` | 8 fett, weiss | Akzent-Füllung, Linie unten, Umbruch | – |
| `text` | 8 | feine Linie unten | – |
| `ganzzahl` | 8 | feine Linie unten | `#,##0` |
| `zahl` | 8 | feine Linie unten | `#,##0.00` |
| `betrag` | 8 | feine Linie unten | `"CHF" #,##0.00` |
| `prozent` | 8 | feine Linie unten | `0.0%` |
| `datum` | 8, zentriert | feine Linie unten | `dd.mm.yyyy` |
| `datum_zeit` | 8 | feine Linie unten | `dd.mm.yyyy hh:mm` |
| `summe` | 8 fett | Linie oben, Doppellinie unten | übernimmt Spaltenformat |
| `hinweis` | 7 kursiv, grau | – | – |

## 2. Geplante API (Oracle 19, bestehende Aufrufe bleiben gültig)

```sql
begin
  as_xlsx.clear_workbook;
  -- Template laden, optional mit eigener Schrift/Farbe
  as_xlsx.use_template( p_font => 'Hack', p_fontsize => 8, p_accent => 'FF2F4F6F' );
  -- eigenen Style definieren oder einen Standard-Style überschreiben
  as_xlsx.define_style( 'warnung'
                      , p_font_bold => true, p_font_rgb => 'FFC00000'
                      , p_fill_rgb  => 'FFFDECEC' );
  as_xlsx.new_sheet( 'Aufträge' );
  as_xlsx.cell( 1, 1, 'Auftragsübersicht', p_style => 'titel' );
  -- query2sheet mit Template: Kopfzeile, Zebra, Formate je Datentyp, Summenzeile
  as_xlsx.query2sheet( 'select auftrag, kunde, datum, menge, betrag from auftraege'
                     , p_row         => 3
                     , p_styled      => true
                     , p_sum_columns => 'MENGE,BETRAG'  -- erzeugt =SUBTOTAL(9,…)
                     , p_col_styles  => 'BETRAG=betrag' );
  as_xlsx.num_formula( 6, 20, 'F19*0.081', p_style => 'betrag' );
  as_xlsx.save( 'MY_DIR', 'auftraege.xlsx' );
end;
```

Kernpunkte:

- `use_template` setzt Standardschrift (Font 0 in `styles.xml`) und registriert die benannten Styles.
- `define_style` speichert einen Style unter einem Namen; intern wird er wie bisher auf eine `xfId` abgebildet.
- `p_style` kommt als zusätzlicher, optionaler Parameter zu `cell`, `hyperlink`, `num_formula`, `str_formula`, `date_formula`, `set_row`, `set_column`.
- `query2sheet` wählt bei `p_styled => true` das Format pro Spaltentyp (Zahl, Datum, Text), legt Zebra-Zeilen, Autofilter, fixierte Kopfzeile und optional eine Summenzeile an.
