# as_xlsx
generate an Excel xlsx file from an Oracle database. Or read data from a xlsx.

This is an updated version of the original, still available at the [AMIS blog](https://technology.amis.nl/languages/oracle-plsql/create-an-excel-file-with-plsql/). Original is at version 20, this one is version 40, so there are some improvements. But I must admit, last time I did change something is several years ago. I can't recall what and why.

And this version does include an updated (finally using XMLtable :stuck_out_tongue_winking_eye:) version of [as_read_xlsx](https://technology.amis.nl/languages/oracle-plsql/read-a-excel-xlsx-with-plsql/)

## Eigene Styles (named styles)

Seit `as_xlsx52` kann ein Workbook benannte Styles haben, die Schrift, Farben,
Rahmen, Zahlformat und Ausrichtung bündeln. Ein Standard-Template (monospaced
Schrift Hack 8, Akzentfarbe `#2F4F6F`, Zebra-Zeilen, CHF-Beträge, Schweizer Datum)
wird mit einem Aufruf aktiviert:

```sql
begin
  as_xlsx.clear_workbook;
  as_xlsx.use_template;                      -- oder: use_template( p_font => 'Consolas', p_accent => 'FFC00000' )
  as_xlsx.query2sheet( 'select * from auftraege'
                     , p_title       => 'Aufträge'
                     , p_styled      => true
                     , p_sum_columns => 'BETRAG'
                     , p_col_styles  => 'BETRAG=betrag' );
  as_xlsx.save( 'MY_DIR', 'auftraege.xlsx' );
end;
```

Details und alle Styles: [usage.md](usage.md#named-styles), Konzept und Muster:
[docs/style-konzept.md](docs/style-konzept.md). Bestehende Aufrufe ohne Styles
verhalten sich unverändert.

