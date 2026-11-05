~~~
begin
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  as_xlsx.cell( 5, 1, 5 );
  as_xlsx.cell( 3, 1, 3 );
  as_xlsx.cell( 2, 2, 45 );
  as_xlsx.cell( 3, 2, 'Anton Scheffer', p_alignment => as_xlsx.get_alignment( p_wraptext => true ) );
  as_xlsx.cell( 1, 4, sysdate, p_fontId => as_xlsx.get_font( 'Calibri', p_rgb => 'FFFF0000' ) );
  as_xlsx.cell( 2, 4, sysdate, p_numFmtId => as_xlsx.get_numFmt( 'dd/mm/yyyy h:mm' ) );
  as_xlsx.cell( 3, 4, sysdate, p_numFmtId => as_xlsx.get_numFmt( as_xlsx.orafmt2excel( 'dd/mon/yyyy' ) ) );
  as_xlsx.cell( 5, 5, 75, p_borderId => as_xlsx.get_border( 'double', 'double', 'double', 'double' ) );
  as_xlsx.cell( 2, 3, 33 );
  as_xlsx.hyperlink( 1, 6, 'http://www.amis.nl', 'Amis site' );
  as_xlsx.cell( 1, 7, 'Some merged cells', p_alignment => as_xlsx.get_alignment( p_horizontal => 'center' ) );
  as_xlsx.mergecells( 1, 7, 3, 7 );
  for i in 1 .. 5
  loop
    as_xlsx.comment( 3, i + 3, 'Row ' || (i+3), 'Anton' );
  end loop;
  as_xlsx.new_sheet;
  as_xlsx.set_row( 1, p_fillId => as_xlsx.get_fill( 'solid', 'FFFF0000' ) ) ;
  for i in 1 .. 5
  loop
    as_xlsx.cell( 1, i, i );
    as_xlsx.cell( 2, i, i * 3 );
    as_xlsx.cell( 3, i, 'x ' || i * 3 );
  end loop;
  as_xlsx.query2sheet( 'select rownum, x.*
, case when mod( rownum, 2 ) = 0 then rownum * 3 end demo
, case when mod( rownum, 2 ) = 1 then ''demo '' || rownum end demo2 from dual x connect by rownum <= 5' );
  as_xlsx.save( 'MY_DIR', 'my.xlsx' );
end;
~~~
~~~
begin
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  as_xlsx.cell( 1, 6, 5 );
  as_xlsx.cell( 1, 7, 3 );
  as_xlsx.cell( 1, 8, 7 );
  as_xlsx.new_sheet;
  as_xlsx.cell( 2, 6, 15, p_sheet => 2 );
  as_xlsx.cell( 2, 7, 13, p_sheet => 2 );
  as_xlsx.cell( 2, 8, 17, p_sheet => 2 );
  as_xlsx.list_validation( 6, 3, 1, 6, 1, 8, p_show_error => true, p_sheet => 1 );
  as_xlsx.defined_name( 2, 6, 2, 8, 'Anton', 2 );
  as_xlsx.list_validation
    ( 6, 1, 'Anton'
    , p_style => 'information'
    , p_title => 'valid values are'
    , p_prompt => '13, 15 and 17'
    , p_show_error => true
    , p_error_title => 'Are you sure?'
    , p_error_txt => 'Valid values are: 13, 15 and 17'
    , p_sheet => 1 );
  as_xlsx.save( 'MY_DIR', 'my.xlsx' );
end;
~~~
~~~
begin
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  as_xlsx.cell( 1, 6, 5 );
  as_xlsx.cell( 1, 7, 3 );
  as_xlsx.cell( 1, 8, 7 );
  as_xlsx.set_autofilter( 1,1, p_row_start => 5, p_row_end => 8 );
  as_xlsx.new_sheet;
  as_xlsx.cell( 2, 6, 5 );
  as_xlsx.cell( 2, 7, 3 );
  as_xlsx.cell( 2, 8, 7 );
  as_xlsx.set_autofilter( 2,2, p_row_start => 5, p_row_end => 8 );
  as_xlsx.save( 'MY_DIR', 'my.xlsx' );
end;
~~~
~~~
begin
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  as_xlsx.setUseXf( false );
  for c in 1 .. 10
  loop
    as_xlsx.cell( c, 1, 'COL' || c );
    as_xlsx.cell( c, 2, 'val' || c );
    as_xlsx.cell( c, 3, c );
  end loop;
  as_xlsx.freeze_rows( 1 );
  as_xlsx.new_sheet;
  for r in 1 .. 10
  loop
    as_xlsx.cell( 1, r, 'ROW' || r );
    as_xlsx.cell( 2, r, 'val' || r );
    as_xlsx.cell( 3, r, r );
  end loop;
  as_xlsx.freeze_cols( 3 );
  as_xlsx.new_sheet;
  as_xlsx.cell( 3, 3, 'Start freeze' );
  as_xlsx.freeze_pane( 3,3 );
  as_xlsx.save( 'MY_DIR', 'my.xlsx' );
end;
~~~
~~~
  select *
  from table( as_xlsx.read( as_xlsx.file2blob( 'MY_DIR', 'test.xlsx' ), '1' ) )
~~~
~~~
begin
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  as_xlsx.add_image( 1, 1, as_barcode.barcode( 'https://github.com/antonscheffer/as_xlsx', 'QR' ) );
  as_xlsx.cell( 1, 8, 'now with png images' );
  as_xlsx.save( 'MY_DIR', 'my.xlsx' );
end;
~~~
~~~
declare
  l_cnt pls_integer;
  l_query sys_refcursor;
begin
  open l_query for
    select date '1900-02-26' + level "Secret Date"
         , to_char( date '1900-02-26' + level, 'yyyy mon dd' ) "Secret String"
    from dual
    connect by level < 8;
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  l_cnt := as_xlsx.query2sheet
             ( p_rc         => l_query
             , p_sheet      => 1
             , p_col        => 5
             , p_row        => 3
             , p_autofilter => true
             , p_date_format => 'yyyy-mmm-dd'
             , p_title      => 'My Secrets'
             , p_title_xfid => as_xlsx.get_xfid( p_alignment => as_xlsx.get_alignment( p_horizontal => 'centerContinuous' ) )
             );
  as_xlsx.set_column_width( p_col   => 5
                          , p_width => 15
                           );
  as_xlsx.set_column_width( p_col   => 6
                          , p_width => 15
                          );
  as_xlsx.cell( 5
              , l_cnt
                 + 3  -- query start row
                 + 2  -- title + headers 
                 + 1  -- interval 
              , 'Rows returned: ' || l_cnt );
  -- make sure you have set as_xlsx.use_dbms_crypto = true; in the package specification
  as_xlsx.save( 'MY_DIR', 'my.xlsx', 'demo' );
end;
~~~
~~~
begin
  as_xlsx.clear_workbook;
  as_xlsx.new_sheet;
  as_xlsx.cell( 1, 1, 3 );
  as_xlsx.cell( 1, 2, 5 );
  as_xlsx.cell( 1, 3, 4 );
  as_xlsx.num_formula( 2, 1, 'A6+B6' );
  as_xlsx.num_formula( 2, 6, 'SUM(A1:A5)' );
  as_xlsx.num_formula( 1, 6, 'SUM(A1:A2)' );
  as_xlsx.date_formula( 5, 1, 'TODAY()', p_numFmtId => as_xlsx.get_numFmt( 'yyyy-mm-dd' ) );
  as_xlsx.str_formula( 5, 3, 'LOWER(TEXT(TODAY(),"DDDD"))' );
  as_xlsx.new_sheet;
  as_xlsx.cell( 1, 1, 13 );
  as_xlsx.cell( 1, 2, 15 );
  as_xlsx.cell( 1, 3, 14 );
  as_xlsx.num_formula( 2, 1, 'A6+B6' );
  as_xlsx.num_formula( 2, 6, 'SUM(A1:A5)' );
  as_xlsx.num_formula( 1, 6, 'SUM(A1:A2)' );
  as_xlsx.save( 'MY_DIR', 'my.xlsx' );
end;
~~~

## Named styles

`use_template` sets the workbook default font (Hack 8) and defines these styles:
`titel`, `untertitel`, `header`, `text`, `ganzzahl`, `zahl`, `betrag`, `prozent`,
`datum`, `datum_zeit`, `summe`, `hinweis`.
Call it (or `set_default_font`) right after `clear_workbook`.

| parameter of `use_template` | default |
|---|---|
| `p_font` / `p_fontsize` | `'Hack'` / `8` |
| `p_accent` | `'FF2F4F6F'` (title, header, total lines, tab colour) |
| `p_zebra` | `'FFF2F4F7'` (every second data row in `query2sheet`) |
| `p_line` | `'FFD0D5DD'` (thin line under data cells) |
| `p_muted` | `'FF6B7280'` (subtitle and notes) |
| `p_currency` | `'CHF'` (style `betrag`) |
| `p_date_format` | `'dd.mm.yyyy'` |

`define_style` adds a style or overrides one. `p_base` starts from an existing style.
Every procedure that writes a cell accepts `p_style`; explicit `p_numFmtId`, `p_fontId`,
`p_fillId`, `p_borderId` and `p_alignment` still win over the style.
~~~
begin
  as_xlsx.clear_workbook;
  as_xlsx.use_template;
  as_xlsx.define_style( 'warnung', p_base => 'text', p_bold => true, p_font_rgb => 'FFC00000', p_fill_rgb => 'FFFDECEC' );
  as_xlsx.new_sheet( 'Styles', p_show_gridlines => false );
  as_xlsx.cell( 1, 1, 'Monatsbericht', p_style => 'titel' );
  as_xlsx.cell( 1, 3, 'Betrag', p_style => 'header' );
  as_xlsx.cell( 1, 4, 1234.5, p_style => 'betrag' );
  as_xlsx.cell( 1, 5, 99.95, p_style => 'betrag' );
  as_xlsx.num_formula( 1, 6, 'SUM(A4:A5)', 1334.45, p_style => 'summe', p_numFmtId => as_xlsx.get_numFmt( '#,##0.00' ) );
  as_xlsx.cell( 1, 8, 'Achtung', p_style => 'warnung' );
  as_xlsx.cell( 2, 4, sysdate, p_style => 'datum' );
  -- p_xfId with get_style_xfid, e.g. for a zebra row
  as_xlsx.cell( 2, 5, sysdate, p_xfId => as_xlsx.get_style_xfid( 'datum', p_zebra => true ) );
  as_xlsx.save( 'MY_DIR', 'styles.xlsx' );
end;
~~~

### Styled query2sheet

`p_styled => true` uses the styles for header, data, title and total row
(`use_template` is applied if no styles exist). It also adds zebra rows, column widths,
a frozen header row and an autofilter. On a new sheet the gridlines are hidden and
the tab gets the accent colour.

* Column style by data type: `NUMBER(p,0)` = `ganzzahl`, `NUMBER(p,s)` = `zahl`,
  other numbers = `text`, `DATE` = `datum`, timestamps = `datum_zeit`, strings = `text`.
* `p_col_styles => 'BETRAG=betrag, QUOTE=prozent'` overrides the style per column.
* `p_sum_columns => 'MENGE, BETRAG'` (or `'ALL'` for every numeric column) adds a total row
  with `=SUBTOTAL(9, ...)`, so the totals follow the autofilter.
~~~
declare
  l_rows number;
begin
  as_xlsx.clear_workbook;
  as_xlsx.use_template( p_accent => 'FFC00000' );
  l_rows := as_xlsx.query2sheet
              ( 'select level nr, cast( level * 12.5 as number(10,2) ) betrag, sysdate - level datum
                 from dual connect by level <= 10'
              , p_title       => 'Beispiel'
              , p_styled      => true
              , p_sum_columns => 'BETRAG'
              , p_col_styles  => 'BETRAG=betrag'
              );
  as_xlsx.save( 'MY_DIR', 'styled.xlsx' );
end;
~~~

## Save an xlsx on your own PC (SQL Developer, SQLcl, Toad)

`SPOOL` only writes text, so it can't produce an xlsx. `query2blob` builds the whole
workbook for one query and returns it as a BLOB from a plain `select ... from dual`.
No directory object and no `UTL_FILE` rights are needed.
~~~
select as_xlsx.query2blob
         ( 'select empno, ename, hiredate, sal from emp'
         , p_title       => 'Mitarbeiter'
         , p_sum_columns => 'SAL'
         , p_col_styles  => 'SAL=betrag'
         ) as xlsx
from dual;
~~~
Parameters (all optional except `p_sql`): `p_title`, `p_sheetname`,
`p_styled` (`'Y'` default, `'N'` = plain export), `p_sum_columns`, `p_col_styles`,
`p_font` (default Hack), `p_fontsize` (default 8), `p_accent` (default `FF2F4F6F`), `p_password`.
Quotes inside the query have to be doubled, or use the `q'[...]'` notation:
`as_xlsx.query2blob( q'[select * from emp where job = 'CLERK']' )`.

**SQL Developer**

1. Run the select with *Run Statement* (Ctrl+Enter).
2. In the *Query Result* grid, double click the `XLSX` cell (or click the pencil icon in it).
3. In the *View Value* dialog, click *Save...* and save the file as `name.xlsx`.

**Toad**: right click the BLOB cell, choose *Save to file*, name it `name.xlsx`.

**SQLcl** has no file dialog: use `as_xlsx.save` to a directory on the server instead.

