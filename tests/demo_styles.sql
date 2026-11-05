-- Demo / test for named styles and styled query2sheet.
-- Needs a directory object; change OUT_DIR to your directory.
set serveroutput on
set define off
declare
  l_cnt number;
  l_rc  sys_refcursor;
begin
  as_xlsx.clear_workbook;
  as_xlsx.use_template;   -- Hack 8, accent FF2F4F6F
  as_xlsx.define_style( 'warnung', p_base => 'text', p_bold => true, p_font_rgb => 'FFC00000', p_fill_rgb => 'FFFDECEC' );
  --
  -- sheet 1: query2sheet with template, total row and per column styles
  as_xlsx.new_sheet( 'Aufträge' );
  as_xlsx.cell( 1, 1, 'Auftragsübersicht', p_style => 'titel' );
  as_xlsx.cell( 1, 2, 'Erstellt am ' || to_char( sysdate, 'dd.mm.yyyy' ) || ' · Beispieldaten', p_style => 'untertitel' );
  open l_rc for
    select 'A-' || ( 1000 + level )                               as auftrag
         , case mod( level, 5 ) when 0 then 'Muster AG' when 1 then 'Beispiel GmbH'
                                when 2 then 'Test & Co' when 3 then 'Alpha SA' else 'Beta Logistik' end as kunde
         , date '2026-10-01' + 2 * level                           as datum
         , cast( mod( level * 7, 40 ) + 1 as number(5) )           as menge
         , cast( round( 10 + mod( level * 37, 240 ) + level / 7, 2 ) as number(10,2) ) as einzelpreis
         , case when mod( level, 3 ) = 0 then 0.1 else 0 end       as rabatt
    from dual
    connect by level <= 12;
  l_cnt := as_xlsx.query2sheet
             ( p_rc          => l_rc
             , p_sheet       => 1
             , p_row         => 4
             , p_styled      => true
             , p_sum_columns => 'MENGE'
             , p_col_styles  => 'EINZELPREIS=betrag, RABATT=prozent'
             );
  -- formula columns next to the query result
  -- explicit parameters override the style: right aligned header
  as_xlsx.cell( 7, 4, 'NETTO', p_style => 'header', p_alignment => as_xlsx.get_alignment( p_vertical => 'center', p_horizontal => 'right' ) );
  for r in 5 .. 4 + l_cnt
  loop
    as_xlsx.num_formula( 7, r, 'D' || r || '*E' || r || '*(1-F' || r || ')'
                       , p_xfId => as_xlsx.get_style_xfid( 'betrag', p_zebra => mod( r - 5, 2 ) = 1 ) );
  end loop;
  as_xlsx.num_formula( 7, 5 + l_cnt, 'SUBTOTAL(9,G5:G' || ( 4 + l_cnt ) || ')', p_style => 'summe', p_numFmtId => as_xlsx.get_numFmt( '\C\H\F #,##0.00' ) );
  as_xlsx.set_column_width( 7, 16 );
  as_xlsx.cell( 1, 7 + l_cnt, 'Netto = Menge × Einzelpreis × (1 − Rabatt), Total = TEILERGEBNIS (beachtet Filter)', p_style => 'hinweis' );
  as_xlsx.cell( 1, 8 + l_cnt, 'Achtung: Beispieldaten', p_style => 'warnung' );
  --
  -- sheet 2: all styles
  as_xlsx.new_sheet( 'Style-Katalog', p_show_gridlines => false );
  as_xlsx.cell( 1, 1, 'Style-Katalog', p_style => 'titel' );
  as_xlsx.cell( 1, 3, 'Style', p_style => 'header' );
  as_xlsx.cell( 2, 3, 'Beispiel', p_style => 'header' );
  as_xlsx.cell( 1, 4,  'text' );       as_xlsx.cell( 2, 4,  'Muster AG', p_style => 'text' );
  as_xlsx.cell( 1, 5,  'ganzzahl' );   as_xlsx.cell( 2, 5,  12345, p_style => 'ganzzahl' );
  as_xlsx.cell( 1, 6,  'zahl' );       as_xlsx.cell( 2, 6,  12345.678, p_style => 'zahl' );
  as_xlsx.cell( 1, 7,  'betrag' );     as_xlsx.cell( 2, 7,  1234.5, p_style => 'betrag' );
  as_xlsx.cell( 1, 8,  'prozent' );    as_xlsx.cell( 2, 8,  0.075, p_style => 'prozent' );
  as_xlsx.cell( 1, 9,  'datum' );      as_xlsx.cell( 2, 9,  date '2026-11-05', p_style => 'datum' );
  as_xlsx.cell( 1, 10, 'datum_zeit' ); as_xlsx.cell( 2, 10, to_date( '05.11.2026 13:30', 'dd.mm.yyyy hh24:mi' ), p_style => 'datum_zeit' );
  as_xlsx.cell( 1, 11, 'summe' );      as_xlsx.num_formula( 2, 11, 'SUM(B5:B7)', 12345 + 12345.678 + 1234.5, p_style => 'summe', p_numFmtId => as_xlsx.get_numFmt( '#,##0.00' ) );
  as_xlsx.cell( 1, 12, 'hinweis' );    as_xlsx.cell( 2, 12, 'Fussnote', p_style => 'hinweis' );
  as_xlsx.cell( 1, 13, 'warnung' );    as_xlsx.cell( 2, 13, 'Achtung', p_style => 'warnung' );
  as_xlsx.cell( 1, 14, 'hyperlink' );  as_xlsx.hyperlink( 2, 14, 'https://github.com/acreonity/as_xlsx', 'as_xlsx auf GitHub' );
  as_xlsx.set_column_width( 1, 14 );
  as_xlsx.set_column_width( 2, 24 );
  --
  -- sheet 3: plain query2sheet, styled, total over all numeric columns, new sheet
  as_xlsx.query2sheet( 'select level as nr, level * 1.5 as wert, -level * 1000 as saldo from dual connect by level <= 5'
                     , p_styled => true, p_sum_columns => 'ALL' );
  as_xlsx.save( 'OUT_DIR', 'demo_styles.xlsx' );
  dbms_output.put_line( 'rows: ' || l_cnt );
end;
/
