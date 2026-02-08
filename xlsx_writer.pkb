set define off

create or replace package body xlsx_writer as -- {{{
-- vi: foldmarker={{{,}}}

  -- NEU_1: Globaler Buffer für LOB-Schreibvorgänge (Performance)
  g_buffer varchar2(32767);
  
  -- NEU_2: Cache für Spaltenbuchstaben (Performance)
  type t_col_cache is table of varchar2(3) index by pls_integer;
  g_col_cache t_col_cache;

  -- NEU_1: Hilfsprozedur zum Leeren des Buffers in das BLOB
  procedure flush_data(b in out nocopy blob) is
  begin
    if g_buffer is not null then
       dbms_lob.append(b, utl_raw.cast_to_raw(g_buffer));
       g_buffer := null;
    end if;
  end flush_data;

  procedure ap (b in out nocopy blob, v in varchar2) is -- {{{
  begin
    -- NEU_1: Buffer-Logik implementiert
    -- Sammelt Strings bis 32k, schreibt dann erst ins BLOB (reduziert I/O drastisch)
    if length(g_buffer) + length(v) > 32000 then
       dbms_lob.append(b, utl_raw.cast_to_raw(g_buffer));
       g_buffer := v;
    else
       g_buffer := g_buffer || v;
    end if;
  end ap; -- }}}

  function start_xml_blob return blob is -- {{{
    ret blob;
  begin
    dbms_lob.createtemporary(ret, true);
    
    ap(ret,q'{<?xml version="1.0" encoding="utf-8"?>
}');
    
    -- NEU_1: Buffer leeren, bevor das BLOB zurückgegeben wird
    flush_data(ret);
    return ret;

  end start_xml_blob; -- }}}

  procedure add_attr(b          in out nocopy blob, -- {{{
                     attr_name         varchar2,
                     attr_value        varchar2) is
  begin
    ap(b, ' ' || attr_name || '="' || attr_value || '"');
  end add_attr; -- }}}

  procedure warning(text varchar2) is -- {{{
  begin
    dbms_output.put_line('! warnining: ' || text);
  end warning; -- }}}

  function start_book return book_r is -- {{{
    ret book_r;
  begin

    ret.sheets                   := new sheet_t           ();
    ret.cell_styles              := new cell_style_t      ();
    ret.borders                  := new border_t          ();
    ret.fonts                    := new font_t            ();
    ret.fills                    := new fill_t            ();
    ret.num_fmts                 := new num_fmt_t         ();
    ret.shared_strings           := new shared_string_t   ();
    ret.medias                   := new media_t           ();
    ret.calc_chain_elems         := new calc_chain_elem_t ();
    ret.drawings                 := new drawing_t         ();
    ret.content_type_vmlDrawing  := false;

    return ret;

  end start_book; -- }}}

  function  add_sheet(xlsx in out nocopy book_r, -- {{{
                      name_       varchar2) return integer is

    ret sheet_r;
  begin

 -- sheetname must not contain any of : [ ]
    ret.name_      := translate(name_, ':[]', '   ');

    ret.col_widths := new col_width_t();
    ret.sheet_rels := new sheet_rel_t();
    ret.validations := new validation_t(); -- NEU_7: Initialisierung für Dropdowns

    xlsx.sheets.extend;
    xlsx.sheets(xlsx.sheets.count) := ret;

    return xlsx.sheets.count;

  end add_sheet; -- }}}

  procedure freeze_sheet      (xlsx        in out nocopy book_r, -- {{{
                               sheet       in     integer,
                               split_x     in     integer := null,
                               split_y     in     integer := null) is
  begin

    xlsx.sheets(sheet).split_x := split_x;
    xlsx.sheets(sheet).split_y := split_y;

  end freeze_sheet; -- }}}

  procedure add_sheet_rel     (xlsx        in out nocopy book_r, -- {{{
                               sheet              integer,
                               raw_               varchar2)
  is
  begin

    xlsx.sheets(sheet).sheet_rels.extend;
    xlsx.sheets(sheet).sheet_rels(xlsx.sheets(sheet).sheet_rels.count). raw_ := raw_;

  end add_sheet_rel; -- }}}

  procedure add_media         (xlsx        in out nocopy book_r, -- {{{
                               b                  blob,
                               name_              varchar2) is
  begin
  
    xlsx.medias.extend;
    xlsx.medias(xlsx.medias.count).b     := b;
    xlsx.medias(xlsx.medias.count).name_ := name_;

  end add_media; -- }}}

  procedure add_drawing       (xlsx        in out nocopy book_r, -- {{{
                               raw_               varchar2) is
  begin

    xlsx.drawings.extend;
    xlsx.drawings(xlsx.drawings.count).raw_ := raw_;
  end add_drawing; -- }}}

  procedure add_checkbox      (xlsx        in out nocopy book_r, -- {{{
                               sheet              integer,
                               col_left           integer,
                               row_top            integer,
                               text               varchar2 := null,
                               checked            boolean  := false) is
  begin

    if xlsx.sheets(sheet).vml_drawings is null then
       xlsx.sheets(sheet).vml_drawings := new vml_drawing_t();
       xlsx.sheets(sheet).vml_drawings.extend;
    -- assumption assmpt_01: at most one vml drawing per sheet.
       xlsx.sheets(sheet).vml_drawings(1).checkboxes := new checkbox_t();
    end if;

    xlsx.sheets(sheet).vml_drawings(1).checkboxes.extend;
    xlsx.sheets(sheet).vml_drawings(1).checkboxes(
    xlsx.sheets(sheet).vml_drawings(1).checkboxes.count).col_left := col_left;

    xlsx.sheets(sheet).vml_drawings(1).checkboxes(
    xlsx.sheets(sheet).vml_drawings(1).checkboxes.count).row_top  := row_top;

    xlsx.sheets(sheet).vml_drawings(1).checkboxes(
    xlsx.sheets(sheet).vml_drawings(1).checkboxes.count).text     := text;

    xlsx.sheets(sheet).vml_drawings(1).checkboxes(
    xlsx.sheets(sheet).vml_drawings(1).checkboxes.count).checked  := checked;

    xlsx.content_type_vmlDrawing := true;

  end add_checkbox; -- }}}

  -- {{{ rows and columns

  function col_to_letter(c integer) return varchar2 is -- {{{
    -- NEU_2: Variable für Cache-Ergebnis
    l_res varchar2(10);
  begin
    -- NEU_2: Cache prüfen
    if g_col_cache.exists(c) then
       return g_col_cache(c);
    end if;

    if c < 27 then
       l_res := substr('ABCDEFGHIJKLMNOPQRSTUVWXYZ', c, 1);
    else
       l_res := col_to_letter(trunc((c-1)/26)) || col_to_letter(mod((c-1), 26)+1);
    end if;
    
    -- NEU_2: Ergebnis in Cache speichern
    g_col_cache(c) := l_res;
    return l_res;

  end col_to_letter; -- }}}

  procedure col_width         (xlsx    in out nocopy book_r,-- {{{
                               sheet          integer,
                               col            integer,
                               width          number
                               ) is
  begin
    col_width(xlsx, sheet, col, col, width);
  end col_width; -- }}}

  procedure col_width         (xlsx    in out nocopy book_r,-- {{{
                               sheet          integer,
                               start_col      integer,
                               end_col        integer,
                               width          number
                               ) is
    r col_width_r;
  begin

    r.start_col := start_col;
    r.end_col   := end_col;
    r.width     := width;

    xlsx.sheets(sheet).col_widths.extend;
    xlsx.sheets(sheet).col_widths(xlsx.sheets(sheet).col_widths.count) := r;

  end col_width; -- }}}

  function does_row_exist(xlsx  in out nocopy book_r, -- {{{
                          sheet        integer,
                          r            integer) return boolean is
  begin

     if xlsx.sheets(sheet).rows_.exists(r) then
        return true;
     end if;

     return false;

  end does_row_exist; -- }}}

  procedure add_row(xlsx     in out nocopy book_r, -- {{{
                    sheet           integer,
                    r               integer,
                    height          number := null) is
  begin

      if does_row_exist(xlsx, sheet, r) then
         raise_application_error(-20800, 'row ' || r || ' already exists');
      end if;

      xlsx.sheets(sheet).rows_(r).height := height;

  end add_row; -- }}}

  -- NEU_5: Interne Prozedur für schnelle Massendaten (z.B. aus SQL)
  -- Überspringt Checks auf Existenz der Zeile/Zelle für bessere Performance
  procedure add_cell_fast     (xlsx    in out nocopy book_r, 
                               sheet         integer,
                               r             integer,
                               c             integer,
                               style_id      integer  := 0,
                               text          varchar2 := null,
                               value_        number   := null,
                               formula       varchar2 := null,
                               date_         date     := null) is
  begin
     xlsx.sheets(sheet).rows_(r).cells(c).style_id := style_id;
     
     if date_ is not null then
        xlsx.sheets(sheet).rows_(r).cells(c).value_ := date_ - date '1899-12-30';
     else
        xlsx.sheets(sheet).rows_(r).cells(c).value_ := value_;
     end if;
     
     xlsx.sheets(sheet).rows_(r).cells(c).formula  := formula;
     
     if formula is not null then
        xlsx.calc_chain_elems.extend;
        xlsx.calc_chain_elems(xlsx.calc_chain_elems.count).cell_reference := col_to_letter(c) || r;
        xlsx.calc_chain_elems(xlsx.calc_chain_elems.count).sheet          := sheet;
     end if;

     if text is not null then
        xlsx.shared_strings.extend;
        xlsx.shared_strings(xlsx.shared_strings.count).val := replace(replace(replace(text, '&', '&amp;'), '>', '&gt;' ), '<', '&lt;' );
        xlsx.sheets(sheet).rows_(r).cells(c).shared_string_id := xlsx.shared_strings.count-1;
     end if;
  end add_cell_fast;

  procedure add_cell          (xlsx    in out nocopy book_r, -- {{{
                               sheet         integer,
                               r             integer,
                               c             integer,
                               style_id      integer  :=    0,
                               text          varchar2 := null,
                               value_        number   := null,
                               formula       varchar2 := null) is
  begin

    if not does_row_exist(xlsx, sheet, r) then
       add_row(xlsx, sheet, r);
    end if;

    if xlsx.sheets(sheet).rows_(r).cells.exists(c) then
       warning('Cell ' || c || ' in row ' || r || ' already exists.');
    end if;

    if style_id is null then
       raise_application_error(-20800, 'style id is null for cell ' || r || '/' || c);
    end if;

    xlsx.sheets(sheet).rows_(r).cells(c).style_id := style_id;
    xlsx.sheets(sheet).rows_(r).cells(c).value_   := value_;
    xlsx.sheets(sheet).rows_(r).cells(c).formula  := formula;

    if formula is not null then -- {{{

       xlsx.calc_chain_elems.extend;
       xlsx.calc_chain_elems(xlsx.calc_chain_elems.count).cell_reference := col_to_letter(c) || r;
       xlsx.calc_chain_elems(xlsx.calc_chain_elems.count).sheet          := sheet;

    end if; -- }}}

    if text is not null then -- {{{
       xlsx.shared_strings.extend;
       xlsx.shared_strings(xlsx.shared_strings.count).val := replace(
                                                             replace(
                                                             replace(text, '&', '&amp;'),
                                                                           '>', '&gt;' ),
                                                                           '<', '&lt;' );
 
       xlsx.sheets(sheet).rows_(r).cells(c).shared_string_id := xlsx.shared_strings.count-1;
    end if; -- }}}

  end add_cell; -- }}}

  procedure add_cell          (xlsx        in out nocopy book_r, -- {{{
                               sheet              integer,
                               r                  integer,
                               c                  integer,
                               date_              date,
                               style_id           integer  :=    0) is
  begin

    add_cell(xlsx         => xlsx,
             sheet        => sheet,
             r            => r,
             c            => c,
             style_id     => style_id,
             value_       => date_ - date '1899-12-30');

  end add_cell; -- }}}

  -- }}}

  -- {{{ related to styles
  --
  function add_font         (xlsx     in out nocopy book_r, -- {{{
                             name            varchar2,
                             size_           number,
                             color           varchar2 := null,
                             b               boolean  := false,
                             i               boolean  := false,
                             u               boolean  := false) return integer is
  begin

      xlsx.fonts.extend;
      xlsx.fonts(xlsx.fonts.count).name  := name; 
      xlsx.fonts(xlsx.fonts.count).size_ := size_; 
      xlsx.fonts(xlsx.fonts.count).color := color; 
      xlsx.fonts(xlsx.fonts.count).b     := b; 
      xlsx.fonts(xlsx.fonts.count).i     := i; 
      xlsx.fonts(xlsx.fonts.count).u     := u; 

      return xlsx.fonts.count; 

  end add_font; -- }}}

  function add_fill        (xlsx     in out nocopy book_r, -- {{{
                            raw_            varchar2) return integer is

  begin
    xlsx.fills.extend;
    xlsx.fills(xlsx.fills.count).raw_ := raw_;
    return xlsx.fills.count + 1; 
  end add_fill; -- }}}

  function add_border      (xlsx     in out nocopy book_r, -- {{{
                            raw_            varchar2) return integer is
  begin
    xlsx.borders.extend;
    xlsx.borders(xlsx.borders.count).raw_ := raw_;
    return xlsx.borders.count; 
  end add_border; -- }}}

  function add_cell_style  (xlsx         in out nocopy book_r, -- {{{
                            font_id             integer  := 0,
                            fill_id             integer  := 0,
                            border_id           integer  := 0,
                            num_fmt_id          integer  := 0,
                            vertical_alignment  varchar2 := null,
                            wrap_text           boolean  := null
                          ) return integer is

    rec cell_style_r;
  begin

    rec.font_id            := font_id;
    rec.fill_id            := fill_id;
    rec.border_id          := border_id;
    rec.num_fmt_id         := num_fmt_id;
    rec.vertical_alignment := vertical_alignment;
    rec.wrap_text          := wrap_text;

    xlsx.cell_styles.extend;
    xlsx.cell_styles(xlsx.cell_styles.count) := rec;

    return xlsx.cell_styles.count; 

  end add_cell_style; -- }}}

  function add_num_fmt     (xlsx     in out nocopy book_r, -- {{{
                            raw_            varchar2,
                            return_id       integer) return integer is
  begin
    xlsx.num_fmts.extend;
    xlsx.num_fmts(xlsx.num_fmts.count).raw_ := raw_;
    return return_id;
  end add_num_fmt; -- }}}

  -- }}}

  -- {{{ blob generators

  function docProps_app return blob is -- {{{
    ret blob;
  begin

    ret := start_xml_blob;
    ap(ret, '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">');
    ap(ret, '</Properties>');
    
    -- NEU_1: Buffer leeren
    flush_data(ret);
    return ret;

  end docProps_app; -- }}}

  function xl_worksheets_sheet(xlsx in out nocopy book_r, -- {{{
                              sheet        integer) return blob is

    ret blob;
  begin

    ret := start_xml_blob;

    ap(ret, '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">');


    if xlsx.sheets(sheet).split_x is not null or -- {{{
       xlsx.sheets(sheet).split_y is not null then

       ap(ret, '<sheetViews>');        
       ap(ret, '<sheetView workbookViewId="0">');        

       ap(ret, '<pane');

       if xlsx.sheets(sheet).split_x is not null then
          add_attr(ret, 'xSplit', xlsx.sheets(sheet).split_x);
       end if; 

       if xlsx.sheets(sheet).split_y is not null then
          add_attr(ret, 'ySplit', xlsx.sheets(sheet).split_y);
       end if; 

       add_attr(ret, 'topLeftCell', col_to_letter(nvl(xlsx.sheets(sheet).split_x, 1) + 1) || (nvl(xlsx.sheets(sheet).split_y, 1) + 1));
       ap(ret, ' state="frozen"/>');

       if xlsx.sheets(sheet).split_y is not null then
          ap(ret, '<selection pane="bottomLeft" activeCell="B6" sqref="B6" />');
       else
          ap(ret, '<selection pane="topRight" activeCell="B6" sqref="B6" />');
       end if;

       ap(ret, '</sheetView>');        
       ap(ret, '</sheetViews>');        

    end if; -- }}}

    if xlsx.sheets(sheet).col_widths.count > 0 then -- {{{
      ap(ret, '<cols>');

      for i in 1 .. xlsx.sheets(sheet).col_widths.count loop -- {{{

        ap(ret, '<col');

        add_attr(ret, 'min'        , xlsx.sheets(sheet).col_widths(i).start_col);
        add_attr(ret, 'max'        , xlsx.sheets(sheet).col_widths(i).end_col  );
        add_attr(ret, 'width'      , xlsx.sheets(sheet).col_widths(i).width    );
        add_attr(ret, 'customWidth', 1);

        ap (ret, '/>');

      end loop; -- }}}

      ap(ret, '</cols>');
    end if; -- }}}

    ap(ret, '<sheetData>'); -- {{{

    declare
      r pls_integer;
      c pls_integer;
    begin
      r := xlsx.sheets(sheet).rows_.first;
      while r is not null loop -- {{{
 
        ap(ret, '<row');
        add_attr(ret, 'r', r);

        if xlsx.sheets(sheet).rows_(r).height is not null then
           add_attr(ret, 'ht', xlsx.sheets(sheet).rows_(r).height);
           add_attr(ret, 'customHeight', 1);
        end if;

        ap(ret, '>');

        c := xlsx.sheets(sheet).rows_(r).cells.first;
        while c is not null loop -- {{{

          ap(ret, '<c');

          -- NEU_3_FIX: Das Attribut 'r' MUSS zwingend gesetzt werden (z.B. r="A1").
          -- Wenn dies fehlt, interpretiert Excel dies als Sparse Matrix und schiebt Daten nach links, wenn Zellen leer sind. Kann aber abhängig vom Kontext, wenn es entfernt werden kann Performanz stark verbessern.
          add_attr(ret, 'r', col_to_letter(c) || r);
          
          add_attr(ret, 's', xlsx.sheets(sheet).rows_(r).cells(c).style_id);

          if xlsx.sheets(sheet).rows_(r).cells(c).shared_string_id is not null then
             add_attr(ret, 't', 's'); -- type is string
          end if;

          ap(ret, '>');

          if xlsx.sheets(sheet).rows_(r).cells(c).formula is not null then
             -- NEU_4: Formeln werden escaped, damit Zeichen wie < oder & nicht das XML zerstören.
             ap(ret, '<f>' || dbms_xmlgen.convert(xlsx.sheets(sheet).rows_(r).cells(c).formula, dbms_xmlgen.entity_encode) || '</f>');
          end if;

          if xlsx.sheets(sheet).rows_(r).cells(c).value_ is not null then
             -- NEU_9: KRITISCH! Explizite Umwandlung auf Punkt-Separator nur hier im XML.
             -- Verhindert globale Änderung der Session durch set_nls.
             ap(ret, '<v>' || to_char(xlsx.sheets(sheet).rows_(r).cells(c).value_, 'TM9', 'NLS_NUMERIC_CHARACTERS=''.,''') || '</v>');
          end if;

          if xlsx.sheets(sheet).rows_(r).cells(c).shared_string_id is not null then
             ap(ret, '<v>' || xlsx.sheets(sheet).rows_(r).cells(c).shared_string_id || '</v>');
          end if;

          ap(ret, '</c>');

          c := xlsx.sheets(sheet).rows_(r).cells.next(c);

        end loop; -- }}}
        
        r := xlsx.sheets(sheet).rows_.next(r);
 
        ap(ret, '</row>');
      end loop; -- }}}
    end;

    ap(ret, '</sheetData>'); -- }}}

  -- {{{ data validations xml generation (NEU_7)
    if xlsx.sheets(sheet).validations is not null and xlsx.sheets(sheet).validations.count > 0 then
       ap(ret, '<dataValidations count="' || xlsx.sheets(sheet).validations.count || '">');
       
       for v in 1 .. xlsx.sheets(sheet).validations.count loop
           ap(ret, '<dataValidation type="' || xlsx.sheets(sheet).validations(v).type_ || '"' ||
                   ' allowBlank="1" showInputMessage="' || case when xlsx.sheets(sheet).validations(v).show_input_msg then '1' else '0' end || '"' ||
                   ' showErrorMessage="' || case when xlsx.sheets(sheet).validations(v).show_error_msg then '1' else '0' end || '"' ||
                   ' sqref="' || xlsx.sheets(sheet).validations(v).sqref || '">');
           
           if xlsx.sheets(sheet).validations(v).formula1 is not null then
               ap(ret, '<formula1>' || xlsx.sheets(sheet).validations(v).formula1 || '</formula1>');
           end if;
           
           ap(ret, '</dataValidation>');
       end loop;
       
       ap(ret, '</dataValidations>');
    end if;
    -- }}}

    ap(ret, '<pageMargins left="0.36000000000000004" right="0.2" top="1" bottom="1" header="0.5" footer="0.5" />');
    ap(ret, '<pageSetup paperSize="9" scale="67" orientation="landscape" horizontalDpi="4294967292" verticalDpi="4294967292" />');
  
    for d in 1 .. xlsx.drawings.count loop 
      ap(ret, '<drawing r:id="rId' || d || '" /> ');
    end loop;

    if xlsx.sheets(sheet).vml_drawings is not null then
       ap(ret, '<legacyDrawing r:id="rel_vml_drawing_' || sheet || '" />');
    end if;

    ap(ret, '</worksheet>');

    -- NEU_1: Buffer leeren
    flush_data(ret);
    return ret;

  end xl_worksheets_sheet; -- }}}

  function xl_styles(xlsx in book_r) return blob is -- {{{
    ret blob;
    tag_alignment boolean := false;
  begin
    ret := start_xml_blob;

    ap(ret, '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">');

    if xlsx.num_fmts.count > 0 then -- {{{
       ap(ret, '<numFmts>');
       for n in 1 .. xlsx.num_fmts.count loop
           ap(ret, '<numFmt ' || xlsx.num_fmts(n).raw_ || ' />');
       end loop;
       ap(ret, '</numFmts>');
    end if; -- }}}

    ap(ret, '<fonts>'); -- {{{
    ap(ret, '<font><sz val="11"/><name val="Calibri" /></font>'); -- default font.
    for f in 1 .. xlsx.fonts.count loop -- {{{
      ap(ret, '<font><name val="' || xlsx.fonts(f).name  || '"/>' ||
                    '  <sz val="' || xlsx.fonts(f).size_ || '"/>');
      if xlsx.fonts(f).color is not null then
         ap(ret, '<color ' || xlsx.fonts(f).color || '/>');
      end if;
      if xlsx.fonts(f).i then
         ap(ret, '<i/>');
      end if;
      if xlsx.fonts(f).b then
         ap(ret, '<b/>');
      end if;
      if xlsx.fonts(f).u then
         ap(ret, '<u/>');
      end if;
      ap(ret, '</font>');
    end loop; -- }}}
    ap(ret, '</fonts>'); -- }}}

    ap(ret, '<fills>'); -- {{{
    ap(ret, q'{
      <fill><patternFill patternType="none"    /></fill>
      <fill><patternFill patternType="gray125" /></fill>
    }');
    for f in 1 .. xlsx.fills.count loop -- {{{
      ap(ret, '<fill>' || xlsx.fills(f).raw_ || '</fill>');
    end loop; -- }}}
    ap(ret, '</fills>'); -- }}}

    ap(ret, '<borders>'); -- {{{
    ap(ret, '<border><left/><right/><top/><bottom/><diagonal/></border>');
    for b in 1 .. xlsx.borders.count loop -- {{{
        ap(ret, '<border>' || xlsx.borders(b).raw_ || '</border>');
    end loop; -- }}}
    ap(ret, '</borders>'); -- }}}

    -- <cellStyleXfs>
    ap(ret, q'{<cellStyleXfs>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
    </cellStyleXfs>}'); 

    ap(ret, '<cellXfs>'); -- {{{ cell styles
    ap(ret, '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />');

    for c in 1 .. xlsx.cell_styles.count loop -- {{{
        ap(ret, '<xf');
        add_attr(ret, 'numFmtId', xlsx.cell_styles(c).num_fmt_id);
        add_attr(ret, 'fillId'  , xlsx.cell_styles(c).fill_id   );
        add_attr(ret, 'fontId'  , xlsx.cell_styles(c).font_id   );
        add_attr(ret, 'borderId', xlsx.cell_styles(c).border_id );
        ap(ret, '>');
        if xlsx.cell_styles(c).vertical_alignment is not null then
           tag_alignment := true;
           ap(ret, '<alignment vertical="' || xlsx.cell_styles(c).vertical_alignment || '"');
        end if;
        if xlsx.cell_styles(c).wrap_text is not null then
           if not tag_alignment then
              tag_alignment := true;
              ap(ret, '<alignment');
            end if;
            ap(ret, ' wrapText="' || case when xlsx.cell_styles(c).wrap_text then '1' else '0' end || '"');
        end if;
        if tag_alignment then
           ap(ret, '/>');
        end if;
        ap(ret, '</xf>');
    end loop; -- }}}

  ap(ret, '</cellXfs>'); -- }}}
    ap(ret, q'{<cellStyles count="33">
      <cellStyle name="Normal" xfId="0" builtinId="0" />
   </cellStyles>}');
    ap(ret, q'{<dxfs count="0" />}');
    ap(ret, q'{</styleSheet>}');
    
    -- NEU_1: Buffer leeren
    flush_data(ret);
    return ret;
  end xl_styles; -- }}}

  function xl_sharedStrings(xlsx book_r) return blob is -- {{{
    ret blob;
  begin
    ret := start_xml_blob();
    ap(ret, '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="24" uniqueCount="24">');
    for s in 1 .. xlsx.shared_strings.count loop
      ap(ret, '<si><t xml:space="preserve">' || xlsx.shared_strings(s).val || '</t></si>');
    end loop;
    ap(ret, '</sst>');
    
    -- NEU_1: Buffer leeren
    flush_data(ret);
    return ret;
  end xl_sharedStrings; -- }}}

  function xl_workbook(xlsx in out nocopy book_r -- {{{
  ) return blob is
    ret blob := start_xml_blob;
  begin
    ap(ret, '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'); -- {{{

--  NEU_4: rupBuild="14420" aktualisiert, um Dynamic Array Kompatibilität (Excel 2016+) zu gewährleisten.
    ap(ret, '  <fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="14420" />'); 

    ap(ret, '<sheets>'); -- {{{
    for s in 1 .. xlsx.sheets.count loop -- {{{
        ap(ret, '<sheet');
        add_attr(ret, 'name'   , xlsx.sheets(s).name_);
        add_attr(ret, 'sheetId', s                   );
        add_attr(ret, 'r:id'   ,'rId' || s           );
        ap(ret, '/>');
     end loop; -- }}}
    ap(ret, '</sheets>'); -- }}}
    ap(ret, '<calcPr calcId="145621" />'); 
    ap(ret, '</workbook>'); -- }}}

    -- NEU_1: Buffer leeren
    flush_data(ret);
    return ret;
  end xl_workbook; -- }}}

  function rels_rels      (xlsx in out nocopy book_r) return blob is -- {{{
    ret blob;
  begin
    ret := start_xml_blob();
    ap(ret, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
      ap(ret, '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"  />');
      ap(ret, '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"   Target="docProps/core.xml" />');
      ap(ret, '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"      Target="xl/workbook.xml"   />');
    ap(ret, '</Relationships>');

    -- NEU_1: Buffer leeren
    flush_data(ret);
    return ret;
  end rels_rels; -- }}}

  function docProps_core  (xlsx in out nocopy book_r) return blob is -- {{{
    ret blob := start_xml_blob;
  begin
    ap(ret, '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">');
    ap(ret, '</cp:coreProperties>');

    -- NEU_1: Buffer leeren
    flush_data(ret);
    return ret;
  end docProps_core; -- }}}

  function xl_rels_workbook           (xlsx in out nocopy book_r) return blob is -- {{{
    ret blob := start_xml_blob;
    rId integer := 1;
    procedure add_relationship(type_ varchar2, target varchar2) is -- {{{
    begin
      ap(ret, '<Relationship');
      add_attr(ret, 'Id'    , 'rId' || rId); rId := rId + 1;
      add_attr(ret, 'Type'  ,  type_      );
      add_attr(ret, 'Target',  target     );
      ap(ret, '/>');
    end add_relationship; -- }}}
  begin
    ap(ret, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
      for s in 1 .. xlsx.sheets.count loop
        add_relationship('http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', 'worksheets/sheet' || s || '.xml');
      end loop;
      add_relationship('http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings', 'sharedStrings.xml');
      add_relationship('http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'       , 'styles.xml'       );
    ap(ret, '</Relationships>');

    -- NEU_1: Buffer leeren
    flush_data(ret);
    return ret;
  end xl_rels_workbook; -- }}}

  function xl_drawings_rels_drawing1  (xlsx in out nocopy book_r) return blob is -- {{{
    ret blob := start_xml_blob;
  begin
    ap(ret, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
    for m in 1 .. xlsx.medias.count loop
      ap(ret, '<Relationship Id="rId' || m || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' || xlsx.medias(m).name_ || '" />');
    end loop;
    ap(ret, '</Relationships>');

    -- NEU_1: Buffer leeren
    flush_data(ret);
    return ret;
  end xl_drawings_rels_drawing1; -- }}}

  function xl_worksheets_rels_sheet(xlsx in out nocopy book_r, sheet integer)  return blob is -- {{{
    ret blob;
  begin
    for r in 1 .. xlsx.sheets(sheet).sheet_rels.count loop
      if ret is null then
         ret := start_xml_blob;
         ap(ret, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
      end if;
      ap(ret, xlsx.sheets(sheet).sheet_rels(r).raw_);
    end loop;

    if xlsx.sheets(sheet).vml_drawings is not null then
       if ret is null then
          ret := start_xml_blob;
          ap(ret, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
       end if;
       ap(ret, '<Relationship Id="rel_vml_drawing_' || sheet ||  '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing' || sheet || '.vml" />');
    end if;

    if ret is not null then
       ap(ret, '</Relationships>');
       -- NEU_1: Buffer leeren
       flush_data(ret);
    end if;
    return ret;
  end xl_worksheets_rels_sheet; -- }}}

  function content_types(xlsx in out nocopy book_r) return blob is -- {{{
    ret blob := start_xml_blob;
  begin
    ap(ret, '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">');
      ap(ret, '<Default Extension="png"  ContentType="image/png"                                                />');
      ap(ret, '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />');
      if xlsx.content_type_vmlDrawing then
      ap(ret, '<Default Extension="vml"  ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing" />'); 
      end if;
      ap(ret, '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />');
      for s in 1 ..xlsx.sheets.count loop
        ap(ret, '<Override PartName="/xl/worksheets/sheet' || s || '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />');
      end loop;
      ap(ret, '<Override PartName="/xl/styles.xml"            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"        />');
      ap(ret, '<Override PartName="/xl/sharedStrings.xml"     ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />');
      ap(ret, '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"                     />');
      ap(ret, '<Override PartName="/xl/calcChain.xml"         ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"     />');
      ap(ret, '<Override PartName="/docProps/core.xml"        ContentType="application/vnd.openxmlformats-package.core-properties+xml"                    />');
      ap(ret, '<Override PartName="/docProps/app.xml"         ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"         />');
    ap(ret, '</Types>');

    -- NEU_1: Buffer leeren
    flush_data(ret);
    return ret;
  end content_types; -- }}}

  function xl_calcChain(xlsx in out nocopy book_r) return blob is -- {{{
    ret blob := start_xml_blob;
  begin
    ap(ret, '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
    for e in 1 .. xlsx.calc_chain_elems.count loop
        ap(ret, '<c');
        add_attr(ret, 'r', xlsx.calc_chain_elems(e).cell_reference);
        add_attr(ret, 'i', xlsx.calc_chain_elems(e).sheet         );
        ap(ret, '/>');
    end loop;
    ap(ret, '</calcChain>');
    
    -- NEU_1: Buffer leeren
    flush_data(ret);
    return ret;
  end xl_calcChain; -- }}}

  function vml_drawing(v vml_drawing_r) return blob is -- {{{
    ret blob := start_xml_blob;
  begin
    ap(ret, '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">');
    for cb in 1 .. v.checkboxes.count loop
      ap(ret, '<v:shape
         type="#_x0000_t201" 
         filled="f"
         fillcolor="window [65]"
         stroked="f" 
         strokecolor="windowText [64]"
         o:insetmode="auto"
       >');
      ap(ret, '<v:textbox style="mso-direction-alt:auto" o:singleclick="f">
      <div style="text-align:left">
        <font face="Tahoma" size="160" color="auto">' || v.checkboxes(cb).text || '</font>
      </div>
    </v:textbox>');
    ap(ret, '<x:ClientData ObjectType="Checkbox">
      <x:SizeWithCells />
      <x:Anchor>' ||
        (v.checkboxes(cb).col_left -1) || ',' || 
        '0,'                                  || 
        (v.checkboxes(cb).row_top  -1) || ',' || 
        '0,'                                  || 
        (v.checkboxes(cb).col_left -1) || ',' || 
        '103,'                                || 
        (v.checkboxes(cb).row_top  -1) || ',' || 
        '17                                   || 
      </x:Anchor> 
      <x:AutoFill>False</x:AutoFill>
      <x:AutoLine>False</x:AutoLine>
      <x:TextVAlign>Center</x:TextVAlign>
      <x:Checked>' || case when v.checkboxes(cb).checked then 1 else 0 end || '</x:Checked>
      <x:NoThreeD />
    </x:ClientData>');
      ap(ret, '</v:shape>');
    end loop;
    ap(ret, '</xml>');
    
    -- NEU_1: Buffer leeren
    flush_data(ret);
    return ret;
  end vml_drawing; -- }}}

  -- {{{ implementation add_data_validation (NEU_7)
  procedure add_data_validation(
      xlsx            in out nocopy book_r,
      sheet           integer,
      sqref           varchar2,
      formula1        varchar2,
      type_           varchar2 := 'list',
      show_input_msg  boolean := true,
      show_error_msg  boolean := true
  ) is
      rec validation_r;
  begin
      rec.type_           := type_;
      rec.formula1        := formula1;
      rec.sqref           := sqref;
      rec.show_input_msg  := show_input_msg;
      rec.show_error_msg  := show_error_msg;
      
      if xlsx.sheets(sheet).validations is null then
         xlsx.sheets(sheet).validations := new validation_t();
      end if;

      xlsx.sheets(sheet).validations.extend;
      xlsx.sheets(sheet).validations(xlsx.sheets(sheet).validations.count) := rec;
  end add_data_validation;
  -- }}}

  function create_xlsx(xlsx in out nocopy book_r) return blob is -- {{{
    xlsx_b   blob;
    xb_xl_worksheets_rels_sheet   blob;

    procedure add_blob_to_zip(zip      in out nocopy blob, -- {{{
                              filename        varchar2,
                              b               blob) is
      b_ blob;
    begin
      b_ := b;
      zipper.addFile(zip, filename, b_);
      dbms_lob.freetemporary(b_);
    end add_blob_to_zip; -- }}}

  begin
    dbms_lob.createtemporary(xlsx_b, true);

    for s in 1 .. xlsx.sheets.count loop
       xb_xl_worksheets_rels_sheet  := xl_worksheets_rels_sheet  (xlsx, s);
       if xb_xl_worksheets_rels_sheet is not null then
          zipper.addFile(xlsx_b, 'xl/worksheets/_rels/sheet' || s || '.xml.rels', xb_xl_worksheets_rels_sheet);
       end if;
    end loop;

    for s in 1 .. xlsx.sheets.count loop
        zipper.addFile(xlsx_b, 'xl/worksheets/sheet' || s || '.xml', xl_worksheets_sheet       (xlsx, s));
        if xlsx.sheets(s).vml_drawings is not null then
           add_blob_to_zip(xlsx_b, 'xl/drawings/vmlDrawing' || s || '.vml', vml_drawing(xlsx.sheets(s).vml_drawings(1)));
        end if;
    end loop;

    add_blob_to_zip(xlsx_b, '_rels/.rels'                        , rels_rels                 (xlsx));
    add_blob_to_zip(xlsx_b, 'docProps/app.xml'                   , docProps_app              ()    );
    add_blob_to_zip(xlsx_b, 'docProps/core.xml'                  , docProps_core             (xlsx));
    add_blob_to_zip(xlsx_b, 'xl/_rels/workbook.xml.rels'         , xl_rels_workbook          (xlsx));

    if xlsx.calc_chain_elems.count > 0 then
       add_blob_to_zip(xlsx_b, 'xl/calcChain.xml'                , xl_calcChain              (xlsx));
    end if;

    add_blob_to_zip(xlsx_b, 'xl/drawings/_rels/drawing1.xml.rels', xl_drawings_rels_drawing1 (xlsx));

    for d in 1 .. xlsx.drawings.count loop
      add_blob_to_zip(xlsx_b, 'xl/drawings/drawing' || d || '.xml',  utl_raw.cast_to_raw(xlsx.drawings(d).raw_));
    end loop;

    for m in 1 .. xlsx.medias.count loop
      add_blob_to_zip(xlsx_b, 'xl/media/' || xlsx.medias(m).name_, xlsx.medias(m).b);
    end loop;

    add_blob_to_zip(xlsx_b, 'xl/sharedStrings.xml'               , xl_sharedStrings          (xlsx));
    add_blob_to_zip(xlsx_b, 'xl/styles.xml'                      , xl_styles                 (xlsx));
    add_blob_to_zip(xlsx_b, 'xl/workbook.xml'                    , xl_workbook               (xlsx));
    add_blob_to_zip(xlsx_b, '[Content_Types].xml'                , content_types             (xlsx));

    zipper.finish(xlsx_b);

    -- Direktes Return vermeidet die Warnung PLW-07206
    return xlsx_b;

  end create_xlsx; -- }}}

  function sql_to_xlsx(sql_stmt varchar2) return blob is -- {{{
    -- NEU_8: Flattening (Variablen hochgezogen, Unterfunktionen aufgelöst)
    workbook xlsx_writer.book_r;
    sheet        integer;
    cs_date      integer;
    cursor_             integer;
    res_                integer;
    column_count        integer;
    column_value        varchar2(4000);
    table_desc_         dbms_sql.desc_tab;
    type column_t       is record(name varchar2(30), datatype char(1), max_characters number);
    type columns_t      is table of column_t;
    column_  column_t;
    columns_  columns_t := columns_t();
    cur_row integer := 1;
  begin

    -- NEU_8: direkter Aufruf ohne Prefix
    workbook := start_book;
    sheet    := add_sheet(workbook, 'Result set');
    cs_date  := add_cell_style(workbook, num_fmt_id => "m/d/yy h:mm");

    cursor_  := dbms_sql.open_cursor;
    dbms_sql.parse(cursor_, sql_stmt, dbms_sql.native);
    dbms_sql.describe_columns(cursor_, column_count, table_desc_);
    
    for c in 1 .. column_count loop
        dbms_sql.define_column(cursor_, c, column_value, 4000);
    end loop;
    
    res_ := dbms_sql.execute(cursor_);
    
    -- NEU_8: Logic von column_names_and_types integriert
    for c in 1 .. column_count loop
        column_.name     := table_desc_(c).col_name;
        column_.datatype := case table_desc_(c).col_type 
                            when dbms_sql.number_type   then 'N'
                            when dbms_sql.date_type     then 'D'
                            when dbms_sql.varchar2_type then 'C'
                            when dbms_sql.char_type     then 'C'
                            else '??' end;
        columns_.extend;
        columns_(c) := column_;
    end loop;

    -- NEU_8: Logic von header integriert
    for c in 1 .. column_count loop
        add_cell(workbook, sheet, 1, c, text => columns_(c).name);
        if columns_(c).datatype = 'D' then
           col_width(workbook, sheet, c, 17);
        end if;
    end loop;

    -- NEU_8: Logic von result_set integriert
    loop
        exit when dbms_sql.fetch_rows(cursor_) = 0;
        cur_row := 1 + cur_row;
        workbook.sheets(sheet).rows_(cur_row).height := null; -- NEU_5 Init

        for c in 1 .. column_count loop
            dbms_sql.column_value(cursor_, c, column_value);
            -- NEU_5: Nutzung add_cell_fast
            if columns_(c).datatype = 'N' then
               add_cell_fast(workbook, sheet, cur_row, c, value_ => column_value);
            elsif columns_(c).datatype = 'D' then
               add_cell_fast(workbook, sheet, cur_row, c, date_  => column_value, style_id => cs_date);
            elsif columns_(c).datatype = 'C' then
               add_cell_fast(workbook, sheet, cur_row, c, text   => column_value);
               columns_(c).max_characters := greatest(nvl(columns_(c).max_characters, 0), nvl(length(column_value), 0));
            end if;
        end loop;
    end loop;

    for c in 1 .. column_count loop
        if columns_(c).datatype = 'C' then
           if columns_(c).max_characters > 13 then
               col_width(workbook, sheet, c, least(columns_(c).max_characters, 50) * 0.95);
           end if;
        end if;
    end loop;

    return create_xlsx(workbook);

  end sql_to_xlsx; -- }}}

  -- NEU_9: Globaler NLS-Block ENTFERNT! Das verhindert Seiteneffekte in der Session.
  -- Die notwendige Punkt-Konvertierung geschieht jetzt lokal in xl_worksheets_sheet.
end xlsx_writer; -- }}}
/