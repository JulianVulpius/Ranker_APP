set define off

create or replace package xlsx_writer as -- {{{
-- vi: foldmarker={{{,}}}

  "0"                          constant integer :=  1;
  "0.00"                       constant integer :=  2;
  "#.##0"                      constant integer :=  3;
  "#.##0.00"                   constant integer :=  4;

  "0%"                         constant integer :=  9;
  "0.00%"                      constant integer := 10;
  "0.00e+00"                   constant integer := 11;
  "# ?/?"                      constant integer := 12;
  "# ??/??"                    constant integer := 13;
  "mm-dd-yy"                   constant integer := 14; 
  "d-mmm-yy"                   constant integer := 15;
  "d-mmm"                      constant integer := 16;
  "mmm-yy"                     constant integer := 17;
  "h:mm am/pm"                 constant integer := 18;
  "h:mm:ss am/pm"              constant integer := 19;
  "h:mm"                       constant integer := 20;
  "h:mm:ss"                    constant integer := 21;
  "m/d/yy h:mm"                constant integer := 22;

  "#,##0 ;(#,##0)"             constant integer := 37;
  "#,##0 ;[red](#,##0)"        constant integer := 38;
  "#,##0.00;(#,##0.00)"        constant integer := 39;
  "#,##0.00;[red](#,##0.00)"   constant integer := 40;

  "mm:ss"                      constant integer := 45;
  "[h]:mm:ss"                  constant integer := 46;
  "mmss.0"                     constant integer := 47;
  "##0.0e+0"                   constant integer := 48;
  "@"                          constant integer := 49;

  -- {{{ types

  -- {{{ related to styles

  type border_r             is record(raw_        varchar2(1000));
  type border_t             is table of border_r;

  type fill_r               is record(raw_        varchar2(1000));
  type fill_t               is table of fill_r;

  type font_r               is record(name        varchar2(100),
                                      size_       number       ,
                                      color       varchar2(100), 
                                      b           boolean      ,
                                      i           boolean      ,
                                      u           boolean
                                    );

  type font_t               is table of font_r;

  type num_fmt_r            is record(raw_        varchar2(1000));
  type num_fmt_t            is table of num_fmt_r;

  type cell_style_r         is record(font_id             integer,
                                      fill_id             integer,
                                      border_id           integer,
                                      num_fmt_id          integer,
                                      vertical_alignment  varchar2(10),
                                      wrap_text           boolean
                                    );

  type cell_style_t         is table of cell_style_r;

  -- }}}

  type shared_string_r      is record(val        varchar2(4000));
  type shared_string_t      is table of shared_string_r;

  -- {{{ sheet types

  type col_width_r          is record(start_col        integer,
                                      end_col          integer,
                                      width            number
                               );
  type col_width_t          is table of col_width_r;

  type cell_r               is record(style_id         integer,
                                      shared_string_id integer,
                                      value_           number,
                                      formula          varchar2(4000));
               
  type cell_t               is table of cell_r index by pls_integer;

  type row_r                is record(r          integer,
                                      height     number,
                                      cells      cell_t);

  type row_t                is table of row_r index by pls_integer;

  type sheet_rel_r          is record(raw_       varchar2(4000));
  type sheet_rel_t          is table of sheet_rel_r;

  -- NEU_7: Typen für Datenvalidierung (Dropdowns) müssen VOR sheet_r definiert sein
  type validation_r is record(
      type_           varchar2(20),   
      formula1        varchar2(4000), 
      sqref           varchar2(4000), 
      show_input_msg  boolean,
      show_error_msg  boolean
  );
  type validation_t is table of validation_r;

  type checkbox_r           is record(text               varchar2(4000),
                                      checked            boolean,
                                      col_left           integer,
                                      row_top            integer);

  type checkbox_t           is table of checkbox_r;

  type vml_drawing_r        is record(checkboxes          checkbox_t);

  type vml_drawing_t        is table of vml_drawing_r;

  type sheet_r              is record(col_widths     col_width_t,
                                      name_          varchar2(100),
                                      rows_          row_t,
                                      split_x        integer,
                                      split_y        integer,
                                      sheet_rels     sheet_rel_t,
                                      vml_drawings   vml_drawing_t,
                                      validations    validation_t -- NEU_7: Integriert im Sheet Record
                                );

  type sheet_t              is table of sheet_r;

  -- }}}

  type media_r              is record(b                blob,
                                      name_            varchar2(100));

  type media_t              is table of media_r;

  type calc_chain_elem_r    is record(cell_reference   varchar2(10),
                                      sheet            integer);

  type calc_chain_elem_t    is table of calc_chain_elem_r;

  type drawing_r            is record(raw_             varchar2(30000));

  type drawing_t            is table of drawing_r;


  -- {{{ the book!

  type book_r               is record(sheets                   sheet_t,
                                      cell_styles              cell_style_t,
                                      borders                  border_t,
                                      fonts                    font_t,
                                      fills                    fill_t,
                                      num_fmts                 num_fmt_t,
                                      shared_strings           shared_string_t,
                                      medias                   media_t,
                                      calc_chain_elems         calc_chain_elem_t,
                                      drawings                 drawing_t,
                                      content_type_vmlDrawing  boolean
                                      );

  -- }}}

  -- }}}

  function  start_book                                       return book_r;
  
  -- NEU_6: 'nocopy' Hint hinzugefügt für Performanz bei großen Objekten
  function  add_sheet         (xlsx         in out nocopy book_r,
                               name_        in     varchar2) return integer;

  -- NEU_6: nocopy
  function add_cell_style     (xlsx         in out nocopy book_r,
                               font_id             integer  := 0,
                               fill_id             integer  := 0,
                               border_id           integer  := 0,
                               num_fmt_id          integer  := 0,
                               vertical_alignment  varchar2 := null,
                               wrap_text           boolean  := null
                            ) return integer;

  -- NEU_6: nocopy
  function add_border         (xlsx         in out nocopy book_r,
                               raw_                varchar2) return integer;

  -- NEU_6: nocopy
  function add_num_fmt        (xlsx         in out nocopy book_r,
                               raw_                varchar2,
                               return_id           integer) return integer;

  -- NEU_6: nocopy
  function add_font           (xlsx         in out nocopy book_r,
                               name                varchar2,
                               size_               number,
                               color               varchar2   := null,
                               b                   boolean    := false,
                               i                   boolean    := false,
                               u                   boolean    := false) return integer;

  -- NEU_6: nocopy
  function add_fill           (xlsx         in out nocopy book_r,
                               raw_                varchar2) return integer;
          
  -- NEU_6: nocopy
  procedure col_width         (xlsx         in out nocopy book_r,
                               sheet              integer,
                               col                integer,
                               width              number
                               );

  -- NEU_6: nocopy
  procedure col_width         (xlsx         in out nocopy book_r,
                               sheet              integer,
                               start_col          integer,
                               end_col            integer,
                               width              number
                               );

  -- NEU_6: nocopy
  procedure add_row           (xlsx        in out nocopy book_r,
                               sheet              integer,
                               r                  integer,
                               height             number := null);

  -- NEU_6: nocopy
  procedure freeze_sheet      (xlsx        in out nocopy book_r,
                               sheet       in     integer,
                               split_x     in     integer := null,
                               split_y     in     integer := null);
                             

  -- NEU_6: nocopy
  procedure add_cell          (xlsx        in out nocopy book_r,
                               sheet              integer,
                               r                  integer,
                               c                  integer,
                               style_id           integer  :=    0,
                               text               varchar2 := null,
                               value_             number   := null,
                               formula            varchar2 := null);

  -- NEU_6: nocopy
  procedure add_cell          (xlsx        in out nocopy book_r,
                               sheet              integer,
                               r                  integer,
                               c                  integer,
                               date_              date,
                               style_id           integer  :=    0);

  -- NEU_6: nocopy
  procedure add_sheet_rel     (xlsx        in out nocopy book_r,
                               sheet              integer,
                               raw_               varchar2);

  -- NEU_6: nocopy
  procedure add_media         (xlsx        in out nocopy book_r,
                               b                  blob,
                               name_              varchar2);        

  -- NEU_6: nocopy
  procedure add_drawing       (xlsx        in out nocopy book_r,
                               raw_               varchar2);

  -- NEU_6: nocopy
  procedure add_checkbox      (xlsx        in out nocopy book_r,
                               sheet              integer,
                               col_left           integer,
                               row_top            integer,
                               text               varchar2 := null,
                               checked            boolean  := false);
                          

  -- NEU_6: nocopy
  function create_xlsx        (xlsx        in out nocopy book_r) return blob;

  function col_to_letter(c integer) return varchar2;

  function sql_to_xlsx(sql_stmt varchar2) return blob;

  -- NEU_7: Public Interface für Dropdowns (Data Validation)
  -- NEU_6: Auch hier nocopy nutzen
  procedure add_data_validation(
      xlsx            in out nocopy book_r,
      sheet           integer,
      sqref           varchar2,
      formula1        varchar2,
      type_           varchar2 := 'list',
      show_input_msg  boolean := true,
      show_error_msg  boolean := true
  );

end xlsx_writer; -- }}}
/