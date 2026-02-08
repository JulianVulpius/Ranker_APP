create or replace PACKAGE BODY "DBS_XLSX_PARSER" is
    g_worksheets_path_prefix constant varchar2(14) := 'xl/worksheets/';

    --==================================================================================================================
    function get_date( p_xlsx_date_number in number ) return date is
    begin
        return
            case when p_xlsx_date_number > 61
                      then DATE'1900-01-01' - 2 + p_xlsx_date_number
                      else DATE'1900-01-01' - 1 + p_xlsx_date_number
            end;
    end get_date;

    --==================================================================================================================
    procedure get_blob_content(
        p_xlsx_name    in            varchar2,
        p_xlsx_content in out nocopy blob )
    is
    begin
        if p_xlsx_name is not null then
            BEGIN
                SELECT blob_content
                INTO p_xlsx_content
                FROM wwv_flow_files
                --FROM apex_application_temp_files
                WHERE name = p_xlsx_name;
            EXCEPTION
            WHEN NO_DATA_FOUND THEN
                BEGIN
                    SELECT blob_content
                    INTO p_xlsx_content
                    --FROM wwv_flow_files
                    FROM apex_application_temp_files
                    WHERE name = p_xlsx_name;
                EXCEPTION
                WHEN NO_DATA_FOUND THEN
                    null;
                END;
            END;
        end if;
    end get_blob_content;

    --==================================================================================================================
    function extract_worksheet(
        p_xlsx           in blob,
        p_worksheet_name in varchar2 ) return blob
    is
        l_worksheet blob;
    begin
        if p_xlsx is null or p_worksheet_name is null then
           return null;
        end if;

        l_worksheet := apex_zip.get_file_content(
            p_zipped_blob => p_xlsx,
            p_file_name   => g_worksheets_path_prefix || p_worksheet_name || '.xml' );

        if l_worksheet is null then
            raise_application_error(-20000, 'WORKSHEET "' || p_worksheet_name || '" DOES NOT EXIST');
        end if;
        return l_worksheet;
    end extract_worksheet;

    --==================================================================================================================
    procedure extract_shared_strings(
        p_xlsx           in blob,
        p_strings        in out nocopy wwv_flow_global.vc_arr2 )
    is
        l_shared_strings blob;
    begin
        l_shared_strings := apex_zip.get_file_content(
            p_zipped_blob => p_xlsx,
            p_file_name   => 'xl/sharedStrings.xml' );

        if l_shared_strings is null then
            return;
        end if;

        select nvl(shared_string,shared_string2)
          bulk collect into p_strings
          from xmltable(
              xmlnamespaces( default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' ),
              '//si'
              passing xmltype.createxml( l_shared_strings, nls_charset_id('AL32UTF8'), null )
              columns
                 shared_string varchar2(4000)   path 't/text()',
                 shared_string2 varchar2(4000)   path 'string-join(r/t/text()," ")' );

    end extract_shared_strings;

    --==================================================================================================================
    procedure extract_date_styles(
        p_xlsx           in blob,
        p_format_codes   in out nocopy wwv_flow_global.vc_arr2 )
    is
        l_stylesheet blob;
    begin
        l_stylesheet := apex_zip.get_file_content(
            p_zipped_blob => p_xlsx,
            p_file_name   => 'xl/styles.xml' );

        if l_stylesheet is null then
            return;
        end if;

        select lower( n.formatCode )
        bulk collect into p_format_codes
        from
            xmltable(
                xmlnamespaces( default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' ),
                '//cellXfs/xf'
                passing xmltype.createxml( l_stylesheet, nls_charset_id('AL32UTF8'), null )
                columns
                   numFmtId number path '@numFmtId' ) s,
            xmltable(
                xmlnamespaces( default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' ),
                '//numFmts/numFmt'
                passing xmltype.createxml( l_stylesheet, nls_charset_id('AL32UTF8'), null )
                columns
                   formatCode varchar2(255) path '@formatCode',
                   numFmtId   number        path '@numFmtId' ) n
        where s.numFmtId = n.numFmtId ( + );

    end extract_date_styles;

    --==================================================================================================================
    function convert_ref_to_col#( p_col_ref in varchar2 ) return pls_integer is
        l_colpart  varchar2(10);
        l_linepart varchar2(10);
    begin
        l_colpart := replace(translate(p_col_ref,'1234567890','__________'), '_');
        if length( l_colpart ) = 1 then
            return ascii( l_colpart ) - 64;
        else
            return ( ascii( substr( l_colpart, 1, 1 ) ) - 64 ) * 26 + ( ascii( substr( l_colpart, 2, 1 ) ) - 64 );
        end if;
    end convert_ref_to_col#;

    --==================================================================================================================
    procedure reset_row( p_parsed_row in out nocopy dbs_xlsx_row_t ) is
    begin
        -- reset row
        p_parsed_row.col01 := null; p_parsed_row.col02 := null; p_parsed_row.col03 := null; p_parsed_row.col04 := null; p_parsed_row.col05 := null;
        p_parsed_row.col06 := null; p_parsed_row.col07 := null; p_parsed_row.col08 := null; p_parsed_row.col09 := null; p_parsed_row.col10 := null;
        p_parsed_row.col11 := null; p_parsed_row.col12 := null; p_parsed_row.col13 := null; p_parsed_row.col14 := null; p_parsed_row.col15 := null;
        p_parsed_row.col16 := null; p_parsed_row.col17 := null; p_parsed_row.col18 := null; p_parsed_row.col19 := null; p_parsed_row.col20 := null;
        p_parsed_row.col21 := null; p_parsed_row.col22 := null; p_parsed_row.col23 := null; p_parsed_row.col24 := null; p_parsed_row.col25 := null;
        p_parsed_row.col26 := null; p_parsed_row.col27 := null; p_parsed_row.col28 := null; p_parsed_row.col29 := null; p_parsed_row.col30 := null;
        p_parsed_row.col31 := null; p_parsed_row.col32 := null; p_parsed_row.col33 := null; p_parsed_row.col34 := null; p_parsed_row.col35 := null;
        p_parsed_row.col36 := null; p_parsed_row.col37 := null; p_parsed_row.col38 := null; p_parsed_row.col39 := null; p_parsed_row.col40 := null;
        p_parsed_row.col41 := null; p_parsed_row.col42 := null; p_parsed_row.col43 := null; p_parsed_row.col44 := null; p_parsed_row.col45 := null;
        p_parsed_row.col46 := null; p_parsed_row.col47 := null; p_parsed_row.col48 := null; p_parsed_row.col49 := null; p_parsed_row.col50 := null;
        p_parsed_row.col51 := null; p_parsed_row.col52 := null; p_parsed_row.col53 := null; p_parsed_row.col54 := null; p_parsed_row.col55 := null;
        p_parsed_row.col56 := null; p_parsed_row.col57 := null; p_parsed_row.col58 := null; p_parsed_row.col59 := null; p_parsed_row.col60 := null;
        p_parsed_row.col61 := null; p_parsed_row.col62 := null; p_parsed_row.col63 := null; p_parsed_row.col64 := null; p_parsed_row.col65 := null;
        p_parsed_row.col66 := null; p_parsed_row.col67 := null; p_parsed_row.col68 := null; p_parsed_row.col69 := null; p_parsed_row.col70 := null;
        p_parsed_row.col71 := null; p_parsed_row.col72 := null; p_parsed_row.col73 := null; p_parsed_row.col74 := null; p_parsed_row.col75 := null;
        p_parsed_row.col76 := null; p_parsed_row.col77 := null; p_parsed_row.col78 := null; p_parsed_row.col79 := null; p_parsed_row.col80 := null;
        p_parsed_row.col81 := null; p_parsed_row.col82 := null; p_parsed_row.col83 := null; p_parsed_row.col84 := null; p_parsed_row.col85 := null;
        p_parsed_row.col86 := null; p_parsed_row.col87 := null; p_parsed_row.col88 := null; p_parsed_row.col89 := null; p_parsed_row.col90 := null;
        p_parsed_row.col91 := null; p_parsed_row.col92 := null; p_parsed_row.col93 := null; p_parsed_row.col94 := null; p_parsed_row.col95 := null;
        p_parsed_row.col96 := null; p_parsed_row.col97 := null; p_parsed_row.col98 := null; p_parsed_row.col99 := null; p_parsed_row.col100 := null;
        p_parsed_row.col101 := null; p_parsed_row.col102 := null; p_parsed_row.col103 := null; p_parsed_row.col104 := null; p_parsed_row.col105 := null;
        p_parsed_row.col106 := null; p_parsed_row.col107 := null; p_parsed_row.col108 := null; p_parsed_row.col109 := null; p_parsed_row.col110 := null;
        p_parsed_row.col111 := null; p_parsed_row.col112 := null; p_parsed_row.col113 := null; p_parsed_row.col114 := null; p_parsed_row.col115 := null;
        p_parsed_row.col116 := null; p_parsed_row.col117 := null; p_parsed_row.col118 := null; p_parsed_row.col119 := null; p_parsed_row.col120 := null;
        p_parsed_row.col121 := null; p_parsed_row.col122 := null; p_parsed_row.col123 := null; p_parsed_row.col124 := null; p_parsed_row.col125 := null;
        p_parsed_row.col126 := null; p_parsed_row.col127 := null; p_parsed_row.col128 := null; p_parsed_row.col129 := null; p_parsed_row.col130 := null;
        p_parsed_row.col131 := null; p_parsed_row.col132 := null; p_parsed_row.col133 := null; p_parsed_row.col134 := null; p_parsed_row.col135 := null;
        p_parsed_row.col136 := null; p_parsed_row.col137 := null; p_parsed_row.col138 := null; p_parsed_row.col139 := null; p_parsed_row.col140 := null;
        p_parsed_row.col141 := null; p_parsed_row.col142 := null; p_parsed_row.col143 := null; p_parsed_row.col144 := null; p_parsed_row.col145 := null;
        p_parsed_row.col146 := null; p_parsed_row.col147 := null; p_parsed_row.col148 := null; p_parsed_row.col149 := null; p_parsed_row.col150 := null;
        p_parsed_row.col151 := null; p_parsed_row.col152 := null; p_parsed_row.col153 := null; p_parsed_row.col154 := null; p_parsed_row.col155 := null;
        p_parsed_row.col156 := null; p_parsed_row.col157 := null; p_parsed_row.col158 := null; p_parsed_row.col159 := null; p_parsed_row.col160 := null;
        p_parsed_row.col161 := null; p_parsed_row.col162 := null; p_parsed_row.col163 := null; p_parsed_row.col164 := null; p_parsed_row.col165 := null;
        p_parsed_row.col166 := null; p_parsed_row.col167 := null; p_parsed_row.col168 := null; p_parsed_row.col169 := null; p_parsed_row.col170 := null;
        p_parsed_row.col171 := null; p_parsed_row.col172 := null; p_parsed_row.col173 := null; p_parsed_row.col174 := null; p_parsed_row.col175 := null;
        p_parsed_row.col176 := null; p_parsed_row.col177 := null; p_parsed_row.col178 := null; p_parsed_row.col179 := null; p_parsed_row.col180 := null;
        p_parsed_row.col181 := null; p_parsed_row.col182 := null; p_parsed_row.col183 := null; p_parsed_row.col184 := null; p_parsed_row.col185 := null;
        p_parsed_row.col186 := null; p_parsed_row.col187 := null; p_parsed_row.col188 := null; p_parsed_row.col189 := null; p_parsed_row.col190 := null;
        p_parsed_row.col191 := null; p_parsed_row.col192 := null; p_parsed_row.col193 := null; p_parsed_row.col194 := null; p_parsed_row.col195 := null;
        p_parsed_row.col196 := null; p_parsed_row.col197 := null; p_parsed_row.col198 := null; p_parsed_row.col199 := null; p_parsed_row.col200 := null;
    end reset_row;

    --==================================================================================================================
    function parse(
        p_xlsx_name      in varchar2 default null,
        p_xlsx_content   in blob     default null,
        p_worksheet_name in varchar2 default 'sheet1',
        p_max_rows       in number   default 1000000 ) return dbs_xlsx_tab_t pipelined
    is
        l_worksheet           blob;
        l_xlsx_content        blob;

        l_shared_strings      wwv_flow_global.vc_arr2;
        l_format_codes        wwv_flow_global.vc_arr2;

        l_parsed_row          dbs_xlsx_row_t := dbs_xlsx_row_t();
        l_first_row           boolean     := true;
        l_value               varchar2(32767);

        l_line#               pls_integer := 1;
        l_real_col#           pls_integer;
        l_row_has_content     boolean := false;
    begin
        if p_xlsx_content is null then
            get_blob_content( p_xlsx_name, l_xlsx_content );
        else
            l_xlsx_content := p_xlsx_content;
        end if;

        if l_xlsx_content is null then
            return;
        end if;

        l_worksheet := extract_worksheet(
            p_xlsx           => l_xlsx_content,
            p_worksheet_name => p_worksheet_name );

        extract_shared_strings(
            p_xlsx    => l_xlsx_content,
            p_strings => l_shared_strings );

        extract_date_styles(
            p_xlsx    => l_xlsx_content,
            p_format_codes => l_format_codes );

        -- the actual XML parsing starts here
        for i in (
            select
                r.xlsx_row,
                c.xlsx_col#,
                c.xlsx_col,
                c.xlsx_col_type,
                c.xlsx_col_style,
                c.xlsx_val
            from xmltable(
                xmlnamespaces( default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' ),
                '//row'
                passing xmltype.createxml( l_worksheet, nls_charset_id('AL32UTF8'), null )
                columns
                     xlsx_row number   path '@r',
                     xlsx_cols xmltype path '.'
            ) r, xmltable (
                xmlnamespaces( default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' ),
                '//c'
                passing r.xlsx_cols
                columns
                     xlsx_col#      for ordinality,
                     xlsx_col       varchar2(15)   path '@r',
                     xlsx_col_type  varchar2(15)   path '@t',
                     xlsx_col_style varchar2(15)   path '@s',
                     xlsx_val       varchar2(4000) path 'v/text()'
            ) c
            where p_max_rows is null or r.xlsx_row <= p_max_rows
        ) loop
            if i.xlsx_col# = 1 then
                l_parsed_row.linenr := l_line#;
                if not l_first_row then
                    pipe row( l_parsed_row );
                    l_line# := i.xlsx_row;--l_line# + 1;
                    reset_row( l_parsed_row );
                    l_row_has_content := false;
                else
                    l_first_row := false;
                end if;
            end if;

            if i.xlsx_col_type = 's' then
                if l_shared_strings.exists( i.xlsx_val + 1) then
                    l_value := l_shared_strings( i.xlsx_val + 1);
                else
                    l_value := '[Data Error: N/A]' ;
                end if;
            else
                if l_format_codes.exists( i.xlsx_col_style + 1 ) and (
                    instr( l_format_codes( i.xlsx_col_style + 1 ), 'd' ) > 0 and
                    instr( l_format_codes( i.xlsx_col_style + 1 ), 'm' ) > 0 )
                then
                    l_value := to_char( get_date( i.xlsx_val ), c_date_format );
                else
                    l_value := i.xlsx_val;
                end if;
            end if;

            pragma inline( convert_ref_to_col#, 'YES' );
            l_real_col# := convert_ref_to_col#( i.xlsx_col );

            if l_real_col# between 1 and 200 then
                l_row_has_content := true;
            end if;

            -- we currently support 50 columns - but this can easily be increased. Just add additional lines
            -- as follows:
            -- when l_real_col# = {nn} then l_parsed_row.col{nn} := l_value;
            case
                when l_real_col# =  1 then l_parsed_row.col01 := l_value;
                when l_real_col# =  2 then l_parsed_row.col02 := l_value;
                when l_real_col# =  3 then l_parsed_row.col03 := l_value;
                when l_real_col# =  4 then l_parsed_row.col04 := l_value;
                when l_real_col# =  5 then l_parsed_row.col05 := l_value;
                when l_real_col# =  6 then l_parsed_row.col06 := l_value;
                when l_real_col# =  7 then l_parsed_row.col07 := l_value;
                when l_real_col# =  8 then l_parsed_row.col08 := l_value;
                when l_real_col# =  9 then l_parsed_row.col09 := l_value;
                when l_real_col# = 10 then l_parsed_row.col10 := l_value;
                when l_real_col# = 11 then l_parsed_row.col11 := l_value;
                when l_real_col# = 12 then l_parsed_row.col12 := l_value;
                when l_real_col# = 13 then l_parsed_row.col13 := l_value;
                when l_real_col# = 14 then l_parsed_row.col14 := l_value;
                when l_real_col# = 15 then l_parsed_row.col15 := l_value;
                when l_real_col# = 16 then l_parsed_row.col16 := l_value;
                when l_real_col# = 17 then l_parsed_row.col17 := l_value;
                when l_real_col# = 18 then l_parsed_row.col18 := l_value;
                when l_real_col# = 19 then l_parsed_row.col19 := l_value;
                when l_real_col# = 20 then l_parsed_row.col20 := l_value;
                when l_real_col# = 21 then l_parsed_row.col21 := l_value;
                when l_real_col# = 22 then l_parsed_row.col22 := l_value;
                when l_real_col# = 23 then l_parsed_row.col23 := l_value;
                when l_real_col# = 24 then l_parsed_row.col24 := l_value;
                when l_real_col# = 25 then l_parsed_row.col25 := l_value;
                when l_real_col# = 26 then l_parsed_row.col26 := l_value;
                when l_real_col# = 27 then l_parsed_row.col27 := l_value;
                when l_real_col# = 28 then l_parsed_row.col28 := l_value;
                when l_real_col# = 29 then l_parsed_row.col29 := l_value;
                when l_real_col# = 30 then l_parsed_row.col30 := l_value;
                when l_real_col# = 31 then l_parsed_row.col31 := l_value;
                when l_real_col# = 32 then l_parsed_row.col32 := l_value;
                when l_real_col# = 33 then l_parsed_row.col33 := l_value;
                when l_real_col# = 34 then l_parsed_row.col34 := l_value;
                when l_real_col# = 35 then l_parsed_row.col35 := l_value;
                when l_real_col# = 36 then l_parsed_row.col36 := l_value;
                when l_real_col# = 37 then l_parsed_row.col37 := l_value;
                when l_real_col# = 38 then l_parsed_row.col38 := l_value;
                when l_real_col# = 39 then l_parsed_row.col39 := l_value;
                when l_real_col# = 40 then l_parsed_row.col40 := l_value;
                when l_real_col# = 41 then l_parsed_row.col41 := l_value;
                when l_real_col# = 42 then l_parsed_row.col42 := l_value;
                when l_real_col# = 43 then l_parsed_row.col43 := l_value;
                when l_real_col# = 44 then l_parsed_row.col44 := l_value;
                when l_real_col# = 45 then l_parsed_row.col45 := l_value;
                when l_real_col# = 46 then l_parsed_row.col46 := l_value;
                when l_real_col# = 47 then l_parsed_row.col47 := l_value;
                when l_real_col# = 48 then l_parsed_row.col48 := l_value;
                when l_real_col# = 49 then l_parsed_row.col49 := l_value;
                when l_real_col# = 50 then l_parsed_row.col50 := l_value;
                when l_real_col# = 51 then l_parsed_row.col51 := l_value;
                when l_real_col# = 52 then l_parsed_row.col52 := l_value;
                when l_real_col# = 53 then l_parsed_row.col53 := l_value;
                when l_real_col# = 54 then l_parsed_row.col54 := l_value;
                when l_real_col# = 55 then l_parsed_row.col55 := l_value;
                when l_real_col# = 56 then l_parsed_row.col56 := l_value;
                when l_real_col# = 57 then l_parsed_row.col57 := l_value;
                when l_real_col# = 58 then l_parsed_row.col58 := l_value;
                when l_real_col# = 59 then l_parsed_row.col59 := l_value;
                when l_real_col# = 60 then l_parsed_row.col60 := l_value;
                when l_real_col# = 61 then l_parsed_row.col61 := l_value;
                when l_real_col# = 62 then l_parsed_row.col62 := l_value;
                when l_real_col# = 63 then l_parsed_row.col63 := l_value;
                when l_real_col# = 64 then l_parsed_row.col64 := l_value;
                when l_real_col# = 65 then l_parsed_row.col65 := l_value;
                when l_real_col# = 66 then l_parsed_row.col66 := l_value;
                when l_real_col# = 67 then l_parsed_row.col67 := l_value;
                when l_real_col# = 68 then l_parsed_row.col68 := l_value;
                when l_real_col# = 69 then l_parsed_row.col69 := l_value;
                when l_real_col# = 70 then l_parsed_row.col70 := l_value;
                when l_real_col# = 71 then l_parsed_row.col71 := l_value;
                when l_real_col# = 72 then l_parsed_row.col72 := l_value;
                when l_real_col# = 73 then l_parsed_row.col73 := l_value;
                when l_real_col# = 74 then l_parsed_row.col74 := l_value;
                when l_real_col# = 75 then l_parsed_row.col75 := l_value;
                when l_real_col# = 76 then l_parsed_row.col76 := l_value;
                when l_real_col# = 77 then l_parsed_row.col77 := l_value;
                when l_real_col# = 78 then l_parsed_row.col78 := l_value;
                when l_real_col# = 79 then l_parsed_row.col79 := l_value;
                when l_real_col# = 80 then l_parsed_row.col80 := l_value;
                when l_real_col# = 81 then l_parsed_row.col81 := l_value;
                when l_real_col# = 82 then l_parsed_row.col82 := l_value;
                when l_real_col# = 83 then l_parsed_row.col83 := l_value;
                when l_real_col# = 84 then l_parsed_row.col84 := l_value;
                when l_real_col# = 85 then l_parsed_row.col85 := l_value;
                when l_real_col# = 86 then l_parsed_row.col86 := l_value;
                when l_real_col# = 87 then l_parsed_row.col87 := l_value;
                when l_real_col# = 88 then l_parsed_row.col88 := l_value;
                when l_real_col# = 89 then l_parsed_row.col89 := l_value;
                when l_real_col# = 90 then l_parsed_row.col90 := l_value;
                when l_real_col# = 91 then l_parsed_row.col91 := l_value;
                when l_real_col# = 92 then l_parsed_row.col92 := l_value;
                when l_real_col# = 93 then l_parsed_row.col93 := l_value;
                when l_real_col# = 94 then l_parsed_row.col94 := l_value;
                when l_real_col# = 95 then l_parsed_row.col95 := l_value;
                when l_real_col# = 96 then l_parsed_row.col96 := l_value;
                when l_real_col# = 97 then l_parsed_row.col97 := l_value;
                when l_real_col# = 98 then l_parsed_row.col98 := l_value;
                when l_real_col# = 99 then l_parsed_row.col99 := l_value;
                when l_real_col# = 100 then l_parsed_row.col100 := l_value;
                when l_real_col# = 101 then l_parsed_row.col101 := l_value;
                when l_real_col# = 102 then l_parsed_row.col102 := l_value;
                when l_real_col# = 103 then l_parsed_row.col103 := l_value;
                when l_real_col# = 104 then l_parsed_row.col104 := l_value;
                when l_real_col# = 105 then l_parsed_row.col105 := l_value;
                when l_real_col# = 106 then l_parsed_row.col106 := l_value;
                when l_real_col# = 107 then l_parsed_row.col107 := l_value;
                when l_real_col# = 108 then l_parsed_row.col108 := l_value;
                when l_real_col# = 109 then l_parsed_row.col109 := l_value;
                when l_real_col# = 110 then l_parsed_row.col110 := l_value;
                when l_real_col# = 111 then l_parsed_row.col111 := l_value;
                when l_real_col# = 112 then l_parsed_row.col112 := l_value;
                when l_real_col# = 113 then l_parsed_row.col113 := l_value;
                when l_real_col# = 114 then l_parsed_row.col114 := l_value;
                when l_real_col# = 115 then l_parsed_row.col115 := l_value;
                when l_real_col# = 116 then l_parsed_row.col116 := l_value;
                when l_real_col# = 117 then l_parsed_row.col117 := l_value;
                when l_real_col# = 118 then l_parsed_row.col118 := l_value;
                when l_real_col# = 119 then l_parsed_row.col119 := l_value;
                when l_real_col# = 120 then l_parsed_row.col120 := l_value;
                when l_real_col# = 121 then l_parsed_row.col121 := l_value;
                when l_real_col# = 122 then l_parsed_row.col122 := l_value;
                when l_real_col# = 123 then l_parsed_row.col123 := l_value;
                when l_real_col# = 124 then l_parsed_row.col124 := l_value;
                when l_real_col# = 125 then l_parsed_row.col125 := l_value;
                when l_real_col# = 126 then l_parsed_row.col126 := l_value;
                when l_real_col# = 127 then l_parsed_row.col127 := l_value;
                when l_real_col# = 128 then l_parsed_row.col128 := l_value;
                when l_real_col# = 129 then l_parsed_row.col129 := l_value;
                when l_real_col# = 130 then l_parsed_row.col130 := l_value;
                when l_real_col# = 131 then l_parsed_row.col131 := l_value;
                when l_real_col# = 132 then l_parsed_row.col132 := l_value;
                when l_real_col# = 133 then l_parsed_row.col133 := l_value;
                when l_real_col# = 134 then l_parsed_row.col134 := l_value;
                when l_real_col# = 135 then l_parsed_row.col135 := l_value;
                when l_real_col# = 136 then l_parsed_row.col136 := l_value;
                when l_real_col# = 137 then l_parsed_row.col137 := l_value;
                when l_real_col# = 138 then l_parsed_row.col138 := l_value;
                when l_real_col# = 139 then l_parsed_row.col139 := l_value;
                when l_real_col# = 140 then l_parsed_row.col140 := l_value;
                when l_real_col# = 141 then l_parsed_row.col141 := l_value;
                when l_real_col# = 142 then l_parsed_row.col142 := l_value;
                when l_real_col# = 143 then l_parsed_row.col143 := l_value;
                when l_real_col# = 144 then l_parsed_row.col144 := l_value;
                when l_real_col# = 145 then l_parsed_row.col145 := l_value;
                when l_real_col# = 146 then l_parsed_row.col146 := l_value;
                when l_real_col# = 147 then l_parsed_row.col147 := l_value;
                when l_real_col# = 148 then l_parsed_row.col148 := l_value;
                when l_real_col# = 149 then l_parsed_row.col149 := l_value;
                when l_real_col# = 150 then l_parsed_row.col150 := l_value;
                when l_real_col# = 151 then l_parsed_row.col151 := l_value;
                when l_real_col# = 152 then l_parsed_row.col152 := l_value;
                when l_real_col# = 153 then l_parsed_row.col153 := l_value;
                when l_real_col# = 154 then l_parsed_row.col154 := l_value;
                when l_real_col# = 155 then l_parsed_row.col155 := l_value;
                when l_real_col# = 156 then l_parsed_row.col156 := l_value;
                when l_real_col# = 157 then l_parsed_row.col157 := l_value;
                when l_real_col# = 158 then l_parsed_row.col158 := l_value;
                when l_real_col# = 159 then l_parsed_row.col159 := l_value;
                when l_real_col# = 160 then l_parsed_row.col160 := l_value;
                when l_real_col# = 161 then l_parsed_row.col161 := l_value;
                when l_real_col# = 162 then l_parsed_row.col162 := l_value;
                when l_real_col# = 163 then l_parsed_row.col163 := l_value;
                when l_real_col# = 164 then l_parsed_row.col164 := l_value;
                when l_real_col# = 165 then l_parsed_row.col165 := l_value;
                when l_real_col# = 166 then l_parsed_row.col166 := l_value;
                when l_real_col# = 167 then l_parsed_row.col167 := l_value;
                when l_real_col# = 168 then l_parsed_row.col168 := l_value;
                when l_real_col# = 169 then l_parsed_row.col169 := l_value;
                when l_real_col# = 170 then l_parsed_row.col170 := l_value;
                when l_real_col# = 171 then l_parsed_row.col171 := l_value;
                when l_real_col# = 172 then l_parsed_row.col172 := l_value;
                when l_real_col# = 173 then l_parsed_row.col173 := l_value;
                when l_real_col# = 174 then l_parsed_row.col174 := l_value;
                when l_real_col# = 175 then l_parsed_row.col175 := l_value;
                when l_real_col# = 176 then l_parsed_row.col176 := l_value;
                when l_real_col# = 177 then l_parsed_row.col177 := l_value;
                when l_real_col# = 178 then l_parsed_row.col178 := l_value;
                when l_real_col# = 179 then l_parsed_row.col179 := l_value;
                when l_real_col# = 180 then l_parsed_row.col180 := l_value;
                when l_real_col# = 181 then l_parsed_row.col181 := l_value;
                when l_real_col# = 182 then l_parsed_row.col182 := l_value;
                when l_real_col# = 183 then l_parsed_row.col183 := l_value;
                when l_real_col# = 184 then l_parsed_row.col184 := l_value;
                when l_real_col# = 185 then l_parsed_row.col185 := l_value;
                when l_real_col# = 186 then l_parsed_row.col186 := l_value;
                when l_real_col# = 187 then l_parsed_row.col187 := l_value;
                when l_real_col# = 188 then l_parsed_row.col188 := l_value;
                when l_real_col# = 189 then l_parsed_row.col189 := l_value;
                when l_real_col# = 190 then l_parsed_row.col190 := l_value;
                when l_real_col# = 191 then l_parsed_row.col191 := l_value;
                when l_real_col# = 192 then l_parsed_row.col192 := l_value;
                when l_real_col# = 193 then l_parsed_row.col193 := l_value;
                when l_real_col# = 194 then l_parsed_row.col194 := l_value;
                when l_real_col# = 195 then l_parsed_row.col195 := l_value;
                when l_real_col# = 196 then l_parsed_row.col196 := l_value;
                when l_real_col# = 197 then l_parsed_row.col197 := l_value;
                when l_real_col# = 198 then l_parsed_row.col198 := l_value;
                when l_real_col# = 199 then l_parsed_row.col199 := l_value;
                when l_real_col# = 200 then l_parsed_row.col200 := l_value;
                else null;
            end case;

        end loop;
        if l_row_has_content then
            l_parsed_row.linenr := l_line#;
            pipe row( l_parsed_row );
        end if;

        return;
    end parse;

    --==================================================================================================================
    function get_worksheets(
        p_xlsx_content   in blob     default null,
        p_xlsx_name      in varchar2 default null ) return apex_t_varchar2 pipelined
    is
        l_zip_files           apex_zip.t_files;
        l_xlsx_content        blob;
    begin
        if p_xlsx_content is null then
            get_blob_content( p_xlsx_name, l_xlsx_content );
        else
            l_xlsx_content := p_xlsx_content;
        end if;

        l_zip_files := apex_zip.get_files(
            p_zipped_blob => l_xlsx_content );

        for i in 1 .. l_zip_files.count loop
            if substr( l_zip_files( i ), 1, length( g_worksheets_path_prefix ) ) = g_worksheets_path_prefix then
                pipe row( rtrim( substr( l_zip_files ( i ), length( g_worksheets_path_prefix ) + 1 ), '.xml' ) );
            end if;
        end loop;

        return;
    end get_worksheets;

end dbs_xlsx_parser;