create or replace PACKAGE BODY           "PREISDATENBANK_PKG" AS

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

  PROCEDURE Import_Muster(p_blob_id number, p_typ_id number,p_date date) AS
    v_namespace number;
  BEGIN

    select GetNamespace(p_blob_id) into v_namespace from dual;

    if v_namespace = 1 then
      Import_Muster1(p_blob_id,p_typ_id,p_date);
    end if;

    if v_namespace = 2 then
      Import_Muster2(p_blob_id,p_typ_id,p_date);
    end if;

  END Import_Muster;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

  PROCEDURE Import_Muster1(p_blob_id number, p_typ_id number,p_date date) AS

    v_kennung   pd_import_typ.kennung%type;
    v_name      pd_import_typ.name%type;

  BEGIN

    select  kennung,name
    into    v_kennung,v_name
    from    pd_import_typ
    where   id = p_typ_id;

    for i in 
            (
            SELECT  v_kennung code,v_name name,null mparent,null description,v_kennung kennung,null einheit
            FROM    pd_import_x83 x
            where   x.id = p_blob_id
            union all    
            SELECT  v_kennung || '.' || xt.code,xt.name,v_kennung mparent,null description,v_kennung kennung,null einheit
            FROM    pd_import_x83 x,
                    XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/200407' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                                PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                COLUMNS
                                   code     VARCHAR2(2000)  PATH '@RNoPart',
                                   name     VARCHAR2(4000)  PATH 'n:LblTx',
                                   subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                            ) xt
            where   x.id = p_blob_id
            union all
            SELECT  v_kennung || '.' || xt.code||'.'||xtd.code2,xtd.name2,v_kennung || '.' || xt.code,null description,null,null
            FROM    pd_import_x83 x,
                    XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/200407' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                                PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                COLUMNS
                                   code     VARCHAR2(2000)  PATH '@RNoPart',
                                   name     VARCHAR2(4000)  PATH 'n:LblTx',
                                   subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                            ) xt,
                    xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/200407'),'/BoQCtgy'
                                   PASSING xt.subitem
                                   COLUMNS 
                                           code2        VARCHAR2(2000)  PATH '@RNoPart',
                                           name2        VARCHAR2(4000)  PATH 'LblTx',
                                           subsubitem   XMLTYPE         PATH 'BoQBody/Itemlist/Item'
                            ) xtd
            where   x.id = p_blob_id
            union all
            SELECT  v_kennung || '.' || xt.code||'.'||xtd.code2||'.'||xtd2.code3,xtd2.name3,v_kennung || '.' || xt.code||'.'||xtd.code2,xtd2.desc3,
                    case 
                        when xtd2.name4 like '%MLV%' or xtd2.name4 like '%MVL%' then v_kennung||substr(xtd2.name4,instr(xtd2.name4,'MLV')+instr(xtd2.name4,'MVL')+7,16)||' '||xtd2.name5
                        when xtd2.name5 like '%MLV%' or xtd2.name5 like '%MVL%' then v_kennung||substr(xtd2.name5,instr(xtd2.name5,'MLV')+instr(xtd2.name5,'MVL')+7,16)
                        else v_kennung||'_'||xt.code||xtd.code2||xtd2.code3 
                    end code,
                    xtd2.einheit
            FROM    pd_import_x83 x,
                    XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/200407' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                                PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                COLUMNS
                                   code     VARCHAR2(2000)  PATH '@RNoPart',
                                   name     VARCHAR2(4000)  PATH 'n:LblTx',
                                   subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                            ) xt,
                    xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/200407'),'/BoQCtgy'
                               PASSING xt.subitem
                               COLUMNS 
                                       code2        VARCHAR2(2000)  PATH '@RNoPart',
                                       name2        VARCHAR2(4000)  PATH 'LblTx',
                                       subsubitem   XMLTYPE         PATH 'BoQBody/Itemlist/Item'
                            ) xtd,
                    xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/200407'),'/Item'
                               PASSING xtd.subsubitem
                               COLUMNS 
                                       code3        VARCHAR2(200)  PATH '@RNoPart',
                                       name3        VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt',
                                       name4        VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[1]',
                                       name5        VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[2]',
                                       desc3        VARCHAR2(4000) PATH 'Description/CompleteText/DetailTxt',
                                       einheit      VARCHAR2(20)   PATH 'QU'
                            ) xtd2
            where x.id = p_blob_id
            )
    loop
        MERGE INTO PD_MUSTER_LVS e
        USING ( SELECT  i.code id,p_typ_id typ,i.mparent mojp FROM dual) h
        ON (e.code = h.id and e.MUSTER_TYP_ID = h.typ and (e.PARENT_ID = h.mojp or (e.PARENT_ID is null and h.mojp is null)))
        WHEN MATCHED THEN
            UPDATE 
            SET     e.name = i.name,
                    e.DESCRIPTION = i.description,
                    e.position_kennung = trim(i.kennung),
                    e.einheit = i.einheit,
                    e.stand_datum = p_date
        WHEN NOT MATCHED THEN
            INSERT (parent_id, code,name,description,muster_typ_id,position_kennung,STAND_DATUM,einheit)
            VALUES (i.mparent, i.code,i.name,i.description,p_typ_id,trim(i.kennung),p_date,i.einheit);
    end loop;

    delete from pd_import_x83 where id = p_blob_id;
  exception when others then
        DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.IMPORT_MUSTER: Fehler bei import: ' || SQLCODE || ': ' || SQLERRM || ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'IMPORT');
        delete from pd_import_x83 where id = p_blob_id;
        raise_application_error(-20000,SQLERRM);
  END Import_Muster1;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

  PROCEDURE Import_Muster2(p_blob_id number, p_typ_id number,p_date date) AS

    v_kennung   pd_import_typ.kennung%type;
    v_name      pd_import_typ.name%type;

  BEGIN

    select  kennung,name
    into    v_kennung,v_name
    from    pd_import_typ
    where   id = p_typ_id;

    for i in 
            (
            SELECT  v_kennung code,v_name name,null mparent,null description,v_kennung kennung,null einheit
            FROM    pd_import_x83 x
            where   x.id = p_blob_id
            union all    
            SELECT  v_kennung || '.' || xt.code,xt.name,v_kennung mparent,null description,v_kennung kennung,null einheit
            FROM    pd_import_x83 x,
                    XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/DA83/3.2' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                                PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                COLUMNS
                                   code     VARCHAR2(2000)  PATH '@RNoPart',
                                   name     VARCHAR2(4000)  PATH 'n:LblTx',
                                   subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                            ) xt
            where   x.id = p_blob_id
            union all
            SELECT  v_kennung || '.' || xt.code||'.'||xtd.code2,xtd.name2,v_kennung || '.' || xt.code,null description,null,null
            FROM    pd_import_x83 x,
                    XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/DA83/3.2' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                                PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                COLUMNS
                                    code     VARCHAR2(2000)  PATH '@RNoPart',
                                    name     VARCHAR2(4000)  PATH 'n:LblTx',
                                    subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                            ) xt,
                    xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/DA83/3.2'),'/BoQCtgy'
                                PASSING xt.subitem
                                COLUMNS 
                                    code2        VARCHAR2(2000)  PATH '@RNoPart',
                                    name2        VARCHAR2(4000)  PATH 'LblTx',
                                    subsubitem   XMLTYPE         PATH 'BoQBody/Itemlist/Item'
                            ) xtd
            where   x.id = p_blob_id
            union all
            SELECT  v_kennung || '.' || xt.code||'.'||xtd.code2||'.'||xtd2.code3,xtd2.name3,v_kennung || '.' || xt.code||'.'||xtd.code2,xtd2.desc3,
                    case 
                        when xtd2.name4 like '%MLV%' or xtd2.name4 like '%MVL%' then v_kennung||substr(xtd2.name4,instr(xtd2.name4,'MLV')+instr(xtd2.name4,'MVL')+7,16)||' '||xtd2.name5
                        when xtd2.name5 like '%MLV%' or xtd2.name5 like '%MVL%' then v_kennung||substr(xtd2.name5,instr(xtd2.name5,'MLV')+instr(xtd2.name5,'MVL')+7,16)
                        else v_kennung||'_'||xt.code||xtd.code2||xtd2.code3 
                    end code,
                    xtd2.einheit
            FROM    pd_import_x83 x,
                    XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/DA83/3.2' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                                PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                COLUMNS
                                    code     VARCHAR2(2000)  PATH '@RNoPart',
                                    name     VARCHAR2(4000)  PATH 'n:LblTx',
                                    subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                            ) xt,
                    xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/DA83/3.2'),'/BoQCtgy'
                                PASSING xt.subitem
                                    COLUMNS 
                                        code2        VARCHAR2(2000)  PATH '@RNoPart',
                                        name2        VARCHAR2(4000)  PATH 'LblTx',
                                        subsubitem   XMLTYPE         PATH 'BoQBody/Itemlist/Item'
                            ) xtd,
                    xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/DA83/3.2'),'/Item'
                                PASSING xtd.subsubitem
                                    COLUMNS 
                                        code3        VARCHAR2(200)  PATH '@RNoPart',
                                        name3        VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt',
                                        name4        VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[1]',
                                        name5        VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[2]',
                                        desc3        VARCHAR2(4000) PATH 'Description/CompleteText/DetailTxt',
                                        einheit      VARCHAR2(20)   PATH 'QU'
                            ) xtd2
            where x.id = p_blob_id
            )
    loop
        MERGE INTO PD_MUSTER_LVS e
        USING (SELECT i.code id,p_typ_id typ,i.mparent mojp FROM dual) h
        ON (e.code = h.id and e.MUSTER_TYP_ID = h.typ and (e.PARENT_ID = h.mojp or (e.PARENT_ID is null and h.mojp is null)))
        WHEN MATCHED THEN
            UPDATE 
            SET     e.name = i.name,
                    e.DESCRIPTION = i.description,
                    e.position_kennung = trim(i.kennung),
                    e.einheit = i.einheit,
                    e.stand_datum = p_date
        WHEN NOT MATCHED THEN
            INSERT (parent_id, code,name,description,muster_typ_id,position_kennung,STAND_DATUM,einheit)
            VALUES (i.mparent, i.code,i.name,i.description,p_typ_id,trim(i.kennung),p_date,i.einheit);
    end loop;

    delete from pd_import_x83 where id = p_blob_id;
  exception when others then
        DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.IMPORT_MUSTER: Fehler bei import: ' || SQLCODE || ': ' || SQLERRM ||
      ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'IMPORT');
      delete from pd_import_x83 where id = p_blob_id;
      raise_application_error(-20000,SQLERRM);
  END Import_Muster2;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

  PROCEDURE import_auftrege(p_blob_id number, p_typ_id number,p_region_id number,p_out out varchar2) AS
    v_namespace number;
  BEGIN

    select GetNamespace(p_blob_id) into v_namespace from dual;

    if v_namespace = 1 then
        import_auftrege1(p_blob_id,p_typ_id,p_region_id,p_out);
    end if;

    if v_namespace = 2 then
        import_auftrege2(p_blob_id,p_typ_id,p_region_id,p_out);
    end if;

    -- CW 12.03.2025: Fehlermeldung um XML-Namespace GAEB_DA_XML/DA83/3.3 abzufangen, der bei den Aufträgen noch nicht implementiert ist
    if v_namespace = 3 then
        DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.IMPORT_AUFTRAEGE: Fehler bei Import: Dateiformat nicht unterstuetzt !','IMPORT');
        delete from pd_import_x86 where id = p_blob_id;
        raise_application_error(-20000,'Fehler bei Import: Dateiformat \GAEB_DA_XML\DA83\3.3 nicht unterstuetzt !');
    end if;

  END import_auftrege;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

  PROCEDURE import_auftrege1(p_blob_id number, p_typ_id number,p_region_id number,p_out out varchar2) AS

    v_kennung               varchar2(10);
    v_auftrag_id            number;
    v_umzetzung_code        varchar2(20);
    v_datum                 date;
    v_check                 number;
    v_einlesung_status      varchar2(1):= 'Y';
    v_message               varchar2(4000);

  BEGIN

    for i in (  SELECT  xt.name,xt.strasse,xt.postleitzahl,xt.Stadt,xt.Telefon,xt.Fax,xt.Name2,xt.Projekt_name || '-' || xt.lv_name projekt_name,xt.Projekt_desc,xt.sap_nr,xt.vertrag_nr,
                        to_date(xt.Projekt_date,'YYYY-MM-DD') Projekt_date, to_number(replace(xt.total,'.',',')) total,nvl(to_char(to_number(xt.KEDITOREN_NUMMER)),'-') KEDITOREN_NUMMER
                FROM    pd_import_x86 x,
                        XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/200407' AS "n"),'/n:GAEB'
                                    PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                    COLUMNS
                                        name          VARCHAR2(100)  PATH 'n:Award/n:CTR/n:Address/n:Name1',
                                        name2         VARCHAR2(100)  PATH 'n:Award/n:CTR/n:Address/n:Name2',
                                        strasse       VARCHAR2(100)  PATH 'n:Award/n:CTR/n:Address/n:Street',
                                        postleitzahl  VARCHAR2(10)   PATH 'n:Award/n:CTR/n:Address/n:PCode',
                                        Stadt         VARCHAR2(50)   PATH 'n:Award/n:CTR/n:Address/n:City',
                                        Telefon       VARCHAR2(20)   PATH 'n:Award/n:CTR/n:Address/n:Phone',
                                        Fax           VARCHAR2(20)   PATH 'n:Award/n:CTR/n:Address/n:Fax',
                                        Projekt_name  VARCHAR2(100)  PATH 'n:PrjInfo/n:NamePrj',
                                        Projekt_desc  VARCHAR2(100)  PATH 'n:PrjInfo/n:LblPrj',
                                        Projekt_date  VARCHAR2(20)   PATH 'n:Award/n:AwardInfo/n:ContrDate',
                                        SAP_nr        VARCHAR2(20)   PATH 'n:Award/n:AwardInfo/n:ContrNo',
                                        vertrag_nr    VARCHAR2(20)   PATH 'n:Award/n:OWN/n:AwardNo',
                                        total         VARCHAR2(20)   PATH 'n:Award/n:BoQ/n:BoQInfo/n:Totals/n:Total',
                                        KEDITOREN_NUMMER VARCHAR2(20)PATH 'n:Award/n:CTR/n:AcctsPayNo',
                                        lv_name       VARCHAR2(100)  PATH 'n:Award/n:BoQ/n:BoQInfo/n:Name') xt
                where   x.id = p_blob_id
              )
    loop

        begin
            select  distinct 1
            into    v_check
            from    PD_AUFTRAEGE
            where   sap_nr = i.sap_nr
            and     vertrag_nr = i.vertrag_nr;
        exception when no_data_found then
            v_check:=0;
        end;

        if v_check > 0  then
            v_einlesung_status := 'N';
            v_message:= v_message||'Der Vertrag mit '||case when i.KEDITOREN_NUMMER = '-' then ' leerem Kreditoren Nummer, ' else null end||
                                   ' SAP Nummer: '||i.sap_nr||' und Vertrag Nummer: '||i.vertrag_nr||' wurde schon eingelesen'||chr(10);

        elsif i.KEDITOREN_NUMMER = '-' or i.sap_nr is null or i.vertrag_nr is null then
            v_message:= v_message||'Der Vertrag mit '||case when i.KEDITOREN_NUMMER = '-' then ' leerem Kreditoren Nummer, ' else null end||
                                                       case when i.sap_nr is null then ' leerem SAP Nummer ' else  'SAP Nummer: '||i.sap_nr end||
                                                       case when i.vertrag_nr is null then 'und mit leerem Vertrag Nummer' else ' und Vertrag Nummer: '||i.vertrag_nr end||
                                                       ' wurde erfolgreich eingelesen'||chr(10);
        else
            v_message:= v_message||'Der Vertrag mit SAP Nummer: '||i.sap_nr||' und Vertrag Nummer: '||i.vertrag_nr||' wurde erfolgreich eingelesen'||chr(10);
        end if;

        MERGE INTO PD_AUFTRAEGE e
        USING (SELECT i.name AUFTRAGNAHMER_NAME, i.Projekt_name PROJEKT_NAME FROM dual) h
        ON (lower(e.AUFTRAGNAHMER_NAME) = lower(h.AUFTRAGNAHMER_NAME) and lower(e.PROJEKT_NAME) = lower(h.PROJEKT_NAME))
        WHEN MATCHED THEN
        UPDATE SET e.strasse = nvl(i.strasse,strasse),
                   e.postleitzahl = nvl(i.postleitzahl,postleitzahl),
                   e.Stadt = nvl(i.Stadt,Stadt),
                   e.Telefon = nvl(i.Telefon,Telefon),
                   e.fax = nvl(i.fax,fax),
                   e.AUFTRAGNAHMER_NAME2 = nvl(i.name2,AUFTRAGNAHMER_NAME2),
                   e.PROJEKT_DESC = nvl(i.Projekt_desc,PROJEKT_DESC),
                   e.REGIONALBEREICH_ID = nvl(p_region_id,REGIONALBEREICH_ID),
                   e.datum = nvl(i.Projekt_date,datum),
                   e.total = nvl(i.total,total),
                   e.KEDITOREN_NUMMER = nvl(i.KEDITOREN_NUMMER,KEDITOREN_NUMMER),
                   e.sap_nr = nvl(i.sap_nr,sap_nr),
                   e.vertrag_nr = nvl(i.vertrag_nr,vertrag_nr),
                   e.DATUM_EINLESUNG = sysdate
        WHEN NOT MATCHED THEN
            INSERT (AUFTRAGNAHMER_NAME,strasse,postleitzahl,stadt,telefon,fax,AUFTRAGNAHMER_NAME2,PROJEKT_NAME,PROJEKT_DESC,REGIONALBEREICH_ID,DATUM,TOTAL,KEDITOREN_NUMMER,SAP_NR,VERTRAG_NR,EINLESUNG_STATUS,DATUM_EINLESUNG)
            VALUES (i.name, i.strasse,i.postleitzahl,i.stadt,i.telefon,i.fax,i.name2,i.Projekt_name,i.Projekt_desc,p_region_id,i.Projekt_date,i.total,i.KEDITOREN_NUMMER,i.sap_nr,i.vertrag_nr,v_einlesung_status,sysdate);

        select  id,datum
        into    v_auftrag_id,v_datum
        from    PD_AUFTRAEGE
        where   lower(AUFTRAGNAHMER_NAME) = lower(i.name) and lower(PROJEKT_NAME) = lower(i.Projekt_name);
    end loop;

    for j in  ( SELECT  xtd2.name2 name,
                        case 
                            when xtd2.name5 like 'MLV%' then trim(xtd2.name5)
                            when xtd2.name4 like 'MLV%' then trim(xtd2.name4)
                            when xtd2.name3 like 'MLV%' then trim(xtd2.name3)
                            when xtd2.name2 like '%MLV-%' or xtd2.name2 like '%MVL-%'then substr(replace(xtd2.name2,'MVL-','MLV-'),instr(replace(xtd2.name2,'MVL-','MLV-'),'MLV'),16)
                            else xt.code||'.'||xtd.code2||'.'||xtd2.code3
                        end code,
                        xtd2.description,
                        to_number(replace(xtd2.menge,'.',',')) menge,
                        xtd2.ME,to_number(replace(xtd2.Einheitspreis,'.',',')) Einheitspreis,
                        to_number(replace(xtd2.Gesamtbetrag,'.',',')) Gesamtbetrag,
                        xt.code||'.'||xtd.code2||'.'||xtd2.code3 position
                FROM   pd_import_x86 x,
                XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/200407' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                            PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                            COLUMNS
                                code     VARCHAR2(2000)  PATH '@RNoPart',
                                name     VARCHAR2(4000)  PATH 'n:LblTx',
                                subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                        ) xt,
               xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/200407'),'/BoQCtgy'
                            PASSING xt.subitem
                            COLUMNS 
                                code2        VARCHAR2(2000)  PATH '@RNoPart',
                                name2        VARCHAR2(4000)  PATH 'LblTx',
                                subsubitem   XMLTYPE         PATH 'BoQBody/Itemlist/Item') xtd,
               xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/200407'),'/Item'
                            PASSING xtd.subsubitem
                            COLUMNS 
                                code3            VARCHAR2(200)  PATH '@RNoPart',
                                name2            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[1]',
                                name3            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[2]',
                                name4            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[3]',
                                name5            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[4]',
                                description      VARCHAR2(4000) PATH 'substring(Description/CompleteText/DetailTxt,1,4000)',
                                Menge            VARCHAR2(20)   PATH 'Qty',
                                ME               VARCHAR2(10)   PATH 'QU',
                                Einheitspreis    VARCHAR2(20)   PATH 'UP',
                                Gesamtbetrag     VARCHAR2(20)   PATH 'IT'
                        ) xtd2
               where    x.id = p_blob_id)
    loop

        begin
            select  NEW_KENNUNG
            into    v_umzetzung_code
            from    pd_muster_umsetzung
            where   replace(replace(replace(OLD_KENNUNG,'_'),' '),'-') = replace(replace(replace(j.code,'_'),' '),'-');
        exception 
            when no_data_found
                then v_umzetzung_code:=null;
        end;

        begin
            select  1
            into    v_check
            from    PD_AUFTRAG_POSITIONEN
            where   auftrag_id = v_auftrag_id
            and     POSITION = j.position;
        exception
            when no_data_found
                then v_check:=0;
        end;

        if v_check = 0 then
            INSERT INTO PD_AUFTRAG_POSITIONEN(AUFTRAG_ID, NAME, CODE, BEZEICHNUNG, MENGE, MENGE_EINHEIT, EINHEITSPREIS, GESAMTBETRAG, UMZETZUNG_CODE,POSITION)
            VALUES (v_auftrag_id, j.name, j.code, j.description, j.menge, j.me, j.Einheitspreis, j.gesamtbetrag, nvl(v_umzetzung_code,j.code),j.position);
        end if;


        IF j.code like 'MLV%' then
            MERGE INTO PD_AUFTRAG_LVS al
            USING (Select substr(nvl(v_umzetzung_code,j.code),1,7) as code, v_auftrag_id as auftrag from dual) h
            ON (al.LV_CODE = h.code and al.AUFTRAG_ID = h.auftrag)
            WHEN NOT MATCHED THEN
                INSERT (LV_CODE, AUFTRAG_ID) VALUES (h.code, h.auftrag);
        END IF;
    end loop;

--DBS_LOGGING.LOG_DEBUG_AT('PREISDATENBdelete','import');
    delete from pd_import_x86 where id = p_blob_id;
    p_out := v_message;

    exception when others then
        DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.IMPORT_AUFTRAEGE: Fehler bei import: ' || SQLCODE || ': ' || SQLERRM ||' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'IMPORT');
        delete from pd_import_x86 where id = p_blob_id;
        raise_application_error(-20000,SQLERRM);
  END import_auftrege1;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

  PROCEDURE import_auftrege2(p_blob_id number, p_typ_id number,p_region_id number,p_out out varchar2) AS

    v_kennung               varchar2(10);
    v_auftrag_id            number;
    v_umzetzung_code        varchar2(20);
    v_datum                 date;
    v_check                 number;
    v_einlesung_status      varchar2(1):= 'Y';
    v_message               varchar2(4000);

  BEGIN

    --DBS_LOGGING.LOG_INFO_AT('PREISDATENBANK_PKG.IMPORT_AUFTRAEGE 1: start import_auftrege2','IMPORT');

    for i in (  SELECT  xt.name,xt.strasse,xt.postleitzahl,xt.Stadt,xt.Telefon,xt.Fax,xt.Name2,xt.Projekt_name || '-' || xt.lv_name projekt_name,xt.Projekt_desc,xt.sap_nr,xt.vertrag_nr,
                        to_date(xt.Projekt_date,'YYYY-MM-DD') Projekt_date, to_number(replace(xt.total,'.',',')) total,nvl(to_char(to_number(xt.KEDITOREN_NUMMER)),'-') KEDITOREN_NUMMER
                FROM   pd_import_x86 x,
                       XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/DA83/3.2' AS "n"),'/n:GAEB'
                                    PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                    COLUMNS
                                        name          VARCHAR2(100)  PATH 'n:Award/n:CTR/n:Address/n:Name1',
                                        name2         VARCHAR2(100)  PATH 'n:Award/n:CTR/n:Address/n:Name2',
                                        strasse       VARCHAR2(100)  PATH 'n:Award/n:CTR/n:Address/n:Street',
                                        postleitzahl  VARCHAR2(10)   PATH 'n:Award/n:CTR/n:Address/n:PCode',
                                        Stadt         VARCHAR2(50)   PATH 'n:Award/n:CTR/n:Address/n:City',
                                        Telefon       VARCHAR2(20)   PATH 'n:Award/n:CTR/n:Address/n:Phone',
                                        Fax           VARCHAR2(20)   PATH 'n:Award/n:CTR/n:Address/n:Fax',
                                        Projekt_name  VARCHAR2(100)  PATH 'n:PrjInfo/n:NamePrj',
                                        Projekt_desc  VARCHAR2(100)  PATH 'n:PrjInfo/n:LblPrj',
                                        Projekt_date  VARCHAR2(20)   PATH 'n:Award/n:AwardInfo/n:ContrDate',
                                        SAP_nr        VARCHAR2(20)   PATH 'n:Award/n:AwardInfo/n:ContrNo',
                                        vertrag_nr    VARCHAR2(20)   PATH 'n:Award/n:OWN/n:AwardNo',
                                        total         VARCHAR2(20)   PATH 'n:Award/n:BoQ/n:BoQInfo/n:Totals/n:Total',
                                        KEDITOREN_NUMMER VARCHAR2(20)PATH 'n:Award/n:CTR/n:AcctsPayNo',
                                        lv_name       VARCHAR2(100)  PATH 'n:Award/n:BoQ/n:BoQInfo/n:Name') xt
                where   x.id = p_blob_id)
    loop
        begin
            select  distinct 1
            into    v_check
            from    PD_AUFTRAEGE
            where   sap_nr = i.sap_nr
            and     vertrag_nr = i.vertrag_nr;
        exception
            when no_data_found then
                v_check:=0;
        end;

        --DBS_LOGGING.LOG_INFO_AT('PREISDATENBANK_PKG.IMPORT_AUFTRAEGE 2: v_check'||v_check,'IMPORT');

        if v_check > 0  then
            v_einlesung_status := 'N';
            v_message:= v_message||'Der Vertrag mit '||case when i.KEDITOREN_NUMMER = '-' then ' leerem Kreditoren Nummer, ' else null end||
                                   ' SAP Nummer: '||i.sap_nr||' und Vertrag Nummer: '||i.vertrag_nr||' wurde schon eingelesen'||chr(10);

        elsif i.KEDITOREN_NUMMER = '-' or i.sap_nr is null or i.vertrag_nr is null then
            v_message:= v_message||'Der Vertrag mit '||case when i.KEDITOREN_NUMMER = '-' then ' leerem Kreditoren Nummer, ' else null end||
                                                       case when i.sap_nr is null then ' leerem SAP Nummer ' else  'SAP Nummer: '||i.sap_nr end||
                                                       case when i.vertrag_nr is null then 'und mit leerem Vertrag Nummer' else ' und Vertrag Nummer: '||i.vertrag_nr end||
                                                       ' wurde erfolgreich eingelesen'||chr(10);
        else
            v_message:= v_message||'Der Vertrag mit SAP Nummer: '||i.sap_nr||' und Vertrag Nummer: '||i.vertrag_nr||' wurde erfolgreich eingelesen'||chr(10);
        end if;

        --DBS_LOGGING.LOG_INFO_AT('PREISDATENBANK_PKG.IMPORT_AUFTRAEGE 3: v_message:'||v_message,'IMPORT');

        MERGE INTO PD_AUFTRAEGE e
        USING (SELECT i.name AUFTRAGNAHMER_NAME, i.Projekt_name PROJEKT_NAME FROM dual) h
        ON (lower(e.AUFTRAGNAHMER_NAME) = lower(h.AUFTRAGNAHMER_NAME) and lower(e.PROJEKT_NAME) = lower(h.PROJEKT_NAME))
        WHEN MATCHED THEN
        UPDATE SET e.strasse = nvl(i.strasse,strasse),
                   e.postleitzahl = nvl(i.postleitzahl,postleitzahl),
                   e.Stadt = nvl(i.Stadt,Stadt),
                   e.Telefon = nvl(i.Telefon,Telefon),
                   e.fax = nvl(i.fax,fax),
                   e.AUFTRAGNAHMER_NAME2 = nvl(i.name2,AUFTRAGNAHMER_NAME2),
                   e.PROJEKT_DESC = nvl(i.Projekt_desc,PROJEKT_DESC),
                   e.REGIONALBEREICH_ID = nvl(p_region_id,REGIONALBEREICH_ID),
                   e.datum = nvl(i.Projekt_date,datum),
                   e.total = nvl(i.total,total),
                   e.KEDITOREN_NUMMER = nvl(i.KEDITOREN_NUMMER,KEDITOREN_NUMMER),
                   e.sap_nr = nvl(i.sap_nr,sap_nr),
                   e.vertrag_nr = nvl(i.vertrag_nr,vertrag_nr),
                   e.DATUM_EINLESUNG = sysdate
        WHEN NOT MATCHED THEN
            INSERT (AUFTRAGNAHMER_NAME,strasse,postleitzahl,stadt,telefon,fax,AUFTRAGNAHMER_NAME2,PROJEKT_NAME,PROJEKT_DESC,REGIONALBEREICH_ID,DATUM,TOTAL,KEDITOREN_NUMMER,SAP_NR,VERTRAG_NR,EINLESUNG_STATUS,DATUM_EINLESUNG)
            VALUES (i.name, i.strasse,i.postleitzahl,i.stadt,i.telefon,i.fax,i.name2,i.Projekt_name,i.Projekt_desc,p_region_id,i.Projekt_date,i.total,i.KEDITOREN_NUMMER,i.sap_nr,i.vertrag_nr,v_einlesung_status,sysdate);

        select  id,datum
        into    v_auftrag_id,v_datum
        from    PD_AUFTRAEGE
        where   lower(AUFTRAGNAHMER_NAME) = lower(i.name) and lower(PROJEKT_NAME) = lower(i.Projekt_name);
    end loop;

    for j in  ( SELECT  xtd2.name2 name,
                        case
                            when xtd2.name5 like 'MLV%' then trim(xtd2.name5)
                            when xtd2.name4 like 'MLV%' then trim(xtd2.name4)
                            when xtd2.name3 like 'MLV%' then trim(xtd2.name3)
                            when xtd2.name2 like '%MLV-%' or xtd2.name2 like '%MVL-%'then substr(replace(xtd2.name2,'MVL-','MLV-'),instr(replace(xtd2.name2,'MVL-','MLV-'),'MLV'),16)
                            else xt.code||'.'||xtd.code2||'.'||xtd2.code3
                        end code,
                        xtd2.description,
                        to_number(replace(xtd2.menge,'.',',')) menge,
                        xtd2.ME,to_number(replace(xtd2.Einheitspreis,'.',',')) Einheitspreis,
                        to_number(replace(xtd2.Gesamtbetrag,'.',',')) Gesamtbetrag,
                        xt.code||'.'||xtd.code2||'.'||xtd2.code3 position
               FROM     pd_import_x86 x,
                        XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/DA83/3.2' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                                    PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                    COLUMNS
                                        code     VARCHAR2(2000)  PATH '@RNoPart',
                                        name     VARCHAR2(4000)  PATH 'n:LblTx',
                                        subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                                ) xt,
                        xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/DA83/3.2'),'/BoQCtgy'
                                    PASSING xt.subitem
                                    COLUMNS 
                                        code2        VARCHAR2(2000)  PATH '@RNoPart',
                                        name2        VARCHAR2(4000)  PATH 'LblTx',
                                        subsubitem   XMLTYPE         PATH 'BoQBody/Itemlist/Item'
                                ) xtd,
                        xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/DA83/3.2'),'/Item'
                                    PASSING xtd.subsubitem
                                    COLUMNS 
                                        code3            VARCHAR2(200)  PATH '@RNoPart',
                                        name2            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[1]',
                                        name3            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[2]',
                                        name4            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[3]',
                                        name5            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[4]',
                                        description      VARCHAR2(4000) PATH 'substring(Description/CompleteText/DetailTxt,1,4000)',
                                        Menge            VARCHAR2(20)   PATH 'Qty',
                                        ME               VARCHAR2(10)   PATH 'QU',
                                        Einheitspreis    VARCHAR2(20)   PATH 'UP',
                                        Gesamtbetrag     VARCHAR2(20)   PATH 'IT'
                                ) xtd2
                where   x.id = p_blob_id)
    loop

        begin
            select  NEW_KENNUNG
            into    v_umzetzung_code
            from    pd_muster_umsetzung
            where   replace(replace(replace(OLD_KENNUNG,'_'),' '),'-') = replace(replace(replace(j.code,'_'),' '),'-');
        exception
            when no_data_found
                then v_umzetzung_code:=null;
        end;

        begin
            select  1
            into    v_check
            from    PD_AUFTRAG_POSITIONEN
            where   auftrag_id = v_auftrag_id
            and     POSITION = j.position;
        exception
            when no_data_found
                then v_check:=0;
        end;

        if v_check = 0 then
            INSERT INTO PD_AUFTRAG_POSITIONEN(AUFTRAG_ID, NAME, CODE, BEZEICHNUNG, MENGE, MENGE_EINHEIT, EINHEITSPREIS, GESAMTBETRAG, UMZETZUNG_CODE,POSITION)
            VALUES (v_auftrag_id, j.name, j.code, j.description, j.menge, j.me, j.Einheitspreis, j.gesamtbetrag, nvl(v_umzetzung_code,j.code),j.position);
        end if;


        IF j.code like 'MLV%' then
            MERGE INTO PD_AUFTRAG_LVS al
            USING (Select substr(nvl(v_umzetzung_code,j.code),1,7) as code, v_auftrag_id as auftrag from dual) h
            ON (al.LV_CODE = h.code and al.AUFTRAG_ID = h.auftrag)
            WHEN NOT MATCHED THEN
                INSERT (LV_CODE, AUFTRAG_ID) VALUES (h.code, h.auftrag);
        END IF;
    end loop;

    delete from pd_import_x86 where id = p_blob_id;
    p_out := v_message;


    --DBS_LOGGING.LOG_INFO_AT('PREISDATENBANK_PKG.IMPORT_AUFTRAEGE: v_message' || v_message ,'IMPORT');

    exception when others then
        DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.IMPORT_AUFTRAEGE: Fehler bei import: ' || SQLCODE || ': ' || SQLERRM ||' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'IMPORT');
        delete from pd_import_x86 where id = p_blob_id;
        raise_application_error(-20000,SQLERRM);
  END import_auftrege2;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

  PROCEDURE import_auftrege_n(p_blob_id number, p_typ_id number,p_region_id number,p_out out varchar2) AS

    v_kennung               varchar2(10);
    v_auftrag_id            number;
    v_umzetzung_code        varchar2(20);
    v_datum                 date;
    v_check                 number;
    v_einlesung_status      varchar2(1):= 'Y';
    v_message               varchar2(4000);

  BEGIN

    for i in (  SELECT  xt.name,xt.strasse,xt.postleitzahl,xt.Stadt,xt.Telefon,xt.Fax,xt.Name2,xt.Projekt_name || '-' || xt.lv_name projekt_name,xt.Projekt_desc,xt.sap_nr,xt.vertrag_nr,
                        to_date(xt.Projekt_date,'YYYY-MM-DD') Projekt_date, to_number(replace(xt.total,'.',',')) total,nvl(to_char(to_number(xt.KEDITOREN_NUMMER)),'-') KEDITOREN_NUMMER
                FROM   pd_import_x86 x,
                       XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/200407' AS "n"),'/n:GAEB'
                                    PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                    COLUMNS
                                        name          VARCHAR2(100)  PATH 'n:Award/n:CTR/n:Address/n:Name1',
                                        name2         VARCHAR2(100)  PATH 'n:Award/n:CTR/n:Address/n:Name2',
                                        strasse       VARCHAR2(100)  PATH 'n:Award/n:CTR/n:Address/n:Street',
                                        postleitzahl  VARCHAR2(10)   PATH 'n:Award/n:CTR/n:Address/n:PCode',
                                        Stadt         VARCHAR2(50)   PATH 'n:Award/n:CTR/n:Address/n:City',
                                        Telefon       VARCHAR2(20)   PATH 'n:Award/n:CTR/n:Address/n:Phone',
                                        Fax           VARCHAR2(20)   PATH 'n:Award/n:CTR/n:Address/n:Fax',
                                        Projekt_name  VARCHAR2(100)  PATH 'n:PrjInfo/n:NamePrj',
                                        Projekt_desc  VARCHAR2(100)  PATH 'n:PrjInfo/n:LblPrj',
                                        Projekt_date  VARCHAR2(20)   PATH 'n:Award/n:AwardInfo/n:ContrDate',
                                        SAP_nr        VARCHAR2(20)   PATH 'n:Award/n:AwardInfo/n:ContrNo',
                                        vertrag_nr    VARCHAR2(20)   PATH 'n:Award/n:OWN/n:AwardNo',
                                        total         VARCHAR2(20)   PATH 'n:Award/n:BoQ/n:BoQInfo/n:Totals/n:Total',
                                        KEDITOREN_NUMMER VARCHAR2(20)PATH 'n:Award/n:CTR/n:AcctsPayNo',
                                        lv_name       VARCHAR2(100)  PATH 'n:Award/n:BoQ/n:BoQInfo/n:Name') xt
                where   x.id = p_blob_id)
    loop
        begin
            select  distinct 1
            into    v_check
            from    PD_AUFTRAEGE
            where   sap_nr = i.sap_nr
            and     vertrag_nr = i.vertrag_nr;
        exception when no_data_found then
            v_check:=0;
        end;

        if v_check > 0  then
            v_einlesung_status := 'N';
            v_message:= v_message||'Der Vertrag mit '||case when i.KEDITOREN_NUMMER = '-' then ' leerem Kreditoren Nummer, ' else null end||
                                   ' SAP Nummer: '||i.sap_nr||' und Vertrag Nummer: '||i.vertrag_nr||' wurde schon eingelesen'||chr(10);

        elsif i.KEDITOREN_NUMMER = '-' or i.sap_nr is null or i.vertrag_nr is null then
            v_message:= v_message||'Der Vertrag mit '||case when i.KEDITOREN_NUMMER = '-' then ' leerem Kreditoren Nummer, ' else null end||
                                                       case when i.sap_nr is null then ' leerem SAP Nummer ' else  'SAP Nummer: '||i.sap_nr end||
                                                       case when i.vertrag_nr is null then 'und mit leerem Vertrag Nummer' else ' und Vertrag Nummer: '||i.vertrag_nr end||
                                                       ' wurde erfolgreich eingelesen'||chr(10);
        else
            v_message:= v_message||'Der Vertrag mit SAP Nummer: '||i.sap_nr||' und Vertrag Nummer: '||i.vertrag_nr||' wurde erfolgreich eingelesen'||chr(10);
        end if;

        MERGE INTO PD_AUFTRAEGE e
        USING (SELECT i.name AUFTRAGNAHMER_NAME, i.Projekt_name PROJEKT_NAME FROM dual) h
        ON (lower(e.AUFTRAGNAHMER_NAME) = lower(h.AUFTRAGNAHMER_NAME) and lower(e.PROJEKT_NAME) = lower(h.PROJEKT_NAME))
        WHEN MATCHED THEN
        UPDATE SET e.strasse = nvl(i.strasse,strasse),
                   e.postleitzahl = nvl(i.postleitzahl,postleitzahl),
                   e.Stadt = nvl(i.Stadt,Stadt),
                   e.Telefon = nvl(i.Telefon,Telefon),
                   e.fax = nvl(i.fax,fax),
                   e.AUFTRAGNAHMER_NAME2 = nvl(i.name2,AUFTRAGNAHMER_NAME2),
                   e.PROJEKT_DESC = nvl(i.Projekt_desc,PROJEKT_DESC),
                   e.REGIONALBEREICH_ID = nvl(p_region_id,REGIONALBEREICH_ID),
                   e.datum = nvl(i.Projekt_date,datum),
                   e.total = nvl(i.total,total),
                   e.KEDITOREN_NUMMER = nvl(i.KEDITOREN_NUMMER,KEDITOREN_NUMMER),
                   e.sap_nr = nvl(i.sap_nr,sap_nr),
                   e.vertrag_nr = nvl(i.vertrag_nr,vertrag_nr),
                   e.DATUM_EINLESUNG = sysdate
        WHEN NOT MATCHED THEN
            INSERT (AUFTRAGNAHMER_NAME,strasse,postleitzahl,stadt,telefon,fax,AUFTRAGNAHMER_NAME2,PROJEKT_NAME,PROJEKT_DESC,REGIONALBEREICH_ID,DATUM,TOTAL,KEDITOREN_NUMMER,SAP_NR,VERTRAG_NR,EINLESUNG_STATUS,DATUM_EINLESUNG)
            VALUES (i.name, i.strasse,i.postleitzahl,i.stadt,i.telefon,i.fax,i.name2,i.Projekt_name,i.Projekt_desc,p_region_id,i.Projekt_date,i.total,i.KEDITOREN_NUMMER,i.sap_nr,i.vertrag_nr,v_einlesung_status,sysdate);

        select  id,datum
        into    v_auftrag_id,v_datum
        from    PD_AUFTRAEGE
        where   lower(AUFTRAGNAHMER_NAME) = lower(i.name) and lower(PROJEKT_NAME) = lower(i.Projekt_name);
    end loop;

    for j in  ( SELECT  xtd2.name2 name,
                        case
                            when xtd2.name5 like 'MLV%' then trim(xtd2.name5)
                            when xtd2.name4 like 'MLV%' then trim(xtd2.name4)
                            when xtd2.name3 like 'MLV%' then trim(xtd2.name3)
                            when xtd2.name2 like '%MLV-%' or xtd2.name2 like '%MVL-%'then substr(replace(xtd2.name2,'MVL-','MLV-'),instr(replace(xtd2.name2,'MVL-','MLV-'),'MLV'),16)
                            else xt.code||'.'||xtd.code2||'.'||xtd2.code3
                        end code,
                        xtd2.description,
                        to_number(replace(xtd2.menge,'.',',')) menge,
                        xtd2.ME,to_number(replace(xtd2.Einheitspreis,'.',',')) Einheitspreis,
                        to_number(replace(xtd2.Gesamtbetrag,'.',',')) Gesamtbetrag,
                        xt.code||'.'||xtd.code2||'.'||xtd2.code3 position
               FROM     pd_import_x86 x,
                        XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/200407' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                            PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                            COLUMNS
                                code     VARCHAR2(2000)  PATH '@RNoPart',
                                name     VARCHAR2(4000)  PATH 'n:LblTx',
                                subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                        ) xt,
               xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/200407'),'/BoQCtgy'
                            PASSING xt.subitem
                            COLUMNS 
                                code2        VARCHAR2(2000)  PATH '@RNoPart',
                                name2        VARCHAR2(4000)  PATH 'LblTx',
                                subsubitem   XMLTYPE         PATH 'BoQBody/Itemlist/Item') xtd,
               xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/200407'),'/Item'
                            PASSING xtd.subsubitem
                            COLUMNS 
                                code3            VARCHAR2(200)  PATH '@RNoPart',
                                name2            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[1]',
                                name3            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[2]',
                                name4            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[3]',
                                name5            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[4]',
                                description      VARCHAR2(4000) PATH 'substring(Description/CompleteText/DetailTxt,1,4000)',
                                Menge            VARCHAR2(20)   PATH 'Qty',
                                ME               VARCHAR2(10)   PATH 'QU',
                                Einheitspreis    VARCHAR2(20)   PATH 'UP',
                                Gesamtbetrag     VARCHAR2(20)   PATH 'IT'
                        ) xtd2
               where   x.id = p_blob_id)
    loop

        begin
            select  NEW_KENNUNG
            into    v_umzetzung_code
            from    pd_muster_umsetzung
            where   replace(replace(replace(OLD_KENNUNG,'_'),' '),'-') = replace(replace(replace(j.code,'_'),' '),'-');
        exception
            when no_data_found
                then v_umzetzung_code:=null;
        end;

        begin
            select  1
            into    v_check
            from    PD_AUFTRAG_POSITIONEN
            where   auftrag_id = v_auftrag_id
            and     POSITION = j.position;
        exception
            when no_data_found
                then v_check:=0;
        end;

        if v_check = 0 then
            INSERT INTO PD_AUFTRAG_POSITIONEN(AUFTRAG_ID, NAME, CODE, BEZEICHNUNG, MENGE, MENGE_EINHEIT, EINHEITSPREIS, GESAMTBETRAG, UMZETZUNG_CODE,POSITION)
            VALUES (v_auftrag_id, j.name, j.code, j.description, j.menge, j.me, j.Einheitspreis, j.gesamtbetrag, nvl(v_umzetzung_code,j.code),j.position);
        end if;


        IF j.code like 'MLV%' then
            MERGE INTO PD_AUFTRAG_LVS al
            USING (Select substr(nvl(v_umzetzung_code,j.code),1,7) as code, v_auftrag_id as auftrag from dual) h
            ON (al.LV_CODE = h.code and al.AUFTRAG_ID = h.auftrag)
            WHEN NOT MATCHED THEN
                INSERT (LV_CODE, AUFTRAG_ID) VALUES (h.code, h.auftrag);
        END IF;
    end loop;



    delete from pd_import_x86 where id = p_blob_id;
    p_out := v_message;

    exception when others then
        DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.IMPORT_AUFTRAEGE: Fehler bei import: ' || SQLCODE || ': ' || SQLERRM ||' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'IMPORT');
        delete from pd_import_x86 where id = p_blob_id;
        raise_application_error(-20000,SQLERRM);
  END import_auftrege_n;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

  PROCEDURE auswertung(p_von date, p_bis date, p_lvs varchar2) as

       TYPE curtype IS REF CURSOR; 
       v_sql_Grundabgaben clob;
       src_cur    curtype;
       curid      NUMBER; 
       desctab    DBMS_SQL.desc_tab2;
       colcnt     NUMBER; 
       namevar    varchar2(4000);
       namevar2   varchar2(4000);
       numvar     NUMBER; 
       datevar    DATE; 
       v_columns  clob;
       v_columnname varchar2(4000);
       l_filename varchar2(100);
       L_BLOB blob;
       l_target_charset VARCHAR2(100) := 'WE8MSWIN1252';--'WE8ISO8859P15';
       L_DEST_OFFSET    INTEGER := 1;
       L_SRC_OFFSET     INTEGER := 1;
       L_LANG_CONTEXT   INTEGER := DBMS_LOB.DEFAULT_LANG_CTX;
       L_WARNING        INTEGER;
       L_LENGTH         INTEGER;

       v_Region clob:='"Region";';
       v_Bezeichnung clob:='"Bezeichnung der Maßnahme";';
       v_Auftragnahmer clob:='"Auftragnehmer";';
       v_LVDatum clob:='"LV-Datum";';
       v_Vergabesumme clob:='"Vergabesumme";';

  Begin

    v_sql_Grundabgaben := 'select r.code "Region", a.projekt_desc "Bezeichnung",a.auftragnahmer_name "Auftragnahmer",a.datum "LVDatum",a.total "Vergabesumme"
                         from pd_auftraege a
                         left join pd_region r on r.id = a.regionalbereich_id
                         where a.datum between to_date('''||p_von||''',''DD.MM.YYYY'') and to_date('''||p_bis||''',''DD.MM.YYYY'')';

    dbms_lob.createtemporary(lob_loc => l_blob, cache => true, dur => dbms_lob.call);

    OPEN src_cur FOR v_sql_Grundabgaben;            -- #DYNAMIC-SQL-IN-PLSQL-CHECKED#

    curid := DBMS_SQL.to_cursor_number (src_cur); 
    DBMS_SQL.describe_columns2 (curid, colcnt, desctab);

    FOR indx IN 1 .. colcnt LOOP 
        IF desctab (indx).col_type = 2 THEN 
            DBMS_SQL.define_column (curid, indx, numvar); 
        ELSIF desctab (indx).col_type = 12 THEN 
            DBMS_SQL.define_column (curid, indx, datevar); 
        ELSE 
            DBMS_SQL.define_column (curid, indx, namevar, 4000); 
        END IF; 
    END LOOP;

    WHILE DBMS_SQL.fetch_rows (curid) > 0 LOOP 
        FOR indx IN 1 .. colcnt LOOP 

            IF (upper(desctab (indx).col_name) = 'REGION') THEN 
                DBMS_SQL.COLUMN_VALUE (curid, indx, namevar);
                v_Region := v_Region||'"'||trim(namevar)||'";';
            ELSIF (upper(desctab (indx).col_name) = 'BEZEICHNUNG') THEN 
                DBMS_SQL.COLUMN_VALUE (curid, indx, namevar); 
                v_Bezeichnung := v_Bezeichnung||'"'||trim(namevar)||'";';
            ELSIF (upper(desctab (indx).col_name) = 'AUFTRAGNAHMER') THEN 
                DBMS_SQL.COLUMN_VALUE (curid, indx, namevar); 
                v_Auftragnahmer := v_Auftragnahmer||'"'||trim(namevar)||'";';
            ELSIF (upper(desctab (indx).col_name) = 'LVDATUM') THEN 
                DBMS_SQL.COLUMN_VALUE (curid, indx, datevar); 
                v_LVDatum := v_LVDatum||'"'||datevar||'";';
            ELSIF (upper(desctab (indx).col_name) = 'VERGABESUMME') THEN 
                DBMS_SQL.COLUMN_VALUE (curid, indx, numvar); 
                v_Vergabesumme := v_Vergabesumme||'"'||numvar||'";';
            END IF; 
          END LOOP;
    END LOOP;

    DBMS_SQL.close_cursor (curid);

    v_columns :=  SUBSTR(v_Region, 1, LENGTH(v_Region) - 1)||chr(13)||chr(10)||
                  SUBSTR(v_Bezeichnung, 1, LENGTH(v_Bezeichnung) - 1)||chr(13)||chr(10)||
                  SUBSTR(v_Auftragnahmer, 1, LENGTH(v_Auftragnahmer) - 1)||chr(13)||chr(10)||
                  SUBSTR(v_LVDatum, 1, LENGTH(v_LVDatum) - 1)||chr(13)||chr(10)||
                  SUBSTR(v_Vergabesumme, 1, LENGTH(v_Vergabesumme) - 1);

     dbms_lob.writeappend(l_blob,UTL_RAW.length(UTL_I18N.STRING_TO_RAW(v_columns||chr(13)||chr(10),l_target_charset)),UTL_I18N.STRING_TO_RAW(v_columns||chr(13)||chr(10),l_target_charset));
     v_columns:=null;

    -- determine length for header
     L_LENGTH := DBMS_LOB.GETLENGTH(L_BLOB);

    -- first clear the header
     HTP.FLUSH;
     HTP.INIT;

    -- create response header
     OWA_UTIL.MIME_HEADER( 'text/csv', FALSE);

     HTP.P('Content-length: ' || L_LENGTH);
     HTP.P('Content-Disposition: attachment; filename="export_my_table.csv"');
     HTP.P('Set-Cookie: fileDownload=true; path=/');

     OWA_UTIL.HTTP_HEADER_CLOSE;

     -- download the BLOB
     WPG_DOCLOAD.DOWNLOAD_FILE( L_BLOB );
    -- rest of the HTML does not render
     DBMS_LOB.FREETEMPORARY(L_BLOB);
    -- stop APEX
     --APEX_APPLICATION.STOP_APEX_ENGINE;
     htmldb_application.g_unrecoverable_error := true;

  exception
    when others then
        WPG_DOCLOAD.DOWNLOAD_FILE( L_BLOB );
        DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.AUSWERTUNG: Fehler bei auswertung: ' || SQLCODE || ': ' || SQLERRM ||
            ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'AUSWERTUNG');
  end;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

PROCEDURE export_ausschreibung_to_excel(p_blob_id number,p_typ_id number,p_region_id number,p_von date, p_bis date,p_regionen varchar2,p_liferant varchar2,p_user_id number) as 
  v_namespace number;
BEGIN
    select GetNamespace(p_blob_id) into v_namespace from dual;

    if v_namespace = 1 then
      export_ausschreibung_to_excel1(p_blob_id,p_typ_id,p_region_id,p_von,p_bis,p_regionen,p_liferant,p_user_id);
    end if;
    if v_namespace = 2 then
      export_ausschreibung_to_excel2(p_blob_id,p_typ_id,p_region_id,p_von,p_bis,p_regionen,p_liferant,p_user_id);
    end if;
    -- CW: nuer Dateityp: http://www.gaeb.de/GAEB_DA_XML/DA83/3.3
    if v_namespace = 3 then
      export_ausschreibung_to_excel3(p_blob_id,p_typ_id,p_region_id,p_von,p_bis,p_regionen,p_liferant,p_user_id);
    end if;
        --x82
    if v_namespace = 4 then
        DBS_LOGGING.LOG_INFO_AT ('X82','X82');
        null;
    end if;

end;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

PROCEDURE export_ausschreibung_to_excel1(p_blob_id number,p_typ_id number,p_region_id number,p_von date, p_bis date,p_regionen varchar2,p_liferant varchar2,p_user_id number) as 

       workbook xlsx_writer.book_r;
       sheet_1  integer;

       xlsx     blob;

       cs_border integer;
       cs_master integer;
       cs_master2 integer;
       cs_parent integer;
       number_format_child integer;
       number_format_parent integer;
       number_format_master integer;
       border_db_full integer;
       font_db  integer;
       fill_master integer;
       fill_parent integer;
       fill_master2 integer;

       c_limit   constant integer := 50;
       c_x_split constant integer := 3;
       c_y_split constant integer := 8;
       c_y_region constant integer := 1;

       TYPE curtype IS REF CURSOR;

       l_filename varchar2(100);
       L_BLOB blob;
       l_target_charset VARCHAR2(100) := 'WE8MSWIN1252';
       L_DEST_OFFSET    INTEGER := 1;
       L_SRC_OFFSET     INTEGER := 1;
       L_LANG_CONTEXT   INTEGER := DBMS_LOB.DEFAULT_LANG_CTX;
       L_WARNING        INTEGER;
       L_LENGTH         INTEGER;
       l_betrag         number;

       l_min number;
       l_avg number;
       l_max number;
       l_median number;
       l_count number;
       l_id  number;
       l_typ number;
       l_parent_id varchar2(20);
       l_master_id varchar2(20);
       l_row number:=2;
begin

    workbook := xlsx_writer.start_book;
    sheet_1  := xlsx_writer.add_sheet  (workbook, 'Ausschreibung LVS');

    font_db := xlsx_writer.add_font      (workbook, 'DB Office', 10);
    fill_master := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="92cddc"/><bgColor indexed="64"/></patternFill>');
    fill_master2 := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="92cddc"/><bgColor indexed="64"/></patternFill>');
    fill_parent := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="daeef3"/><bgColor indexed="64"/></patternFill>');
    border_db_full := xlsx_writer.add_border      (workbook, '<left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/>');
    cs_border := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, font_id => font_db);
    cs_master := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, fill_id => fill_master,font_id => font_db);
    cs_master2 := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, fill_id => fill_master2,font_id => font_db);
    cs_parent := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, fill_id => fill_parent,font_id => font_db);
    number_format_child := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00",font_id => font_db);
    number_format_parent := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00", font_id => font_db, fill_id => fill_parent);
    number_format_master := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00", font_id => font_db, fill_id => fill_master2);

    dbms_lob.createtemporary(lob_loc => l_blob, cache => true, dur => dbms_lob.call);

    xlsx_writer.col_width(workbook, sheet_1, 1, 75);
    xlsx_writer.col_width(workbook, sheet_1, 2, 20);
    xlsx_writer.col_width(workbook, sheet_1, 3, 20);
    xlsx_writer.col_width(workbook, sheet_1, 6, 20);
    xlsx_writer.col_width(workbook, sheet_1, 7, 20);
    xlsx_writer.col_width(workbook, sheet_1, 8, 20);
    xlsx_writer.col_width(workbook, sheet_1, 9, 20);
    xlsx_writer.col_width(workbook, sheet_1, 10, 20);
    xlsx_writer.col_width(workbook, sheet_1, 11, 20);
    xlsx_writer.col_width(workbook, sheet_1, 12, 20);
    xlsx_writer.col_width(workbook, sheet_1, 13, 20);

    xlsx_writer.add_cell(workbook, sheet_1, 1, 1,style_id => cs_master, text => 'NAME');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 2,style_id => cs_master, text => 'CODE');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 3,style_id => cs_master, text => 'KENNUNG');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 4,style_id => cs_master, text => 'EINHEIT');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 5,style_id => cs_master, text => 'MENGE');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 6,style_id => cs_master, text => 'MIN PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 7,style_id => cs_master, text => 'GESAMT MIN PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 8,style_id => cs_master, text => 'MITTELWERT');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 9,style_id => cs_master, text => 'GESAMT MITTELWERT');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 10,style_id => cs_master, text => 'MEDIAN PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 11,style_id => cs_master, text => 'GESAMT MEDIAN PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 12,style_id => cs_master, text => 'MAX PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 13,style_id => cs_master, text => 'GESAMT MAX PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 14,style_id => cs_master, text => 'ANZAHL VERGABEN');

    for i in  ( SELECT  xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,'') name,xt.code master,xt.code||'.'||xtd.code2 parent,xt.name master_name, xtd.name2 parent_name,
                        (case
                            when instr(replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_'),'MLV') = 0
                                then ''
                            when substr(replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_'),8,1) = '_'
                                then replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_')
                            else substr(replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_'),1,7) || '_' || substr(replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_'),8)                            
                        end) kennung,
                        xt.code||'.'||xtd.code2||'.'||xtd2.code3 code,
                        xtd2.description,
                        replace(xtd2.menge,'.',',') menge,
                        xtd2.ME,
                        replace(xtd2.Einheitspreis,'.',',') Einheitspreis,
                        replace(xtd2.Gesamtbetrag,'.',',') Gesamtbetrag
                FROM    pd_import_x86 x,
                XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/200407' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                            PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                            COLUMNS
                                code     VARCHAR2(2000)  PATH '@RNoPart',
                                name     VARCHAR2(4000)  PATH 'n:LblTx',
                                subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                        ) xt,
               xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/200407'),'/BoQCtgy'
                            PASSING xt.subitem
                            COLUMNS 
                                code2        VARCHAR2(2000)  PATH '@RNoPart',
                                name2        VARCHAR2(4000)  PATH 'LblTx',
                                subsubitem   XMLTYPE         PATH 'BoQBody/Itemlist/Item'
                        ) xtd,
               xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/200407'),'/Item'
                            PASSING xtd.subsubitem
                            COLUMNS 
                                code3            VARCHAR2(200)  PATH '@RNoPart',
                                name2            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[1]',
                                name3            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[2]',
                                name4            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[3]',
                                name5            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[4]',
                                description      VARCHAR2(4000) PATH 'substring(Description/CompleteText/DetailTxt,1,4000)',
                                Menge            VARCHAR2(20)   PATH 'Qty',
                                ME               VARCHAR2(10)   PATH 'QU',
                                Einheitspreis    VARCHAR2(20)   PATH 'UP',
                                Gesamtbetrag     VARCHAR2(20)   PATH 'IT'
                        ) xtd2
               where    x.id = p_blob_id
              )
    loop

       begin
            select  min(ap.EINHEITSPREIS),max(EINHEITSPREIS),avg(EINHEITSPREIS),median(EINHEITSPREIS),count(code)
            into    l_min,l_max,l_avg,l_median,l_count
            from    pd_auftrag_positionen ap
            join    pd_auftraege a on a.id = ap.auftrag_id
            where   replace(ap.UMSETZUNG_CODE,' ') = replace(replace(replace(i.kennung,'-'),'_'),' ')
            and     a.EINLESUNG_STATUS = 'Y'
            and     a.REGIONALBEREICH_ID in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
            and     a.KEDITOREN_NUMMER in (select  column_value from table(apex_string.split(p_liferant, ':')))
            and     ap.einheitspreis > 0;--13.09.2022,AR
       exception 
            when no_data_found then 
                l_min:=null;
                l_max:=null;
                l_avg:=null;
                l_median:=null;
                l_id :=null;
       end;

       INSERT INTO pd_ausschreibung
            (blob_id,muster_id,name,code,einheit,menge,min_preis,mittlerer_preis,max_preis,median_preis,parent_id,master_id,parent_name,master_name,kennung,code_count)
       VALUES
            (p_blob_id,l_id,i.name,i.code,i.me,to_number(replace(i.menge,',','.')),l_min,l_avg,l_max,l_median,i.parent,i.master,i.parent_name,i.master_name,i.kennung,l_count);

    end loop;--i in  (SELECT xtd2.name2...

    commit;

    for j in (with  mustern as 
                   (select  pa.parent_id,pa.code,pa.kennung,pa.name,pa.min_preis,pa.mittlerer_preis,pa.max_preis,pa.median_preis,pa.EINHEIT,pa.MENGE,'CHILD' art,
                            pa.min_preis*pa.MENGE gesamt_min,pa.mittlerer_preis * pa.MENGE gesammt_mittel,pa.max_preis * pa.MENGE gesamt_max,pa.median_preis * pa.MENGE gesamt_median,pa.code_count
                    from    pd_ausschreibung pa 
                    where   pa.blob_id = p_blob_id
                    union all
                    select  pa.master_id,pa.parent_id,null,pa.parent_name,sum(pa.min_preis),sum(pa.mittlerer_preis),sum(pa.max_preis),sum(pa.median_preis),null,null,'PARENT' art,
                            sum(pa.min_preis*pa.MENGE),sum(pa.mittlerer_preis*pa.MENGE),sum(pa.max_preis*pa.MENGE),sum(pa.median_preis*pa.MENGE),null
                    from    pd_ausschreibung pa 
                    where   pa.blob_id = p_blob_id
                    group by pa.parent_id,pa.master_id,pa.parent_name
                    union all 
                    select  null,pa.master_id,null,pa.master_name,sum(pa.min_preis),sum(pa.mittlerer_preis),sum(pa.max_preis),sum(pa.median_preis),null,null,'MASTER' art,
                            sum(pa.min_preis*pa.MENGE),sum(pa.mittlerer_preis*pa.MENGE),sum(pa.max_preis*pa.MENGE),sum(pa.median_preis*pa.MENGE),null
                    from    pd_ausschreibung pa 
                    where   pa.blob_id = p_blob_id
                    group by pa.master_id,pa.master_name
                   )
                SELECT  trim(m.code) code,
                        trim(m.name) name,
                        trim(m.kennung) kennung,
                        trim(m.einheit) einheit,
                        m.menge,
                        m.min_preis,
                        m.mittlerer_preis,
                        m.max_preis,
                        m.median_preis,
                        m.gesamt_min,
                        m.gesammt_mittel,
                        m.gesamt_max,
                        m.gesamt_median,
                        m.art,
                        m.code_count,
                        (case m.art
                            when 'CHILD' then cs_border
                            when 'PARENT' then cs_parent
                            else cs_master
                        end) cs_style,
                        (case m.art
                            when 'CHILD' then number_format_child
                            when 'PARENT' then number_format_parent
                            else number_format_master
                        end) number_format
              FROM    mustern m
              START WITH parent_id IS NULL
              CONNECT BY PRIOR m.code = m.parent_id 
              ORDER SIBLINGS BY m.code
             )
    loop

       xlsx_writer.add_cell(workbook, sheet_1, l_row, 1,style_id => j.cs_style, text => j.name);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 2,style_id => j.cs_style, text => j.code);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 3,style_id => j.cs_style, text => j.kennung);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 4,style_id => j.cs_style, text => j.einheit);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 5,style_id => j.cs_style, value_ => j.menge);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 6,style_id => j.number_format, value_ => j.min_preis);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 7,style_id => j.number_format, value_ => j.gesamt_min);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 8,style_id => j.number_format, value_ => j.mittlerer_preis);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 9,style_id => j.number_format, value_ => j.gesammt_mittel);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 10,style_id => j.number_format, value_ => j.median_preis);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 11,style_id => j.number_format, value_ => j.gesamt_median);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 12,style_id => j.number_format, value_ => j.max_preis);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 13,style_id => j.number_format, value_ => j.gesamt_max);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 14,style_id => j.cs_style, value_ => j.code_count);

       l_row := l_row + 1;
    end loop;--j in (with  mustern as (  select pa.parent_id...

    xlsx_writer.freeze_sheet(workbook, sheet_1,0,1);

    xlsx := xlsx_writer.create_xlsx(workbook);



    delete from pd_import_x86 where id = p_blob_id;
    delete from pd_ausschreibung where blob_id = p_blob_id;

    --Mailversand der Auswertung
    SendMailAuswertung(p_user_id => p_user_id,p_anhang => xlsx,p_filename => 'Ausschreibung_LVS.xlsx');

    -- rest of the HTML does not render
    DBMS_LOB.FREETEMPORARY(xlsx);

    exception 
        when others then
            DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.export_ausschreibung_to_excel1: Fehler bei auswertung: ' || SQLCODE || ': ' || SQLERRM ||
            ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'AUSSCHREIBUNG');

            SendMailAuswertungFehler
                (
                p_user_id => p_user_id,
                p_error => 'Fehler bei der Auswertung, bitte wenden Sie sich an den Anwendungsverantwortlichen.'
                );

end export_ausschreibung_to_excel1;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

PROCEDURE export_ausschreibung_to_excel2(p_blob_id number,p_typ_id number,p_region_id number,p_von date, p_bis date,p_regionen varchar2,p_liferant varchar2,p_user_id number) as 

       workbook xlsx_writer.book_r;
       sheet_1  integer;

       xlsx     blob;

       cs_border integer;
       cs_master integer;
       cs_master2 integer;
       cs_parent integer;
       number_format_child integer;
       number_format_parent integer;
       number_format_master integer;
       border_db_full integer;
       font_db  integer;
       fill_master integer;
       fill_parent integer;
       fill_master2 integer;

       c_limit   constant integer := 50;
       c_x_split constant integer := 3;
       c_y_split constant integer := 8;
       c_y_region constant integer := 1;

       TYPE curtype IS REF CURSOR;

       l_filename varchar2(100);
       L_BLOB blob;
       l_target_charset VARCHAR2(100) := 'WE8MSWIN1252';
       L_DEST_OFFSET    INTEGER := 1;
       L_SRC_OFFSET     INTEGER := 1;
       L_LANG_CONTEXT   INTEGER := DBMS_LOB.DEFAULT_LANG_CTX;
       L_WARNING        INTEGER;
       L_LENGTH         INTEGER;
       l_betrag         number;

       l_min number;
       l_avg number;
       l_max number;
       l_median number;
       l_count number;
       l_id  number;
       l_typ number;
       l_parent_id varchar2(20);
       l_master_id varchar2(20);
       l_row number:=2;
       v_trimm number;
begin

    workbook := xlsx_writer.start_book;
    sheet_1  := xlsx_writer.add_sheet  (workbook, 'Ausschreibung LVS');

    font_db := xlsx_writer.add_font      (workbook, 'DB Office', 10);
    fill_master := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="92cddc"/><bgColor indexed="64"/></patternFill>');
    fill_master2 := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="92cddc"/><bgColor indexed="64"/></patternFill>');
    fill_parent := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="daeef3"/><bgColor indexed="64"/></patternFill>');
    border_db_full := xlsx_writer.add_border      (workbook, '<left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/>');
    cs_border := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, font_id => font_db);
    cs_master := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, fill_id => fill_master,font_id => font_db);
    cs_master2 := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, fill_id => fill_master2,font_id => font_db);
    cs_parent := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, fill_id => fill_parent,font_id => font_db);
    number_format_child := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00",font_id => font_db);
    number_format_parent := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00", font_id => font_db, fill_id => fill_parent);
    number_format_master := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00", font_id => font_db, fill_id => fill_master2);

    dbms_lob.createtemporary(lob_loc => l_blob, cache => true, dur => dbms_lob.call);

    xlsx_writer.col_width(workbook, sheet_1, 1, 75);
    xlsx_writer.col_width(workbook, sheet_1, 2, 20);
    xlsx_writer.col_width(workbook, sheet_1, 3, 20);
    xlsx_writer.col_width(workbook, sheet_1, 6, 20);
    xlsx_writer.col_width(workbook, sheet_1, 7, 20);
    xlsx_writer.col_width(workbook, sheet_1, 8, 20);
    xlsx_writer.col_width(workbook, sheet_1, 9, 20);
    xlsx_writer.col_width(workbook, sheet_1, 10, 20);
    xlsx_writer.col_width(workbook, sheet_1, 11, 20);
    xlsx_writer.col_width(workbook, sheet_1, 12, 20);
    xlsx_writer.col_width(workbook, sheet_1, 13, 20);

    xlsx_writer.add_cell(workbook, sheet_1, 1, 1,style_id => cs_master, text => 'NAME');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 2,style_id => cs_master, text => 'CODE');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 3,style_id => cs_master, text => 'KENNUNG');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 4,style_id => cs_master, text => 'EINHEIT');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 5,style_id => cs_master, text => 'MENGE');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 6,style_id => cs_master, text => 'MIN PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 7,style_id => cs_master, text => 'GESAMT MIN PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 8,style_id => cs_master, text => 'MITTELWERT');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 9,style_id => cs_master, text => 'GESAMT MITTELWERT');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 10,style_id => cs_master, text => 'MEDIAN PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 11,style_id => cs_master, text => 'GESAMT MEDIAN PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 12,style_id => cs_master, text => 'MAX PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 13,style_id => cs_master, text => 'GESAMT MAX PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 14,style_id => cs_master, text => 'ANZAHL VERGABEN');

    for i in  ( /*SELECT   xtd2.name2 || nvl(xtd2.name3,'') name,xt.code master,xt.code||'.'||xtd.code2 parent,xt.name master_name, xtd.name2 parent_name,
                        (case
                            when instr(replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_'),'MLV') = 0
                                then ''
                            when substr(replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_'),8,1) = '_'
                                then replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_')
                            else substr(replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_'),1,7) || '_' || substr(replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_'),8)                            
                        end) kennung,
                        xt.code||'.'||xtd.code2||'.'||xtd2.code3 code,
                        xtd2.description,
                        replace(xtd2.menge,'.',',') menge,
                        xtd2.ME,
                        replace(xtd2.Einheitspreis,'.',',') Einheitspreis,
                        replace(xtd2.Gesamtbetrag,'.',',') Gesamtbetrag
               FROM     pd_import_x86 x,
                        XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/DA83/3.2' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                                PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                COLUMNS
                                    code     VARCHAR2(2000)  PATH '@RNoPart',
                                    name     VARCHAR2(4000)  PATH 'n:LblTx',
                                    subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                                ) xt,
                        xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/DA83/3.2'),'/BoQCtgy'
                                PASSING xt.subitem
                                COLUMNS 
                                    code2        VARCHAR2(2000)  PATH '@RNoPart',
                                    name2        VARCHAR2(4000)  PATH 'LblTx',
                                    subsubitem   XMLTYPE         PATH 'BoQBody/Itemlist/Item'
                                ) xtd,
                        xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/DA83/3.2'),'/Item'
                                PASSING xtd.subsubitem
                                COLUMNS 
                                    code3            VARCHAR2(200)  PATH '@RNoPart',
                                    name2            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[1]',
                                    name3            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[2]',
                                    name4            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[3]',
                                    name5            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[4]',
                                    description      VARCHAR2(4000) PATH 'substring(Description/CompleteText/DetailTxt,1,4000)',
                                    Menge            VARCHAR2(20)   PATH 'Qty',
                                    ME               VARCHAR2(10)   PATH 'QU',
                                    Einheitspreis    VARCHAR2(20)   PATH 'UP',
                                    Gesamtbetrag     VARCHAR2(20)   PATH 'IT'
                                ) xtd2
               where    x.id = p_blob_id 
               */

               --format x83

						SELECT   xtd2.name2 || nvl(xtd2.name3,'') name,xt.code master,xt.code||'.'||xtd.code2 parent,xt.name master_name, xtd.name2 parent_name,
												(case
													when instr(replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_'),'MLV') = 0
														then ''
													when substr(replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_'),8,1) = '_'
														then replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_')
													else substr(replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_'),1,7) || '_' || substr(replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_'),8)                            
												end) kennung,
												xt.code||'.'||xtd.code2||'.'||xtd2.code3 code,
												xtd2.description,
												replace(xtd2.menge,'.',',') menge,
												xtd2.ME,
												replace(xtd2.Einheitspreis,'.',',') Einheitspreis,
												replace(xtd2.Gesamtbetrag,'.',',') Gesamtbetrag, 'x83' as file_
										FROM     pd_import_x86 x,
												XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/DA83/3.2' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
														PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
														COLUMNS
															code     VARCHAR2(2000)  PATH '@RNoPart',
															name     VARCHAR2(4000)  PATH 'n:LblTx',
															subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
														) xt,
												xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/DA83/3.2'),'/BoQCtgy'
														PASSING xt.subitem
														COLUMNS 
															code2        VARCHAR2(2000)  PATH '@RNoPart',
															name2        VARCHAR2(4000)  PATH 'LblTx',
															subsubitem   XMLTYPE         PATH 'BoQBody/Itemlist/Item'
														) xtd,
												xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/DA83/3.2'),'/Item'
														PASSING xtd.subsubitem
														COLUMNS 
															code3            VARCHAR2(200)  PATH '@RNoPart',
															name2            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[1]',
															name3            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[2]',
															name4            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[3]',
															name5            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[4]',
															description      VARCHAR2(4000) PATH 'substring(Description/CompleteText/DetailTxt,1,4000)',
															Menge            VARCHAR2(20)   PATH 'Qty',
															ME               VARCHAR2(10)   PATH 'QU',
															Einheitspreis    VARCHAR2(20)   PATH 'UP',
															Gesamtbetrag     VARCHAR2(20)   PATH 'IT'
														) xtd2
										where    x.id = p_blob_id

						union all

						--format x82

						SELECT
						it.textoutltxt name,
						l1.rnopart_lvl1 master,
						l1.rnopart_lvl1 || '.' || l2.rnopart_lvl2 Parent,
						-- it.rnopart_item,
						l1.ctgy_lbl_lvl1 MASTERNAME,      -- Überschrift Kategorie Level 1
						l2.ctgy_lbl_lvl2 Parentname,      -- Überschrift Kategorie Level 2
						CASE
							WHEN INSTR(it.textoutltxt, 'MLV') > 0
							THEN SUBSTR(it.textoutltxt, INSTR(it.textoutltxt, 'MLV'))
							ELSE NULL
						END AS KENNUNG, 
						lpad(l1.rnopart_lvl1,2,0) || '.' || LPAD(l2.rnopart_lvl2,2,0) || '.' || LPAD(
						it.rnopart_item,4,0) code,
						it.description,
						--it.item_id,
						it.qty,
						it.qu,
						it.up,
						null , 'x82' as file_
						FROM pd_import_x86 x
						CROSS JOIN XMLTABLE(
						XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/200407' AS "n"),
						'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
						PASSING XMLType(x.DATEI, nls_charset_id('AL32UTF8'))
						COLUMNS
							rnopart_lvl1   VARCHAR2(20)    PATH '@RNoPart',
							ctgy_lbl_lvl1  VARCHAR2(400)   PATH 'normalize-space(string-join(n:LblTx//n:span, " "))',
							lvl2           XMLTYPE         PATH 'n:BoQBody/n:BoQCtgy'
						) l1
						CROSS JOIN XMLTABLE(
						XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/200407' AS "n"),
						'/n:BoQCtgy'
						PASSING l1.lvl2
						COLUMNS
							rnopart_lvl2   VARCHAR2(20)    PATH '@RNoPart',
							ctgy_lbl_lvl2  VARCHAR2(400)   PATH 'normalize-space(string-join(n:LblTx//n:span, " "))',
							items          XMLTYPE         PATH 'n:BoQBody/n:Itemlist/n:Item'
						) l2
						CROSS JOIN XMLTABLE(
						XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/200407' AS "n"),
						'/n:Item'
						PASSING l2.items
						COLUMNS
							rnopart_item   VARCHAR2(20)    PATH '@RNoPart',
							item_id        VARCHAR2(50)    PATH '@ID',
							qty            VARCHAR2(20)    PATH 'n:Qty',
							qu             VARCHAR2(20)    PATH 'n:QU',
							up             VARCHAR2(20)    PATH 'n:UP',
							description    VARCHAR2(4000)  PATH 'normalize-space(string-join(n:Description/n:CompleteText/n:DetailTxt//n:span, " "))',
							textoutltxt    VARCHAR2(1000)  PATH 'normalize-space(string-join(n:Description/n:CompleteText/n:OutlineText/n:OutlTxt/n:TextOutlTxt//n:span, " "))',
							alltext        VARCHAR2(4000)  PATH 'normalize-space(string-join(.//text(), " "))'
						) it
						WHERE x.id = p_blob_id
						union all

						--format x82 --anderes Format

						SELECT
						it.textoutltxt,
						l1.rnopart_lvl1,
						l1.rnopart_lvl1 || '.' || l2.rnopart_lvl2 ,
						l1.ctgy_lbl_lvl1,      -- Überschrift Kategorie Level 1
						l2.ctgy_lbl_lvl2,      -- Überschrift Kategorie Level 2
							CASE
							WHEN INSTR(it.textoutltxt, 'MLV') > 0
							THEN SUBSTR(it.textoutltxt, INSTR(it.textoutltxt, 'MLV'))
							ELSE NULL
						END  AS KENNUNG,
						lpad(l1.rnopart_lvl1,2,0) || '.' || lpad (l2.rnopart_lvl2,2,0) || '.' || 
						lpad (it.rnopart_item,4,0) ,
						it.description,
						-- it.item_id,
						it.qty,
						it.qu,
						it.up,
						null, 'x82' as file_
						FROM pd_import_x86 x
						CROSS JOIN XMLTABLE(
						XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/DA82/3.2' AS "n"),
						'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
						PASSING XMLType(x.DATEI, nls_charset_id('AL32UTF8'))
						COLUMNS
							rnopart_lvl1   VARCHAR2(20)    PATH '@RNoPart',
							ctgy_lbl_lvl1  VARCHAR2(400)   PATH 'normalize-space(string-join(n:LblTx//n:span, " "))',
							lvl2           XMLTYPE         PATH 'n:BoQBody/n:BoQCtgy'
						) l1
						CROSS JOIN XMLTABLE(
						XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/DA82/3.2' AS "n"),
						'/n:BoQCtgy'
						PASSING l1.lvl2
						COLUMNS
							rnopart_lvl2   VARCHAR2(20)    PATH '@RNoPart',
							ctgy_lbl_lvl2  VARCHAR2(400)   PATH 'normalize-space(string-join(n:LblTx//n:span, " "))',
							items          XMLTYPE         PATH 'n:BoQBody/n:Itemlist/n:Item'
						) l2
						CROSS JOIN XMLTABLE(
						XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/DA82/3.2' AS "n"),
						'/n:Item'
						PASSING l2.items
						COLUMNS
							rnopart_item   VARCHAR2(20)    PATH '@RNoPart',
							item_id        VARCHAR2(50)    PATH '@ID',
							qty            VARCHAR2(20)    PATH 'n:Qty',
							qu             VARCHAR2(20)    PATH 'n:QU',
							up             VARCHAR2(20)    PATH 'n:UP',
							description    VARCHAR2(4000)  PATH 'normalize-space(string-join(n:Description/n:CompleteText/n:DetailTxt//n:span, " "))',
							textoutltxt    VARCHAR2(1000)  PATH 'normalize-space(string-join(n:Description/n:CompleteText/n:OutlineText/n:OutlTxt/n:TextOutlTxt//n:span, " "))',
							alltext        VARCHAR2(4000)  PATH 'normalize-space(string-join(.//text(), " "))'
						) it
						WHERE x.id = p_blob_id
               )
    loop

       begin
            select  min(ap.EINHEITSPREIS),max(EINHEITSPREIS),avg(EINHEITSPREIS),median(EINHEITSPREIS),count(code)
            into    l_min,l_max,l_avg,l_median,l_count
            from    pd_auftrag_positionen ap
            join    pd_auftraege a on a.id = ap.auftrag_id
            where   replace(ap.UMSETZUNG_CODE,' ') = replace(replace(replace(i.kennung,'-'),'_'),' ')
            and     a.EINLESUNG_STATUS = 'Y'
            and     a.REGIONALBEREICH_ID in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
            and     a.KEDITOREN_NUMMER in (select  column_value from table(apex_string.split(p_liferant, ':')))
            and     ap.einheitspreis > 0;--13.09.2022,AR
       exception
            when no_data_found then 
                l_min:=null;
                l_max:=null;
                l_avg:=null;
                l_median:=null;
                l_id :=null;
       end;



        INSERT INTO pd_ausschreibung
            (blob_id,muster_id,name,code,einheit,menge,min_preis,mittlerer_preis,max_preis,median_preis,parent_id,master_id,parent_name,master_name,kennung,code_count,x82, einzelpreis, gesamtpreis )
        VALUES
            (p_blob_id,l_id,i.name,i.code,i.me,to_number(replace(i.menge,',','.')),l_min,l_avg,l_max,l_median,i.parent,i.master,i.parent_name,i.master_name,i.kennung,l_count, case when i.file_= 'x82' then 1 else 0 end, nvl(i.einheitspreis,0), to_number(nvl(i.einheitspreis,0))*to_number (nvl(i.menge,0)));

    end loop;--i in  (SELECT xtd2.name2...

        ---jetzt filtern
        IF FILTER_VERGABESUMME(p_blob_id) is null then --filter der vergabesumme
            delete pd_ausschreibung where blob_id =p_blob_id; --Filterbereich greift 
        end if;

        select GETRIMMTER_MITTELWERT into v_trimm from PD_IMPORT_X86 where id=p_blob_id;
        FILTER_TRIMM_MITTELWERT(v_trimm,p_blob_id); --filter getrimmter mittelwert anteil prozent löschen

    commit;

    for j in (with  mustern as 
                   (select  pa.parent_id,pa.code,pa.kennung,pa.name,pa.min_preis,pa.mittlerer_preis,pa.max_preis,pa.median_preis,pa.EINHEIT,pa.MENGE,'CHILD' art,
                            pa.min_preis*pa.MENGE gesamt_min,pa.mittlerer_preis * pa.MENGE gesammt_mittel,pa.max_preis * pa.MENGE gesamt_max,pa.median_preis * pa.MENGE gesamt_median,pa.code_count
                    from    pd_ausschreibung pa 
                    where   pa.blob_id = p_blob_id
                    union all
                    select  pa.master_id,pa.parent_id,null,pa.parent_name,sum(pa.min_preis),sum(pa.mittlerer_preis),sum(pa.max_preis),sum(pa.median_preis),null,null,'PARENT' art,
                            sum(pa.min_preis*pa.MENGE),sum(pa.mittlerer_preis*pa.MENGE),sum(pa.max_preis*pa.MENGE),sum(pa.median_preis*pa.MENGE),null
                    from    pd_ausschreibung pa 
                    where   pa.blob_id = p_blob_id
                    group by pa.parent_id,pa.master_id,pa.parent_name
                    union all 
                    select  null,pa.master_id,null,pa.master_name,sum(pa.min_preis),sum(pa.mittlerer_preis),sum(pa.max_preis),sum(pa.median_preis),null,null,'MASTER' art,
                            sum(pa.min_preis*pa.MENGE),sum(pa.mittlerer_preis*pa.MENGE),sum(pa.max_preis*pa.MENGE),sum(pa.median_preis*pa.MENGE),null
                    from    pd_ausschreibung pa 
                    where   pa.blob_id = p_blob_id
                    group by pa.master_id,pa.master_name
                   )
              SELECT  trim(m.code) code,
                      trim(m.name) name,
                      trim(m.kennung) kennung,
                      trim(m.einheit) einheit,
                      m.menge,
                      m.min_preis,
                      m.mittlerer_preis,
                      m.max_preis,
                      m.median_preis,
                      m.gesamt_min,
                      m.gesammt_mittel,
                      m.gesamt_max,
                      m.gesamt_median,
                      m.art,
                      m.code_count,
                      (case m.art
                        when 'CHILD' then cs_border
                        when 'PARENT' then cs_parent
                        else cs_master
                      end) cs_style,
                      (case m.art
                        when 'CHILD' then number_format_child
                        when 'PARENT' then number_format_parent
                        else number_format_master
                      end) number_format
              FROM  mustern m
              START WITH parent_id IS NULL
              CONNECT BY PRIOR m.code = m.parent_id 
              ORDER SIBLINGS BY m.code
             )
    loop

       xlsx_writer.add_cell(workbook, sheet_1, l_row, 1,style_id => j.cs_style, text => j.name);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 2,style_id => j.cs_style, text => j.code);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 3,style_id => j.cs_style, text => j.kennung);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 4,style_id => j.cs_style, text => j.einheit);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 5,style_id => j.cs_style, value_ => j.menge);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 6,style_id => j.number_format, value_ => j.min_preis);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 7,style_id => j.number_format, value_ => j.gesamt_min);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 8,style_id => j.number_format, value_ => j.mittlerer_preis);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 9,style_id => j.number_format, value_ => j.gesammt_mittel);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 10,style_id => j.number_format, value_ => j.median_preis);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 11,style_id => j.number_format, value_ => j.gesamt_median);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 12,style_id => j.number_format, value_ => j.max_preis);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 13,style_id => j.number_format, value_ => j.gesamt_max);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 14,style_id => j.cs_style, value_ => j.code_count);

       l_row := l_row + 1;
    end loop;--j in (with  mustern as (  select pa.parent_id...


    workbook:= print_param_worksheet(workbook,p_blob_id,number_format_master) ;---Eingabeparameter ergänzen

    xlsx_writer.freeze_sheet(workbook, sheet_1,0,1);

    xlsx := xlsx_writer.create_xlsx(workbook);


    /*delete from pd_import_x86 where id = p_blob_id;*/
    --aufgrund speicher, lösche ich die Datei 
    update pd_import_x86 
    set datei=null
    where id = p_blob_id;

    delete from pd_import_x86 where datum  < sysdate-365 ;

    delete from pd_ausschreibung where blob_id not in ( 
    select blob_id from pd_import_x86);

    --Mailversand der Auswertung
    SendMailAuswertung(p_user_id => p_user_id,p_anhang => xlsx,p_filename => 'Ausschreibung_LVS.xlsx');

    DBMS_LOB.FREETEMPORARY(xlsx);

    exception 
        when others then
            DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.export_ausschreibung_to_excel2: Fehler bei auswertung: ' || SQLCODE || ': ' || SQLERRM ||
            ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'AUSSCHREIBUNG');

            SendMailAuswertungFehler
                (
                p_user_id => p_user_id,
                p_error => 'Fehler bei der Auswertung, bitte wenden Sie sich an den Anwendungsverantwortlichen.'
                );

end export_ausschreibung_to_excel2;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

PROCEDURE export_ausschreibung_to_excel3(p_blob_id number,p_typ_id number,p_region_id number,p_von date, p_bis date,p_regionen varchar2,p_liferant varchar2,p_user_id number) as 

       workbook xlsx_writer.book_r;
       sheet_1  integer;

       xlsx     blob;

       cs_border integer;
       cs_master integer;
       cs_master2 integer;
       cs_parent integer;
       number_format_child integer;
       number_format_parent integer;
       number_format_master integer;
       border_db_full integer;
       font_db  integer;
       fill_master integer;
       fill_parent integer;
       fill_master2 integer;

       c_limit   constant integer := 50;
       c_x_split constant integer := 3;
       c_y_split constant integer := 8;
       c_y_region constant integer := 1;

       TYPE curtype IS REF CURSOR;

       l_filename varchar2(100);
       L_BLOB blob;
       l_target_charset VARCHAR2(100) := 'WE8MSWIN1252';
       L_DEST_OFFSET    INTEGER := 1;
       L_SRC_OFFSET     INTEGER := 1;
       L_LANG_CONTEXT   INTEGER := DBMS_LOB.DEFAULT_LANG_CTX;
       L_WARNING        INTEGER;
       L_LENGTH         INTEGER;
       l_betrag         number;

       l_min number;
       l_avg number;
       l_max number;
       l_median number;
       l_count number;
       l_id  number;
       l_typ number;
       l_parent_id varchar2(20);
       l_master_id varchar2(20);
       l_row number:=2;
begin

    workbook := xlsx_writer.start_book;
    sheet_1  := xlsx_writer.add_sheet  (workbook, 'Ausschreibung LVS');

    font_db := xlsx_writer.add_font      (workbook, 'DB Office', 10);
    fill_master := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="92cddc"/><bgColor indexed="64"/></patternFill>');
    fill_master2 := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="92cddc"/><bgColor indexed="64"/></patternFill>');
    fill_parent := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="daeef3"/><bgColor indexed="64"/></patternFill>');
    border_db_full := xlsx_writer.add_border      (workbook, '<left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/>');
    cs_border := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, font_id => font_db);
    cs_master := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, fill_id => fill_master,font_id => font_db);
    cs_master2 := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, fill_id => fill_master2,font_id => font_db);
    cs_parent := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, fill_id => fill_parent,font_id => font_db);
    number_format_child := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00",font_id => font_db);
    number_format_parent := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00", font_id => font_db, fill_id => fill_parent);
    number_format_master := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00", font_id => font_db, fill_id => fill_master2);

    dbms_lob.createtemporary(lob_loc => l_blob, cache => true, dur => dbms_lob.call);

    xlsx_writer.col_width(workbook, sheet_1, 1, 75);
    xlsx_writer.col_width(workbook, sheet_1, 2, 20);
    xlsx_writer.col_width(workbook, sheet_1, 3, 20);
    xlsx_writer.col_width(workbook, sheet_1, 6, 20);
    xlsx_writer.col_width(workbook, sheet_1, 7, 20);
    xlsx_writer.col_width(workbook, sheet_1, 8, 20);
    xlsx_writer.col_width(workbook, sheet_1, 9, 20);
    xlsx_writer.col_width(workbook, sheet_1, 10, 20);
    xlsx_writer.col_width(workbook, sheet_1, 11, 20);
    xlsx_writer.col_width(workbook, sheet_1, 12, 20);
    xlsx_writer.col_width(workbook, sheet_1, 13, 20);

    xlsx_writer.add_cell(workbook, sheet_1, 1, 1,style_id => cs_master, text => 'NAME');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 2,style_id => cs_master, text => 'CODE');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 3,style_id => cs_master, text => 'KENNUNG');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 4,style_id => cs_master, text => 'EINHEIT');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 5,style_id => cs_master, text => 'MENGE');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 6,style_id => cs_master, text => 'MIN PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 7,style_id => cs_master, text => 'GESAMT MIN PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 8,style_id => cs_master, text => 'MITTELWERT');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 9,style_id => cs_master, text => 'GESAMT MITTELWERT');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 10,style_id => cs_master, text => 'MEDIAN PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 11,style_id => cs_master, text => 'GESAMT MEDIAN PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 12,style_id => cs_master, text => 'MAX PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 13,style_id => cs_master, text => 'GESAMT MAX PREIS');
    xlsx_writer.add_cell(workbook, sheet_1, 1, 14,style_id => cs_master, text => 'ANZAHL VERGABEN');

    for i in  ( SELECT   xtd2.name2 || nvl(xtd2.name3,'') name,xt.code master,xt.code||'.'||xtd.code2 parent,xt.name master_name, xtd.name2 parent_name,
                        (case
                            when instr(replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_'),'MLV') = 0
                                then ''
                            when substr(replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_'),8,1) = '_'
                                then replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_')
                            else substr(replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_'),1,7) || '_' || substr(replace(substr(replace(xtd2.name2 || nvl(xtd2.name3,''),'MVL-','MLV-'),instr(replace(xtd2.name2 || nvl(xtd2.name3,'') || nvl(xtd2.name4,'') || nvl(xtd2.name5,''),'MVL-','MLV-'),'MLV-'),16),' ','_'),8)                            
                        end) kennung,
                        xt.code||'.'||xtd.code2||'.'||xtd2.code3 code,
                        xtd2.description,
                        replace(xtd2.menge,'.',',') menge,
                        xtd2.ME,
                        replace(xtd2.Einheitspreis,'.',',') Einheitspreis,
                        replace(xtd2.Gesamtbetrag,'.',',') Gesamtbetrag
               FROM     pd_import_x86 x,
                        XMLTABLE(XMLNAMESPACES('http://www.gaeb.de/GAEB_DA_XML/DA83/3.3' AS "n"),'/n:GAEB/n:Award/n:BoQ/n:BoQBody/n:BoQCtgy'
                                PASSING XMLType(x.DATEI,nls_charset_id('UTF8'))
                                COLUMNS
                                    code     VARCHAR2(2000)  PATH '@RNoPart',
                                    name     VARCHAR2(4000)  PATH 'n:LblTx',
                                    subitem  xmltype         PATH 'n:BoQBody/n:BoQCtgy'
                                ) xt,
                        xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/DA83/3.3'),'/BoQCtgy'
                                PASSING xt.subitem
                                COLUMNS 
                                    code2        VARCHAR2(2000)  PATH '@RNoPart',
                                    name2        VARCHAR2(4000)  PATH 'LblTx',
                                    subsubitem   XMLTYPE         PATH 'BoQBody/Itemlist/Item'
                                ) xtd,
                        xmltable(XMLNamespaces (default 'http://www.gaeb.de/GAEB_DA_XML/DA83/3.3'),'/Item'
                                PASSING xtd.subsubitem
                                COLUMNS 
                                    code3            VARCHAR2(200)  PATH '@RNoPart',
                                    name2            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[1]',
                                    name3            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[2]',
                                    name4            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[3]',
                                    name5            VARCHAR2(4000) PATH 'Description/CompleteText/OutlineText/OutlTxt/TextOutlTxt/p/span[4]',
                                    description      VARCHAR2(4000) PATH 'substring(Description/CompleteText/DetailTxt,1,4000)',
                                    Menge            VARCHAR2(20)   PATH 'Qty',
                                    ME               VARCHAR2(10)   PATH 'QU',
                                    Einheitspreis    VARCHAR2(20)   PATH 'UP',
                                    Gesamtbetrag     VARCHAR2(20)   PATH 'IT'
                                ) xtd2
               where    x.id = p_blob_id
               )
    loop

       begin
            select  min(ap.EINHEITSPREIS),max(EINHEITSPREIS),avg(EINHEITSPREIS),median(EINHEITSPREIS),count(code)
            into    l_min,l_max,l_avg,l_median,l_count
            from    pd_auftrag_positionen ap
            join    pd_auftraege a on a.id = ap.auftrag_id
            where   replace(ap.UMSETZUNG_CODE,' ') = replace(replace(replace(i.kennung,'-'),'_'),' ')
            and     a.EINLESUNG_STATUS = 'Y'
            and     a.REGIONALBEREICH_ID in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
            and     a.KEDITOREN_NUMMER in (select  column_value from table(apex_string.split(p_liferant, ':')))
            and     ap.einheitspreis > 0;--13.09.2022,AR
       exception
            when no_data_found then 
                l_min:=null;
                l_max:=null;
                l_avg:=null;
                l_median:=null;
                l_id :=null;
       end;

       INSERT INTO pd_ausschreibung
            (blob_id,muster_id,name,code,einheit,menge,min_preis,mittlerer_preis,max_preis,median_preis,parent_id,master_id,parent_name,master_name,kennung,code_count)
       VALUES
            (p_blob_id,l_id,i.name,i.code,i.me,to_number(replace(i.menge,',','.')),l_min,l_avg,l_max,l_median,i.parent,i.master,i.parent_name,i.master_name,i.kennung,l_count);

    end loop;--i in  (SELECT xtd2.name2...

    commit;

    for j in (with  mustern as 
                   (select  pa.parent_id,pa.code,pa.kennung,pa.name,pa.min_preis,pa.mittlerer_preis,pa.max_preis,pa.median_preis,pa.EINHEIT,pa.MENGE,'CHILD' art,
                            pa.min_preis*pa.MENGE gesamt_min,pa.mittlerer_preis * pa.MENGE gesammt_mittel,pa.max_preis * pa.MENGE gesamt_max,pa.median_preis * pa.MENGE gesamt_median,pa.code_count
                    from    pd_ausschreibung pa 
                    where   pa.blob_id = p_blob_id
                    union all
                    select  pa.master_id,pa.parent_id,null,pa.parent_name,sum(pa.min_preis),sum(pa.mittlerer_preis),sum(pa.max_preis),sum(pa.median_preis),null,null,'PARENT' art,
                            sum(pa.min_preis*pa.MENGE),sum(pa.mittlerer_preis*pa.MENGE),sum(pa.max_preis*pa.MENGE),sum(pa.median_preis*pa.MENGE),null
                    from    pd_ausschreibung pa 
                    where   pa.blob_id = p_blob_id
                    group by pa.parent_id,pa.master_id,pa.parent_name
                    union all 
                    select  null,pa.master_id,null,pa.master_name,sum(pa.min_preis),sum(pa.mittlerer_preis),sum(pa.max_preis),sum(pa.median_preis),null,null,'MASTER' art,
                            sum(pa.min_preis*pa.MENGE),sum(pa.mittlerer_preis*pa.MENGE),sum(pa.max_preis*pa.MENGE),sum(pa.median_preis*pa.MENGE),null
                    from    pd_ausschreibung pa 
                    where   pa.blob_id = p_blob_id
                    group by pa.master_id,pa.master_name
                   )
              SELECT  trim(m.code) code,
                      trim(m.name) name,
                      trim(m.kennung) kennung,
                      trim(m.einheit) einheit,
                      m.menge,
                      m.min_preis,
                      m.mittlerer_preis,
                      m.max_preis,
                      m.median_preis,
                      m.gesamt_min,
                      m.gesammt_mittel,
                      m.gesamt_max,
                      m.gesamt_median,
                      m.art,
                      m.code_count,
                      (case m.art
                        when 'CHILD' then cs_border
                        when 'PARENT' then cs_parent
                        else cs_master
                      end) cs_style,
                      (case m.art
                        when 'CHILD' then number_format_child
                        when 'PARENT' then number_format_parent
                        else number_format_master
                      end) number_format
              FROM  mustern m
              START WITH parent_id IS NULL
              CONNECT BY PRIOR m.code = m.parent_id 
              ORDER SIBLINGS BY m.code
             )
    loop

       xlsx_writer.add_cell(workbook, sheet_1, l_row, 1,style_id => j.cs_style, text => j.name);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 2,style_id => j.cs_style, text => j.code);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 3,style_id => j.cs_style, text => j.kennung);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 4,style_id => j.cs_style, text => j.einheit);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 5,style_id => j.cs_style, value_ => j.menge);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 6,style_id => j.number_format, value_ => j.min_preis);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 7,style_id => j.number_format, value_ => j.gesamt_min);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 8,style_id => j.number_format, value_ => j.mittlerer_preis);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 9,style_id => j.number_format, value_ => j.gesammt_mittel);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 10,style_id => j.number_format, value_ => j.median_preis);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 11,style_id => j.number_format, value_ => j.gesamt_median);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 12,style_id => j.number_format, value_ => j.max_preis);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 13,style_id => j.number_format, value_ => j.gesamt_max);
       xlsx_writer.add_cell(workbook, sheet_1, l_row, 14,style_id => j.cs_style, value_ => j.code_count);

       l_row := l_row + 1;
    end loop;--j in (with  mustern as (  select pa.parent_id...

    xlsx_writer.freeze_sheet(workbook, sheet_1,0,1);

    xlsx := xlsx_writer.create_xlsx(workbook);



    delete from pd_import_x86 where id = p_blob_id;
    delete from pd_ausschreibung where blob_id = p_blob_id;

    --Mailversand der Auswertung
    SendMailAuswertung(p_user_id => p_user_id,p_anhang => xlsx,p_filename => 'Ausschreibung_LVS.xlsx');

    DBMS_LOB.FREETEMPORARY(xlsx);

    exception 
        when others then
            DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.export_ausschreibung_to_excel2: Fehler bei auswertung: ' || SQLCODE || ': ' || SQLERRM ||
            ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'AUSSCHREIBUNG');

            SendMailAuswertungFehler
                (
                p_user_id => p_user_id,
                p_error => 'Fehler bei der Auswertung, bitte wenden Sie sich an den Anwendungsverantwortlichen.'
                );

end export_ausschreibung_to_excel3;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

procedure auswertung_to_excel(p_von date, p_bis date, p_lvs varchar2,p_regionen varchar2,p_liferant varchar2,p_user_id number) as

       workbook xlsx_writer_v2.book_r;
       sheet_1  integer;

       xlsx     blob;

       c_limit   constant integer := 50;
       c_x_split constant integer := 3;
       c_y_split constant integer := 14;
       c_y_region constant integer := 1;
       cs_center_wrapped integer;
       cs_center integer;
       cs_number integer;
       cs_center_bold integer;
       cs_center_bold_white integer;
       cs_center_bold_grey integer;
       cs_border integer;
       cs_master integer;
       cs_master_master integer;
       cs_master_parent integer;
       cs_rot integer;
       cs_gelb integer;
       cs_hellgelb integer;
       cs_hellgruen integer;
       cs_gruen integer;
       cs_noborder integer;
       fill_gelb integer;
       fill_hellgelb integer;
       fill_hellgruen integer;
       fill_gruen integer;
       fill_rot integer;

       font_db  integer;
       font_db_small integer;
       font_db_bold integer;
       fill_db integer;
       fill_db_grey integer;
       font_db_bold_white integer;
       fill_master integer;
       fill_parent integer;
       border_db integer;
       border_db_full integer;

       number_format integer;
       center_number_format integer;
       datum_format integer;

       TYPE curtype IS REF CURSOR;

       v_auftrage t_auftraege := new t_auftraege();
       v_kennungen T_KENNUNG := new T_KENNUNG();
       l_filename varchar2(100);
       L_BLOB blob;
       l_target_charset VARCHAR2(100) := 'WE8MSWIN1252';
       L_DEST_OFFSET    INTEGER := 1;
       L_SRC_OFFSET     INTEGER := 1;
       L_LANG_CONTEXT   INTEGER := DBMS_LOB.DEFAULT_LANG_CTX;
       L_WARNING        INTEGER;
       L_LENGTH         INTEGER;
       l_column         number:=1;
       l_row            number;
    l_betrag         number;
    l_contract_count number := 0;
    l_muster_count   number := 0;
       l_random_number  number := floor(dbms_random.value(1, 1000000000));
begin

    DBS_LOGGING.LOG_INFO_AT('PREISDATENBANK_PKG.auswertung_to_excel START: p_von='||to_char(p_von,'YYYY-MM-DD')||', p_bis='||to_char(p_bis,'YYYY-MM-DD')||', p_lvs='||nvl(p_lvs,'')||', p_regionen='||nvl(p_regionen,'')||', p_liferant='||nvl(p_liferant,'')||', p_user_id='||nvl(to_char(p_user_id),'-'),'AUSWERTUNG');

    workbook := xlsx_writer_v2.start_book;
    sheet_1  := xlsx_writer_v2.add_sheet  (workbook, 'Auswertung Vertrags LVS');

    dbms_lob.createtemporary(lob_loc => l_blob, cache => true, dur => dbms_lob.call);

    -- Style definition
    font_db := xlsx_writer_v2.add_font      (workbook, 'DB Office', 10);
    border_db := xlsx_writer_v2.add_border  (workbook, '<left/><right/><top/><bottom/><diagonal/>');
    border_db_full := xlsx_writer_v2.add_border      (workbook, '<left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/>');
    fill_db:= xlsx_writer_v2.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="00ccff"/><bgColor indexed="64"/></patternFill>');
    fill_db_grey := xlsx_writer_v2.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="d9d9d9"/><bgColor indexed="64"/></patternFill>');
    fill_master := xlsx_writer_v2.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="92cddc"/><bgColor indexed="64"/></patternFill>');
    fill_parent := xlsx_writer_v2.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="daeef3"/><bgColor indexed="64"/></patternFill>');
    fill_gelb := xlsx_writer_v2.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="ffc000"/><bgColor indexed="64"/></patternFill>');
    fill_hellgelb := xlsx_writer_v2.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="ffff00"/><bgColor indexed="64"/></patternFill>');
    fill_hellgruen := xlsx_writer_v2.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="00ff00"/><bgColor indexed="64"/></patternFill>');
    fill_gruen := xlsx_writer_v2.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="00b050"/><bgColor indexed="64"/></patternFill>');
    fill_rot := xlsx_writer_v2.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="ff0000"/><bgColor indexed="64"/></patternFill>');
    font_db_small := xlsx_writer_v2.add_font      (workbook, 'DB Office', 7);
    font_db_bold := xlsx_writer_v2.add_font      (workbook, 'DB Office', 10, b => true);
    font_db_bold_white := xlsx_writer_v2.add_font      (workbook, 'DB Office', 10, color=> 'theme="0"', b => true);

    cs_center_wrapped  := xlsx_writer_v2.add_cell_style(workbook, vertical_alignment => 'center', vertical_horizontal => 'center', wrap_text => true,font_id => font_db_small, border_id => border_db_full);
    cs_center  := xlsx_writer_v2.add_cell_style(workbook, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_db_full);
    cs_center_bold  := xlsx_writer_v2.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_db, border_id => border_db_full);
    cs_center_bold_white := xlsx_writer_v2.add_cell_style(workbook, font_id => font_db_bold_white, fill_id => fill_db,vertical_alignment => 'center', vertical_horizontal => 'center', border_id => border_db_full);
    cs_border := xlsx_writer_v2.add_cell_style(workbook, border_id => border_db_full);
    cs_noborder := xlsx_writer_v2.add_cell_style(workbook, border_id => border_db);
    cs_center_bold_grey := xlsx_writer_v2.add_cell_style(workbook, fill_id => fill_db_grey, border_id => border_db_full, font_id => font_db_bold);
    cs_rot := xlsx_writer_v2.add_cell_style(workbook, fill_id => fill_rot, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db);
    cs_gelb := xlsx_writer_v2.add_cell_style(workbook, fill_id => fill_gelb, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db);
    cs_hellgelb := xlsx_writer_v2.add_cell_style(workbook, fill_id => fill_hellgelb, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db);
    cs_hellgruen := xlsx_writer_v2.add_cell_style(workbook, fill_id => fill_hellgruen, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db);
    cs_gruen := xlsx_writer_v2.add_cell_style(workbook, fill_id => fill_gruen, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db);
    cs_master_parent := xlsx_writer_v2.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_parent, border_id => border_db_full);
    cs_master_master := xlsx_writer_v2.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_master, border_id => border_db_full);

    number_format := xlsx_writer_v2.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer_v2."#.##0.00 €" , font_id => font_db);
    center_number_format := xlsx_writer_v2.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer_v2."#.##0.00 €" , font_id => font_db, vertical_alignment => 'center', vertical_horizontal => 'center');
    datum_format := xlsx_writer_v2.add_cell_style(workbook, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_db_full, num_fmt_id => xlsx_writer_v2."mm-dd-yy");

    xlsx_writer_v2.add_cell(workbook, sheet_1,   1, 1,style_id => cs_center_bold, text => 'I.) Grundangaben');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   1, 2,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   1, 3,style_id => cs_border, text => '');

    xlsx_writer_v2.add_cell(workbook, sheet_1,   2, 1,style_id => cs_border, text => 'Region');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   2, 2,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   2, 3,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   3, 1,style_id => cs_border, text => 'Bezeichnung der Maßnahme');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   3, 2,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   3, 3,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   4, 1,style_id => cs_border, text => 'Auftragnehmer');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   4, 2,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   4, 3,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   5, 1,style_id => cs_border, text => 'LV-Datum');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   5, 2,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   5, 3,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   6, 1,style_id => cs_border, text => 'Vergabesumme');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   6, 2,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   6, 3,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   7, 1,style_id => cs_border, text => 'Vergabevorgangsnummer');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   7, 2,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   7, 3,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   8, 1,style_id => cs_border, text => 'SAP-Kontraktnummer');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   8, 2,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   8, 3,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   9, 1,style_id => cs_border, text => 'Kreditorennummer');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   9, 2,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   9, 3,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   10, 1,style_id => cs_border, text => 'Anzahl Muster-LV Pos. mit Kennung');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   10, 2,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   10, 3,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   11, 1,style_id => cs_border, text => '% Anteil  Muster-LV Pos. mit Kennung');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   11, 2,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   11, 3,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   12, 1,style_id => cs_border, text => 'Muster-LV Code');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   12, 2,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   12, 3,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   13, 1,style_id => cs_center_bold, text => 'II.) Einheitspreise');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   13, 2,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   13, 3,style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   14, 1,style_id => cs_center_bold_grey, text => 'Pos.');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   14, 2,style_id => cs_center_bold_grey, text => 'Pos.-Text');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   14, 3,style_id => cs_center_bold_grey, text => 'Einheit');

    xlsx_writer_v2.col_width(workbook, sheet_1, 1, 10);
    xlsx_writer_v2.col_width(workbook, sheet_1, 2, 50);
    xlsx_writer_v2.col_width(workbook, sheet_1, 3, 10);

    for i in 
    (
    with all_positionen as (select count(*) gesamt_anzahl, AUFTRAG_ID from pd_auftrag_positionen group by AUFTRAG_ID),
         lv_positionen as (select count(*) kennung_anzahl, AUFTRAG_ID from pd_auftrag_positionen where UMZETZUNG_CODE like 'MLV%' group by AUFTRAG_ID)
    select  r.code Region,
            a.projekt_desc Bezeichnung,
            a.auftragnahmer_name Auftragnahmer,
            a.datum LVDatum,
            a.total Vergabesumme,
            a.id,
            a.KEDITOREN_NUMMER,
            a.SAP_NR,
            a.VERTRAG_NR,
            nvl(kennung_anzahl,0)||' von '||gesamt_anzahl anzahl,
            round((nvl(kennung_anzahl,0)/gesamt_anzahl)*100,2)||'%' prozentual_anzahl,
            listagg(lvs.lv_code, ', ') within group (order by lvs.lv_code) lv_code
    from    pd_auftraege a
    join    pd_region r on r.id = a.regionalbereich_id
    join    lv_positionen ap on a.id = ap.auftrag_id
    join    pd_auftrag_lvs lvs on a.id = lvs.auftrag_id
    join    all_positionen allp on a.id = allp.AUFTRAG_ID
    where   a.datum between p_von and p_bis
    and     a.EINLESUNG_STATUS = 'Y'
    and     r.id in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
    and     a.KEDITOREN_NUMMER in (select column_value from table(apex_string.split(p_liferant, ':')))
    --keine Verträge in die Auswertung übernehmen, die nicht einen Preis für die Positionen haben
    and     (   select  count(*)
                from        pd_auftrag_positionen ap
                cross join  pd_muster_lvs m
                where   m.id in (select column_value from table(apex_string.split_numbers(p_lvs, ':')))
                and     ap.auftrag_id = a.id
                and     ap.code like 'M%'
                and     ap.einheitspreis > 0
                and     instr(ap.UMSETZUNG_CODE,m.position_kennung2) > 0
            ) > 0
	--keine Verträge in die Auswertung übernehmen, die nicht einen Preis für die Positionen haben
    group by r.code,a.projekt_desc,a.auftragnahmer_name,a.datum,a.total,a.id,a.KEDITOREN_NUMMER,a.SAP_NR,a.VERTRAG_NR,kennung_anzahl,gesamt_anzahl
    order by a.auftragnahmer_name
    )
    loop

        xlsx_writer_v2.col_width(workbook, sheet_1, c_x_split+l_column+4, 15);                              
        xlsx_writer_v2.add_cell(workbook, sheet_1,   c_y_region, c_x_split+l_column+5,style_id => cs_center_bold_white, text => 'Test');
        xlsx_writer_v2.add_cell(workbook, sheet_1,   c_y_region+12, c_x_split+l_column+5,style_id => cs_center_bold_white, text => '');
        xlsx_writer_v2.add_cell(workbook, sheet_1,   c_y_region+13, c_x_split+l_column+5,style_id => cs_center_bold_grey, text => '');
        xlsx_writer_v2.add_cell(workbook, sheet_1,   1+c_y_region, c_x_split+l_column+5,style_id => cs_center, text => trim(i.Region));

        v_auftrage.extend;
        v_auftrage(v_auftrage.count).auftrag_id := i.id;
        v_auftrage(v_auftrage.count).column_position := c_x_split+l_column+5;

        xlsx_writer_v2.add_cell(workbook, sheet_1,   2+c_y_region, c_x_split+l_column+5,style_id => cs_center_wrapped, text => trim(i.Bezeichnung));
        xlsx_writer_v2.add_cell(workbook, sheet_1,   3+c_y_region, c_x_split+l_column+5,style_id => cs_center, text => trim(i.Auftragnahmer));
        xlsx_writer_v2.add_cell(workbook, sheet_1,   4+c_y_region, c_x_split+l_column+5,style_id => cs_center, text => i.LVDatum);
        xlsx_writer_v2.add_cell(workbook, sheet_1,   5+c_y_region, c_x_split+l_column+5,style_id => center_number_format, value_ => i.Vergabesumme);
        xlsx_writer_v2.add_cell(workbook, sheet_1,   6+c_y_region, c_x_split+l_column+5,style_id => cs_center, text => i.VERTRAG_NR);
        xlsx_writer_v2.add_cell(workbook, sheet_1,   7+c_y_region, c_x_split+l_column+5,style_id => cs_center, text => i.SAP_NR);
        xlsx_writer_v2.add_cell(workbook, sheet_1,   8+c_y_region, c_x_split+l_column+5,style_id => cs_center, text => i.KEDITOREN_NUMMER);

        xlsx_writer_v2.add_cell(workbook, sheet_1,   9+c_y_region, c_x_split+l_column+5,style_id => cs_center, text => i.anzahl);
        xlsx_writer_v2.add_cell(workbook, sheet_1,   10+c_y_region, c_x_split+l_column+5,style_id => cs_center, text => i.prozentual_anzahl);
        xlsx_writer_v2.add_cell(workbook, sheet_1,   11+c_y_region, c_x_split+l_column+5,style_id => cs_center, text => i.lv_code);

        l_column:=l_column+1;
    end loop;

    l_contract_count := v_auftrage.count;
    DBS_LOGGING.LOG_DEBUG_AT('PREISDATENBANK_PKG.auswertung_to_excel: fetched contracts count = '||l_contract_count,'AUSWERTUNG');

    xlsx_writer_v2.add_cell(workbook, sheet_1,   14, 4,style_id => cs_center_bold_grey, text => 'Minimum');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   14, 5,style_id => cs_center_bold_grey, text => 'Mittelwert');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   14, 6,style_id => cs_center_bold_grey, text => 'Median');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   14, 7,style_id => cs_center_bold_grey, text => 'Maximum');
    xlsx_writer_v2.add_cell(workbook, sheet_1,   13, 4,style_id => cs_center_bold_grey, value_ => l_column-1);
    xlsx_writer_v2.add_cell(workbook, sheet_1,   13, 5,style_id => cs_center_bold_grey, text => 'Vergaben');

    l_row:=1;

    xlsx_writer_v2.add_cell(workbook, sheet_1, 8, 8,style_id => cs_rot, text => '= 0');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 9, 8,style_id => cs_gelb, text => '= 1 - 4');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 10, 8,style_id => cs_hellgelb, text => '= 5 - 9');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 11, 8,style_id => cs_hellgruen, text => '= 10 - 24');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 12, 8,style_id => cs_gruen, text => '? 25');

    for j in 
    (
        with  mustern as 
        (select m.id,
                m.code,
                m.position_kennung,
                m.name,
                m.description,
                m.MUSTER_TYP_ID || m.code id_tree, 
                decode(m.parent_id,null,null,m.MUSTER_TYP_ID||m.parent_id) parent_tree,
                m.parent_id,
                m.einheit,
                to_char(round(min(ap.EINHEITSPREIS),2),'999G999G999G999G990D00') MINIMUM_PREIS,
                to_char(round(avg(ap.EINHEITSPREIS),2),'999G999G999G999G990D00') MITTELWERT_PREIS, 
                to_char(round(median(ap.EINHEITSPREIS),2),'999G999G999G999G990D00') MEDIAN_PREIS,
                to_char(round(max(ap.EINHEITSPREIS),2),'999G999G999G999G990D00') MAXIMUM_PREIS,
                COUNT(ap.UMZETZUNG_CODE) ANZAHL
         from      PD_MUSTER_LVS m
         left join ( select avg(EINHEITSPREIS) EINHEITSPREIS,
                            auftrag_id,
                            UMZETZUNG_CODE,
                            UMSETZUNG_CODE
                    from    pd_auftrag_positionen pa
                    join    pd_auftraege a on a.id = pa.auftrag_id and a.datum between p_von and p_bis 
                    and     a.KEDITOREN_NUMMER in (select column_value from table(apex_string.split(p_liferant, ':')))
                    and     pa.code like 'M%'
                    and     a.EINLESUNG_STATUS = 'Y'
                    join    pd_region r on r.id = a.regionalbereich_id
                    where   EINHEITSPREIS > 0
                    and     r.id in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
                    group by auftrag_id,UMZETZUNG_CODE,UMSETZUNG_CODE
                    ) ap on (m.position_kennung2 = ap.UMSETZUNG_CODE)
         group by m.id,m.position_kennung,m.code,m.name,m.description,m.MUSTER_TYP_ID,m.einheit,m.parent_id
         )
         SELECT case 
                    when m.parent_tree is null then 'MASTER'
                    when m.parent_id = '01' then 'PARENT'
                    else 'CHILD' 
                end PARENT_MASTER,
                m.code POSITION,
                m.name as POS_TEXT,
                m.EINHEIT,
                case when m.parent_tree is null then null
                     when m.parent_id = '01' then 'Minimum'
                     else m.MINIMUM_PREIS end MINIMUM_PREIS,
                case when m.parent_tree is null then null
                     when m.parent_id = '01' then 'Mittelwert'
                     else m.MITTELWERT_PREIS end MITTELWERT_PREIS,
                case when m.parent_tree is null then null
                     when m.parent_id = '01' then 'Maximum'
                     else m.MAXIMUM_PREIS end MAXIMUM_PREIS,
                case when m.parent_tree is null then null
                     when m.parent_id = '01' then 'Median'
                     else m.MEDIAN_PREIS end MEDIAN_PREIS,
                case when m.parent_tree is null then null
                     when m.parent_id = '01' then null
                     else m.ANZAHL end ANZAHL,
                case when m.parent_tree is null then 'MASTER'
                     when m.parent_id = '01' then 'PARENT'
                     else m.position_kennung end position_kennung
         FROM mustern m
         START WITH m.id in (select column_value from table(apex_string.split_numbers(p_lvs, ':')))
         CONNECT BY PRIOR id_tree = parent_tree
         ORDER SIBLINGS BY m.code
         ) 
     loop

        IF j.PARENT_MASTER = 'PARENT' then
            cs_master:=xlsx_writer_v2.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_parent, border_id => border_db_full);
        ELSIF j.PARENT_MASTER = 'MASTER' then
            cs_master:=xlsx_writer_v2.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_master, border_id => border_db_full);
        ELSE
            cs_master:=cs_border;
        END IF;

        xlsx_writer_v2.add_cell(workbook, sheet_1, l_row+c_y_split, 1,style_id => cs_master, text => trim(j.POSITION));
        xlsx_writer_v2.add_cell(workbook, sheet_1, l_row+c_y_split, 2,style_id => cs_master, text => trim(j.POS_TEXT));
        xlsx_writer_v2.add_cell(workbook, sheet_1, l_row+c_y_split, 3,style_id => cs_master, text => trim(j.EINHEIT));

        if upper(j.MINIMUM_PREIS)='MINIMUM' or j.PARENT_MASTER in ('PARENT','MASTER') then
            xlsx_writer_v2.col_width(workbook, sheet_1, 4, 15);
            xlsx_writer_v2.add_cell(workbook, sheet_1, l_row+c_y_split, 4,style_id => cs_master, text => trim(j.MINIMUM_PREIS));
        else
            xlsx_writer_v2.add_cell(workbook, sheet_1, l_row+c_y_split, 4,style_id => number_format, value_ => to_number(replace(j.MINIMUM_PREIS,' ')));
        end if;

        if upper(j.MITTELWERT_PREIS)='MITTELWERT' or j.PARENT_MASTER in ('PARENT','MASTER') then
            xlsx_writer_v2.col_width(workbook, sheet_1, 5, 15);
            xlsx_writer_v2.add_cell(workbook, sheet_1, l_row+c_y_split, 5,style_id => cs_master, text => trim(j.MITTELWERT_PREIS));
        else
            xlsx_writer_v2.add_cell(workbook, sheet_1, l_row+c_y_split, 5,style_id => number_format, value_ => to_number(replace(j.MITTELWERT_PREIS,' ')));
        end if;

        if upper(j.MEDIAN_PREIS)='MEDIAN' or j.PARENT_MASTER in ('PARENT','MASTER') then
            xlsx_writer_v2.col_width(workbook, sheet_1, 6, 15);
            xlsx_writer_v2.add_cell(workbook, sheet_1, l_row+c_y_split, 6,style_id => cs_master, text => trim(j.MEDIAN_PREIS));
        else
            xlsx_writer_v2.add_cell(workbook, sheet_1, l_row+c_y_split, 6,style_id => number_format, value_ => to_number(replace(j.MEDIAN_PREIS,' ')));
        end if;

        if upper(j.MAXIMUM_PREIS)='MAXIMUM' or j.PARENT_MASTER in ('PARENT','MASTER') then
            xlsx_writer_v2.col_width(workbook, sheet_1, 7, 15);
            xlsx_writer_v2.add_cell(workbook, sheet_1, l_row+c_y_split, 7,style_id => cs_master, text => trim(j.MAXIMUM_PREIS));
        else
            xlsx_writer_v2.add_cell(workbook, sheet_1, l_row+c_y_split, 7,style_id => number_format, value_ => to_number(replace(j.MAXIMUM_PREIS,' ')));
        end if;

        if j.ANZAHL = 0 then
            cs_master := cs_rot;
        elsif j.ANZAHL > 0 and j.ANZAHL < 5 then
            cs_master := cs_gelb;
        elsif j.ANZAHL > 4 and j.ANZAHL < 10 then
            cs_master := cs_hellgelb;
        elsif j.ANZAHL > 9 and j.ANZAHL < 25 then
            cs_master := cs_hellgruen;
        elsif j.ANZAHL > 24 then
            cs_master := cs_gruen;
        end if;
        xlsx_writer_v2.col_width(workbook, sheet_1, 8,9);
        xlsx_writer_v2.add_cell(workbook, sheet_1, l_row+c_y_split, 8,style_id => cs_master, value_ => j.ANZAHL);

        v_kennungen.extend;
        v_kennungen(v_kennungen.count).kennung := replace(replace(j.position_kennung,'-'),'_');
        v_kennungen(v_kennungen.count).row_position := l_row+c_y_split;

        l_row:=l_row+1;
    end loop;

    l_muster_count := v_kennungen.count;
    DBS_LOGGING.LOG_DEBUG_AT('PREISDATENBANK_PKG.auswertung_to_excel: prepared muster rows count = '||l_muster_count,'AUSWERTUNG');

    for i in 1..v_auftrage.count loop
        l_row := v_kennungen(1).row_position;

        for j in (  
                    with  mustern as 
                    (   select  m.id,
                                m.code,
                                m.position_kennung2 position_kennung,
                                m.MUSTER_TYP_ID || m.code id_tree, 
                                decode(m.parent_id,null,null,m.MUSTER_TYP_ID||m.parent_id) parent_tree,
                                m.parent_id
                     from       PD_MUSTER_LVS m
                     left join (    select  avg(pa.EINHEITSPREIS) EINHEITSPREIS,
                                            pa.auftrag_id,
                                            pa.UMSETZUNG_CODE
                                            from    pd_auftrag_positionen pa
                                            join    pd_auftraege a on (a.id = pa.auftrag_id)
                                            and     a.datum between p_von and p_bis
                                            and     a.KEDITOREN_NUMMER in (select column_value from table(apex_string.split(p_liferant, ':')))
                                            and     pa.code like 'M%'
                                            and     a.EINLESUNG_STATUS = 'Y'
                                            where   pa.EINHEITSPREIS > 0
                                            and     a.regionalbereich_id in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
                                            group by auftrag_id,UMSETZUNG_CODE
                                ) ap on (m.position_kennung = ap.UMSETZUNG_CODE)
                     group by m.id,m.position_kennung2,m.code,m.name,m.description,m.MUSTER_TYP_ID,m.einheit,m.parent_id
                    )
                    select  muster.position_kennung,
                            pos.einheitspreis
                    from
                    (
                        SELECT  case
                                    when m.parent_tree is null then 'MASTER'
                                    when m.parent_id = '01' then 'PARENT'
                                    else m.position_kennung 
                                end position_kennung,
                                m.code
                        FROM   mustern m
                        START WITH m.id in (select column_value from table(apex_string.split_numbers(p_lvs, ':')))
                        CONNECT BY PRIOR id_tree = parent_tree
                    ) muster
                    left join ( select  round(avg(EINHEITSPREIS),2) einheitspreis,
                                        umsetzung_code
                                from    pd_auftrag_positionen
                                where   auftrag_id = v_auftrage(i).auftrag_id
                                and     EINHEITSPREIS > 0
                                group by UMSETZUNG_CODE
                              ) pos on (muster.position_kennung = pos.UMSETZUNG_CODE)
                    order by muster.code
                    )
        loop
            IF j.position_kennung = 'PARENT' then
                cs_master:=cs_master_parent;
            ELSIF j.position_kennung = 'MASTER' then
                cs_master:=cs_master_master;
            ELSE
                cs_master:=number_format;
            END IF;

            if j.einheitspreis is not null then
                xlsx_writer_v2.add_cell(workbook, sheet_1, l_row, v_auftrage(i).column_position,style_id => cs_master, value_ => j.einheitspreis);
            elsif j.position_kennung in ('PARENT','MASTER') then
                xlsx_writer_v2.add_cell(workbook, sheet_1, l_row, v_auftrage(i).column_position,style_id => cs_master, text => '');
            end if;

            l_row := l_row + 1;
        end loop;--j
    end loop;--i in 1..v_auftrage.count

    DBS_LOGGING.LOG_DEBUG_AT('PREISDATENBANK_PKG.auswertung_to_excel: populated price matrix for '||v_auftrage.count||' contracts','AUSWERTUNG');

    xlsx_writer_v2.freeze_sheet(workbook, sheet_1, c_x_split+5, c_y_split);
    xlsx     := xlsx_writer_v2.create_xlsx(workbook,2);

    --Mailversand der Auswertung
    DBS_LOGGING.LOG_INFO_AT('PREISDATENBANK_PKG.auswertung_to_excel: sending mail to user_id='||nvl(to_char(p_user_id),'-')||' with xlsx size='||nvl(to_char(case when xlsx is not null then dbms_lob.getlength(xlsx) end),'0'),'AUSWERTUNG');

    SendMailAuswertung(p_user_id => p_user_id,p_anhang => xlsx,p_filename => 'Auswertung.xlsx');

    if dbms_lob.istemporary(xlsx) = 1 then
        dbms_lob.freetemporary(xlsx);
    end if;
    if dbms_lob.istemporary(l_blob) = 1 then
        dbms_lob.freetemporary(l_blob);
    end if;

exception when others then
        if dbms_lob.istemporary(xlsx) = 1 then
            dbms_lob.freetemporary(xlsx);
        end if;
        if dbms_lob.istemporary(l_blob) = 1 then
            dbms_lob.freetemporary(l_blob);
        end if;

        DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.auswertung_to_excel: Fehler bei auswertung: ' || SQLCODE || ': ' || SQLERRM ||
      ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'AUSWERTUNG');

        SendMailAuswertungFehler
            (
            p_user_id => p_user_id,
            p_error => 'PREISDATENBANK_PKG.export_ausschreibung_to_excel: Fehler bei auswertung: ' || SQLCODE || ': ' || SQLERRM ||' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE
            );

end auswertung_to_excel;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
procedure auswertung_to_excel_2(p_von date,
                                p_bis date,
                                p_lvs varchar2,
                                p_regionen varchar2,
                                p_liferant varchar2,
                                p_user_id number) as

       workbook xlsx_writer_v2.book_r;
       sheet_1  integer;

       xlsx     blob;

       c_x_split constant integer := 3;
       c_y_split constant integer := 14;
       c_y_region constant integer := 1;

       cs_center_wrapped integer;
       cs_center integer;
       cs_center_bold integer;
       cs_center_bold_white integer;
       cs_center_bold_grey integer;
       cs_border integer;
       cs_master_parent integer;
       cs_master_master integer;
       cs_rot integer;
       cs_gelb integer;
       cs_hellgelb integer;
       cs_hellgruen integer;
       cs_gruen integer;

       fill_gelb integer;
       fill_hellgelb integer;
       fill_hellgruen integer;
       fill_gruen integer;
       fill_rot integer;
       fill_db integer;
       fill_db_grey integer;
       fill_master integer;
       fill_parent integer;

       font_db  integer;
        font_db_small integer;
       font_db_bold integer;
       font_db_bold_white integer;
       border_db integer;
       border_db_full integer;

       number_format integer;
       center_number_format integer;
       datum_format integer;

       type t_contract_rec is record(
            region            varchar2(100),
            bezeichnung       varchar2(4000),
            auftragnahmer     varchar2(4000),
            lvdatum           date,
            vergabesumme      number,
            auftrag_id        number,
            kreditorennummer  varchar2(400),
            sap_nr            varchar2(400),
            vertrag_nr        varchar2(400),
            anzahl            varchar2(200),
            prozentual_anzahl varchar2(200),
            lv_code           varchar2(4000),
            column_position   number
       );
       type t_contract_tab is table of t_contract_rec;

       type t_muster_rec is record(
            parent_master     varchar2(10),
            position          varchar2(200),
            pos_text          varchar2(4000),
            einheit           varchar2(50),
            minimum_preis     number,
            mittelwert_preis  number,
            median_preis      number,
            maximum_preis     number,
            anzahl            number,
            position_kennung  varchar2(400),
            sanitized_kennung varchar2(400),
            row_position      number
       );
       type t_muster_tab is table of t_muster_rec;

       type t_price_rec is record(
            auftrag_id       number,
            position_kennung varchar2(400),
            einheitspreis    number
       );
       type t_price_tab is table of t_price_rec;

       type t_price_map is table of number index by varchar2(800);

       l_contracts        t_contract_tab;
       l_mustern          t_muster_tab;
       l_price_rows       t_price_tab;
       l_price_map        t_price_map;

       l_column           number := 1;
       l_contract_count   number;
       l_row_index        number;
       l_col_index        number;
       l_style            integer;
       l_count_style      integer;
       l_price_key        varchar2(800);

       l_blob             blob;

begin

    DBS_LOGGING.LOG_INFO_AT('PREISDATENBANK_PKG.auswertung_to_excel_2 START: p_von='||to_char(p_von,'YYYY-MM-DD')||', p_bis='||to_char(p_bis,'YYYY-MM-DD')||', p_lvs='||nvl(p_lvs,'')||', p_regionen='||nvl(p_regionen,'')||', p_liferant='||nvl(p_liferant,'')||', p_user_id='||nvl(to_char(p_user_id),'-'),'AUSWERTUNG');

    workbook := xlsx_writer_v2.start_book;
    sheet_1  := xlsx_writer_v2.add_sheet  (workbook, 'Auswertung Vertrags LVS');

    dbms_lob.createtemporary(lob_loc => l_blob, cache => true, dur => dbms_lob.call);

    font_db := xlsx_writer_v2.add_font(workbook, 'DB Office', 10);
    font_db_small := xlsx_writer_v2.add_font(workbook, 'DB Office', 7);
    font_db_bold := xlsx_writer_v2.add_font(workbook, 'DB Office', 10, b => true);
    font_db_bold_white := xlsx_writer_v2.add_font(workbook, 'DB Office', 10, color=> 'theme="0"', b => true);

    border_db := xlsx_writer_v2.add_border(workbook, '<left/><right/><top/><bottom/><diagonal/>');
    border_db_full := xlsx_writer_v2.add_border(workbook, '<left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/>' );

    fill_db := xlsx_writer_v2.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="00ccff"/><bgColor indexed="64"/></patternFill>');
    fill_db_grey := xlsx_writer_v2.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="d9d9d9"/><bgColor indexed="64"/></patternFill>');
    fill_master := xlsx_writer_v2.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="92cddc"/><bgColor indexed="64"/></patternFill>');
    fill_parent := xlsx_writer_v2.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="daeef3"/><bgColor indexed="64"/></patternFill>');
    fill_gelb := xlsx_writer_v2.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="ffc000"/><bgColor indexed="64"/></patternFill>');
    fill_hellgelb := xlsx_writer_v2.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="ffff00"/><bgColor indexed="64"/></patternFill>');
    fill_hellgruen := xlsx_writer_v2.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="00ff00"/><bgColor indexed="64"/></patternFill>');
    fill_gruen := xlsx_writer_v2.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="00b050"/><bgColor indexed="64"/></patternFill>');
    fill_rot := xlsx_writer_v2.add_fill(workbook, '<patternFill patternType="solid"><fgColor rgb="ff0000"/><bgColor indexed="64"/></patternFill>');

    cs_center_wrapped := xlsx_writer_v2.add_cell_style(workbook, vertical_alignment => 'center', vertical_horizontal => 'center', wrap_text => true, font_id => font_db_small, border_id => border_db_full);
    cs_center := xlsx_writer_v2.add_cell_style(workbook, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_db_full);
    cs_center_bold := xlsx_writer_v2.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_db, border_id => border_db_full);
    cs_center_bold_white := xlsx_writer_v2.add_cell_style(workbook, font_id => font_db_bold_white, fill_id => fill_db, vertical_alignment => 'center', vertical_horizontal => 'center', border_id => border_db_full);
    cs_center_bold_grey := xlsx_writer_v2.add_cell_style(workbook, fill_id => fill_db_grey, border_id => border_db_full, font_id => font_db_bold);
    cs_border := xlsx_writer_v2.add_cell_style(workbook, border_id => border_db_full);
    cs_master_parent := xlsx_writer_v2.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_parent, border_id => border_db_full);
    cs_master_master := xlsx_writer_v2.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_master, border_id => border_db_full);
    cs_rot := xlsx_writer_v2.add_cell_style(workbook, fill_id => fill_rot, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db);
    cs_gelb := xlsx_writer_v2.add_cell_style(workbook, fill_id => fill_gelb, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db);
    cs_hellgelb := xlsx_writer_v2.add_cell_style(workbook, fill_id => fill_hellgelb, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db);
    cs_hellgruen := xlsx_writer_v2.add_cell_style(workbook, fill_id => fill_hellgruen, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db);
    cs_gruen := xlsx_writer_v2.add_cell_style(workbook, fill_id => fill_gruen, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db);

    number_format := xlsx_writer_v2.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer_v2."#.##0.00 €", font_id => font_db);
    center_number_format := xlsx_writer_v2.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer_v2."#.##0.00 €", font_id => font_db, vertical_alignment => 'center', vertical_horizontal => 'center');
    datum_format := xlsx_writer_v2.add_cell_style(workbook, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_db_full, num_fmt_id => xlsx_writer_v2."mm-dd-yy");

    xlsx_writer_v2.add_cell(workbook, sheet_1,  1, 1, style_id => cs_center_bold, text => 'I.) Grundangaben');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  1, 2, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  1, 3, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  2, 1, style_id => cs_border, text => 'Region');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  2, 2, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  2, 3, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  3, 1, style_id => cs_border, text => 'Bezeichnung der Maßnahme');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  3, 2, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  3, 3, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  4, 1, style_id => cs_border, text => 'Auftragnehmer');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  4, 2, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  4, 3, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  5, 1, style_id => cs_border, text => 'LV-Datum');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  5, 2, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  5, 3, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  6, 1, style_id => cs_border, text => 'Vergabesumme');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  6, 2, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  6, 3, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  7, 1, style_id => cs_border, text => 'Vergabevorgangsnummer');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  7, 2, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  7, 3, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  8, 1, style_id => cs_border, text => 'SAP-Kontraktnummer');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  8, 2, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  8, 3, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  9, 1, style_id => cs_border, text => 'Kreditorennummer');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  9, 2, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1,  9, 3, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 10, 1, style_id => cs_center_bold, text => 'II.) Einheitspreise');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 10, 2, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 10, 3, style_id => cs_border, text => '');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 11, 1, style_id => cs_center_bold_grey, text => 'Pos.');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 11, 2, style_id => cs_center_bold_grey, text => 'Pos.-Text');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 11, 3, style_id => cs_center_bold_grey, text => 'Einheit');

    xlsx_writer_v2.col_width(workbook, sheet_1, 1, 10);
    xlsx_writer_v2.col_width(workbook, sheet_1, 2, 50);
    xlsx_writer_v2.col_width(workbook, sheet_1, 3, 10);

    select  r.code region,
            a.projekt_desc bezeichnung,
            a.auftragnahmer_name auftragnahmer,
            a.datum lvdatum,
            a.total vergabesumme,
            a.id auftrag_id,
            a.keditoren_nummer,
            a.sap_nr,
            a.vertrag_nr,
            nvl(kennung_anzahl,0)||' von '||gesamt_anzahl anzahl,
            round((nvl(kennung_anzahl,0)/gesamt_anzahl)*100,2)||'%' prozentual_anzahl,
            listagg(lvs.lv_code, ', ') within group (order by lvs.lv_code) lv_code,
            null column_position
    bulk collect into l_contracts
    from    pd_auftraege a
            join (select count(*) gesamt_anzahl, auftrag_id from pd_auftrag_positionen group by auftrag_id) all_positionen
                on all_positionen.auftrag_id = a.id
            join (select count(*) kennung_anzahl, auftrag_id
                  from pd_auftrag_positionen
                  where umsetzung_code like 'MLV%'
                  group by auftrag_id) lv_positionen
                on lv_positionen.auftrag_id = a.id
            join pd_region r on r.id = a.regionalbereich_id
            join pd_auftrag_lvs lvs on lvs.auftrag_id = a.id
    where   a.datum between p_von and p_bis
    and     a.einlesung_status = 'Y'
    and     r.id in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
    and     a.keditoren_nummer in (select column_value from table(apex_string.split(p_liferant, ':')))
    and     (
                select count(*)
                from   pd_auftrag_positionen ap
                       cross join pd_muster_lvs m
                where  m.id in (select column_value from table(apex_string.split_numbers(p_lvs, ':')))
                and    ap.auftrag_id = a.id
                and    ap.code like 'M%'
                and    ap.einheitspreis > 0
                and    instr(ap.umsetzung_code, m.position_kennung2) > 0
            ) > 0
    group by r.code, a.projekt_desc, a.auftragnahmer_name, a.datum, a.total, a.id, a.keditoren_nummer, a.sap_nr, a.vertrag_nr, kennung_anzahl, gesamt_anzahl
    order by a.auftragnahmer_name;
    DBS_LOGGING.LOG_DEBUG_AT('PREISDATENBANK_PKG.auswertung_to_excel_2: fetched contracts count = '||l_contracts.count,'AUSWERTUNG');

    if l_contracts.count > 0 then
        for idx in 1..l_contracts.count loop
            l_col_index := c_x_split + l_column + 5;
            l_contracts(idx).column_position := l_col_index;

            xlsx_writer_v2.col_width(workbook, sheet_1, c_x_split + l_column + 4, 15);
            xlsx_writer_v2.add_cell(workbook, sheet_1, c_y_region, l_col_index, style_id => cs_center_bold_white, text => 'Test');
            xlsx_writer_v2.add_cell(workbook, sheet_1, c_y_region + 12, l_col_index, style_id => cs_center_bold_white, text => '');
            xlsx_writer_v2.add_cell(workbook, sheet_1, c_y_region + 13, l_col_index, style_id => cs_center_bold_grey, text => '');

            xlsx_writer_v2.add_cell(workbook, sheet_1, c_y_region + 1, l_col_index, style_id => cs_center, text => trim(l_contracts(idx).region));
            xlsx_writer_v2.add_cell(workbook, sheet_1, c_y_region + 2, l_col_index, style_id => cs_center_wrapped, text => trim(l_contracts(idx).bezeichnung));
            xlsx_writer_v2.add_cell(workbook, sheet_1, c_y_region + 3, l_col_index, style_id => cs_center, text => trim(l_contracts(idx).auftragnahmer));
            xlsx_writer_v2.add_cell(workbook, sheet_1, c_y_region + 4, l_col_index, style_id => cs_center, text => l_contracts(idx).lvdatum);
            xlsx_writer_v2.add_cell(workbook, sheet_1, c_y_region + 5, l_col_index, style_id => center_number_format, value_ => l_contracts(idx).vergabesumme);
            xlsx_writer_v2.add_cell(workbook, sheet_1, c_y_region + 6, l_col_index, style_id => cs_center, text => l_contracts(idx).vertrag_nr);
            xlsx_writer_v2.add_cell(workbook, sheet_1, c_y_region + 7, l_col_index, style_id => cs_center, text => l_contracts(idx).sap_nr);
            xlsx_writer_v2.add_cell(workbook, sheet_1, c_y_region + 8, l_col_index, style_id => cs_center, text => l_contracts(idx).kreditorennummer);
            xlsx_writer_v2.add_cell(workbook, sheet_1, c_y_region + 9, l_col_index, style_id => cs_center, text => l_contracts(idx).anzahl);
            xlsx_writer_v2.add_cell(workbook, sheet_1, c_y_region + 10, l_col_index, style_id => cs_center, text => l_contracts(idx).prozentual_anzahl);
            xlsx_writer_v2.add_cell(workbook, sheet_1, c_y_region + 11, l_col_index, style_id => cs_center, text => l_contracts(idx).lv_code);

            l_column := l_column + 1;
        end loop;
    end if;

    l_contract_count := l_column - 1;

    xlsx_writer_v2.add_cell(workbook, sheet_1, 14, 4, style_id => cs_center_bold_grey, text => 'Minimum');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 14, 5, style_id => cs_center_bold_grey, text => 'Mittelwert');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 14, 6, style_id => cs_center_bold_grey, text => 'Median');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 14, 7, style_id => cs_center_bold_grey, text => 'Maximum');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 13, 4, style_id => cs_center_bold_grey, value_ => l_contract_count);
    xlsx_writer_v2.add_cell(workbook, sheet_1, 13, 5, style_id => cs_center_bold_grey, text => 'Vergaben');

    xlsx_writer_v2.add_cell(workbook, sheet_1, 8, 8, style_id => cs_rot, text => '= 0');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 9, 8, style_id => cs_gelb, text => '= 1 - 4');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 10, 8, style_id => cs_hellgelb, text => '= 5 - 9');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 11, 8, style_id => cs_hellgruen, text => '= 10 - 24');
    xlsx_writer_v2.add_cell(workbook, sheet_1, 12, 8, style_id => cs_gruen, text => '≥ 25');

    select parent_master,
           position,
           pos_text,
           einheit,
           minimum_preis,
           mittelwert_preis,
           median_preis,
           maximum_preis,
           anzahl,
           position_kennung,
           sanitized_kennung,
           null row_position
    bulk collect into l_mustern
    from (
            with mustern as (
                select  m.id,
                        m.code,
                        m.position_kennung,
                        m.position_kennung2,
                        m.name,
                        m.muster_typ_id || m.code id_tree,
                        decode(m.parent_id, null, null, m.muster_typ_id || m.parent_id) parent_tree,
                        m.parent_id,
                        m.einheit,
                        round(min(ap.einheitspreis), 2) minimum_preis,
                        round(avg(ap.einheitspreis), 2) mittelwert_preis,
                        round(median(ap.einheitspreis), 2) median_preis,
                        round(max(ap.einheitspreis), 2) maximum_preis,
                        count(ap.umsetzung_code) anzahl
                from    pd_muster_lvs m
                left join (
                            select avg(pa.einheitspreis) einheitspreis,
                                   pa.auftrag_id,
                                   pa.umsetzung_code
                            from   pd_auftrag_positionen pa
                                   join pd_auftraege a on a.id = pa.auftrag_id
                                   join pd_region r on r.id = a.regionalbereich_id
                            where  a.datum between p_von and p_bis
                            and    a.keditoren_nummer in (select column_value from table(apex_string.split(p_liferant, ':')))
                            and    pa.code like 'M%'
                            and    a.einlesung_status = 'Y'
                            and    pa.einheitspreis > 0
                            and    r.id in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
                            group by pa.auftrag_id, pa.umsetzung_code
                          ) ap on m.position_kennung2 = ap.umsetzung_code
                group by m.id, m.code, m.name, m.position_kennung, m.position_kennung2, m.muster_typ_id, m.parent_id, m.einheit, m.muster_typ_id || m.code, decode(m.parent_id, null, null, m.muster_typ_id || m.parent_id)
            )
            select case
                        when m.parent_tree is null then 'MASTER'
                        when m.parent_id = '01' then 'PARENT'
                        else 'CHILD'
                   end parent_master,
                   m.code position,
                   m.name pos_text,
                   m.einheit,
                   case when m.parent_tree is null then null else m.minimum_preis end minimum_preis,
                   case when m.parent_tree is null then null else m.mittelwert_preis end mittelwert_preis,
                   case when m.parent_tree is null then null else m.median_preis end median_preis,
                   case when m.parent_tree is null then null else m.maximum_preis end maximum_preis,
                   case when m.parent_tree is null or m.parent_id = '01' then null else m.anzahl end anzahl,
                   case when m.parent_tree is null then 'MASTER'
                        when m.parent_id = '01' then 'PARENT'
                        else m.position_kennung end position_kennung,
                   case when m.parent_tree is null then 'MASTER'
                        when m.parent_id = '01' then 'PARENT'
                        else replace(replace(m.position_kennung,'-'),'_') end sanitized_kennung
            from mustern m
            start with m.id in (select column_value from table(apex_string.split_numbers(p_lvs, ':')))
            connect by prior id_tree = parent_tree
            order siblings by m.code
         );

    xlsx_writer_v2.col_width(workbook, sheet_1, 4, 15);
    xlsx_writer_v2.col_width(workbook, sheet_1, 5, 15);
    xlsx_writer_v2.col_width(workbook, sheet_1, 6, 15);
    xlsx_writer_v2.col_width(workbook, sheet_1, 7, 15);
    xlsx_writer_v2.col_width(workbook, sheet_1, 8, 9);

    if l_mustern.count > 0 then
        for idx in 1..l_mustern.count loop
            l_row_index := c_y_split + idx;

            if l_mustern(idx).parent_master = 'PARENT' then
                l_style := cs_master_parent;
            elsif l_mustern(idx).parent_master = 'MASTER' then
                l_style := cs_master_master;
            else
                l_style := cs_border;
            end if;

            xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 1, style_id => l_style, text => trim(l_mustern(idx).position));
            xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 2, style_id => l_style, text => trim(l_mustern(idx).pos_text));
            xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 3, style_id => l_style, text => trim(l_mustern(idx).einheit));

            if l_mustern(idx).parent_master = 'PARENT' then
                xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 4, style_id => l_style, text => 'Minimum');
                xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 5, style_id => l_style, text => 'Mittelwert');
                xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 6, style_id => l_style, text => 'Median');
                xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 7, style_id => l_style, text => 'Maximum');
            elsif l_mustern(idx).parent_master = 'MASTER' then
                xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 4, style_id => l_style, text => '');
                xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 5, style_id => l_style, text => '');
                xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 6, style_id => l_style, text => '');
                xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 7, style_id => l_style, text => '');
            else
                if l_mustern(idx).minimum_preis is not null then
                    xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 4, style_id => number_format, value_ => l_mustern(idx).minimum_preis);
                end if;
                if l_mustern(idx).mittelwert_preis is not null then
                    xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 5, style_id => number_format, value_ => l_mustern(idx).mittelwert_preis);
                end if;
                if l_mustern(idx).median_preis is not null then
                    xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 6, style_id => number_format, value_ => l_mustern(idx).median_preis);
                end if;
                if l_mustern(idx).maximum_preis is not null then
                    xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 7, style_id => number_format, value_ => l_mustern(idx).maximum_preis);
                end if;
            end if;

            if l_mustern(idx).anzahl is null then
                l_count_style := l_style;
                xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 8, style_id => l_count_style, text => '');
            elsif l_mustern(idx).anzahl = 0 then
                l_count_style := cs_rot;
                xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 8, style_id => l_count_style, value_ => l_mustern(idx).anzahl);
            elsif l_mustern(idx).anzahl between 1 and 4 then
                l_count_style := cs_gelb;
                xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 8, style_id => l_count_style, value_ => l_mustern(idx).anzahl);
            elsif l_mustern(idx).anzahl between 5 and 9 then
                l_count_style := cs_hellgelb;
                xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 8, style_id => l_count_style, value_ => l_mustern(idx).anzahl);
            elsif l_mustern(idx).anzahl between 10 and 24 then
                l_count_style := cs_hellgruen;
                xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 8, style_id => l_count_style, value_ => l_mustern(idx).anzahl);
            else
                l_count_style := cs_gruen;
                xlsx_writer_v2.add_cell(workbook, sheet_1, l_row_index, 8, style_id => l_count_style, value_ => l_mustern(idx).anzahl);
            end if;

            l_mustern(idx).row_position := l_row_index;
        end loop;
    end if;

    select auftrag_id,
           umsetzung_code position_kennung,
           round(avg(einheitspreis), 2) einheitspreis
    bulk collect into l_price_rows
    from pd_auftrag_positionen pa
         join pd_auftraege a on a.id = pa.auftrag_id
         join pd_region r on r.id = a.regionalbereich_id
    where a.datum between p_von and p_bis
    and   a.einlesung_status = 'Y'
    and   a.keditoren_nummer in (select column_value from table(apex_string.split(p_liferant, ':')))
    and   r.id in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
    and   pa.einheitspreis > 0
    and   pa.code like 'M%'
    group by auftrag_id, umsetzung_code;
    DBS_LOGGING.LOG_DEBUG_AT('PREISDATENBANK_PKG.auswertung_to_excel_2: fetched price rows count = '||l_price_rows.count,'AUSWERTUNG');

    if l_price_rows.count > 0 then
        for idx in 1..l_price_rows.count loop
            if l_price_rows(idx).position_kennung is not null then
                l_price_key := l_price_rows(idx).auftrag_id || '|' || replace(replace(l_price_rows(idx).position_kennung,'-'),'_');
                l_price_map(l_price_key) := l_price_rows(idx).einheitspreis;
            end if;
        end loop;
    end if;

    if l_contracts.count > 0 and l_mustern.count > 0 then
        for idx in 1..l_contracts.count loop
            for jdx in 1..l_mustern.count loop
                if l_mustern(jdx).parent_master = 'PARENT' then
                    xlsx_writer_v2.add_cell(workbook, sheet_1, l_mustern(jdx).row_position, l_contracts(idx).column_position, style_id => cs_master_parent, text => '');
                elsif l_mustern(jdx).parent_master = 'MASTER' then
                    xlsx_writer_v2.add_cell(workbook, sheet_1, l_mustern(jdx).row_position, l_contracts(idx).column_position, style_id => cs_master_master, text => '');
                else
                    l_price_key := l_contracts(idx).auftrag_id || '|' || l_mustern(jdx).sanitized_kennung;
                    if l_price_map.exists(l_price_key) then
                        xlsx_writer_v2.add_cell(workbook, sheet_1, l_mustern(jdx).row_position, l_contracts(idx).column_position, style_id => number_format, value_ => l_price_map(l_price_key));
                    end if;
                end if;
            end loop;
        end loop;
    end if;

    xlsx_writer_v2.freeze_sheet(workbook, sheet_1, c_x_split + 5, c_y_split);
    xlsx := xlsx_writer_v2.create_xlsx(workbook,2);

    DBS_LOGGING.LOG_INFO_AT('PREISDATENBANK_PKG.auswertung_to_excel_2: sending mail to user_id='||nvl(to_char(p_user_id),'-')||' with xlsx size='||nvl(to_char(nvl(dbms_lob.getlength(xlsx),0)),'0'),'AUSWERTUNG');

    SendMailAuswertung(p_user_id => p_user_id, p_anhang => xlsx, p_filename => 'Auswertung.xlsx');

    if dbms_lob.istemporary(xlsx) = 1 then
        dbms_lob.freetemporary(xlsx);
    end if;
    if dbms_lob.istemporary(l_blob) = 1 then
        dbms_lob.freetemporary(l_blob);
    end if;

exception
    when others then
        if dbms_lob.istemporary(xlsx) = 1 then
            dbms_lob.freetemporary(xlsx);
        end if;
        if dbms_lob.istemporary(l_blob) = 1 then
            dbms_lob.freetemporary(l_blob);
        end if;

        dbs_logging.log_error_at('PREISDATENBANK_PKG.auswertung_to_excel_2: Fehler bei auswertung: ' || SQLCODE || ': ' || SQLERRM ||
            ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'AUSWERTUNG');

        SendMailAuswertungFehler(
            p_user_id => p_user_id,
            p_error   => 'PREISDATENBANK_PKG.auswertung_to_excel_2: Fehler bei auswertung: ' || SQLCODE || ': ' || SQLERRM || ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE
        );
end auswertung_to_excel_2;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

procedure auswertung_to_excel_leser(p_von date, p_bis date, p_lvs varchar2,p_regionen varchar2,p_liferant varchar2,p_user_id number) as

       workbook xlsx_writer.book_r;
       sheet_1  integer;

       xlsx     blob;

       c_limit   constant integer := 50;
       c_x_split constant integer := 3;
       c_y_split constant integer := 11;
       c_y_region constant integer := 1;
       cs_center_wrapped integer;
       cs_center integer;
       cs_number integer;
       cs_center_bold integer;
       cs_center_bold_white integer;
       cs_center_bold_grey integer;
       cs_border integer;
       cs_master integer;
       cs_master_master integer;
       cs_master_parent integer;
       cs_rot integer;
       cs_gelb integer;
       cs_hellgelb integer;
       cs_hellgruen integer;
       cs_gruen integer;
       cs_noborder integer;
       fill_gelb integer;
       fill_hellgelb integer;
       fill_hellgruen integer;
       fill_gruen integer;
       fill_rot integer;

       font_db  integer;
       font_db_small integer;
       font_db_bold integer;
       fill_db integer;
       fill_db_grey integer;
       font_db_bold_white integer;
       fill_master integer;
       fill_parent integer;
       border_db integer;
       border_db_full integer;

       number_format integer;
       center_number_format integer;
       datum_format integer;

       TYPE curtype IS REF CURSOR;

       v_auftrage t_auftraege := new t_auftraege();
       v_kennungen T_KENNUNG := new T_KENNUNG();
       l_filename varchar2(100);
       L_BLOB blob;
       l_target_charset VARCHAR2(100) := 'WE8MSWIN1252';
       L_DEST_OFFSET    INTEGER := 1;
       L_SRC_OFFSET     INTEGER := 1;
       L_LANG_CONTEXT   INTEGER := DBMS_LOB.DEFAULT_LANG_CTX;
       L_WARNING        INTEGER;
       L_LENGTH         INTEGER;
       l_column         number:=1;
       l_row            number;
       l_betrag         number;
       l_random_number  number := floor(dbms_random.value(1, 1000000000));
begin

    workbook := xlsx_writer.start_book;
    sheet_1  := xlsx_writer.add_sheet  (workbook, 'Auswertung Vertrags LVS');

    dbms_lob.createtemporary(lob_loc => l_blob, cache => true, dur => dbms_lob.call);

    -- Style definition

    font_db := xlsx_writer.add_font      (workbook, 'DB Office', 10);
    border_db := xlsx_writer.add_border  (workbook, '<left/><right/><top/><bottom/><diagonal/>');
    border_db_full := xlsx_writer.add_border      (workbook, '<left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/>');
    fill_db:= xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="00ccff"/><bgColor indexed="64"/></patternFill>');
    fill_db_grey := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="d9d9d9"/><bgColor indexed="64"/></patternFill>');
    fill_master := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="92cddc"/><bgColor indexed="64"/></patternFill>');
    fill_parent := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="daeef3"/><bgColor indexed="64"/></patternFill>');
    fill_gelb := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="ffc000"/><bgColor indexed="64"/></patternFill>');
    fill_hellgelb := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="ffff00"/><bgColor indexed="64"/></patternFill>');
    fill_hellgruen := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="00ff00"/><bgColor indexed="64"/></patternFill>');
    fill_gruen := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="00b050"/><bgColor indexed="64"/></patternFill>');
    fill_rot := xlsx_writer.add_fill      (workbook, '<patternFill patternType="solid"><fgColor rgb="ff0000"/><bgColor indexed="64"/></patternFill>');
    font_db_small := xlsx_writer.add_font      (workbook, 'DB Office', 7);
    font_db_bold := xlsx_writer.add_font      (workbook, 'DB Office', 10, b => true);
    font_db_bold_white := xlsx_writer.add_font      (workbook, 'DB Office', 10, color=> 'theme="0"', b => true);

    --cs_center_wrapped  := xlsx_writer.add_cell_style(workbook, vertical_alignment => 'center', vertical_horizontal => 'center', wrap_text => true,font_id => font_db_small, border_id => border_db_full);
    --cs_center  := xlsx_writer.add_cell_style(workbook, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_db_full);
    cs_center_bold  := xlsx_writer.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_db, border_id => border_db_full);
    --cs_center_bold_white := xlsx_writer.add_cell_style(workbook, font_id => font_db_bold_white, fill_id => fill_db,vertical_alignment => 'center', vertical_horizontal => 'center', border_id => border_db_full);
    cs_border := xlsx_writer.add_cell_style(workbook, border_id => border_db_full);
    cs_noborder := xlsx_writer.add_cell_style(workbook, border_id => border_db);
    cs_center_bold_grey := xlsx_writer.add_cell_style(workbook, fill_id => fill_db_grey, border_id => border_db_full, font_id => font_db_bold);
    --cs_rot := xlsx_writer.add_cell_style(workbook, fill_id => fill_rot, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db);
    --cs_gelb := xlsx_writer.add_cell_style(workbook, fill_id => fill_gelb, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db);
    --cs_hellgelb := xlsx_writer.add_cell_style(workbook, fill_id => fill_hellgelb, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db);
    --cs_hellgruen := xlsx_writer.add_cell_style(workbook, fill_id => fill_hellgruen, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db);
    --cs_gruen := xlsx_writer.add_cell_style(workbook, fill_id => fill_gruen, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db);
    cs_master_parent := xlsx_writer.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_parent, border_id => border_db_full);
    cs_master_master := xlsx_writer.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_master, border_id => border_db_full);

    number_format := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00", font_id => font_db);
    --center_number_format := xlsx_writer.add_cell_style(workbook, border_id => border_db_full, num_fmt_id => xlsx_writer."#.##0.00" , font_id => font_db, vertical_alignment => 'center', vertical_horizontal => 'center');
    --datum_format := xlsx_writer.add_cell_style(workbook, vertical_alignment => 'center', vertical_horizontal => 'center', font_id => font_db, border_id => border_db_full, num_fmt_id => xlsx_writer."mm-dd-yy");

    xlsx_writer.add_cell(workbook, sheet_1,   1, 1,style_id => cs_center_bold, text => 'I.) Grundangaben');
    xlsx_writer.add_cell(workbook, sheet_1,   1, 2,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   1, 3,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   2, 1,style_id => cs_border, text => 'Region');
    xlsx_writer.add_cell(workbook, sheet_1,   2, 2,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   2, 3,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   3, 1,style_id => cs_border, text => 'Bezeichnung der Maßnahme');
    xlsx_writer.add_cell(workbook, sheet_1,   3, 2,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   3, 3,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   4, 1,style_id => cs_border, text => 'Auftragnehmer');
    xlsx_writer.add_cell(workbook, sheet_1,   4, 2,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   4, 3,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   5, 1,style_id => cs_border, text => 'LV-Datum');
    xlsx_writer.add_cell(workbook, sheet_1,   5, 2,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   5, 3,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   6, 1,style_id => cs_border, text => 'Vergabesumme');
    xlsx_writer.add_cell(workbook, sheet_1,   6, 2,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   6, 3,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   7, 1,style_id => cs_border, text => 'Vergabevorgangsnummer');
    xlsx_writer.add_cell(workbook, sheet_1,   7, 2,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   7, 3,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   8, 1,style_id => cs_border, text => 'SAP-Kontraktnummer');
    xlsx_writer.add_cell(workbook, sheet_1,   8, 2,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   8, 3,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   9, 1,style_id => cs_border, text => 'Kreditorennummer');
    xlsx_writer.add_cell(workbook, sheet_1,   9, 2,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   9, 3,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   10, 1,style_id => cs_center_bold, text => 'II.) Einheitspreise');
    xlsx_writer.add_cell(workbook, sheet_1,   10, 2,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   10, 3,style_id => cs_border, text => '');
    xlsx_writer.add_cell(workbook, sheet_1,   11, 1,style_id => cs_center_bold_grey, text => 'Pos.');
    xlsx_writer.add_cell(workbook, sheet_1,   11, 2,style_id => cs_center_bold_grey, text => 'Pos.-Text');
    xlsx_writer.add_cell(workbook, sheet_1,   11, 3,style_id => cs_center_bold_grey, text => 'Einheit');

    xlsx_writer.col_width(workbook, sheet_1, 1, 10);
    xlsx_writer.col_width(workbook, sheet_1, 2, 50);
    xlsx_writer.col_width(workbook, sheet_1, 3, 10);

    for i in 
    (
    with all_positionen as (select count(*) gesamt_anzahl, AUFTRAG_ID from pd_auftrag_positionen group by AUFTRAG_ID),
         lv_positionen as (select count(*) kennung_anzahl, AUFTRAG_ID from pd_auftrag_positionen where UMZETZUNG_CODE like 'MLV%' group by AUFTRAG_ID)
    select  r.code Region,
            a.projekt_desc Bezeichnung,
            a.auftragnahmer_name Auftragnahmer,
            a.datum LVDatum,
            a.total Vergabesumme,
            a.id,
            a.KEDITOREN_NUMMER,
            a.SAP_NR,
            a.VERTRAG_NR,
            nvl(kennung_anzahl,0)||' von '||gesamt_anzahl anzahl,
            round((nvl(kennung_anzahl,0)/gesamt_anzahl)*100,2)||'%' prozentual_anzahl,
            listagg(lvs.lv_code, ', ') within group (order by lvs.lv_code) lv_code
    from    pd_auftraege a
    join    pd_region r on r.id = a.regionalbereich_id
    join    lv_positionen ap on a.id = ap.auftrag_id
    join    pd_auftrag_lvs lvs on a.id = lvs.auftrag_id
    join    all_positionen allp on a.id = allp.AUFTRAG_ID
    where   a.datum between p_von and p_bis
    and     a.EINLESUNG_STATUS = 'Y'
    and     r.id in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
    and     a.KEDITOREN_NUMMER in (select column_value from table(apex_string.split(p_liferant, ':')))
    --keine Verträge in die Auswertung übernehmen, die nicht einen Preis für die Positionen haben
    and     (   select  count(*)
                from        pd_auftrag_positionen ap
                cross join  pd_muster_lvs m
                where   m.id in (select column_value from table(apex_string.split_numbers(p_lvs, ':')))
                and     ap.auftrag_id = a.id
                and     ap.code like 'M%'
                and     ap.einheitspreis > 0
                and     (instr(ap.UMSETZUNG_CODE,m.position_kennung2) > 0)
            ) > 0
	--keine Verträge in die Auswertung übernehmen, die nicht einen Preis für die Positionen haben
    group by r.code,a.projekt_desc,a.auftragnahmer_name,a.datum,a.total,a.id,a.KEDITOREN_NUMMER,a.SAP_NR,a.VERTRAG_NR,kennung_anzahl,gesamt_anzahl
    order by a.auftragnahmer_name
    )
    loop
        l_column:=l_column+1;
    end loop;

    xlsx_writer.add_cell(workbook, sheet_1,   11, 4,style_id => cs_center_bold_grey, text => 'Minimum');
    xlsx_writer.add_cell(workbook, sheet_1,   11, 5,style_id => cs_center_bold_grey, text => 'Mittelwert');
    xlsx_writer.add_cell(workbook, sheet_1,   11, 6,style_id => cs_center_bold_grey, text => 'Median');
    xlsx_writer.add_cell(workbook, sheet_1,   11, 7,style_id => cs_center_bold_grey, text => 'Maximum');
    xlsx_writer.add_cell(workbook, sheet_1,   10, 4,style_id => cs_center_bold_grey, value_ => l_column-1);
    xlsx_writer.add_cell(workbook, sheet_1,   10, 5,style_id => cs_center_bold_grey, text => 'Vergaben');
    l_row:=1;

    xlsx_writer.add_cell(workbook, sheet_1, 5, 8,style_id => cs_rot, text => '= 0');
    xlsx_writer.add_cell(workbook, sheet_1, 6, 8,style_id => cs_gelb, text => '= 1 - 4');
    xlsx_writer.add_cell(workbook, sheet_1, 7, 8,style_id => cs_hellgelb, text => '= 5 - 9');
    xlsx_writer.add_cell(workbook, sheet_1, 8, 8,style_id => cs_hellgruen, text => '= 10 - 24');
    xlsx_writer.add_cell(workbook, sheet_1, 9, 8,style_id => cs_gruen, text => '? 25');

    for j in (
              with  mustern as (select  m.id,
                                        m.code,
                                        m.position_kennung,
                                        m.name,
                                        m.description,
                                        m.MUSTER_TYP_ID || m.code id_tree, 
                                        decode(m.parent_id,null,null,m.MUSTER_TYP_ID||m.parent_id) parent_tree,
                                        m.parent_id,
                                        m.einheit,
                                        to_char(round(min(ap.EINHEITSPREIS),2),'999G999G999G999G990D00') MINIMUM_PREIS,
                                        to_char(round(avg(ap.EINHEITSPREIS),2),'999G999G999G999G990D00') MITTELWERT_PREIS, 
                                        to_char(round(median(ap.EINHEITSPREIS),2),'999G999G999G999G990D00') MEDIAN_PREIS,
                                        to_char(round(max(ap.EINHEITSPREIS),2),'999G999G999G999G990D00') MAXIMUM_PREIS,
                                        COUNT(ap.UMSETZUNG_CODE) ANZAHL
                              from PD_MUSTER_LVS m
                              left join (select avg(EINHEITSPREIS) EINHEITSPREIS,auftrag_id,UMSETZUNG_CODE 
                                          from pd_auftrag_positionen pa
                                          join pd_auftraege a on a.id = pa.auftrag_id and a.datum between p_von and p_bis
                                          and a.KEDITOREN_NUMMER in (select column_value  
                                                      from table(apex_string.split(p_liferant, ':')))
                                          and  pa.code like 'M%'
                                          and  a.EINLESUNG_STATUS = 'Y'
                                         left join pd_region r on r.id = a.regionalbereich_id
                                         where EINHEITSPREIS > 0
                                         and     r.id in (select column_value from table(apex_string.split_numbers(p_regionen, ':')))
                                         group by auftrag_id,UMSETZUNG_CODE) ap on
                                            (m.position_kennung2 = ap.UMSETZUNG_CODE)
                              group by m.id,m.position_kennung,m.code,m.name,m.description,m.MUSTER_TYP_ID,m.einheit,m.parent_id)
                                        SELECT case when m.parent_tree is null then 'MASTER'
                                                    when m.parent_id = '01' then 'PARENT'
                                                    else 'CHILD' end PARENT_MASTER,
                                               m.code POSITION,
                                               m.name as POS_TEXT,
                                               m.EINHEIT,
                                               case when m.parent_tree is null then null
                                                    when m.parent_id = '01' then 'Minimum'
                                                    else m.MINIMUM_PREIS end MINIMUM_PREIS,
                                               case when m.parent_tree is null then null
                                                    when m.parent_id = '01' then 'Mittelwert'
                                                    else m.MITTELWERT_PREIS end MITTELWERT_PREIS,
                                               case when m.parent_tree is null then null
                                                    when m.parent_id = '01' then 'Median'
                                                    else m.MEDIAN_PREIS end MEDIAN_PREIS,
                                               case when m.parent_tree is null then null
                                                    when m.parent_id = '01' then 'Maximum'
                                                    else m.MAXIMUM_PREIS end MAXIMUM_PREIS,
                                               case when m.parent_tree is null then null
                                                    when m.parent_id = '01' then null
                                                    else m.ANZAHL end ANZAHL,
                                               case when m.parent_tree is null then 'MASTER'
                                                    when m.parent_id = '01' then 'PARENT'
                                                    else m.position_kennung end position_kennung
                                        FROM mustern m
                                        START WITH m.id in (select column_value  
                                                      from table(apex_string.split_numbers(p_lvs, ':')))
                                        CONNECT BY PRIOR id_tree = parent_tree
                                        ORDER SIBLINGS BY m.code
    ) loop

                        IF j.PARENT_MASTER = 'PARENT' then
                            cs_master:=xlsx_writer.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_parent, border_id => border_db_full);
                        ELSIF j.PARENT_MASTER = 'MASTER' then
                            cs_master:=xlsx_writer.add_cell_style(workbook, font_id => font_db_bold, fill_id => fill_master, border_id => border_db_full);
                        ELSE
                            cs_master:=cs_border;
                        END IF;

                        xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 1,style_id => cs_master, text => trim(j.POSITION));
                        xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 2,style_id => cs_master, text => trim(j.POS_TEXT));
                        xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 3,style_id => cs_master, text => trim(j.EINHEIT));

                        if upper(j.MINIMUM_PREIS)='MINIMUM' or j.PARENT_MASTER in ('PARENT','MASTER') then
                                xlsx_writer.col_width(workbook, sheet_1, 4, 15);
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 4,style_id => cs_master, text => trim(j.MINIMUM_PREIS));
                        else
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 4,style_id => number_format, value_ => to_number(replace(j.MINIMUM_PREIS,' ')));
                        end if;

                        if upper(j.MITTELWERT_PREIS)='MITTELWERT' or j.PARENT_MASTER in ('PARENT','MASTER') then
                                xlsx_writer.col_width(workbook, sheet_1, 5, 15);
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 5,style_id => cs_master, text => trim(j.MITTELWERT_PREIS));
                        else
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 5,style_id => number_format, value_ => to_number(replace(j.MITTELWERT_PREIS,' ')));
                        end if;

                        if upper(j.MEDIAN_PREIS)='MEDIAN' or j.PARENT_MASTER in ('PARENT','MASTER') then
                                xlsx_writer.col_width(workbook, sheet_1, 6, 15);
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 6,style_id => cs_master, text => trim(j.MEDIAN_PREIS));
                        else
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 6,style_id => number_format, value_ => to_number(replace(j.MEDIAN_PREIS,' ')));
                        end if;

                        if upper(j.MAXIMUM_PREIS)='MAXIMUM' or j.PARENT_MASTER in ('PARENT','MASTER') then
                                xlsx_writer.col_width(workbook, sheet_1, 7, 15);
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 7,style_id => cs_master, text => trim(j.MAXIMUM_PREIS));
                        else
                                xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 7,style_id => number_format, value_ => to_number(replace(j.MAXIMUM_PREIS,' ')));
                        end if;

                        if j.ANZAHL = 0 then
                            cs_master := cs_rot;
                        elsif j.ANZAHL > 0 and j.ANZAHL < 5 then
                            cs_master := cs_gelb;
                        elsif j.ANZAHL > 4 and j.ANZAHL < 10 then
                            cs_master := cs_hellgelb;
                        elsif j.ANZAHL > 9 and j.ANZAHL < 25 then
                            cs_master := cs_hellgruen;
                        elsif j.ANZAHL > 24 then
                            cs_master := cs_gruen;
                        end if;
                        xlsx_writer.col_width(workbook, sheet_1, 8,9);
                        xlsx_writer.add_cell(workbook, sheet_1, l_row+c_y_split, 8,style_id => cs_master, value_ => j.ANZAHL);

                        v_kennungen.extend;
                        v_kennungen(v_kennungen.count).kennung := j.position_kennung;
                        v_kennungen(v_kennungen.count).row_position := l_row+c_y_split;

                        l_row:=l_row+1;
    end loop;

    xlsx_writer.freeze_sheet(workbook, sheet_1, c_x_split+5, c_y_split);
    xlsx     := xlsx_writer.create_xlsx(workbook);

    --Mailversand der Auswertung
    SendMailAuswertung(p_user_id,p_anhang => xlsx,p_filename => 'Auswertung.xlsx');

    DBMS_LOB.FREETEMPORARY(xlsx);

exception when others then
        DBS_LOGGING.LOG_ERROR_AT('PREISDATENBANK_PKG.auswertung_to_excel_leser: Fehler bei auswertung: ' || SQLCODE || ': ' || SQLERRM ||
      ' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE,'AUSWERTUNG');

        SendMailAuswertungFehler
            (
            p_user_id => p_user_id,
            p_error => 'PREISDATENBANK_PKG.auswertung_to_excel_leser: Fehler bei auswertung: ' || SQLCODE || ': ' || SQLERRM ||' Stacktrace: ' || DBMS_UTILITY.FORMAT_ERROR_BACKTRACE
            );

end auswertung_to_excel_leser;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

function print_param_worksheet(p_ws  IN OUT xlsx_writer.book_r, p_id IN NUMBER, p_style number)
  --es geht hier um die Eingabeparameter, die beim ersten mal eingegeben wurden 
  -- diese werden auf einer extra Registerkarte ausgegeben 

return xlsx_writer.book_r

is

   v_list varchar(2000 char);
   v_sheet number;
     -- Werte laden
  l_von               VARCHAR2(200);
  l_bis               VARCHAR2(200);
  l_getrimmter_mw     VARCHAR2(200); -- falls Zahl/Datum: TO_CHAR unten anpassen
  l_region            VARCHAR2(200);
  l_lvs               VARCHAR2(200);
  l_lieferant         VARCHAR2(200);
  l_lieferant_liste   VARCHAR2(200);
  l_vergabesumme      VARCHAR2(200); -- falls Zahl: TO_CHAR unten anpassen
  l_ausschreibung_lv  VARCHAR2(200);
  -- Labels und Werte als Collections
  l_names SYS.ODCIVARCHAR2LIST;
  l_vals  SYS.ODCIVARCHAR2LIST;

 begin

  	v_sheet  := XLSX_WRITER.add_sheet  (p_ws, 'input_parameter');

  /*  
    --Liste der Parameter
     SELECT  VON|| ':'|| BIS|| ':'|| GETRIMMTER_MITTELWERT|| ':'|| REGION|| ':'|| LVS|| ':'|| LIEFERANT|| ':'|| LIEFERANT_LISTE|| ':'|| VERGABESUMME|| ':'|| AUSSCHREIBUNG_LV
    into v_list
    from  PD_IMPORT_X86 where id=p_id;
*/

  -- Einzelwerte aus der Tabelle lesen
  SELECT 
    TO_CHAR(VON),                           -- bei DATE: TO_CHAR(VON,'DD.MM.YYYY')
    TO_CHAR(BIS),                           -- bei DATE: TO_CHAR(BIS,'DD.MM.YYYY')
    TO_CHAR(GETRIMMTER_MITTELWERT),         -- bei NUMBER: TO_CHAR(..., 'FM999G999D00')
    TO_CHAR(REGION),
    TO_CHAR(LVS),
    TO_CHAR(LIEFERANT),
    substr (LIEFERANT_LISTE,0,20) || case when length (LIEFERANT_LISTE) >20 then  '...' end,
    TO_CHAR(VERGABESUMME) || ' bis ' || VERGABESUMME2,                  -- Format ggf. anpassen
    TO_CHAR(AUSSCHREIBUNG_LV)
  INTO 
    l_von, l_bis, l_getrimmter_mw, l_region, l_lvs, l_lieferant, l_lieferant_liste, l_vergabesumme, l_ausschreibung_lv
  FROM PD_IMPORT_X86
  WHERE id = p_id;

  -- Labels zusammenstellen
   l_names := SYS.ODCIVARCHAR2LIST(
    'Von', 
    'Bis', 
    'GETRIMMTER_MITTELWERT', 
    'Region', 
    'LVS', 
    'Lieferant', 
    'Lieferant Liste', 
    'Vergabesumme', 
    'Ausschreibung LV'
  );

  -- Werte zusammenstellen
  l_vals := SYS.ODCIVARCHAR2LIST(
    NVL(l_von, ''),
    NVL(l_bis, ''),
    NVL(l_getrimmter_mw, ''),
    NVL(l_region, ''),
    NVL(l_lvs, ''),
    NVL(l_lieferant, ''),
    NVL(l_lieferant_liste, ''),
    NVL(l_vergabesumme, ''),
    NVL(l_ausschreibung_lv, '')
  );

    xlsx_writer.add_cell(p_ws, v_sheet, 1, 1, text => 'Eingegebene Parameter sind: ');
  -- Ausgabe: pro i eine Zeile, Spalte 1 = Name:, Spalte 2 = Wert
  FOR i IN 1 .. l_names.COUNT LOOP
    xlsx_writer.add_cell(p_ws, v_sheet, i+1, 1, text => l_names(i) || ':');
    xlsx_writer.add_cell(p_ws, v_sheet, i+1, 2, text => l_vals(i));
  END LOOP;
 	return p_ws;
end;





FUNCTION GetNamespace(p_blob_id in number) return number as
  v_return number;
begin

  select  case
           /* when upper (name) like '%.X82%' 
                then 4

                */
            when dbms_lob.instr(datei, utl_raw.cast_to_raw('http://www.gaeb.de/GAEB_DA_XML/DA83/3.2'), 1, 1) > 0 
                or dbms_lob.instr(datei, utl_raw.cast_to_raw('http://www.gaeb.de/GAEB_DA_XML/200407'), 1, 1) > 0  --x82 Format 28.02.2026
              then 2
            -- CW 12.03.2025: Neuer Typ 3
            when dbms_lob.instr(datei, utl_raw.cast_to_raw('http://www.gaeb.de/GAEB_DA_XML/DA83/3.3'), 1, 1) > 0 
              then 3
            else 1
          end
  into    v_return
  from    pd_import_x86
  where   id = p_blob_id;

  return v_return;

exception
  when others
    then return 1;
end;

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

PROCEDURE SendMailAuswertung(p_user_id number,p_anhang blob,p_filename varchar2) as
    v_mail                  dbs_email.rt_mail;
    r_mail_attachment       dbs_email.rt_attachment;
    r_mail_attachment_null  dbs_email.rt_attachment;
    t_mail_attachments      dbs_email.tt_attachments := dbs_email.tt_attachments();
    v_attachment_cnt        pls_integer := 0;
    v_mail_id               number;
    v_workspace_id          number;
begin

    v_workspace_id := apex_util.find_security_group_id (p_workspace => 'PREISDB');
    apex_util.set_security_group_id (p_security_group_id => v_workspace_id);

    --Empfänger-Mailadresse holen
    select  mail 
    into    v_mail.send_to
    from    dbs_user 
    where   id = p_user_id; 

    --Mail versenden 
    v_mail_id := apex_mail.send(
        p_to        => v_mail.send_to,
        p_from      => 'noreply@deutschebahn.com',
        p_body      => 'Sehr geehrte(r) Anwender(in),<br><br>anbei erhalten Sie Ihre gewünschte Auswertung aus der Preisdatenbank.',
        p_body_html => 'Sehr geehrte(r) Anwender(in),<br><br>anbei erhalten Sie Ihre gewünschte Auswertung aus der Preisdatenbank.',
        p_subj      => 'Ihre Auswertung aus der Preisdatenbank',
        p_cc        => null,
        p_bcc       => null,
        p_replyto   => null
        );

    --Anhang der Mail hinzufügen
    r_mail_attachment.content := p_anhang;
    r_mail_attachment.file_name := p_filename;
    t_mail_attachments.extend;
    t_mail_attachments(t_mail_attachments.last) := r_mail_attachment;

    for i in 1..t_mail_attachments.count loop
        v_attachment_cnt := v_attachment_cnt + 1;
        r_mail_attachment := r_mail_attachment_null;
        r_mail_attachment := t_mail_attachments(i);

        if r_mail_attachment.filebrowse_value is not null then
            select blob_content, filename
            into r_mail_attachment.content, r_mail_attachment.file_name
            from apex_application_temp_files
            where name = r_mail_attachment.filebrowse_value;
        elsif r_mail_attachment.file_name is null then
                r_mail_attachment.file_name := 'ATT' || to_char(v_attachment_cnt, 'FM009');
        end if;

        apex_mail.add_attachment(
            p_mail_id       => v_mail_id,
            p_attachment    => r_mail_attachment.content,
            p_filename      => r_mail_attachment.file_name,
            p_mime_type     => 'application/octet-stream'
        );
    end loop;

    apex_mail.push_queue;

end SendMailAuswertung;

PROCEDURE SendMailAuswertungFehler(p_user_id number,p_error varchar2) as
    v_mail                  dbs_email.rt_mail;
    r_mail_attachment       dbs_email.rt_attachment;
    r_mail_attachment_null  dbs_email.rt_attachment;
    t_mail_attachments      dbs_email.tt_attachments := dbs_email.tt_attachments();
    v_attachment_cnt        pls_integer := 0;
    v_mail_id               number;
    v_workspace_id          number;
begin

    v_workspace_id := apex_util.find_security_group_id (p_workspace => 'PREISDB');
    apex_util.set_security_group_id (p_security_group_id => v_workspace_id);

    --Empfänger-Mailadresse holen
    select  mail 
    into    v_mail.send_to
    from    dbs_user 
    where   id = p_user_id; 

    --Mail versenden 
    v_mail_id := apex_mail.send(
        p_to        => v_mail.send_to,
        p_from      => 'noreply@deutschebahn.com',
        p_body      => p_error,
        p_body_html => p_error,
        p_subj      => 'Ihre Auswertung aus der Preisdatenbank (Fehler)',
        p_cc        => null,
        p_bcc       => null,
        p_replyto   => null
        );

    apex_mail.push_queue;

end SendMailAuswertungFehler;

END PREISDATENBANK_PKG;