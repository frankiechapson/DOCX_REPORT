create or replace function F_DOCUMENT_MERGE_BLOB ( I_ZIP_BLOB           in blob
                                                 , I_SELECTS            in T_STRING_LIST
                                                 ) return blob as

    V_CLOB                  clob;
    V_DOC_BLOB              blob;
    V_NEW_BLOB              blob;
    V_PARSER                DBMS_XMLParser.Parser;  
    V_XML_DOC               DBMS_XMLDom.DomDocument;  
    V_ZIP_FILES             as_zip.file_list;

begin

    V_DOC_BLOB := as_zip.get_file ( I_ZIP_BLOB, 'word/document.xml' );
    V_CLOB     := BLOB_TO_CLOB    ( V_DOC_BLOB );

    V_PARSER  := dbms_xmlparser.newParser;  
    dbms_xmlparser.parseClob( V_PARSER, V_CLOB );  
    V_XML_DOC := dbms_xmlparser.getDocument( V_PARSER );  

    for L_I in 1..I_SELECTS.count
    loop
        if L_I = I_SELECTS.count then
            V_XML_DOC := F_DOCUMENT_MERGE ( V_XML_DOC, I_SELECTS( L_I ), 1 );
        else
            V_XML_DOC := F_DOCUMENT_MERGE ( V_XML_DOC, I_SELECTS( L_I ), 0 );
        end if;
    end loop;


    dbms_xmldom.writetoclob( V_XML_DOC, V_CLOB );
    dbms_xmldom.freeDocument(V_XML_DOC);  

    V_CLOB := replace( V_CLOB, '<w:instrText ' , '<w:t '  );
    V_CLOB := replace( V_CLOB, '</w:instrText>', '</w:t>' );

    V_DOC_BLOB := clob_to_blob( V_CLOB );

    V_ZIP_FILES  := as_zip.get_file_list( I_ZIP_BLOB );
    for L_I in V_ZIP_FILES.first() .. V_ZIP_FILES.last
    loop
        if V_ZIP_FILES( L_I ) = 'word/document.xml' then
            as_zip.add1file  ( V_NEW_BLOB, 'word/document.xml', V_DOC_BLOB );
        else
            as_zip.add1file  ( V_NEW_BLOB, V_ZIP_FILES( L_I ), as_zip.get_file( I_ZIP_BLOB, V_ZIP_FILES( L_I ) ) );
        end if;
    end loop;
    as_zip.finish_zip( V_NEW_BLOB );
    
    return V_NEW_BLOB;

end;
/
