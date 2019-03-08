create or replace procedure P_DOCUMENT_MERGE ( I_INP_DIRECTORY      in varchar2
                                             , I_INP_FILE_NAME      in varchar2
                                             , I_OUT_DIRECTORY      in varchar2 := null
                                             , I_OUT_FILE_NAME      in varchar2 := null
                                             , I_SELECTS            in T_STRING_LIST
                                             ) as

    V_OUT_DIRECTORY         varchar2( 1000 ) := nvl( I_OUT_DIRECTORY, I_INP_DIRECTORY );
    V_OUT_FILE_NAME         varchar2( 1000 ) := nvl( I_OUT_FILE_NAME, I_INP_FILE_NAME );
    V_ZIP_BLOB              blob;

begin

    V_ZIP_BLOB := as_zip.file2blob( I_INP_DIRECTORY, I_INP_FILE_NAME );

    V_ZIP_BLOB := F_DOCUMENT_MERGE_BLOB ( V_ZIP_BLOB, I_SELECTS );

    as_zip.save_zip  ( V_ZIP_BLOB, V_OUT_DIRECTORY, V_OUT_FILE_NAME );

end;
/
