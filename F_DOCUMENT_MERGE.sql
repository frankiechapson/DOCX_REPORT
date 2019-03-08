create or replace function F_DOCUMENT_MERGE ( I_DOCUMENT        in xmldom.DomDocument
                                            , I_SQL             in varchar2 
                                            , I_REMOVE_FIELDS   in number := 0  -- default no
                                            ) return xmldom.DomDocument as 

    V_DOCUMENT          xmldom.DomDocument := I_DOCUMENT;
    V_NODE_LIST         xmldom.DomNodeList;  
    V_COLUMN_NAME       varchar2(   500 );   -- field name
    V_NODE_LIST_I       integer := 0;
    V_I                 integer;
    V_DATA              sys_refcursor;
    V_CURSOR            integer;
    V_ROW_COUNT         integer;
    V_FETCHED_I         integer := 0;
    V_COLUMNS           integer;
    V_DESC              dbms_sql.desc_tab;
    V_COL               varchar2(   500 );   -- field name
    V_STR               varchar2( 32000 );   -- field value
    V_NOF_REPLACED      integer := 0;
    V_WE_ARE_IN_ROW     boolean := false; 
    V_MULTIPLE_DATA     boolean := false;

    V_NODE              xmldom.DomNode;  
    V_LOOP_NODE         xmldom.DomNode;    -- the TR
    V_LOOP_PARENT       xmldom.DomNode;    -- the parent of TR
    V_NEW_ROW           xmldom.DomNode;  
    V_NEW_NODE          xmldom.DomNode;  
    V_BEGIN_NODE        xmldom.DomNode;     
    V_COLUMN_NODE       xmldom.DomNode;  
    V_COLNAM_NODE       xmldom.DomNode;  
    V_END_NODE          xmldom.DomNode;  

    ------------------------------------------------------------------------------------------------------------------------
    function GO_FOR( I_NODE_NAME in varchar2, I_ATTR_NAME in varchar2, I_ATTR_VALUE in varchar2 ) return  xmldom.DomNode is
    ------------------------------------------------------------------------------------------------------------------------
        L_NODE              xmldom.DomNode;    
        L_ATTR_NODE_MAP     xmldom.Domnamednodemap;  
        L_ATTR_NODE         xmldom.Domnode;     
        L_ATTR_NAME         varchar2( 5000 );  
        L_ATTR_VALUE        varchar2( 5000 );      
        L_ATTR_LENGTH       number;  
    begin
        loop
            exit when V_NODE_LIST_I >= xmldom.getLength( V_NODE_LIST ) ;
            L_NODE := xmldom.item( V_NODE_LIST, V_NODE_LIST_I );            

            if xmldom.getNodeName( L_NODE ) = I_NODE_NAME then

                L_ATTR_NODE_MAP := xmldom.getAttributes( L_NODE );
   
                if xmldom.isNull( L_ATTR_NODE_MAP ) = false then
                    L_ATTR_LENGTH := xmldom.getLength( L_ATTR_NODE_MAP );  
                    --Loop through attributes  
                    FOR i in 0..L_ATTR_LENGTH-1   
                    LOOP  
                        L_ATTR_NODE  := xmldom.item( L_ATTR_NODE_MAP, i );  
                        L_ATTR_NAME  := xmldom.getNodeName ( L_ATTR_NODE );  
                        L_ATTR_VALUE := xmldom.getNodeValue( L_ATTR_NODE );  
                        if L_ATTR_NAME = I_ATTR_NAME and L_ATTR_VALUE = I_ATTR_VALUE then
                            exit;
                        end if;
                    END LOOP;          
                END IF;  

                if L_ATTR_NAME = I_ATTR_NAME and L_ATTR_VALUE = I_ATTR_VALUE then
                    exit;
                end if;

            end if;

            V_NODE_LIST_I := V_NODE_LIST_I + 1;

        end loop;

        return L_NODE;

    exception when others then return null;
    end;

    ------------------------------------------------------------------------------------------------------------------------
    procedure REMOVE_FIELDS is
    ------------------------------------------------------------------------------------------------------------------------
    begin
        if nvl( xmldom.getLength( V_NODE_LIST ), 0 )  > 0 then
            loop
                V_BEGIN_NODE  := GO_FOR( 'w:fldChar', 'w:fldCharType', 'begin' );

                if not xmldom.isNull( V_BEGIN_NODE ) then

                    -- we found a field!         
                    V_BEGIN_NODE  := xmldom.getParentNode ( V_BEGIN_NODE  );             
                    V_COLUMN_NODE := xmldom.getNextSibling( V_BEGIN_NODE  );
                    V_COLNAM_NODE := GO_FOR( 'w:instrText', 'xml:space', 'preserve' );
                    if not xmldom.isNull( V_COLNAM_NODE ) then

                        V_COLUMN_NAME := trim( xmldom.Getnodevalue( xmldom.getFirstChild( V_COLNAM_NODE ) ) );  
                        V_END_NODE    := xmldom.getNextSibling( V_COLUMN_NODE );
             
                        if V_COLUMN_NAME in ( '/', chr( 92 ) ) then  
                            V_NODE        := DBms_xmldom.removechild( dbms_xmldom.getparentnode ( V_BEGIN_NODE  ), V_BEGIN_NODE  ); 
                            V_NODE        := DBms_xmldom.removechild( dbms_xmldom.getparentnode ( V_COLUMN_NODE ), V_COLUMN_NODE );
                            V_NODE        := DBms_xmldom.removechild( dbms_xmldom.getparentnode ( V_END_NODE    ), V_END_NODE    );
                        else
                            V_NODE        := DBms_xmldom.removechild( dbms_xmldom.getparentnode ( V_BEGIN_NODE ), V_BEGIN_NODE  ); 
                            V_NODE        := DBms_xmldom.removechild( dbms_xmldom.getparentnode ( V_END_NODE   ), V_END_NODE    );
                        end if;
                    end if;
                end if;

                exit when xmldom.isNull( V_BEGIN_NODE );

            end loop;
        end if;
    exception when others then null;
    end;

begin

    open V_DATA for I_SQL;
    V_CURSOR := dbms_sql.to_cursor_number( V_DATA );
    dbms_sql.describe_columns( V_CURSOR, V_COLUMNS, V_DESC );
    -- define every column type to string
    for V_I in 1..V_COLUMNS 
    loop
        dbms_sql.define_column( V_CURSOR, V_I, V_STR, 4000 );
    end loop;

    V_NODE_LIST := xmldom.getElementsByTagName( V_DOCUMENT, '*' );  

    -- get the data and merge with the document
    if nvl( xmldom.getLength( V_NODE_LIST ), 0 )  > 0 then

        loop
           
            V_ROW_COUNT := V_FETCHED_I;
            V_FETCHED_I := V_FETCHED_I + dbms_sql.fetch_rows( V_CURSOR );
            exit when V_FETCHED_I = 0;
            exit when V_FETCHED_I > 1 and V_NOF_REPLACED = 0;  -- if we fetched the second row and there were no replace in the first one, then exit;

            if V_FETCHED_I > 1 and V_MULTIPLE_DATA then
                V_NODE      := xmldom.clonenode( V_LOOP_NODE, true );  -- the row
                V_NEW_ROW   := xmldom.appendChild( V_LOOP_PARENT , V_NODE  );
                V_NODE_LIST := xmldom.getElementsByTagName( V_DOCUMENT, '*' );  
            end if;  
              
            -- browse the xml for fields:        
            loop
                V_BEGIN_NODE  := GO_FOR( 'w:fldChar', 'w:fldCharType', 'begin' );
             
                if not xmldom.isNull( V_BEGIN_NODE ) then

                    -- we found a field!         
                    V_BEGIN_NODE  := xmldom.getParentNode ( V_BEGIN_NODE  );
             
                    V_COLUMN_NODE := xmldom.getNextSibling( V_BEGIN_NODE  );
                    V_COLNAM_NODE := GO_FOR( 'w:instrText', 'xml:space', 'preserve' );
                    if not xmldom.isNull( V_COLNAM_NODE ) then
                        V_COLUMN_NAME := trim( xmldom.Getnodevalue( xmldom.getFirstChild( V_COLNAM_NODE ) ) );  
                        
                        V_END_NODE    := xmldom.getNextSibling( V_COLUMN_NODE );
                        
                        if V_COLUMN_NAME = '/' then  
                            -- this marker shows the start of the loop
                            V_WE_ARE_IN_ROW := true;
                        
                            if V_NOF_REPLACED = 0 then
                                V_LOOP_NODE    := xmldom.clonenode( xmldom.getParentNode( xmldom.getParentNode( xmldom.getParentNode( V_BEGIN_NODE ) ) ), true );  -- the row
                                V_LOOP_PARENT  := xmldom.getParentNode( xmldom.getParentNode( xmldom.getParentNode( xmldom.getParentNode( V_BEGIN_NODE ) ) ) );  -- to append child
                            end if;
                        
                        elsif V_COLUMN_NAME = chr( 92 ) then  
                            -- \ marker shows the end of the loop
                            V_WE_ARE_IN_ROW := false;
                            if V_MULTIPLE_DATA then
                                V_ROW_COUNT    := V_FETCHED_I;
                                V_FETCHED_I    := V_FETCHED_I + dbms_sql.fetch_rows( V_CURSOR );
                                if V_FETCHED_I > V_ROW_COUNT then
                                    V_NODE      := xmldom.clonenode( V_LOOP_NODE, true );  -- the row
                                    V_NEW_ROW   := xmldom.appendChild( V_LOOP_PARENT , V_NODE  );
                                    V_NODE_LIST := xmldom.getElementsByTagName( V_DOCUMENT, '*' );  
                                end if;
                            end if;
                        
                        else
                        
                            V_I := 1; 
                            loop
                                exit when V_I > V_COLUMNS or upper( V_DESC( V_I ).col_name ) = upper( V_COLUMN_NAME );
                                V_I := V_I + 1;
                            end loop;
                        
                            if V_I <= V_COLUMNS and upper( V_DESC( V_I ).col_name ) = upper( V_COLUMN_NAME ) then      
                                -- replace the field to the value
                                dbms_sql.column_value( V_CURSOR, V_I, V_STR );   
                                xmldom.setNodeValue( xmldom.getFirstChild( V_COLNAM_NODE ), V_STR );
                                V_NOF_REPLACED := V_NOF_REPLACED + 1;
                                if V_WE_ARE_IN_ROW then
                                    V_MULTIPLE_DATA := true;
                                end if;
                        
                            end if;
                        
                        end if;

                    end if;

                else 
                    exit;   -- no more "BEGIN"
                end if;

                exit when V_FETCHED_I = V_ROW_COUNT or xmldom.isNull( V_BEGIN_NODE );
        
            end loop;
        
            exit when V_FETCHED_I = V_ROW_COUNT or xmldom.isNull( V_BEGIN_NODE );
        
        end loop;

        -- and finally remove the field...
        if I_REMOVE_FIELDS > 0 then 
            V_NODE_LIST   := xmldom.getElementsByTagName( V_DOCUMENT, '*' );  
            V_NODE_LIST_I := 0;
            REMOVE_FIELDS;
        end if;

    end if;

    dbms_sql.close_cursor( V_CURSOR );

    return V_DOCUMENT;

end;
/


