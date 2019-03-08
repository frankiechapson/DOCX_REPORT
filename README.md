
# DOCX report from PL/SQL

## Oracle PL/SQL solution to generate docx reports

## Why?

Because the printing is a real problem from Oracle or it needs BI Publisher.
This is a not too elegant solution for this problem, but it is working...

## How?

I tested only with MS Word 2013

It uses document fields, so first of all you have to switch on 
**File** tab then click **Options** then **Advanced** and under **Show document content** then check **Show field codes instead of their values** to display field code in document.
So we can see the fields now!

If you open the **document_merge_test_template.docx** you can see an example for usage.

You can insert fields into the doc using crtl+F9. After pressing ctrl+F9 you have to enter a field name between \{\}. This field name will be a column name of a select, so use them according this.

**VERY IMPORTANT!** You can use any formatting on the fields but keep the field name in one piece! If you use different formatting for different parts of the field, the field name will break into pieces in the XML, so my program will not recognize it.
Save and test the template docx often and if a field did not work, then remove the latest inserted field and insert it again and try the merge again!

## Tables
We can create tables. The program can repeat a table row. So we have to insert a special {/} field into the first cell of the row (what we want to repeat). And we have to insert a special closer mark \{\\} into the last cell of the row.
And finally we have to insert the necessary fields into the cells of the table row.   
You can see example for tables in the **document_merge_test_template.docx** file as well.

## How to install?
The right order:
1. as_zip.sql
2. create or replace type T_STRING_LIST as table of varchar2( 32000 );
3. F_DOCUMENT_MERGE.sql
4. F_DOCUMENT_MERGE_BLOB.sql
5. P_DOCUMENT_MERGE.sql


## How to call?


    begin
        P_DOCUMENT_MERGE ( 'INP_FILE_ORA_DIR'    
                         , 'document_merge_test_template.docx'
                         , 'OUT_FILE_ORA_DIR'
                         , 'document_merge_test_output.docx'
                         , T_STRING_LIST( 'select ''Tóth Feri'' as PERSON_NAME, 33 as PERSON_ID, ''Kutya'' TEAM_NAME, ''K'' TEAM_CODE from dual' 
                                        , 'select 11 as A1, 12 as A2, 13 as A3, 14 as A4 from dual union select 21, 22, 23, 24 from dual union  select 31, 32, 33, 34 from dual'
                                        , 'select 11 as F1, 12 as F2, 13 as F3, 14 as F4 from dual union select 21, 22, 23, 24 from dual' 
                                        ) 
                         );
    end;
    

If you use APEX, you can call **F_DOCUMENT_MERGE** which parameter is a zipped BLOB (docx file) instead of file and directory names.



## Behind the scenes

I used **as_zip** package created by **Anton Scheffer**. Thanks to him for it!

The **P_DOCUMENT_MERGE** procedures calls the **F_DOCUMENT_MERGE** with each SQL and finally clean the rest unnecessary nodes from the XML.

We need a simple **T_STRING_LIST** type:

    create or replace type T_STRING_LIST as table of varchar2( 32000 );

That's it!

I know the coding of F_DOCUMENT_MERGE function is not nice and not elegant, but I did not find better and working way, and I did not want to spend too much time with it.


