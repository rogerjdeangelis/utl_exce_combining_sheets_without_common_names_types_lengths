# utl_exce_combining_sheets_without_common_names_types_lengths
Combining excel workbooks without common names types lengths.  Keywords: sas sql join merge big data analytics macros oracle teradata mysql sas communities stackoverflow statistics artificial inteligence AI Python R Java Javascript WPS Matlab SPSS Scala Perl C C# Excel MS Access JSON graphics maps NLP natural language processing machine learning igraph DOSUBL DOW loop stackoverflow SAS community.
    Combining excel workbooks without common names types lengths

    github
    https://tinyurl.com/ycp5tond
    https://github.com/rogerjdeangelis/utl_exce_combining_sheets_without_common_names_types_lengths

    do_over macro
    https://github.com/rogerjdeangelis/utl_sql_looping_or_using_arrays_in_sql_do_over_macro

    rename macro
    https://github.com/rogerjdeangelis/utl_rename_coordinated_lists_of_variables

    optlen macro
    https://github.com/rogerjdeangelis/utl_optlen

    We need to extract meta data that will allow us to combine very
    defferent sheets, with similar data.

    Original Topic:Import Multiple Excel Files where Columns have different formats


    SAS forum
    https://tinyurl.com/y8xr65yk
    https://communities.sas.com/t5/General-SAS-Programming/Import-Multiple-Excel-Files-where-Columns-have-different-formats/m-p/453325

    INPUT
    =====

      Algorithm

        Computed Meta data has to be examined and decisions on common names, type and lengths are needed.

        An estimate of maximum number of rows and columns simplifies the code.
        You can over specify.

           1. Get excel valid and 'invalid' column names from both workbooks.

           2. Get type and length for all sheets in mutiple workbooks using passthru to excel

           3. Build final meta data and decide on best names, types and lengths
              For simplicity I will use the column names in first sheet.
              However you can easily use other names using the meta data.

           4. Import the sheets and append all workbook sheets into on table.


     TWO WORKBOOKS

         Note $GENDER$ would become _GENDER_


      d:/xls/havOne.xlsx

          +---------------------------------------
          |     A      |    B       |     C      |
          +---------------------------------------
       1  | NAME       |   SEX      |    AGE     |   * these are the names I want;
          +------------+------------+------------+
       2  | ALFRED     |    M       |    15      |
          +------------+------------+------------+
       3  | ALICE      |    F       |    13      |
          +------------+------------+------------+
           ...

       [havOne]


      d:/xls/havTwo.xlsx

          +-------------------------------------------------
          |     A                |    B       |     C      |
          +-------------------------------------------------  * 123 is a numeric name
       1  | NAM                  | $GENDER$   |    123     |  * different names, types and lengths;
          +----------------------+------------+------------+
       2  | ALFRED               |  MALE      |    14yr    |
          +----------------------+------------+------------+
       3  | Mrs or Ms Alice      |  FEMALE    |    13      |
          +----------------------+------------+------------+
          ...

      [havTwo]


     MTADIC THIS META DATA

                     COLUMN_
          TABLE      POSITION    ANSWER   Excel SQL output before SAS

        havOneNam       F1       name
        havOneNam       F2       sex      Column Names
        havOneNam       F3       age

        havOneChr       F1       5
        havOneChr       F2       5        Column Type
        havOneChr       F3       0
        havOneLen       F1       7

        havOneLen       F2       1        Column Length
        havOneLen       F3       2

        havTwoNam       F1       nam
        havTwoNam       F2       $gender$
        havTwoNam       F3       123

        havTwoChr       F1       5
        havTwoChr       F2       5
        havTwoChr       F3       3
        havTwoLen       F1       16
        havTwoLen       F2       6
        havTwoLen       F3       4



     WORK.WANT total obs=10

         GRP       NAME                SEX       AGE

      havOneFix    Alfred              M         14     Workbook havOne
      havOneFix    Alice               F         13
      havOneFix    Barbara             F         13
      havOneFix    Carol               F         14
      havOneFix    Henry               M         14

      havTwoFix    Barbara             FEMALE    13     Workbook havOne
      havTwoFix    Carol               FEMALE    14yr
      havTwoFix    Henry               MALE      14yr
      havTwoFix    Mrs or Ms Alfred    MALE      14yr
      havTwoFix    Mrs or Ms Alice     FEMALE    13

          Variables in Creation Order

    #    Variable    Type    Len

    1    GRP         Char      9  added variable

    2    NAME        Char     16  * from second sheet (numeric in first)
    3    SEX         Char      6  * from second sheet
    4    AGE         Char      8  * from second sheet


    PROCESS
    =======

    * Build column name meta data;

    libname xelOne "d:/xls/havOne.xlsx" header=no;
    libname xelTwo "d:/xls/havTwo.xlsx" header=no;

    data mtaColNam;

       retain grp;
       length f1-f26 $32;
       array fs[*] $32 f1-f26;
       set
           xelOne.'havOne$A1:Z1'n(in=uno)
           xelTwo.'havTwo$A1:Z1'n
       ;
       if uno then grp='havOneNam';
       else grp='havTwoNam';

       popn=26-cmiss(of f:);
       call symputx('popn',popn);

       drop popn;
    run;quit;

    libname xelOne clear;
    libname xelTwo clear;

    /*
    havColNam total obs=2

    Obs    F1    F2         F3    F4 ... F26

     1    name   sex        age
     2    nam    $gender$   123

    %put &=popn;

    popn=3

    */

    * get type of all variables;
    %array(vars,values=f1-f&popn.)
    proc sql dquote=ansi;
      connect to excel (Path="d:/xls/havOne.xlsx" header=no);
        create table mtaOneChr as
        select * from connection to Excel
            (
             select
                  'havOneChr'  as grp,
                  %do_over(vars,phrase=format(count(*) + sum(isnumeric(?)),'####0') as ?,between=comma)
             from
                  [havOne$A2:C99]
            );
        disconnect from Excel;
    quit;

    /*
    WORK.MTAONECHR total obs=1
                       Number of Characters in each column
                       If 1 or more then use a character input format

      Obs       GRP       F1    F2    F3
       1     havOneChr    5     5     0
    */

    proc sql dquote=ansi;
      connect to excel (Path="d:/xls/havTwo.xlsx" header=no);
        create table mtaTwoChr as
        select * from connection to Excel
            (
             select
                  'havTwoChr'  as grp,
                  %do_over(vars,phrase=format(count(*) + sum(isnumeric(?)),'####0')  as ?,between=comma)
             from
                  [havTwo$A2:C99]
            );
        disconnect from Excel;
    quit;

    * get lengths;

    proc sql dquote=ansi;
      connect to excel (Path="d:/xls/havOne.xlsx" header=no);
        create table mtaOneLen as
        select * from connection to Excel
            (
             Select
                  'havOneLen'  as grp,
                  %do_over(vars,phrase=format(max(len(?)),'####') as ?,between=comma)
             from
                  [havOne$A2:C99]
            );
        disconnect from Excel;
    quit;

    /*
    WORK.MTAONELEN total obs=1
                   Length in Bytes for each column
    Obs       GRP        F1    F2    F3
     1      havOneLen    7     1     2
    */

    proc sql dquote=ansi;
      connect to excel (Path="d:/xls/havTwo.xlsx" header=no);
        create table mtaTwoLen as
        select * from connection to Excel
            (
             Select
                  'havTwoLen'  as grp,
                  %do_over(vars,phrase=format(max(len(?)),'####') as ?,between=comma)
             from
                  [havTwo$A2:C99]
            );
        disconnect from Excel;
    quit;

    data havCmb;

      set
        mtaColNam
        mtaOneChr
        mtaOneLen
        mtaTwoChr
        mtaTwoLen
       ;

      keep grp f1-f&popn.;

    run;quit;

    /*
     WORK.HAVCMB total obs=6

          GRP       F1      F2          F3
       havOneNam    name    sex         age
       havTwoNam    nam     $gender$    123

       havOneChr    5       5           0
       havOneLen    7       1           2

       havTwoChr    5       5           3
       havTwoLen    16      6           4
    */

    proc sort data=havCmb out=havCmbSrt;
    by grp;
    run;quit;

    proc transpose data=havCmbSrt out=mtaDic
       (drop=_label_ rename=(grp=table _name_=column_position col1=answer));
    by grp;
    var f:;
    run;quit;


    /*
     WORK.MTADIC total obs=18

                   COLUMN_
        TABLE      POSITION    ANSWER

      havOneChr       F1       5
      havOneChr       F2       5
      havOneChr       F3       0
      havOneLen       F1       7
     ....

    */

    * Fix first workbook;
    * fix change age to char in sheet havOne with length $1024;

    * note we could query meta data and set lengths using do_over on first select clause;
    %array(f12,values=f1-f2);
    proc sql;
      connect to excel (Path="d:/xls/havOne.xlsx" header=no);
        create table fixHavOne as
        select * from connection to Excel
            (
             Select
                  'havOneFix'  as grp,
                  %do_over(f12,phrase=?,between=comma)
                 ,format(f3,'#####') as f3
             from
                  [havOne$A2:C99]
            );
        disconnect from Excel;
    quit;

    #    Variable    Type     Len

    1    GRP         Char    1024   * you can fix this above or later using meta data;
    2    F1          Char     255
    3    F2          Char     255
    4    F3          Char    1024

    * Fix second workbook - all data is character;
    proc sql;
      connect to excel (Path="d:/xls/havTwo.xlsx" header=no);
        create table fixHavTwo as
        select * from connection to Excel
            (
             Select
                  'havTwoFix'  as grp
                  ,*
             from
                  [havTwo$A2:C99]
            );
        disconnect from Excel;
    quit;

    * uses the longer lengths;
    * get the new names from meta data;

    /*
    #    Variable    Type     Len

    1    GRP         Char    1024
    2    F1          Char     255
    3    F2          Char     255
    4    F3          Char     255
    */

    * note sql uses the longer lengths with union;
    Proc sql;
      select answer into :answer separated by " " from mtaDic where table="havOneNam";
      select column_position into :position separated by " " from mtaDic where table="havOneNam";
      create
        table want (rename=(%utl_renamel(old=&position,new=&answer))) as
      select
        *
      from
        fixHavOne
      union
        corr
      select
        *
      from
        fixHavTwo
    ;quit;

    /*
    #    Variable    Type     Len

    1    GRP         Char    1024
    2    NAME        Char     255
    3    SEX         Char     255
    4    AGE         Char    1024
    */

    %utl_optlen(inp=want,out=want);

    /*
    #    Variable    Type    Len

    1    GRP         Char      9
    2    NAME        Char     16
    3    SEX         Char      6
    4    AGE         Char      4
    */


    OUTPUT
    ======

    Up to 40 obs from mtaDic total obs=18

                        COLUMN_
    Obs      TABLE      POSITION    ANSWER

      1    havOneChr       F1       5
      2    havOneChr       F2       5
      3    havOneChr       F3       0
      4    havOneLen       F1       7
      5    havOneLen       F2       1
      6    havOneLen       F3       2
      7    havOneNam       F1       name
      8    havOneNam       F2       sex
      9    havOneNam       F3       age
     10    havTwoChr       F1       5
     11    havTwoChr       F2       5
     12    havTwoChr       F3       3
     13    havTwoLen       F1       16
     14    havTwoLen       F2       6
     15    havTwoLen       F3       4
     16    havTwoNam       F1       nam
     17    havTwoNam       F2       $gender$
     18    havTwoNam       F3       123


    WORK.WANT  total obs=10

    Obs       GRP       NAME                SEX       AGE

      1    havOneFix    Alfred              M         14
      2    havOneFix    Alice               F         13
      3    havOneFix    Barbara             F         13
      4    havOneFix    Carol               F         14
      5    havOneFix    Henry               M         14
      6    havTwoFix    Barbara             FEMALE    13
      7    havTwoFix    Carol               FEMALE    14yr
      8    havTwoFix    Henry               MALE      14yr
      9    havTwoFix    Mrs or Ms Alfred    MALE      14yr
     10    havTwoFix    Mrs or Ms Alice     FEMALE    13

    OUTPUT

    *                _               _       _
     _ __ ___   __ _| | _____     __| | __ _| |_ __ _
    | '_ ` _ \ / _` | |/ / _ \   / _` |/ _` | __/ _` |
    | | | | | | (_| |   <  __/  | (_| | (_| | || (_| |
    |_| |_| |_|\__,_|_|\_\___|   \__,_|\__,_|\__\__,_|

    ;

    %utlfkil(d:/xls/havOne.xlsx);
    %utlfkil(d:/xls/havTwo.xlsx);

    libname xelOne "d:/xls/havOne.xlsx";
    libname xelTwo "d:/xls/havTwo.xlsx";

    options validvarname=any;
    data xelOne.havOne(keep=name sex age)

         xelTwo.havTwo(keep=nam sex agec  rename=( sex='$gender$'n agec='123'n))
       ;
       retain name nam nam sex age;
       length agec $12 nam $16 sex $8;
       set sashelp.class(obs=5);
       output xelOne.havOne;
       if age>13 then agec=cats(put(age,2.),'yr');
       else agec=put(age,2.);
       if name=: 'A' then nam="Mrs or Ms "!!name;
       else nam=name;
       if sex='M' then sex="MALE";
       else sex="FEMALE";
       output xelTwo.havTwo;

    run;quit;
    options validvarname=upcase;

    libname xelOne clear;
    libname xelTwo clear;

    *          _       _   _
     ___  ___ | |_   _| |_(_) ___  _ __
    / __|/ _ \| | | | | __| |/ _ \| '_ \
    \__ \ (_) | | |_| | |_| | (_) | | | |
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|

    ;

    see above



