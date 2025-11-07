# utl-altair-slc-convert-pdf-tables-to-excel-sheets
Altair slc convert pdf tables to excel sheets
    %let pgm=utl-altair-slc-convert-pdf-tables-to-excel-sheets;

    %stop_submission;

    RE: Altair slc convert pdf tables to excel sheets

    Too long to post here, see github

    github
    https://github.com/rogerjdeangelis/utl-altair-slc-convert-pdf-tables-to-excel-sheets

    community.altair
    https://community.altair.com/discussion/19154?tab=all

    This solution requires macros array and do_over see

    macros
    https://github.com/rogerjdeangelis/utl-macros-used-in-many-of-rogerjdeangelis-repositories

    /*               _     _
     _ __  _ __ ___ | |__ | | ___ _ __ ___
    | `_ \| `__/ _ \| `_ \| |/ _ \ `_ ` _ \
    | |_) | | | (_) | |_) | |  __/ | | | | |
    | .__/|_|  \___/|_.__/|_|\___|_| |_| |_|
    |_|

    CONVERT PDF TABLES TO EXCEL SHEETS
    */

    /*****************************************************************************************/
    /*     INPUT              | 3 INTERMEDIATE CSV FILES  |      OUTPUT (three sheets)       */
    /* =================      |========================   |   ========================       */
    /* d:/pdf/tables.pdf      |   d:/csv/table_0.csv      |   d:/xls/tables.xlsx             */
    /*                        |                           |                                  */
    /*                        |   NAME SEX AGE            |   -------------------+           */
    /* This is title1         |   Bill M 12               |   | A1| fx    |NAME  |           */
    /* This is title2         |   Jack M 12               |   --------------------------+    */
    /* This is title3         |   Joe M 12                |   [_] |    A  |    B |    C |    */
    /*                        |                           |   --------------------------|    */
    /* AGE=12                 |   d:/csv/table_1.csv      |    1  | NAME  | SEX  | AGE  |    */
    /*                        |                           |    -- |-------+------+------|    */
    /* Obs    NAME    SEX     |   NAME SEX AGE            |    2  | Bill  | M    | 12   |    */
    /*                        |   Alice F 13              |    -- |-------+------+------|    */
    /*  1     Bill     M      |   Barb F 13               |    3  | Jack  | M    | 12   |    */
    /*  2     Jack     M      |   Kate F 13               |    -- |-------+------+------|    */
    /*  3     Joe      M      |                           |    4  | Joe   | M    | 12   |    */
    /*                        |   d:/csv/table_1.csv      |    -- |-------+------+------|    */
    /*                        |                           |   [TABLE_0]                      */
    /* AGE=13                 |   NAME SEX AGE            |                                  */
    /*                        |   Alfred M 14             |   -------------------+           */
    /* Obs    NAME     SEX    |   Carol F 14              |   | A1| fx    |NAME  |           */
    /*                        |   Henry M 14              |   --------------------------+    */
    /*  4     Alice     F     |                           |   [_] |    A  |    B |    C |    */
    /*  5     Barb      F     |                           |   --------------------------|    */
    /*  6     Kate      F     |                           |    1  | NAME  | SEX  | AGE  |    */
    /*                        |                           |    -- |-------+------+------|    */
    /*                        |                           |    2  | Alice | F    | 13   |    */
    /* AGE=14                 |                           |    -- |-------+------+------|    */
    /*                        |                           |    3  | Barb  | F    | 13   |    */
    /* Obs     NAME     SEX   |                           |    -- |-------+------+------|    */
    /*                        |                           |    4  | Kate  | F    | 13   |    */
    /*  7     Alfred     M    |                           |    -- |-------+------+------|    */
    /*  8     Carol      F    |                           |   [TABLE_1]                      */
    /*  9     Henry      M    |                           |                                  */
    /*                        |                           |   -------------------+           */
    /* This is footnote4      |                           |   | A1| fx    |NAME  |           */
    /* This is footnote5      |                           |   --------------------------+    */
    /* This is footnote6      |                           |   [_] |    A  |    B |   C  |    */
    /*                        |                           |   --------------------------|    */
    /*                        |                           |    1  | NAME  | SEX  | AGE  |    */
    /* NAME SEX AGE           |                           |    -- |-------+------+------|    */
    /* Bill M 12              |                           |    2  | Alfred| M    | 14   |    */
    /* Jack M 12              |                           |    -- |-------+------+------|    */
    /* Joe M 12               |                           |    3  | Carol | F    | 14   |    */
    /*                        |                           |    -- |-------+------+------|    */
    /* NAME SEX AGE           |                           |    4  | Henry | M    | 14   |    */
    /* Alice F 13             |                           |    -- |-------+------+------|    */
    /* Barb F 13              |                           |   [TABLE_2]                      */
    /* Kate F 13              |                           |                                  */
    /*                        |                           |                                  */
    /* NAME SEX AGE           |                           |                                  */
    /* Alfred M 14            |                           |                                  */
    /* Carol F 14             |                           |                                  */
    /* Henry M 14             |                           |                                  */
    /*****************************************************************************************/

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    /*---- create pdf file                           ---*/
    &_init_;
    libname sd1 "d:/sd1";
    data sd1.have;
      input
        name$
        sex$ age;
    cards4;
    Bill    M 12
    Jack    M 12
    Joe     M 12
    Alice   F 13
    Barb    F 13
    Kate    F 13
    Alfred  M 14
    Carol   F 14
    Henry   M 14
    ;;;;
    run;quit;

    %utlfkil(d:/pdf/tables.pdf);

    ods pdf file="d:/pdf/tables.pdf"  ;
    proc report data=sd1.have;
    by age;
    title1 "This is title1";
    title2 "This is title2";
    title3 "This is title3";
    footnote1 "This is footnote1";
    footnote2 "This is footnote2";
    footnote3 "This is footnote3";
    run;quit;
    ods pdf close;

    /*---- reset titles and footnotes                ---*/
    title;
    footnote;

    /*
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    */

    1                                          Altair SLC       15:09 Friday, November  7, 2025

    NOTE: Copyright 2002-2025 World Programming, an Altair Company
    NOTE: Altair SLC 2026 (05.26.01.00.000758)
          Licensed to Roger DeAngelis
    NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

    NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
    NOTE: AUTOEXEC source line

    ** KNOWN BUG***

    1       +  ï»¿;;;;
               ^
    ERROR: Expected a statement keyword : found "?"
    NOTE: AUTOEXEC processing completed

    1         /*---- create pdf file                           ---*/
    2         &_init_;
    NOTE: Directory contains files of mixed engine types
    3         libname sd1 "d:/sd1";
    NOTE: Library sd1 assigned as follows:
          Engine:        WPD
          Physical Name: d:\sd1

    4         data sd1.have;
    5           input
    6             name$
    7             sex$ age;
    8         cards4;

    NOTE: Data set "SD1.have" has 9 observation(s) and 3 variable(s)
    NOTE: The data step took :
          real time : 0.010
          cpu time  : 0.015


    9         Bill    M 12
    10        Jack    M 12
    11        Joe     M 12
    12        Alice   F 13
    13        Barb    F 13
    14        Kate    F 13
    15        Alfred  M 14
    16        Carol   F 14
    17        Henry   M 14
    18        ;;;;
    19        run;quit;
    20
    21        %utlfkil(d:/pdf/tables.pdf);
    22
    23        ods pdf file="d:/pdf/tables.pdf"  ;
    24        proc report data=sd1.have;
    25        by age;
    26        title1 "This is title1";
    27        title2 "This is title2";
    28        title3 "This is title3";
    29        footnote1 "This is footnote1";
    30        footnote2 "This is footnote2";
    31        footnote3 "This is footnote3";
    32        run;quit;
    NOTE: Writing file d:\pdf\tables.pdf
    NOTE: 9 observations were read from "SD1.have"
    NOTE: Procedure report step took :
          real time : 0.080
          cpu time  : 0.187

    33        ods pdf close;
    34
    35        /*---- reset titles and footnotes               ---*/
    36        title;
    37        footnote;
    38
    ERROR: Error printed on page 1

    NOTE: Submitted statements took :
          real time : 0.190
          cpu time  : 0.281
    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    /*---- create csv files and excel sheets                  ---*/

    &_init_;

    /*---- only needed for testing                            ---*/

    %utlfkil(d:/csv/table_0.csv);
    %utlfkil(d:/csv/table_1.csv);
    %utlfkil(d:/csv/table_2.csv);

    options set=PYTHONHOME "D:\python310";
    proc python;
    submit;
    import pyperclip
    from tabula.io import read_pdf
    import pandas as pd

    # Read all tables from all pages of the PDF into a list of DataFrames
    tables = read_pdf('d:/pdf/tables.pdf', pages='all', multiple_tables=True)

    # Print number of tables extracted
    print(f"Number of tables extracted: {len(tables)}")
    # Print the first table
    print(tables[0])
    print(tables[1])
    print(tables[2])

    for i, table in enumerate(tables):
        table.to_csv(f'd:/csv/table_{i}.csv',index=False)

    last_index = str(len(tables)-1)
    print(last_index)

    pyperclip.copy(last_index)
    endsubmit;
    run;quit;

    /*--- use clipboard to pass number of tables              ---*/

    filename clp clipbrd ;
    data _null_;
     infile clp;
     input;
     call symputx("last_index",_infile_,"G");
    run;quit;

    %put xxxxxxxxxxx &=last_index;

    proc datasets lib=work;
     delete table_:;
    run;quit;

    %utlfkil(d:/xls/tables.xlsx);

    libname xls excel "d:/xls/tables.xlsx";

    %array(idx,values=0-&last_index)

    /*--- loop pver csv files                                 ---*/

    %do_over(idx,phrase=%str(
       proc import datafile="d:/csv/table_?.csv"
           out=table_?
           dbms=dlm
           replace;
           delimiter=' ';
           getnames=yes;
       run;quit;
       proc print data=table_?;
       run;quit;
       )
    );

    /*--- copy wpd datasets to excel sheets                   ---*/

    proc copy in=work out=xls;
      select table_:;
    run;quit;

    libname xls clear;
    /*           _               _
      ___  _   _| |_ _ __  _   _| |_
     / _ \| | | | __| `_ \| | | | __|
    | (_) | |_| | |_| |_) | |_| | |_
     \___/ \__,_|\__| .__/ \__,_|\__|
     _       _      |_|                      _ _       _                          __ _ _
    (_)_ __ | |_ ___ _ __ _ __ ___   ___  __| (_) __ _| |_ ___    ___ _____   __ / _(_) | ___  ___
    | | `_ \| __/ _ \ `__| `_ ` _ \ / _ \/ _` | |/ _` | __/ _ \  / __/ __\ \ / /| |_| | |/ _ \/ __|
    | | | | | ||  __/ |  | | | | | |  __/ (_| | | (_| | ||  __/ | (__\__ \\ V / |  _| | |  __/\__ \
    |_|_| |_|\__\___|_|  |_| |_| |_|\___|\__,_|_|\__,_|\__\___|  \___|___/ \_/  |_| |_|_|\___||___/
    */

    d:/csv/table_0.csv

    NAME SEX AGE
    Bill M 12
    Jack M 12
    Joe M 12

    d:/csv/table_1.csv

    NAME SEX AGE
    Alice F 13
    Barb F 13
    Kate F 13

    d:/csv/table_1.csv

    NAME SEX AGE
    Alfred M 14
    Carol F 14
    Henry M 14

    FROM PYTHON LIST
    ================

    Altair SLC

    The PYTHON Procedure

    Number of tables extracted: 3

      NAME SEX AGE
    0    Bill M 12
    1    Jack M 12
    2     Joe M 12

      NAME SEX AGE
    0   Alice F 13
    1    Barb F 13
    2    Kate F 13

      NAME SEX AGE
    0  Alfred M 14
    1   Carol F 14
    2   Henry M 14

    /*                 _      _               _
      _____  _____ ___| | ___| |__   ___  ___| |_ ___
     / _ \ \/ / __/ _ \ |/ __| `_ \ / _ \/ _ \ __/ __|
    |  __/>  < (_|  __/ |\__ \ | | |  __/  __/ |_\__ \
     \___/_/\_\___\___|_||___/_| |_|\___|\___|\__|___/
    */

    d:/xls/tables.xlsx

    -------------------+
    | A1| fx    |NAME  |
    --------------------------+
    [_] |    A  |    B |    C |
    --------------------------|
     1  | NAME  | SEX  | AGE  |
     -- |-------+------+------|
     2  | Bill  | M    | 12   |
     -- |-------+------+------|
     3  | Jack  | M    | 12   |
     -- |-------+------+------|
     4  | Joe   | M    | 12   |
     -- |-------+------+------|
    [TABLE_0]

    -------------------+
    | A1| fx    |NAME  |
    --------------------------+
    [_] |    A  |    B |    C |
    --------------------------|
     1  | NAME  | SEX  | AGE  |
     -- |-------+------+------|
     2  | Alice | F    | 13   |
     -- |-------+------+------|
     3  | Barb  | F    | 13   |
     -- |-------+------+------|
     4  | Kate  | F    | 13   |
     -- |-------+------+------|
    [TABLE_1]

    -------------------+
    | A1| fx    |NAME  |
    --------------------------+
    [_] |    A  |    B |   C  |
    --------------------------|
     1  | NAME  | SEX  | AGE  |
     -- |-------+------+------|
     2  | Alfred| M    | 14   |
     -- |-------+------+------|
     3  | Carol | F    | 14   |
     -- |-------+------+------|
     4  | Henry | M    | 14   |
     -- |-------+------+------|
    [TABLE_2]



    /*
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    */

    1                                          Altair SLC       15:40 Friday, November  7, 2025

    NOTE: Copyright 2002-2025 World Programming, an Altair Company
    NOTE: Altair SLC 2026 (05.26.01.00.000758)
          Licensed to Roger DeAngelis
    NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

    NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
    NOTE: AUTOEXEC source line
    1       +  ï»¿;;;;
               ^
    ERROR: Expected a statement keyword : found "?"
    NOTE: AUTOEXEC processing completed

    1
    2         /*--- clipboard num_tables to a macro variable num_tables ---*/
    3         filename clp clipbrd ;
    4         data _null_;
    5          infile clp;
    6          input;
    7          call symputx("last_index",_infile_,"G");
    8         run;

    NOTE: The infile clp is:
          Clipboard

    NOTE: 1 record was read from file clp
          The minimum record length was 1
          The maximum record length was 1
    NOTE: The data step took :
          real time : 0.003
          cpu time  : 0.000


    8       !     quit;
    9
    10        %put xxxxxxxxxxx &=last_index;
    xxxxxxxxxxx last_index=2

                                               Altair SLC       15:40 Friday, November  7, 2025    1

                                         The DATASETS Procedure

                                               Directory

                                   Libref           WORK
                                   Engine           WPD
                                   Physical Name    d:\wpswrk\_TD3492
    11
    12        proc datasets lib=work;
    NOTE: No matching members in directory
    13         delete table_:;
    14        run;quit;
    NOTE: Procedure datasets step took :
          real time : 0.023
          cpu time  : 0.015


    15
    16        %utlfkil(d:/xls/tables.xlsx);
    The file d:/xls/tables.xlsx does not exist
    17
    18        libname xls excel "d:/xls/tables.xlsx";
    NOTE: Library xls assigned as follows:
          Engine:        OLEDB

    2                                          Altair SLC       15:40 Friday, November  7, 2025

          Physical Name: d:/xls/tables.xlsx

    19
    20        %array(idx,values=0-&last_index)
    21
    22        %do_over(idx,phrase=%str(
    NOTE: View opening spill file for output observations.
    23           proc import datafile="d:/csv/table_?.csv"
    24               out=table_?
    25               dbms=dlm
    26               replace;
    27               delimiter=' ';
    28               getnames=yes;
    29           run;quit;
    30           proc print data=table_?;
    31           run;quit;
    32           )
    33        );
    NOTE: Procedure import step took :
          real time : 0.002
          cpu time  : 0.000


    34        data table_0;
    35          infile 'd:\csv\table_0.csv' delimiter=' ' MISSOVER DSD firstobs=2 LRECL=32760;
    36          informat 'NAME'n $4.;
    37          informat 'SEX'n $1.;
    38          informat 'AGE'n BEST32.;
    39          format 'NAME'n $4.;
    40          format 'SEX'n $1.;
    41          format 'AGE'n BEST12.;
    42          label 'NAME'n = 'NAME';
    43          label 'SEX'n = 'SEX';
    44          label 'AGE'n = 'AGE';
    45          input    'NAME'n $
    46            'SEX'n $
    47            'AGE'n
    48          ;
    49          run;

    NOTE: The infile 'd:\csv\table_0.csv' is:
          Filename='d:\csv\table_0.csv',
          Owner Name=T7610\Roger,
          File size (bytes)=46,
          Create Time=14:01:48 Nov 07 2025,
          Last Accessed=15:40:33 Nov 07 2025,
          Last Modified=15:40:08 Nov 07 2025,
          Lrecl=32760, Recfm=V

    NOTE: 3 records were read from file 'd:\csv\table_0.csv'
          The minimum record length was 8
          The maximum record length was 9
    NOTE: Data set "WORK.table_0" has 3 observation(s) and 3 variable(s)

    3                                          Altair SLC       15:40 Friday, November  7, 2025

    NOTE: The data step took :
          real time : 0.005
          cpu time  : 0.015


    50        FILENAME ##IMPORT CLEAR;
    NOTE: 3 observations were read from "WORK.table_0"
    NOTE: Procedure print step took :
          real time : 0.011
          cpu time  : 0.015


    NOTE: Procedure import step took :
          real time : 0.000
          cpu time  : 0.000


    51        data table_1;
    52          infile 'd:\csv\table_1.csv' delimiter=' ' MISSOVER DSD firstobs=2 LRECL=32760;
    53          informat 'NAME'n $5.;
    54          informat 'SEX'n $1.;
    55          informat 'AGE'n BEST32.;
    56          format 'NAME'n $5.;
    57          format 'SEX'n $1.;
    58          format 'AGE'n BEST12.;
    59          label 'NAME'n = 'NAME';
    60          label 'SEX'n = 'SEX';
    61          label 'AGE'n = 'AGE';
    62          input    'NAME'n $
    63            'SEX'n $
    64            'AGE'n
    65          ;
    66          run;

    NOTE: The infile 'd:\csv\table_1.csv' is:
          Filename='d:\csv\table_1.csv',
          Owner Name=T7610\Roger,
          File size (bytes)=48,
          Create Time=14:01:48 Nov 07 2025,
          Last Accessed=15:40:33 Nov 07 2025,
          Last Modified=15:40:08 Nov 07 2025,
          Lrecl=32760, Recfm=V

    NOTE: 3 records were read from file 'd:\csv\table_1.csv'
          The minimum record length was 9
          The maximum record length was 10
    NOTE: Data set "WORK.table_1" has 3 observation(s) and 3 variable(s)
    NOTE: The data step took :
          real time : 0.003
          cpu time  : 0.000


    67        FILENAME ##IMPORT CLEAR;

    4                                          Altair SLC       15:40 Friday, November  7, 2025

    NOTE: 3 observations were read from "WORK.table_1"
    NOTE: Procedure print step took :
          real time : 0.010
          cpu time  : 0.000


    NOTE: Procedure import step took :
          real time : 0.000
          cpu time  : 0.000


    68        data table_2;
    69          infile 'd:\csv\table_2.csv' delimiter=' ' MISSOVER DSD firstobs=2 LRECL=32760;
    70          informat 'NAME'n $6.;
    71          informat 'SEX'n $1.;
    72          informat 'AGE'n BEST32.;
    73          format 'NAME'n $6.;
    74          format 'SEX'n $1.;
    75          format 'AGE'n BEST12.;
    76          label 'NAME'n = 'NAME';
    77          label 'SEX'n = 'SEX';
    78          label 'AGE'n = 'AGE';
    79          input    'NAME'n $
    80            'SEX'n $
    81            'AGE'n
    82          ;
    83          run;

    NOTE: The infile 'd:\csv\table_2.csv' is:
          Filename='d:\csv\table_2.csv',
          Owner Name=T7610\Roger,
          File size (bytes)=51,
          Create Time=14:01:48 Nov 07 2025,
          Last Accessed=15:40:33 Nov 07 2025,
          Last Modified=15:40:08 Nov 07 2025,
          Lrecl=32760, Recfm=V

    NOTE: 3 records were read from file 'd:\csv\table_2.csv'
          The minimum record length was 10
          The maximum record length was 11
    NOTE: Data set "WORK.table_2" has 3 observation(s) and 3 variable(s)
    NOTE: The data step took :
          real time : 0.003
          cpu time  : 0.015


    84        FILENAME ##IMPORT CLEAR;
    NOTE: 3 observations were read from "WORK.table_2"
    NOTE: Procedure print step took :
          real time : 0.011
          cpu time  : 0.000



    5                                                                                                                         Altair SLC

    85        &_init_;
    86
    87        /*---- only needed for testing                            ---*/
    88
    89        %utlfkil(d:/csv/table_0.csv);
    90        %utlfkil(d:/csv/table_1.csv);
    91        %utlfkil(d:/csv/table_2.csv);
    92
    93        options set=PYTHONHOME "D:\python310";
    94        proc python;
    95        submit;
    96        import pyperclip
    97        from tabula.io import read_pdf
    98        import pandas as pd
    99
    100       # Read all tables from all pages of the PDF into a list of DataFrames
    101       tables = read_pdf('d:/pdf/tables.pdf', pages='all', multiple_tables=True)
    102
    103       # Print number of tables extracted
    104       print(f"Number of tables extracted: {len(tables)}")
    105       # Print the first table
    106       print(tables[0])
    107       print(tables[1])
    108       print(tables[2])
    109
    110       for i, table in enumerate(tables):
    111           table.to_csv(f'd:/csv/table_{i}.csv',index=False)
    112
    113       last_index = str(len(tables)-1)
    114       print(last_index)
    115
    116       pyperclip.copy(last_index)
    117       endsubmit;

    NOTE: Submitting statements to Python:


    118       run;quit;
    NOTE: Procedure python step took :
          real time : 3.083
          cpu time  : 0.015


    119
    120       /*--- use clipboard to pass number of tables              ---*/
    121
    122       filename clp clipbrd ;
    123       data _null_;
    124        infile clp;
    125        input;
    126        call symputx("last_index",_infile_,"G");
    127       run;

    NOTE: The infile clp is:
          Clipboard

    NOTE: 1 record was read from file clp
          The minimum record length was 1
          The maximum record length was 1
    NOTE: The data step took :
          real time : 0.001
          cpu time  : 0.000


    6                                                                                                                         Altair SLC


    127     !     quit;
    128
    129       %put xxxxxxxxxxx &=last_index;
    xxxxxxxxxxx last_index=2

    Altair SLC

    The DATASETS Procedure

                Directory

    Libref           WORK
    Engine           WPD
    Physical Name    d:\wpswrk\_TD3492

                                   Members

                Member     Member
      Number    Name       Type          File Size      Date Last Modified

    ----------------------------------------------------------------------

           1    SASMACR    CATALOG           36864      07NOV2025:15:40:33
           2    TABLE_0    DATA               8192      07NOV2025:15:40:33
           3    TABLE_1    DATA               8192      07NOV2025:15:40:33
           4    TABLE_2    DATA               8192      07NOV2025:15:40:33
    130
    131       proc datasets lib=work;
    132        delete table_:;
    133       run;quit;
    NOTE: Deleting "WORK.TABLE_0" (memtype="DATA")
    NOTE: Deleting "WORK.TABLE_1" (memtype="DATA")
    NOTE: Deleting "WORK.TABLE_2" (memtype="DATA")
    NOTE: Procedure datasets step took :
          real time : 0.039
          cpu time  : 0.015


    134
    135       %utlfkil(d:/xls/tables.xlsx);
    The file d:/xls/tables.xlsx does not exist
    136
    137       libname xls excel "d:/xls/tables.xlsx";
    NOTE: Library xls assigned as follows:
          Engine:        OLEDB
          Physical Name: d:/xls/tables.xlsx

    138
    139       %array(idx,values=0-&last_index)
    140
    141       /*--- loop pver csv files                                 ---*/
    142
    143       %do_over(idx,phrase=%str(
    NOTE: View opening spill file for output observations.
    144          proc import datafile="d:/csv/table_?.csv"
    145              out=table_?
    146              dbms=dlm
    147              replace;
    148              delimiter=' ';
    149              getnames=yes;
    150          run;quit;
    151          proc print data=table_?;
    152          run;quit;
    153          )
    154       );
    NOTE: Procedure import step took :
          real time : 0.001
          cpu time  : 0.000


    155       data table_0;
    156         infile 'd:\csv\table_0.csv' delimiter=' ' MISSOVER DSD firstobs=2 LRECL=32760;
    157         informat 'NAME'n $4.;
    158         informat 'SEX'n $1.;
    159         informat 'AGE'n BEST32.;
    160         format 'NAME'n $4.;
    161         format 'SEX'n $1.;
    162         format 'AGE'n BEST12.;
    163         label 'NAME'n = 'NAME';
    164         label 'SEX'n = 'SEX';
    165         label 'AGE'n = 'AGE';
    166         input    'NAME'n $
    167           'SEX'n $
    168           'AGE'n

    7                                                                                                                         Altair SLC

    169         ;
    170         run;

    NOTE: The infile 'd:\csv\table_0.csv' is:
          Filename='d:\csv\table_0.csv',
          Owner Name=T7610\Roger,
          File size (bytes)=46,
          Create Time=14:01:48 Nov 07 2025,
          Last Accessed=15:40:36 Nov 07 2025,
          Last Modified=15:40:36 Nov 07 2025,
          Lrecl=32760, Recfm=V

    NOTE: 3 records were read from file 'd:\csv\table_0.csv'
          The minimum record length was 8
          The maximum record length was 9
    NOTE: Data set "WORK.table_0" has 3 observation(s) and 3 variable(s)
    NOTE: The data step took :
          real time : 0.005
          cpu time  : 0.000


    171       FILENAME ##IMPORT CLEAR;
    NOTE: 3 observations were read from "WORK.table_0"
    NOTE: Procedure print step took :
          real time : 0.028
          cpu time  : 0.000


    NOTE: Procedure import step took :
          real time : 0.001
          cpu time  : 0.000


    172       data table_1;
    173         infile 'd:\csv\table_1.csv' delimiter=' ' MISSOVER DSD firstobs=2 LRECL=32760;
    174         informat 'NAME'n $5.;
    175         informat 'SEX'n $1.;
    176         informat 'AGE'n BEST32.;
    177         format 'NAME'n $5.;
    178         format 'SEX'n $1.;
    179         format 'AGE'n BEST12.;
    180         label 'NAME'n = 'NAME';
    181         label 'SEX'n = 'SEX';
    182         label 'AGE'n = 'AGE';
    183         input    'NAME'n $
    184           'SEX'n $
    185           'AGE'n
    186         ;
    187         run;

    NOTE: The infile 'd:\csv\table_1.csv' is:
          Filename='d:\csv\table_1.csv',
          Owner Name=T7610\Roger,
          File size (bytes)=48,
          Create Time=14:01:48 Nov 07 2025,
          Last Accessed=15:40:37 Nov 07 2025,
          Last Modified=15:40:36 Nov 07 2025,
          Lrecl=32760, Recfm=V

    NOTE: 3 records were read from file 'd:\csv\table_1.csv'
          The minimum record length was 9
          The maximum record length was 10
    NOTE: Data set "WORK.table_1" has 3 observation(s) and 3 variable(s)

    8                                                                                                                         Altair SLC

    NOTE: The data step took :
          real time : 0.005
          cpu time  : 0.000


    188       FILENAME ##IMPORT CLEAR;
    NOTE: 3 observations were read from "WORK.table_1"
    NOTE: Procedure print step took :
          real time : 0.015
          cpu time  : 0.015


    NOTE: Procedure import step took :
          real time : 0.001
          cpu time  : 0.015


    189       data table_2;
    190         infile 'd:\csv\table_2.csv' delimiter=' ' MISSOVER DSD firstobs=2 LRECL=32760;
    191         informat 'NAME'n $6.;
    192         informat 'SEX'n $1.;
    193         informat 'AGE'n BEST32.;
    194         format 'NAME'n $6.;
    195         format 'SEX'n $1.;
    196         format 'AGE'n BEST12.;
    197         label 'NAME'n = 'NAME';
    198         label 'SEX'n = 'SEX';
    199         label 'AGE'n = 'AGE';
    200         input    'NAME'n $
    201           'SEX'n $
    202           'AGE'n
    203         ;
    204         run;

    NOTE: The infile 'd:\csv\table_2.csv' is:
          Filename='d:\csv\table_2.csv',
          Owner Name=T7610\Roger,
          File size (bytes)=51,
          Create Time=14:01:48 Nov 07 2025,
          Last Accessed=15:40:37 Nov 07 2025,
          Last Modified=15:40:36 Nov 07 2025,
          Lrecl=32760, Recfm=V

    NOTE: 3 records were read from file 'd:\csv\table_2.csv'
          The minimum record length was 10
          The maximum record length was 11
    NOTE: Data set "WORK.table_2" has 3 observation(s) and 3 variable(s)
    NOTE: The data step took :
          real time : 0.004
          cpu time  : 0.000


    205       FILENAME ##IMPORT CLEAR;
    NOTE: 3 observations were read from "WORK.table_2"
    NOTE: Procedure print step took :
          real time : 0.023
          cpu time  : 0.000


    206
    207       /*--- copy wpd datasets to excel sheets                   ---*/
    208
    209       proc copy in=work out=xls;

    9                                                                                                                         Altair SLC

    210         select table_:;
    211       run;quit;
    NOTE: Copying Member "WORK.TABLE_0" to XLS.TABLE_0 (memtype=DATA)
    NOTE: 3 observations read from input dataset WORK.TABLE_0
    NOTE: Output dataset XLS.TABLE_0 has 3 observations and 3 variables
    NOTE: Member "WORK.TABLE_0.DATA" (memtype=DATA) copied
    NOTE: Copying Member "WORK.TABLE_1" to XLS.TABLE_1 (memtype=DATA)
    NOTE: 3 observations read from input dataset WORK.TABLE_1
    NOTE: Output dataset XLS.TABLE_1 has 3 observations and 3 variables
    NOTE: Member "WORK.TABLE_1.DATA" (memtype=DATA) copied
    NOTE: Copying Member "WORK.TABLE_2" to XLS.TABLE_2 (memtype=DATA)
    NOTE: 3 observations read from input dataset WORK.TABLE_2
    NOTE: Output dataset XLS.TABLE_2 has 3 observations and 3 variables
    NOTE: Member "WORK.TABLE_2.DATA" (memtype=DATA) copied
    NOTE: 3 members copied
    NOTE: Procedure copy step took :
          real time : 2.871
          cpu time  : 2.203


    NOTE: Libref XLS has been deassigned.
    212
    213       libname xls clear;
    214
    ERROR: Error printed on page 1

    NOTE: Submitted statements took :
          real time : 7.846
          cpu time  : 3.500

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
