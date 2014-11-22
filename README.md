xlsx-pivot-to-csv
=================

Extract the complete data source of a xlsx pivot table and export it to a csv file.

This project is a temporary solution since the current version of Apache POI (3.11-beta) does not allow to access the complete source for a given pivot table. 

The xlsx-pivot-to-csv code unzips the xlsx file in memory and reads the xml-formatted data sources for the pivot table(s). The data is converted using the pivot table's column definitions. Finally, the data is exported as a csv file.

You can find more details at the bottom regarding the content of an unzipped xlsx file.

###Run the tool###

    $> lein run --xlsx xlsx-filename ;; simplest form
    $> lein run --xlsx xlsx-filename [--out output --defs def-file --recs rec-file]

###Command line arguments###

1 **--xlsx**: mandatory - the xlsx filename

2 **--out**: optional - the name of the output file (defaults to the xlsx filename with extension .csv instead of .xlsx)

3 **--defs**: optional - the name of the pivots cache definition file (defaults to xl/pivotCache/pivotCacheRecords1.xml, which is the definition of the first pivot table)

4 **--recs**: optional - the name of the records cache file (defaults to xl/pivotCache/pivotCacheRecords1.xml)

###Content of a xlsx file###

    $> mv pivot_example.xlsx pivot_example.zip
    $> unzip pivot_example
    $> tree pivot_example

<pre>
.
├── [Content_Types].xml
├── docProps
│   ├── app.xml
│   └── core.xml
├── pivot_table_example.zip
├── _rels
└── xl
    ├── pivotCache
    │   ├── pivotCacheDefinition1.xml
    │   ├── pivotCacheRecords1.xml
    │   └── _rels
    │       └── pivotCacheDefinition1.xml.rels
    ├── pivotTables
    │   ├── pivotTable1.xml
    │   └── _rels
    │       └── pivotTable1.xml.rels
    ├── _rels
    │   └── workbook.xml.rels
    ├── sharedStrings.xml
    ├── styles.xml
    ├── theme
    │   └── theme1.xml
    ├── workbook.xml
    └── worksheets
        ├── _rels
        │   └── sheet1.xml.rels
        └── sheet1.xml

11 directories, 16 files
</pre>