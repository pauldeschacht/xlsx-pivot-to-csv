xlsx-pivot-to-csv
=================

Extract the complete data source from a xlsx pivot table and export it to a csv file

This project is a temporary fix. The current version of Apache POI (3.11-beta) does not allow to access the complete data source for a given pivot table. 

The xlsx-pivot-to-csv code unzips the xlsx file and reads the raw (xml-formatted) data source for the pivot table(s). The data is converted using the pivot's column definitions. Finally, the data is exported as a csv file.
