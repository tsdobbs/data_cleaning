# data_cleaning

Miscellaneous scripts for basic data cleaning using Python.

### European Delimiter Conversion
Currently the only script is for cleaning an Excel sheet in which some entries in a column are entered in the American style `#,###.##`, and some are entered in the European style `#.###,##`. If the script finds an American-style number stored in the column, it attempts to convert every cell in the column to American-style.