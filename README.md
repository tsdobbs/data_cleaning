﻿# data_cleaning

Miscellaneous scripts for basic data cleaning using Python.

### European Delimiter Conversion
_Requirements:_ Uses xlrd, xlwt, and xlutils
Currently the only script is for cleaning an Excel sheet in which some entries in a column are entered in the American style `#,###.##`, and some are entered in the European style `#.###,##`. If the script finds an American-style number stored in the column, it attempts to convert every cell in the column to American-style.

Known issue: Currently, xlrd is not able to read Excel format data if it is a .xlsx (post-2003) file. Since Excel stores dates as specially-formatted numbers, this means that running this script on a .xlsx file will also convert your dates to numbers which have to be fixed later.
