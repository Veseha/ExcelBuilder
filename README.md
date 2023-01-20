# ExcelBuilder
C# App/library for automating the creation of Excel books with short commands, based on OpenXML and cross-platform
All information about how to use this check in comments in Program.cs, it's so simple that you can figure it out in five minutes. But, anyway, here short information about commands:

Excellib - app for generate Excel books
  NewExFile - Create EX file and book  (name_of_doc, StyleSheet) [if we put "auto" to name_of_doc arg we get name with format ddMMyyyy.xlsx]
  AddSheet - Create Sheet  (name_sheet), new sheet add to the end of sheets and become a active sheet
  SetActSheet - Set Active Sheet (name_sheet)
  AddCell - Create cells (row, column, label, style, (sheet_name)=activeSheet)
  AddMerge - Create merge between cells (cell_from, cell_to, (sheet_name)=activeSheet)
Esyle - Stylesheet generate
  All cell style based on Font, Fill, Border and in the end we "mix" this parameters to create CellFormat
  All components connecting with ID from 0
  Color sheme - HEX: 11223344: 11-Bright, 22-R, 33-G, 44-B
  Important! InitBase create default: Fill(0 - transpared, 1 - grid), Border (0), Cell(0) (yes, it's bug, may be fix in the future)
  NewFont (size, color, fontname, isbold, isitalic, isunderline)
  NewFill (color)
  NewBorder (style) 
  NewCellFormat (idFONT, idFILL, idBORDER)
And, optional - DB - create a connection to MS SQL DB (optional)
  1. You need to open connection with command CreateSqlConnection(connection_line)
  2. To retrieve data from a query, run the command db.GetData("SELECT * FROM table")
     Data format - Tuple<int, List<String[]>>
     .Item1 - The number of columns in the resulting table
     .Item2 - List of rows, each of which consists of an array of N(string) columns - received values
     At the moment, all data are automatically converted to string format. Support for dynamic data format is planned for the future.
  3. At the end of all operations with DB you need to close the connection. CloseConn()

Free to use for personal and non-commercial tasks, for commercial please contact with me on instagram (same nickname) or telegram (@hahaclassic)
