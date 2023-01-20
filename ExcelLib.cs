using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using Funcs;
using System.Data;
using System.Data.SqlClient;

namespace ExcelLibrary;
class ExcelLib
{   
    SpreadsheetDocument? document; 
    WorkbookPart? workbookPart; 
    WorkbookStylesPart? wbsp;  
    Sheets? sheets;
    List<string> worksheetNamesList = new List<string>(); 
    List<WorksheetPart> worksheetPartList = new List<WorksheetPart>(); 
    List<SheetData>? sheetDataList = new List<SheetData>(); 
    Dictionary<int, Dictionary<string, double>> width_settings = new Dictionary<int, Dictionary<string, double>>();
    public String? name_doc;
    String Active_sheet_name = "-1";

    public void NewExFile(string name_docc, Stylesheet name_stylesheet) //-------INITIALIZATION--------
    {   
            // Create Spread Sheet Document  
        string name_doc = name_docc;  
        if(name_docc == "auto"){
            name_doc = DateTime.Today.ToString("ddMMyyyy") + ".xlsx";
        }
        Console.WriteLine(name_doc);
        document = SpreadsheetDocument.Create(name_doc, SpreadsheetDocumentType.Workbook);
           // Create FV
        //fv = new FileVersion();
        //fv.ApplicationName = "Microsoft Office Excel";

            // Create workbookPart
        workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook(); 

        sheets = workbookPart.Workbook.AppendChild(new Sheets());

        wbsp = workbookPart.AddNewPart<WorkbookStylesPart>();
        wbsp.Stylesheet = name_stylesheet;
        wbsp.Stylesheet.Save();
    }
    public void AddSheet(string name_sheet, int number_of_sheet = -1)
    {
        if(number_of_sheet == -1){
            number_of_sheet = worksheetPartList.Count;
        }
        
                // Create WorkSheet
        Console.WriteLine("Adding new WorkSheet with number " + number_of_sheet.ToString());
        //worksheetPartList.Add(workbookPart.AddNewPart<WorksheetPart>());
        worksheetPartList.Add(document.WorkbookPart.AddNewPart<WorksheetPart>());
        worksheetPartList[number_of_sheet].Worksheet = new Worksheet(new SheetData());
        

        Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPartList[number_of_sheet]), SheetId = Convert.ToUInt32(number_of_sheet+1), Name = name_sheet };
        
        //Sheets sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>()!;
        //string relationshipId = document.WorkbookPart.GetIdOfPart(worksheetPartList[number_of_sheet]);

        // Get a unique ID for the new worksheet.
        //uint sheetId = Convert.ToUInt32(number_of_sheet+1);
        // if (sheets.Elements<Sheet>().Count() > 0)
        // {
        //     sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
        // }
        // Give the new worksheet a name.
        //string sheetName = name_sheet;

        // Append the new worksheet and associate it with the workbook.
        //Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
        //sheets.Append(sheet);

        // --------------------------------- old part ----------------------------------
        worksheetNamesList.Add(name_sheet);
        Active_sheet_name = name_sheet;
        sheets.Append(sheet);
        sheetDataList.Add(worksheetPartList[number_of_sheet].Worksheet.GetFirstChild<SheetData>());
        width_settings[number_of_sheet] = new Dictionary<string, double>();
        width_settings[number_of_sheet]["Created"] = 1;
        Console.WriteLine("Adding Sheet success");
    }
    public void SetWidthEx(int count_columns, List<int> param_width_column, List<Double> width_column, string name_of_sheet = "-1")
    {
        if (name_of_sheet == "-1"){
            name_of_sheet = ActiveSheet();
        }
        int number_of_sheet = ConvertFromNameToID(name_of_sheet);

        Console.WriteLine("Adding SetOfWidth starting");
        Columns? lstColumns = worksheetPartList[number_of_sheet].Worksheet.GetFirstChild<Columns>()!;
        Boolean needToInsertColumns = false;
                
        if (lstColumns == null)
        {
            lstColumns = new Columns();
            needToInsertColumns = true;
        }
            // Change width of columns
        for (int i = 1; i < count_columns; i+=1){
            Console.WriteLine(Convert.ToUInt32(param_width_column[i]));
            Console.WriteLine(width_column[i]);
            Console.WriteLine("-----");
            lstColumns.Append(new Column() {Min = Convert.ToUInt32(param_width_column[i]), Max = Convert.ToUInt32(param_width_column[i]), Width = width_column[i], CustomWidth = true});
        }
           
        if (needToInsertColumns)
        {
            worksheetPartList[number_of_sheet].Worksheet.InsertAt(lstColumns, 0);
        }  
        Console.WriteLine("Adding SetOfWidth success");
    }      
    public void AddCell(string column_, uint row_, string label_, uint style_, string name_of_sheet_ = "-1", CellValues type = CellValues.String)
    {   
        if (name_of_sheet_ == "-1")
            name_of_sheet_ = ActiveSheet();
        int number_of_sheet_ = ConvertFromNameToID(name_of_sheet_);

        SharedStringTablePart shareStringPart;
            if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            //Insert the text into the SharedStringTablePart.
        int index = InsertSharedStringItem(label_, shareStringPart);
        Cell cell = InsertCellInWorksheet(column_, row_, worksheetPartList[number_of_sheet_]);
        Console.WriteLine(index.ToString());
        cell.CellValue = new CellValue(ReplaceHexadecimalSymbols(label_));
        cell.DataType = new EnumValue<CellValues>(type);
        cell.StyleIndex = style_;

        // Save the new worksheet.
        worksheetPartList[number_of_sheet_].Worksheet.Save();
        double weith_of_cell = CalcWeithOfCell(cell);
        if(width_settings[number_of_sheet_].ContainsKey(column_)){
            if(width_settings[number_of_sheet_][column_] < weith_of_cell & weith_of_cell <= 50){
                width_settings[number_of_sheet_][column_] = weith_of_cell;
            }
        }
        else{
            width_settings[number_of_sheet_][column_] = weith_of_cell;
        }
        Console.WriteLine("Finish adding cell");

                
    }
    private double CalcWeithOfCell(Cell celll){
        //---
        //If u have trouble with auto-width, change here some integers
        //---
        UInt32[] numberStyles = new UInt32[] { 5, 6, 7, 8 }; //styles that will add extra chars
        UInt32[] boldStyles = new UInt32[] { 1, 2, 3, 4, 6, 7, 8 }; //styles that will bold
        double maxColWidth =0;
        var cell = celll;
                var cellValue = cell.CellValue == null ? string.Empty : cell.CellValue.InnerText;
                var cellTextLength = cellValue.Length;

                if (cell.StyleIndex != null && numberStyles.Contains(cell.StyleIndex))
                {
                    int thousandCount = (int)Math.Truncate((double)cellTextLength / 4);

                    //add 3 for '.00' 
                    cellTextLength += (3 + thousandCount);
                }

                if (cell.StyleIndex != null && boldStyles.Contains(cell.StyleIndex))
                {
                    //add an extra char for bold - not 100% acurate but good enough for what i need.
                    cellTextLength += 1;
                }
                double maxWidth = 7;
                double width = Math.Truncate((cellTextLength * maxWidth + 5) / maxWidth * 256) / 256;

                return width;
            }
    private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        // If the part does not contain a SharedStringTable, create one.
        if (shareStringPart.SharedStringTable == null)
        {
            shareStringPart.SharedStringTable = new SharedStringTable();
        }

        int i = 0;

        // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                return i;
            }

            i++;
        }

        // The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        shareStringPart.SharedStringTable.Save();

        return i;
    }
    private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
    {
        Worksheet worksheet = worksheetPart.Worksheet;
        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        string cellReference = columnName + rowIndex;

        Row row;
        if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
        {
            row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }
        else
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
        {
            return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
        }
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell refCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (cell.CellReference.Value.Length == cellReference.Length)
                {
                  if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                  {
                    refCell = cell;
                    break;
                  }
                }
            }

            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);

            worksheet.Save();
            return newCell;
        }
    } 
    static string ReplaceHexadecimalSymbols(string txt)
        {
            // this function delete all shit-chars from your string
            string r = "[\x00-\x08\x0B\x0C\x0E-\x1F\x26]";
            return Regex.Replace(txt, r, "", RegexOptions.Compiled);
        }
    public void AddMerge(string from_, string to_, string name_sheet = "-1")
    {
        if (name_sheet == "-1") name_sheet = ActiveSheet();
        Mergee.MergeTwoCells(name_doc, name_sheet, from_, to_, document);
    }
    public void CloseEx(){
        foreach (KeyValuePair<int, Dictionary<string, double>> i in width_settings){
            int sheet_ = i.Key;
            List<int> list_col = new List<int>();
            List<double> list_width = new List<double>();
            foreach (KeyValuePair<string, double> j in i.Value){
                string column_ = j.Key;
                double width_ = j.Value;
                list_col.Add(GetColumnNumber(column_));
                list_width.Add(width_);
                //Console.WriteLine(width_);
            }
            SetWidthEx(width_settings[i.Key].Count, list_col, list_width, worksheetNamesList[sheet_]);
        }
        workbookPart.Workbook.Save();
        document.Close();
    }
    public static int GetColumnNumber(string name)
{
    int number = 0;
    int pow = 1;
    for (int i = name.Length - 1; i >= 0; i--)
    {
        number += (name[i] - 'A' + 1) * pow;
        pow *= 26;
    }

    return number;}
    private int ConvertFromNameToID(string name_sheet){
        if(name_sheet == "-1"){
            name_sheet = ActiveSheet();
        }
        for (int i = 0; i < worksheetNamesList.Count; i++){
            if(name_sheet == worksheetNamesList[i]){
                return i;
            }
        }
        return -1;
    }
    private string ActiveSheet(){
        //return worksheetNamesList[worksheetNamesList.Count - 1];
        return Active_sheet_name;
    }
    public void SetActSheet(string sheet_name){
        Active_sheet_name = sheet_name;
    }
}

// class DB{

//     public SqlConnection connection = new SqlConnection();
//     public List<String[]>? list_of_data = new List<string[]>();
//     public void CreateSqlConnection(string connection_line)
//     {
//         connection = new SqlConnection(connection_line);
//         connection.Open();
//     }
//     public Tuple<int, List<String[]>> GetData(string queryString)
//     {
//         SqlCommand command = new SqlCommand(queryString, connection);
//         SqlDataReader reader = command.ExecuteReader();
//         int cnt = reader.FieldCount;
        
//         while (reader.Read())
//         {
//             ReadSingleRow((IDataRecord)reader, cnt);
//         }
//         Tuple<int, List<String[]>> tp = new Tuple<int, List<string[]>>(cnt, list_of_data);
//         reader.Close();
//         return tp;
//     }
//     public void CloseConn()
//     {
//         connection.Close();
//     }
//     // public object CreateList(IDataRecord dataRecord, int column){
//     //     for(int i = 0; i < column; i++){
//     //         Console.WriteLine(dataRecord[i].GetType());
//     //     }
//     // }
//     public void ReadSingleRow(IDataRecord dataRecord, int column)
//     {
//         string[] temp_mass = new string[column];
//         for(int i = 0; i < column; i++){
//             // dataRecord[i].GetType() temp_value = dataRecord[i]
//             temp_mass[i] = String.Format("{0}", dataRecord[i]);
//             //Console.WriteLine(dataRecord[i].GetType());
//         }
//         list_of_data.Add(temp_mass);
//         //Console.WriteLine(String.Format("{0}, {1}, {2}, {3}", dataRecord[0], dataRecord[1], dataRecord[2], dataRecord[3]));
//     }
// }

// class Etime{ // Excel Time
//     static public string GetNumberOfWeekAndFromTo(){
//             string header_w = "WEEK №";
//             int firstDayOfYear = (int)new DateTime(DateTime.Now.Year, 1, 1).DayOfWeek;
//             int weekNumber = (DateTime.Now.DayOfYear + firstDayOfYear) / 7;
            
//             header_w += weekNumber.ToString() + " (From " +  DateTime.Today.AddDays(-7).ToString("yyyy-MM-dd") + " till " + DateTime.Today.AddDays(-1).ToString("yyyy-MM-dd") + ")";
//             return header_w;
//     }
// }

// class Estyle{ // Excel Functions
//     public string connectionString = "Data Source=localhost;Initial Catalog=BeelineStat;User ID=sa;Password=S2580s2580s2580@";
//     static List<Font> fonts = new List<Font>();
//     static List<Fill> fills = new List<Fill>();
//     static List<Border> borders = new List<Border>();
//     static List<CellFormat> cellformats = new List<CellFormat>();
//     static public void InitBase(){
//         //NewFont();
//         NewFill();
//         NewFill(color:"FFFFFFFF");
//         NewBorder();
//         //Estyle.NewCellFormat(0, 0, 0); //standart
//         //Estyle.NewCellFormat(1, 2, 0); //
//     }
//     static public void NewFont(Double size = 11.0, string color = "00000000", string fontname = "Calibri", bool isBold = false, bool isItalic = false, bool isUnderline = false){
//         int IdFont = fonts.Count;
//         fonts.Add(new Font());
//         fonts[IdFont].AddChild(new FontSize(){Val = size});
//         fonts[IdFont].AddChild(new Color(){Rgb = new HexBinaryValue() { Value = color}});
//         fonts[IdFont].AddChild(new FontName(){ Val = fontname});
//         if(isBold) fonts[IdFont].AddChild(new Bold());
//         if(isItalic) fonts[IdFont].AddChild(new Italic());
//         if(isUnderline) fonts[IdFont].AddChild(new Underline());
//         Console.WriteLine("Created new font, ID: " + IdFont.ToString());
//     }
//     static public void NewFill(string color = "0", string pattern = "solid"){
//         int IdFill = fills.Count;
//         fills.Add(new Fill());
//         // fills[IdFill].AddChild(new PatternFill(
//         //                    new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFAAAAAA" } }
//         //                       )
//         //                   { PatternType = PatternValues.Solid });
//         if(color == "0"){
//             //fills[IdFill].AddChild(new PatternFill(){ PatternType = PatternValues.None });
//             fills[IdFill].AddChild(new PatternFill(new ForegroundColor(){Rgb = new HexBinaryValue(){Value = "00000000"}}
//                               )
//                           { PatternType = PatternValues.Solid });
//         }
//         else{
//             fills[IdFill].AddChild(new PatternFill(new ForegroundColor(){Rgb = new HexBinaryValue(){Value = color}}
//                                         ){PatternType = PatternValues.Solid});
//         }
//         Console.WriteLine("Created new fill, ID: " + IdFill.ToString());
//     }
//     static public void NewBorder(string style = "", BorderStyleValues custom_st = BorderStyleValues.Thin){
//         int IdBorder = borders.Count;
//         borders.Add(new Border());
//         if(style == ""){
//             borders[IdBorder].AddChild(new LeftBorder());
//             borders[IdBorder].AddChild(new RightBorder());
//             borders[IdBorder].AddChild(new TopBorder());
//             borders[IdBorder].AddChild(new BottomBorder());
//             borders[IdBorder].AddChild(new DiagonalBorder());
//             Console.WriteLine("Created new default border, ID: " + IdBorder.ToString());
//         }
//         else{
//             BorderStyleValues st = custom_st;
//             if(style == "thin"){
//                 st = BorderStyleValues.Thin;
//             }
//             if(style == "medium"){
//                 st = BorderStyleValues.Medium;
//             }
//             if(style == "dotted"){
//                 st = BorderStyleValues.Dotted;
//             }
//             if(style == "custom"){
//                 st = custom_st;
//             }
//             borders[IdBorder].AddChild(new LeftBorder(new Color() { Auto = true }) {Style = st});
//             borders[IdBorder].AddChild(new RightBorder(new Color() { Auto = true }) {Style = st});
//             borders[IdBorder].AddChild(new TopBorder(new Color() { Auto = true }) {Style = st});
//             borders[IdBorder].AddChild(new BottomBorder(new Color() { Auto = true }) {Style = st});
//             borders[IdBorder].AddChild(new DiagonalBorder());
//             Console.WriteLine("Created new border, ID: " + IdBorder.ToString());
//         }
//     }
//     static public void NewCellFormat(int Idfont = 0, int IdFill=0, int IdBorder=0, int isCenter = 0){
//         int IdCell = cellformats.Count;
//         cellformats.Add(new CellFormat());
//         if(isCenter == 1){
//             Console.WriteLine("imhere");
//             cellformats[IdCell] = new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center}){ FontId = Convert.ToUInt32(Idfont), FillId = Convert.ToUInt32(IdFill), BorderId = Convert.ToUInt32(IdBorder)};  
//         }
//         else if(isCenter == 2){
//             cellformats[IdCell] = new CellFormat(new Alignment() {  Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Center, WrapText = true }){ FontId = Convert.ToUInt32(Idfont), FillId = Convert.ToUInt32(IdFill), BorderId = Convert.ToUInt32(IdBorder)};  
//         }
//         else {
//             cellformats[IdCell] = new CellFormat(){ FontId = Convert.ToUInt32(Idfont), FillId = Convert.ToUInt32(IdFill), BorderId = Convert.ToUInt32(IdBorder)};  
//         }
//     }
//     static public Stylesheet CreateStyleSheet(){

//         return new Stylesheet(
//             new Fonts(
//                 fonts
//             ),
//             new Fills(
//                 fills
//             ),
//             new Borders(
//                 borders
//             ),
//             new CellFormats(
//                 cellformats
//             )

//         );
//     }
    
// }

// class Mergee{

//     public static void MergeTwoCells(string docName, string sheetName, string cell1Name, string cell2Name, SpreadsheetDocument document)
//     {
//         // Open the document for editing.
//             Console.WriteLine("Start Merge Cells " + cell1Name + " and " + cell2Name);
//             Worksheet worksheet = GetWorksheet(document, sheetName);
//             if (worksheet == null || string.IsNullOrEmpty(cell1Name) || string.IsNullOrEmpty(cell2Name))
//             {
//                 return;
//             }

//             // Verify if the specified cells exist, and if they do not exist, create them.
//             //for (int i = 0; i < )
//             CreateSpreadsheetCellIfNotExist(worksheet, cell1Name);
//             CreateSpreadsheetCellIfNotExist(worksheet, cell2Name);

//             MergeCells mergeCells;
//             if (worksheet.Elements<MergeCells>().Count() > 0)
//             {
//                 mergeCells = worksheet.Elements<MergeCells>().First();
//             }
//             else
//             {
//                 mergeCells = new MergeCells();

//                 // Insert a MergeCells object into the specified position.
//                 if (worksheet.Elements<CustomSheetView>().Count() > 0)
//                 {
//                     worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
//                 }
//                 else if (worksheet.Elements<DataConsolidate>().Count() > 0)
//                 {
//                     worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
//                 }
//                 else if (worksheet.Elements<SortState>().Count() > 0)
//                 {
//                     worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
//                 }
//                 else if (worksheet.Elements<AutoFilter>().Count() > 0)
//                 {
//                     worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
//                 }
//                 else if (worksheet.Elements<Scenarios>().Count() > 0)
//                 {
//                     worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
//                 }
//                 else if (worksheet.Elements<ProtectedRanges>().Count() > 0)
//                 {
//                     worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
//                 }
//                 else if (worksheet.Elements<SheetProtection>().Count() > 0)
//                 {
//                     worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
//                 }
//                 else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0)
//                 {
//                     worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
//                 }
//                 else
//                 {
//                     worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
//                 }
//             }

//             // Create the merged cell and append it to the MergeCells collection.
//             MergeCell mergeCell = new MergeCell() { Reference = new StringValue(cell1Name + ":" + cell2Name) };
//             mergeCells.Append(mergeCell);

//             Console.WriteLine("Finish merging");

            
        
//     }
//     // Given a Worksheet and a cell name, verifies that the specified cell exists.
//     // If it does not exist, creates a new cell. 
//     private static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName)
//     {
//         string columnName = GetColumnName(cellName);
//         uint rowIndex = GetRowIndex(cellName);

//         IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => r.RowIndex.Value == rowIndex);

//         // If the Worksheet does not contain the specified row, create the specified row.
//         // Create the specified cell in that row, and insert the row into the Worksheet.
//         if (rows.Count() == 0)
//         {
//             Row row = new Row() { RowIndex = new UInt32Value(rowIndex) };
//             Cell cell = new Cell() { CellReference = new StringValue(cellName) };
//             row.Append(cell);
//             worksheet.Descendants<SheetData>().First().Append(row);
//             worksheet.Save();
//         }
//         else
//         {
//             Row row = rows.First();

//             IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

//             // If the row does not contain the specified cell, create the specified cell.
//             if (cells.Count() == 0)
//             {
//                 Cell cell = new Cell() { CellReference = new StringValue(cellName) };
//                 row.Append(cell);
//                 worksheet.Save();
//             }
//         }
//     }

//     // Given a SpreadsheetDocument and a worksheet name, get the specified worksheet.
//     private static Worksheet GetWorksheet(SpreadsheetDocument document, string worksheetName)
//     {
//         IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
//         WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
//         if (sheets.Count() == 0)
//             return null;
//         else
//             return worksheetPart.Worksheet;
//     }

//     // Given a cell name, parses the specified cell to get the column name.
//     private static string GetColumnName(string cellName)
//     {
//         // Create a regular expression to match the column name portion of the cell name.
//         Regex regex = new Regex("[A-Za-z]+");
//         Match match = regex.Match(cellName);

//         return match.Value;
//     }
//     // Given a cell name, parses the specified cell to get the row index.
//     private static uint GetRowIndex(string cellName)
//     {
//         // Create a regular expression to match the row index portion the cell name.
//         Regex regex = new Regex(@"\d+");
//         Match match = regex.Match(cellName);

//         return uint.Parse(match.Value);
//     }
// }