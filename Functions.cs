using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.SqlClient;

namespace Funcs;

class DB{

    public SqlConnection connection = new SqlConnection();
    public List<String[]>? list_of_data = new List<string[]>();
    public void CreateSqlConnection(string connection_line)
    {
        connection = new SqlConnection(connection_line);
        connection.Open();
    }
    public Tuple<int, List<String[]>> GetData(string queryString)
    {
        SqlCommand command = new SqlCommand(queryString, connection);
        SqlDataReader reader = command.ExecuteReader();
        int cnt = reader.FieldCount;
        
        while (reader.Read())
        {
            ReadSingleRow((IDataRecord)reader, cnt);
        }
        Tuple<int, List<String[]>> tp = new Tuple<int, List<string[]>>(cnt, list_of_data);
        reader.Close();
        return tp;
    }
    public void CloseConn()
    {
        connection.Close();
    }
    // public object CreateList(IDataRecord dataRecord, int column){
    //     for(int i = 0; i < column; i++){
    //         Console.WriteLine(dataRecord[i].GetType());
    //     }
    // }
    public void ReadSingleRow(IDataRecord dataRecord, int column)
    {
        string[] temp_mass = new string[column];
        for(int i = 0; i < column; i++){
            // dataRecord[i].GetType() temp_value = dataRecord[i]
            temp_mass[i] = String.Format("{0}", dataRecord[i]);
            //Console.WriteLine(dataRecord[i].GetType());
        }
        list_of_data.Add(temp_mass);
        //Console.WriteLine(String.Format("{0}, {1}, {2}, {3}", dataRecord[0], dataRecord[1], dataRecord[2], dataRecord[3]));
    }
}

class Estyle{ // Excel Functions
    public string connectionString = "Data Source=localhost;Initial Catalog=BeelineStat;User ID=sa;Password=S2580s2580s2580@";
    static List<Font> fonts = new List<Font>();
    static List<Fill> fills = new List<Fill>();
    static List<Border> borders = new List<Border>();
    static List<CellFormat> cellformats = new List<CellFormat>();
    static public void InitBase(){
        //NewFont();
        NewFill();
        NewFill(color:"FFFFFFFF");
        NewBorder();
        //Estyle.NewCellFormat(0, 0, 0); //standart
        //Estyle.NewCellFormat(1, 2, 0); //
    }
    static public void NewFont(Double size = 11.0, string color = "00000000", string fontname = "Calibri", bool isBold = false, bool isItalic = false, bool isUnderline = false){
        int IdFont = fonts.Count;
        fonts.Add(new Font());
        fonts[IdFont].AddChild(new FontSize(){Val = size});
        fonts[IdFont].AddChild(new Color(){Rgb = new HexBinaryValue() { Value = color}});
        fonts[IdFont].AddChild(new FontName(){ Val = fontname});
        if(isBold) fonts[IdFont].AddChild(new Bold());
        if(isItalic) fonts[IdFont].AddChild(new Italic());
        if(isUnderline) fonts[IdFont].AddChild(new Underline());
        Console.WriteLine("Created new font, ID: " + IdFont.ToString());
    }
    static public void NewFill(string color = "0", string pattern = "solid"){
        int IdFill = fills.Count;
        fills.Add(new Fill());
        // fills[IdFill].AddChild(new PatternFill(
        //                    new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFAAAAAA" } }
        //                       )
        //                   { PatternType = PatternValues.Solid });
        if(color == "0"){
            //fills[IdFill].AddChild(new PatternFill(){ PatternType = PatternValues.None });
            fills[IdFill].AddChild(new PatternFill(new ForegroundColor(){Rgb = new HexBinaryValue(){Value = "00000000"}}
                              )
                          { PatternType = PatternValues.Solid });
        }
        else{
            fills[IdFill].AddChild(new PatternFill(new ForegroundColor(){Rgb = new HexBinaryValue(){Value = color}}
                                        ){PatternType = PatternValues.Solid});
        }
        Console.WriteLine("Created new fill, ID: " + IdFill.ToString());
    }
    static public void NewBorder(string style = "", BorderStyleValues custom_st = BorderStyleValues.Thin){
        int IdBorder = borders.Count;
        borders.Add(new Border());
        if(style == ""){
            borders[IdBorder].AddChild(new LeftBorder());
            borders[IdBorder].AddChild(new RightBorder());
            borders[IdBorder].AddChild(new TopBorder());
            borders[IdBorder].AddChild(new BottomBorder());
            borders[IdBorder].AddChild(new DiagonalBorder());
            Console.WriteLine("Created new default border, ID: " + IdBorder.ToString());
        }
        else{
            BorderStyleValues st = custom_st;
            if(style == "thin"){
                st = BorderStyleValues.Thin;
            }
            if(style == "medium"){
                st = BorderStyleValues.Medium;
            }
            if(style == "dotted"){
                st = BorderStyleValues.Dotted;
            }
            if(style == "custom"){
                st = custom_st;
            }
            borders[IdBorder].AddChild(new LeftBorder(new Color() { Auto = true }) {Style = st});
            borders[IdBorder].AddChild(new RightBorder(new Color() { Auto = true }) {Style = st});
            borders[IdBorder].AddChild(new TopBorder(new Color() { Auto = true }) {Style = st});
            borders[IdBorder].AddChild(new BottomBorder(new Color() { Auto = true }) {Style = st});
            borders[IdBorder].AddChild(new DiagonalBorder());
            Console.WriteLine("Created new border, ID: " + IdBorder.ToString());
        }
    }
    static public void NewCellFormat(int Idfont = 0, int IdFill=0, int IdBorder=0, int isCenter = 0){
        int IdCell = cellformats.Count;
        cellformats.Add(new CellFormat());
        if(isCenter == 1){
            Console.WriteLine("imhere");
            cellformats[IdCell] = new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center}){ FontId = Convert.ToUInt32(Idfont), FillId = Convert.ToUInt32(IdFill), BorderId = Convert.ToUInt32(IdBorder)};  
        }
        else if(isCenter == 2){
            cellformats[IdCell] = new CellFormat(new Alignment() {  Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Center, WrapText = true }){ FontId = Convert.ToUInt32(Idfont), FillId = Convert.ToUInt32(IdFill), BorderId = Convert.ToUInt32(IdBorder)};  
        }
        else {
            cellformats[IdCell] = new CellFormat(){ FontId = Convert.ToUInt32(Idfont), FillId = Convert.ToUInt32(IdFill), BorderId = Convert.ToUInt32(IdBorder)};  
        }
    }
    static public Stylesheet CreateStyleSheet(){

        return new Stylesheet(
            new Fonts(
                fonts
            ),
            new Fills(
                fills
            ),
            new Borders(
                borders
            ),
            new CellFormats(
                cellformats
            )

        );
    }
    
}

class Mergee{

    public static void MergeTwoCells(string docName, string sheetName, string cell1Name, string cell2Name, SpreadsheetDocument document)
    {
        // Open the document for editing.
            Console.WriteLine("Start Merge Cells " + cell1Name + " and " + cell2Name);
            Worksheet worksheet = GetWorksheet(document, sheetName);
            if (worksheet == null || string.IsNullOrEmpty(cell1Name) || string.IsNullOrEmpty(cell2Name))
            {
                return;
            }

            // Verify if the specified cells exist, and if they do not exist, create them.
            //for (int i = 0; i < )
            CreateSpreadsheetCellIfNotExist(worksheet, cell1Name);
            CreateSpreadsheetCellIfNotExist(worksheet, cell2Name);

            MergeCells mergeCells;
            if (worksheet.Elements<MergeCells>().Count() > 0)
            {
                mergeCells = worksheet.Elements<MergeCells>().First();
            }
            else
            {
                mergeCells = new MergeCells();

                // Insert a MergeCells object into the specified position.
                if (worksheet.Elements<CustomSheetView>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                }
                else if (worksheet.Elements<DataConsolidate>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
                }
                else if (worksheet.Elements<SortState>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
                }
                else if (worksheet.Elements<AutoFilter>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
                }
                else if (worksheet.Elements<Scenarios>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
                }
                else if (worksheet.Elements<ProtectedRanges>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
                }
                else if (worksheet.Elements<SheetProtection>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
                }
                else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
                }
                else
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                }
            }

            // Create the merged cell and append it to the MergeCells collection.
            MergeCell mergeCell = new MergeCell() { Reference = new StringValue(cell1Name + ":" + cell2Name) };
            mergeCells.Append(mergeCell);

            Console.WriteLine("Finish merging");

            
        
    }
    // Given a Worksheet and a cell name, verifies that the specified cell exists.
    // If it does not exist, creates a new cell. 
    private static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName)
    {
        string columnName = GetColumnName(cellName);
        uint rowIndex = GetRowIndex(cellName);

        IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => r.RowIndex.Value == rowIndex);

        // If the Worksheet does not contain the specified row, create the specified row.
        // Create the specified cell in that row, and insert the row into the Worksheet.
        if (rows.Count() == 0)
        {
            Row row = new Row() { RowIndex = new UInt32Value(rowIndex) };
            Cell cell = new Cell() { CellReference = new StringValue(cellName) };
            row.Append(cell);
            worksheet.Descendants<SheetData>().First().Append(row);
            worksheet.Save();
        }
        else
        {
            Row row = rows.First();

            IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

            // If the row does not contain the specified cell, create the specified cell.
            if (cells.Count() == 0)
            {
                Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                row.Append(cell);
                worksheet.Save();
            }
        }
    }

    // Given a SpreadsheetDocument and a worksheet name, get the specified worksheet.
    private static Worksheet GetWorksheet(SpreadsheetDocument document, string worksheetName)
    {
        IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
        if (sheets.Count() == 0)
            return null;
        else
            return worksheetPart.Worksheet;
    }

    // Given a cell name, parses the specified cell to get the column name.
    private static string GetColumnName(string cellName)
    {
        // Create a regular expression to match the column name portion of the cell name.
        Regex regex = new Regex("[A-Za-z]+");
        Match match = regex.Match(cellName);

        return match.Value;
    }
    // Given a cell name, parses the specified cell to get the row index.
    private static uint GetRowIndex(string cellName)
    {
        // Create a regular expression to match the row index portion the cell name.
        Regex regex = new Regex(@"\d+");
        Match match = regex.Match(cellName);

        return uint.Parse(match.Value);
    }
}