using Funcs;
using ExcelLibrary;


namespace appGenerateExcel
{
    class Program
    {
        static void Main(string[] args)
        {

            // ----------------------------------------------------------------------------------------------------------

                // DB - create a connection to MS SQL DB (optional)
            //DB db = new DB();
                // 1. You need to open connection with command CreateSqlConnection(connection_line)
            //string connection_line = "Data Source=location_of_db;Initial Catalog=NameDB;User ID=NAME_user;Password=Password_of_db";
            //db.CreateSqlConnection(connection_line);
                // 2. To retrieve data from a query, run the command db.GetData("SELECT * FROM table")
                //    Data format - Tuple<int, List<String[]>>
                //    .Item1 - The number of columns in the resulting table
                //    .Item2 - List of rows, each of which consists of an array of N(string) columns - received values
                //    At the moment, all data are automatically converted to string format. Support for dynamic data format is planned for the future.
            //string example = db.GetData("SELECT id FROM table").Item2[0][0];
                // 3. At the end of all operations with DB you need to close the connection. CloseConn()
            //db.CloseConn();

            // ----------------------------------------------------------------------------------------------------------

                // Esyle - Stylesheet generate
                // All cell style based on Font, Fill, Border and in the end we "mix" this parameters to create CellFormat
                // All components connecting with ID from 0
                // Color sheme - HEX: 11223344: 11-Bright, 22-R, 33-G, 44-B
            Estyle.InitBase();
                // Important! Init base create default: Fill(0 - transpared, 1 - grid), Border (0), Cell(0) (yes, it's bug, may be fix in the future)
                // Some colors constant:
            string darkred = "FFd10000";
            string darkblue = "FF021373";
            string gray_back = "FFBCBCBC";
            string yellow_neon = "FFeff705";
            string black = "00000000";
                // Fontpart (size, color, fontname, isbold, isitalic, isunderline)
            Estyle.NewFont(size:14.0, color:darkred, isBold:true);        // 0
            Estyle.NewFont(size:14.0, color:darkred, isBold:true);        // 1
            Estyle.NewFont(size:14.0, color:darkblue, isBold:true);       // 2
            Estyle.NewFont(size:11.0, color:black, isBold:true);          // 3
            Estyle.NewFont(size:11.0, color:black);                       // 4

                // Fillpart (color)             // 0 white, 1 grid
            Estyle.NewFill(color:gray_back);    // 2
            Estyle.NewFill(color:yellow_neon);  // 3

                // Borderpart (style)   // 0 common
            Estyle.NewBorder("medium"); // 1
            Estyle.NewBorder("dotted"); // 2
            Estyle.NewBorder("thin");   // 3

                //Cellformat (idFONT, idFILL, idBORDER)
            Estyle.NewCellFormat(0, 2, 0, 1);   // 0 not delete, don't use
            Estyle.NewCellFormat(1, 2, 0, 1);   // 1 header blue 14
            Estyle.NewCellFormat(2, 2, 0, 1);   // 2 header red 14
            Estyle.NewCellFormat(3, 3, 0, 1);   // 3 subheader yellow_back bold 11
            Estyle.NewCellFormat(4, 3, 0, 1);   // 4 subheader yellow_back 11
            Estyle.NewCellFormat(3, 2, 0, 1);   // 5 subheader gray_back bold 11
            Estyle.NewCellFormat(4, 0, 2, 0);   // 6 cell white_back 11

            // ----------------------------------------------------------------------------------------------------------
                // Excellib - app for generate Excel books
            ExcelLib app = new ExcelLib(); 
                // NewExFile - Create EX file and book  (name_of_doc, StyleSheet) [if we put "auto" to name_of_doc arg we get name with format ddMMyyyy.xlsx]
            app.NewExFile("Excel.xlsx", Estyle.CreateStyleSheet()); 
                // AddSheet - Create Sheet  (name_sheet)
                // new sheet add to the end of sheets and become a active sheet
            app.AddSheet("Sheet1"); 
            app.AddSheet("Sheet2"); // this is an active sheet at the moment
                // SetActSheet - Set Active Sheet (name_sheet)
            app.SetActSheet("Sheet1");
                // AddCell - Create cells (row, column, label, style, (sheet_name)=activeSheet)
            app.AddCell("A", 1, "This", 3);
            app.AddCell("B", 2, "is", 2);
            app.AddCell("C", 3, "coooool", 1);
                // AddMerge - Create merge between cells (cell_from, cell_to, (sheet_name)=activeSheet)
            app.AddMerge("C3", "H3");
                // CloseEx - close the book            
            app.CloseEx();            
        }
    
    }
}
