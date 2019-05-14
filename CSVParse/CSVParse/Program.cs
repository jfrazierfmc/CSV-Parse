using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.VisualBasic.FileIO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace CSVParse
{
    class Program
    {
        static void Main(string[] args)
        {
            //string dirPath = @"K:\ImportSource Rebates";
            string dirPath = @"K:\ImportSource Trace";
            //string filePath; Test
            //string fileFormat;
            //string indexString;
            //string customerCode;
            string month;
            string year;
            //string dateFormat;
            //int startRow;

            //Set Month and Year
            month =  "APRIL"; //"MAY"; //"JUNE"; //"JULY"; //"AUGUST"; //"SEPTEMBER"; //"OCTOBER"; //"NOVEMBER"; //"DECEMBER"; //"JANUARY"; //"FEBRUARY"; //"MARCH"; //
            year = "2019";//"2018";//"2017";//"2014";//"2015";//

            try
            {
                DialogResult answer;
                answer = MessageBox.Show("Run Validate - Choose Yes; Import - Choose No; Else Cancel", "Validate or Import", MessageBoxButtons.YesNoCancel);

                if (answer != DialogResult.Cancel)
                {
                    List<string> dirs = new List<string>(Directory.EnumerateDirectories(dirPath));
                    foreach (var dir in dirs)
                    {
                        if (answer == DialogResult.Yes)
                        {
                            if (!ValidateDir(dir))
                            {
                                Console.WriteLine("  {0} is invalid.", dir);
                            }
                            else
                            {
                                ProcessDir(dir, year, month);
                            }
                        }
                        //if (answer == DialogResult.No)
                        //{
                        //    ProcessDir(dir, year, month);

                        //}
                    }
                    Console.WriteLine("{0} directories found.", dirs.Count);

                }

            }
            catch (UnauthorizedAccessException UAEx)
            {
                Console.WriteLine(UAEx.Message);
            }
            catch (PathTooLongException PathEx)
            {
                Console.WriteLine(PathEx.Message);
            }
            catch (Exception Ex)
            {
                Console.WriteLine(Ex.Message);
            }
        }

        static bool ValidateDir(string directory)
        {
            bool isValid = true;
            string[] fileEntries = Directory.GetFiles(directory);
            if (fileEntries.Length > 0)
            {
                isValid = false;
                string dirName = directory.Substring(directory.LastIndexOf("\\") + 1);
                Profile profile = GetProfile(dirName);
                if (profile != null)
                {
                    foreach (string fileName in fileEntries)
                    {
                        Console.WriteLine("Validating {0}.", fileName);
                        isValid = true;
                        if (!ValidateFile(profile, fileName))
                        {
                            isValid = false;
                            break;
                        }
                    }

                }
            }

            return isValid;
        }

        static bool ValidateFile(Profile profile, string fileName)
        {
            //TODO - Get Actual Headers then compare to Expected and update if just moved around
            bool isValid = false;

            if (profile.FileFormat == "EXCEL")
            {
                isValid = ValidateExcel(fileName, profile);
            }
            else
            {
                isValid = ValidateCsv(fileName, profile);
            }

            return isValid;
        }

        static bool ValidateCsv(string filePath, Profile profile)
        {
            bool isValid = false;
            string[] headersExpected = profile.Headers.Split('~');

            FileStream stream = File.Open(filePath, FileMode.Open);
            StreamReader reader = new StreamReader(stream);
            int startRow = profile.HeaderRow;

            string delimiter;

            switch (profile.FileFormat)
            {
                case "CSV":
                    delimiter = ",";
                    break;
                case "TAB":
                    delimiter = "\t";
                    break;
                default:
                    delimiter = ",";
                    break;
            }
            int rowCount = 1;
            int col = 0;
            using (var myFile = new TextFieldParser(reader)) //filePath
            {
                myFile.TextFieldType = FieldType.Delimited;
                myFile.SetDelimiters(delimiter);

                while (!myFile.EndOfData)
                {
                    string[] fieldArray;
                    try
                    {
                        fieldArray = myFile.ReadFields();
                        if (rowCount == startRow)
                        {
                            isValid = true;
                            for (col = 0; col < fieldArray.Length; col++)
                            {
                                if (fieldArray[col].Trim() != headersExpected[col].Trim())
                                {
                                    isValid = false;
                                    break;
                                }
                            }
                        }
                    }
                    catch (Microsoft.VisualBasic.FileIO.MalformedLineException ex)
                    {
                        continue;
                    }
                    rowCount++;
                }
            }
            reader.Close();

            return isValid;
        }

        static bool ValidateExcel(string filePath, Profile profile)
        {
            //string[] fileEntries = Directory.GetFiles(directory);
            //int startRow = profile.HeaderRow;
            bool isValid = true;
            string[] headersExpected = profile.Headers.Split('~');

            
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            //Excel.ProtectedViewWindow pvw;

            string str;
            int rCnt = profile.HeaderRow;
            int cCnt = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(filePath, 0, true);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                str = getCellValue(range, rCnt, cCnt).Replace("\n", "").Replace("\r", "");
                if (str.Trim() != headersExpected[cCnt - 1].Trim().Replace("\n", "").Replace("\r", ""))
                {
                    isValid = false;
                    break;
                }
            }

            xlWorkBook.Close(false);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            return isValid;
        }

        static void ProcessDir(string directory, string year, string month)
        {
            string[] fileEntries = Directory.GetFiles(directory);
            if (fileEntries.Length > 0)
            {
                string dirName = directory.Substring(directory.LastIndexOf("\\") + 1);
                string dirProcessed = directory + "\\Processed";
                string dirError = directory + "\\Error";
                Profile profile = GetProfile(dirName);
                if (profile != null)
                {
                    foreach (string fileName in fileEntries)
                    {
                        //                    ProcessFile(dirName, fileType, fileFormat, headerRow, dateFormat, headers, mapping, fileName, year, month);
                        try
                        {
                            ProcessFile(profile, fileName, year, month);
                            MoveFile(fileName, directory, "Processed");
                        }
                        catch
                        {
                            MoveFile(fileName, directory, "Error");
                        }


                    }

                }
            }
        }

        static void MoveFile(string fileName, string directory, string status)
        {
            string destDir = directory + "\\" +status;
            string destFile = fileName.Replace(directory, destDir);
            if (!Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            Directory.Move(fileName, destFile);

        }
        static Profile GetProfile(string directory)
        {
            string connectionString =
               "Data Source=(local);Initial Catalog=Commissions;"
               + "Integrated Security=true";

            Profile profile = null;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.CommandText = "SELECT [ProfileSettingID],[ProfileID],[FileType],[FileFormat],[HeaderRow],[StartRow],[DateFormat],[Headers],[Mapping],[Folder] FROM[dbo].[ProfileSetting] WHERE [Folder] = @P0";
                    cmd.Parameters.AddWithValue("@P0", directory);
                    cmd.Connection = connection;

                    SqlDataReader reader = cmd.ExecuteReader();

                    if (reader.HasRows)
                    {
                        profile = new Profile();
                        while (reader.Read())
                        {
                            profile.ProfileID = reader.GetString(1);
                            profile.FileType = reader.GetString(2);
                            profile.FileFormat = reader.GetString(3);
                            profile.HeaderRow = reader.GetInt32(4);
                            //profile.StartRow = reader.GetInt32(5);
                            if (reader.IsDBNull(5))
                            {
                                profile.StartRow = profile.HeaderRow + 1;
                            }
                            else
                            {
                                profile.StartRow = reader.GetInt32(5);
                            }
                            profile.DateFormat = reader.GetString(6);
                            profile.Headers = reader.GetString(7);
                            profile.Mapping = reader.GetString(8);
                        }
                    }
                    else
                    {
                        Console.WriteLine("No rows found.");
                    }
                    reader.Close();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            return profile;
        }
        class Profile
        {
            public string ProfileID { get; set; }
            public string FileType { get; set; }
            public string FileFormat { get; set; }
            public int HeaderRow { get; set; }
            public int StartRow { get; set; }
            public string DateFormat { get; set; }
            public string Headers { get; set; }
            public string Mapping { get; set; }

        }
        //static void ProcessFile(string profileID, string fileType, string fileFormat, int headerRow, string dateFormat, string headers, string mapping, string fileName, string year, string month)
        static void ProcessFile(Profile profile, string fileName, string year, string month)
        {
            Console.WriteLine("{0}", fileName);

            if (profile.FileFormat == "EXCEL")
            {
                LoadExcel(fileName, profile.ProfileID, profile.Mapping, profile.StartRow, month, year, profile.DateFormat, profile.FileType);
            }
            else
            {
                LoadFile(fileName, profile.FileFormat, profile.ProfileID, profile.Mapping, profile.StartRow, month, year, profile.FileType);
            }

        }
        static void Test()
        {

            string filePath;
            string fileFormat;
            string indexString;
            string customerCode;
            string month;
            string year;
            string dateFormat;
            int startRow;
            month = "APRIL";
            year = "2016";

            indexString = ",,,7,,,,,8,9,0,,,,,,,,,,,1,3,,5,,6,,,,4,,,,,,,,,,";
            //            filePath = @"K:\Excel\Excel.xlsx";
            //            filePath = @"K:\Excel\Excel.xls";
            filePath = @"K:\CSV\CSV.csv";
            startRow = 3;
            dateFormat = "D";
            customerCode = "MOORE";
            LoadExcel(filePath, customerCode, indexString, startRow, month, year, dateFormat,"");

            indexString = ",,,7,,,,,8,9,0,,,,,,,,,,,1,3,,5,,6,,,,4,,,,,,,,,,";
            fileFormat = "CSV";
            filePath = @"K:\CSV\CSV.csv";
            startRow = 3;
            customerCode = "MOORE";
            LoadFile(filePath, fileFormat, customerCode, indexString, startRow, month, year,"");
            //LoadCsv(filePath);

            fileFormat = "TAB";
            filePath = @"K:\Tab\Tab.csv";
            indexString = "0,29";
            LoadFile(filePath, fileFormat, customerCode, indexString, startRow, month, year,"");
            //LoadTab(filePath);
        }

        static void LoadFile(string filePath, string fileFormat, string customerCode, string indexString, int startRow, string month, string year, string fileType)
        {
            FileStream stream = File.Open(filePath, FileMode.Open);
            StreamReader reader = new StreamReader(stream);
            string[] indexes = indexString.Split(',');
            string delimiter;
            switch (fileFormat)
            {
                case "CSV":
                    delimiter = ",";
                    break;
                case "TAB":
                    delimiter = "\t";
                    break;
                default:
                    delimiter = ",";
                    break;
            }
            int rowCount = 1;
            using (var myFile = new TextFieldParser(reader)) //filePath
            {
                myFile.TextFieldType = FieldType.Delimited;
                myFile.SetDelimiters(delimiter);

                while (!myFile.EndOfData)
                {
                    string[] fieldArray;
                    try
                    {
                        fieldArray = myFile.ReadFields();
                        if (rowCount >= startRow)
                        {
                            bool inserted;
                            inserted = InsertRow(fieldArray, customerCode, indexes, month, year, "S", fileType);
                            if (!inserted)
                                break;
                        }
                    }
                    catch (Microsoft.VisualBasic.FileIO.MalformedLineException ex)
                    {
                        continue;
                    }
                    rowCount++;
                }
            }
            reader.Close();
        }

        static bool InsertRow(string[] fieldArray, string customerCode, string[] indexes, string month, string year, string dateFormat, string fileType)
        {
            int length = 41;
            string[] insertSource = new string[length];

            for (int i = 0; i < length; i++)
            {
                insertSource[i] = getValue(fieldArray, indexes, i);
            }

            if (!hasData(insertSource)) //No data so exit
                return false;

            string connectionString =
                "Data Source=(local);Initial Catalog=Commissions;"
                + "Integrated Security=true";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.CommandText = getSql(customerCode, insertSource, month, year, fileType);
                    AddParameters(cmd, customerCode, insertSource, month, year, dateFormat);
                    cmd.Connection = connection;
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }
            return true;

        }

        static bool hasData(string[] insertSource)
        {
            int threshold = 3;
            int dataCount = 0;

            for (int i = 0; i < insertSource.Length; i++)
            {
                if (insertSource[i] != String.Empty)
                    dataCount++;
                if (dataCount >= threshold)
                    return true;
            }
            return false;

        }

        static string getSql(string customerCode, string[] insertSource, string month, string year, string fileType)
        {
            string sql = @"INSERT INTO [dbo].[TraceImport] 
           ([PROFILE_ID] 
           ,[DOCUMENT_ID]
           ,[CONTRACT_NUMBER]
           ,[SHIPTO_NAME]
           ,[SHIPTO_LOCATION_ID_NUMBER]
           ,[SHIPTO_OTHER_NAME]
           ,[SHIPTO_ADDRESS_ONE]
           ,[SHIPTO_ADDRESS_TWO]
           ,[SHIPTO_CITY]
           ,[SHIPTO_STATE]
           ,[SHIPTO_ZIP]
           ,[SHIPTO_COUNTRY]
           ,[BILLTO_NAME]
           ,[BILLTO_LOCATION_ID_NUMBER]
           ,[BILLTO_OTHER_NAME]
           ,[BILLTO_ADDRESS_ONE]
           ,[BILLTO_ADDRESS_TWO]
           ,[BILLTO_CITY]
           ,[BILLTO_STATE]
           ,[BILLTO_ZIP]
           ,[BILLTO_COUNTRY]
           ,[ITEM_ONE]
           ,[ITEM_TWO]
           ,[ITEM_THREE]
           ,[QUANTITY_SOLD]
           ,[QUANTITY_RETURNED]
           ,[UNIT_OF_MEASURE]
           ,[ITEM_COST]
           ,[ACQ_COST EXtended]
           ,[INVOICE_DATE]
           ,[INVOICE]
           ,[RECORD_COUNTER]
           ,[DistWarehouseID]
           ,[REBATE AMOUNT]
           ,[report month]
           ,[MEMO_NUM]
           ,[HIN GPO ID]
           ,[Report YEar]
           ,[GLN]
           ,[MARKET TYPE]
           ,[MARKET CODE])
     VALUES (
            @P0, @P1, @P2, @P3, @P4, @P5, @P6, @P7, @P8, @P9,@P10, @P11, @P12, @P13, @P14, @P15, @P16, @P17, @P18, @P19,
            @P20, @P21, @P22, @P23, @P24, @P25, @P26, @P27, @P28, @P29,@P30, @P31, @P32, @P33, @P34, @P35, @P36, @P37, @P38, @P39, @P40)";


            if (fileType == "T")
                return sql.Replace("\r", "").Replace("\n", "");
            else
                return sql.Replace("\r", "").Replace("\n", "").Replace("[TraceImport]", "[TraceImportRebate]");

        }

        static void AddParameters(SqlCommand cmd, string customerCode, string[] insertSource, string month, string year, string dateFormat)
        {
            cmd.Parameters.AddWithValue("@P0", customerCode);
            cmd.Parameters.AddWithValue("@P1", insertSource[1].Trim());
            cmd.Parameters.AddWithValue("@P2", insertSource[2].Replace("'", "").Trim()); //Contract Number
            cmd.Parameters.AddWithValue("@P3", insertSource[3].Trim());
            cmd.Parameters.AddWithValue("@P4", insertSource[4].Trim());
            cmd.Parameters.AddWithValue("@P5", insertSource[5].Trim());
            cmd.Parameters.AddWithValue("@P6", insertSource[6].Trim().Replace("null", ""));
            cmd.Parameters.AddWithValue("@P7", insertSource[7].Trim().Replace("null", ""));
            cmd.Parameters.AddWithValue("@P8", insertSource[8].Trim());
            cmd.Parameters.AddWithValue("@P9", insertSource[9].Trim());
            if (insertSource[10].Replace("'", "").Length > 5)
                cmd.Parameters.AddWithValue("@P10", insertSource[10].Trim().Replace("'", "").Replace("-", "").Substring(0, 5)); //Ship To Zip Code
            else
                cmd.Parameters.AddWithValue("@P10", insertSource[10].Trim().Replace("'", "").Replace("-", "")); //Ship To Zip Code
            cmd.Parameters.AddWithValue("@P11", insertSource[11].Trim());
            cmd.Parameters.AddWithValue("@P12", insertSource[12].Trim());
            cmd.Parameters.AddWithValue("@P13", insertSource[13].Trim());
            cmd.Parameters.AddWithValue("@P14", insertSource[14].Trim());
            cmd.Parameters.AddWithValue("@P15", insertSource[15].Trim().Replace("null", ""));
            cmd.Parameters.AddWithValue("@P16", insertSource[16].Trim().Replace("null", ""));
            cmd.Parameters.AddWithValue("@P17", insertSource[17].Trim());
            cmd.Parameters.AddWithValue("@P18", insertSource[18].Trim());
            if (insertSource[19].Replace("'", "").Length > 5)
                cmd.Parameters.AddWithValue("@P19", insertSource[19].Trim().Replace("'", "").Replace("-", "").Substring(0, 5)); //Ship To Zip Code
            else
                cmd.Parameters.AddWithValue("@P19", insertSource[19].Trim().Replace("'", "").Replace("-", "")); //Ship To Zip Code
            cmd.Parameters.AddWithValue("@P20", insertSource[20].Trim());
            cmd.Parameters.AddWithValue("@P21", insertSource[21].Replace("'", "").Trim()); //Item One
            cmd.Parameters.AddWithValue("@P22", insertSource[22].Trim());
            cmd.Parameters.AddWithValue("@P23", insertSource[23].Trim());
            cmd.Parameters.AddWithValue("@P24", getRoundedValue(insertSource[24], 2)); //Quantity Sold
            cmd.Parameters.AddWithValue("@P25", insertSource[25].Trim());
            cmd.Parameters.AddWithValue("@P26", insertSource[26].Trim());
            cmd.Parameters.AddWithValue("@P27", getRoundedValue(insertSource[27], 2)); //Item Cost
            cmd.Parameters.AddWithValue("@P28", getRoundedValue(insertSource[28], 2)); //Extended Cost
            cmd.Parameters.AddWithValue("@P29", getDateValue(insertSource[29], dateFormat)); //Invoice Date
            cmd.Parameters.AddWithValue("@P30", insertSource[30].Trim());
            cmd.Parameters.AddWithValue("@P31", insertSource[31].Trim());
            cmd.Parameters.AddWithValue("@P32", insertSource[32]);
            cmd.Parameters.AddWithValue("@P33", getRoundedValue(insertSource[33], 2)); //Rebate Amount
            cmd.Parameters.AddWithValue("@P34", month);
            cmd.Parameters.AddWithValue("@P35", insertSource[35].Trim());
            cmd.Parameters.AddWithValue("@P36", insertSource[36].Trim());
            cmd.Parameters.AddWithValue("@P37", year);
            cmd.Parameters.AddWithValue("@P38", insertSource[38].Trim());
            cmd.Parameters.AddWithValue("@P39", insertSource[39].Trim());
            cmd.Parameters.AddWithValue("@P40", insertSource[40].Trim());
        }

        static string getRoundedValue(string numberString, int decimals)
        {
            string str = string.Empty;
            try
            {
                double n = Math.Round(double.Parse(numberString.Replace("$", "")), decimals);
                str = n.ToString();
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.Message);
            }
            if (str == string.Empty)
                str = numberString;

            return str;

        }
        static string getDateValue(string dateString, string dateFormat)
        {
            string str = string.Empty;
            if (dateFormat == "D")
            {
                try
                {
                    double date = double.Parse(dateString);
                    str = String.Format("{0:MM/dd/yyyy}", DateTime.FromOADate(date));
                }
                catch (Exception ex)
                {

                }
            }
            if (str == string.Empty)
                str = dateString;

            return str;
        }

        static string getValue(string[] fieldArray, string[] indexes, int position)
        {
            string str;

            if (position >= indexes.Length)
                str = string.Empty;
            else if (indexes[position] == string.Empty)
                str = string.Empty;
            else if (indexes[position] != string.Empty && int.Parse(indexes[position]) <= fieldArray.Length - 1)
                str = fieldArray[int.Parse(indexes[position])].Replace("\"", "").Trim();
            else if (int.Parse(indexes[position]) > fieldArray.Length - 1)
                str = string.Empty;
            else
                str = string.Empty;

            return str;
        }
        //static void LoadTab(string filePath)
        //{
        //    FileStream stream = File.Open(filePath, FileMode.Open);

        //    StreamReader reader = new StreamReader(stream);

        //    string s = "0,29";
        //    string[] indexes = s.Split(',');

        //    using (var myTabFile = new TextFieldParser(reader)) //filePath
        //    {
        //        myTabFile.TextFieldType = FieldType.Delimited;
        //        myTabFile.SetDelimiters("\t");

        //        while (!myTabFile.EndOfData)
        //        {
        //            string[] fieldArray;
        //            try
        //            {
        //                string profileID = "";
        //                string unitCost = "";

        //                fieldArray = myTabFile.ReadFields();
        //                profileID = fieldArray[int.Parse(indexes[0])];
        //                unitCost = fieldArray[int.Parse(indexes[1])];
        //            }
        //            catch (Microsoft.VisualBasic.FileIO.MalformedLineException ex)
        //            {
        //                continue;
        //            }
        //        }
        //    }
        //}
        //    static void LoadCsv(string filePath)
        //{
        //    FileStream stream = File.Open(filePath, FileMode.Open);

        //    StreamReader reader = new StreamReader(stream);

        //    using (var myCsvFile = new TextFieldParser(reader)) //filePath
        //    {
        //        myCsvFile.TextFieldType = FieldType.Delimited;
        //        myCsvFile.SetDelimiters(",");

        //        while (!myCsvFile.EndOfData)
        //        {
        //            string[] fieldArray;
        //            try
        //            {
        //                fieldArray = myCsvFile.ReadFields();
        //            }
        //            catch (Microsoft.VisualBasic.FileIO.MalformedLineException ex)
        //            {
        //                continue;
        //            }
        //        }
        //    }

        //}

        static void LoadExcel(string filePath, string customerCode, string indexString, int startRow, string month, string year, string dateFormat, string fileType)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            Excel.ProtectedViewWindow pvw;

            string str;
            int rCnt = 0;
            int cCnt = 0;

            string[] indexes = indexString.Split(',');

            xlApp = new Excel.Application();
            //xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow;
            //xlWorkBook = xlApp.Workbooks.Open("csharp.net-informations.xls", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkBook = xlApp.Workbooks.Open(filePath, 0, true);
            //pvw = xlApp.ProtectedViewWindows.Open(filePath);
            //xlWorkBook = pvw.Workbook;
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            try { 

                string[] fieldArray = new string[range.Columns.Count];
                for (rCnt = startRow; rCnt <= range.Rows.Count; rCnt++)
                {
                    for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                    {
                        str = getCellValue(range, rCnt, cCnt);
                        fieldArray[cCnt - 1] = str;
                    }
                    bool inserted;
                    inserted = InsertRow(fieldArray, customerCode, indexes, month, year, dateFormat, fileType);
                    if (!inserted)
                        break;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

            }
            finally
            {
                xlWorkBook.Close(false);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }


        }
        static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
        static string getCellValue(Excel.Range range, int row, int column)
        {
            object cellValue = (range.Cells[row, column] as Excel.Range).Value2;
            if (cellValue != null)
            {
                return Convert.ToString(cellValue);
            }
            else
            {
                return string.Empty;
            }
        }

        

    }
}