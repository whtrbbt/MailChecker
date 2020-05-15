using Microsoft.VisualBasic.FileIO;
using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;


namespace CSVUtility
{
    public static class CSVUtility
    {
        public static void ToCSV(this DataTable dtDataTable, string strFilePath)
        // Сохраняет DataTable в файл CSV
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers  

            //Паттерн для поиска разделителя в полях таблицы
            string pattern = ";+";
            
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(";");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        //if (value.Contains(';'))
                        //{
                        //value = String.Format("\\{0}\\", value);
                        value = Regex.Replace(value, @"\n+", " ");    
                        value = Regex.Replace(value, pattern, ":");
                            sw.Write(value);
                        //}
                        //else
                        //{
                        //    sw.Write(dr[i].ToString());
                        //}
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(";");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        public static void ToXLSX(DataTable dtDataTable, string strFilePath, string tmplFileName )
        // Схраняет DataTable в файл XLXS
        {
            Excel.Application exc = new Microsoft.Office.Interop.Excel.Application();
            Excel.XlReferenceStyle RefStyle = exc.ReferenceStyle;
            Excel.Workbook wb = null;

            try
            {
                wb = exc.Workbooks.Add(tmplFileName);
            }
            catch (SystemException ex)
            {
                throw new Exception("Не удалось загрузить шаблон для экспорта" + tmplFileName + "\n" + ex.Message);
            }


            Excel.Worksheet whs1 = wb.Worksheets.get_Item(1) as Excel.Worksheet;

            //Заполняем заголовок таблицы
            Excel.Range header;
            Int32 columnCount;
            columnCount = dtDataTable.Columns.Count;

            object[,] objHeaderData = new Object[1, columnCount];

            for (int i = 0; i < columnCount; i++)
            {
                objHeaderData[0, i] = dtDataTable.Columns[i].ToString();
            }
            header = whs1.get_Range("A1", "A1");
            header = header.get_Resize(1, columnCount);
            header.Value = objHeaderData;
            
            int r = 2; //Отступаем одну строку с заголовком
            int c = 1;
            foreach (DataRow dr in dtDataTable.Rows)
            {
                foreach (DataColumn dc in dtDataTable.Columns)
                {
                    Excel.Range excelcell = whs1.Cells[r, c];
                    excelcell.NumberFormat = "@";
                    excelcell.Value2 = dr[dc].ToString();
                    c++;
                }
                c = 1;
                r++;
            }
            whs1.Columns.AutoFit();
            wb.SaveAs(strFilePath);
            exc.Quit();

        }

        public static DataTable GetDataTableFromCSVFile(string file_path)
        // Преобразует данные их CSV файла в DataTable
        {
            Console.WriteLine(file_path);

            //string[] lines = System.IO.File.ReadAllLines(file_path);

            //// Display the file contents by using a foreach loop.
            //System.Console.WriteLine("Contents of PAY_DOC.CSV = ");
            //foreach (string line in lines)
            //{
            //    // Use a tab to indent each line of the file.
            //    Console.WriteLine("\t" + line);
            //}

            DataTable csvData = new DataTable();
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(file_path))
                {
                    csvReader.SetDelimiters(new string[] { ";" });
                    csvReader.HasFieldsEnclosedInQuotes = false;
                    string[] colFields = csvReader.ReadFields();
                    Console.WriteLine("Количество столбцов: {0}", colFields.Length);
                    foreach (string column in colFields)
                    {
                        DataColumn datecolumn = new DataColumn(column);
                        //Console.WriteLine("Поле: {0}", column);
                        datecolumn.AllowDBNull = true;
                        csvData.Columns.Add(datecolumn);
                    }
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        //Making empty value as null
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] == "")
                            {
                                fieldData[i] = null;
                            }
                        }
                        csvData.Rows.Add(fieldData);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: {0}", ex);
                return null;
            }
            return csvData;
        }

        public static DataTable GetDataTableFromXLXS (string file_path, bool header = true)
        {
            DataTable dtResult = null;
            int totalSheet = 0; //No of sheets on excel file  
            using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file_path + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))
            {
                objConn.Open();
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = string.Empty;
                if (dt != null)
                {
                    var tempDataTable = (from dataRow in dt.AsEnumerable()
                                         where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                         select dataRow).CopyToDataTable();
                    dt = tempDataTable;
                    totalSheet = dt.Rows.Count;
                    sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                }
                cmd.Connection = objConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds, "excelData");
                dtResult = ds.Tables["excelData"];
                objConn.Close();
                return dtResult; //Returning Dattable  
            }
        }

        public static void InsertDataIntoSQLServerUsingSQLBulkCopy(DataTable dataTable, string tableName, string connectionString)
        // Вставляет данные из dataTable в таблицу tableName в базе данных MS SQL
        {
            using (SqlConnection dbConnection = new SqlConnection(connectionString))
            {
                dbConnection.Open();
                using (SqlBulkCopy s = new SqlBulkCopy(dbConnection))
                {
                    s.DestinationTableName = tableName;
                    s.EnableStreaming = true;
                    s.BatchSize = 10000;
                    s.BulkCopyTimeout = 0;
                    s.NotifyAfter = 100;
                    s.SqlRowsCopied += delegate (object sender, SqlRowsCopiedEventArgs e)
                    {
                        Console.WriteLine(e.RowsCopied.ToString("#,##0") + " rows copied.");
                    };
                    foreach (var column in dataTable.Columns)
                    {
                        s.ColumnMappings.Add(column.ToString(), column.ToString());

                        Console.WriteLine();
                    }
                    s.WriteToServer(dataTable);
                }
                dbConnection.Close();
            }
        }


    }  
 
}

