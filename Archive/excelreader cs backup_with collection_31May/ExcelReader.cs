using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Remya.ExcelReader
{
    public class ExcelReader
    {
       
            private static IDictionary<string, ExcelDocData> _cache;

            private string filePath { get; set; }

            static ExcelReader()
            {
                _cache = new Dictionary<string, ExcelDocData>();

            }
            public static DataSet OpenExcel(string filePath)
        {

            FileStream stream = new FileStream(filePath, FileMode.Open);
            IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });
            return result;
           
        }

        private static void PopulateExcelDictionary(string filePath, DataSet result)
        {
            ExcelDocData excelData = new ExcelDocData();
            foreach (DataTable table in result.Tables)
            {
                
                List<Datacollection> dataCol = new List<Datacollection>();
                dataCol = PopulateSheetDataInCollection(table);
                excelData.sheetData.Add(table.TableName, dataCol);
            }
            _cache.Add(filePath, excelData);
        }

        public static void Load(string dataFilePath)
        {
            DataSet result = OpenExcel(dataFilePath);
            PopulateExcelDictionary(dataFilePath, result);
        }

        public static List<Datacollection> PopulateSheetDataInCollection(DataTable table)
            {

                List<Datacollection> dataCol = new List<Datacollection>();
                //Iterate through the rows and columns of the Table
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        Datacollection dcCellValue = new Datacollection()
                        {
                            rowNumber = row,
                            colName = table.Columns[col].ColumnName,
                            colValue = table.Rows[row - 1][col].ToString()
                        };
                        //Add all the details for each row
                        dataCol.Add(dcCellValue);
                    }
                }
                return dataCol;


            }
            public static string ReadData(string filepath, string sheetName, int rowNumber, string columnName)
            {
                try
                {
                //Load the file if it doesn't exist already
                    if (!_cache.ContainsKey(filepath))
                    {
                        Load(filepath);
                    }
                    //Retriving Data using LINQ 
                    string data = (from colData in _cache[filepath].sheetData[sheetName]
                                   where colData.colName == columnName && colData.rowNumber == rowNumber
                                   select colData.colValue).SingleOrDefault();

                    //var datas = dataCol.Where(x => x.colName == columnName && x.rowNumber == rowNumber).SingleOrDefault().colValue;
                    return data.ToString();
                }
                catch (Exception e)
                {
                    return null;
                }
            }


        }
        public class Datacollection
        {
            public int rowNumber { get; set; }
            public string colName { get; set; }
            public string colValue { get; set; }
        }
        public class ExcelDocData
        {
            public IDictionary<string, List<Datacollection>> sheetData;
            public ExcelDocData()
            {
                sheetData = new Dictionary<string, List<Datacollection>>();

            }
        }
   
}
