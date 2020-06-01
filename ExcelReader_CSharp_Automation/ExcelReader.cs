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
               excelData.sheetData.Add(table.TableName, table);
            }
            _cache.Add(filePath, excelData);
        }

        public static int GetRowCount(string filepath, string sheetName, string columnName)
        {
            try
            {
                //Load the file if it doesn't exist already
                if (!_cache.ContainsKey(filepath))
                {
                    Load(filepath);
                }
                //Retriving Data 
                int data = _cache[filepath].sheetData[sheetName].Rows.Count;

                return data;
            }
            catch (Exception e)
            {
                return 0;
            }
        }

        public static void Load(string dataFilePath)
        {
            if (!_cache.ContainsKey(dataFilePath))
             {
                DataSet result = OpenExcel(dataFilePath);
                PopulateExcelDictionary(dataFilePath, result);
            }
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
                //Retriving Data 
                string data = _cache[filepath].sheetData[sheetName].Rows[rowNumber][columnName].ToString();
                             
                return data;
            }
            catch (Exception e)
            {
                return null;
            }
        }


    }
       public class ExcelDocData
        {
            public IDictionary<string, DataTable> sheetData;
            public ExcelDocData()
            {
                sheetData = new Dictionary<string, DataTable>();

            }
        }
   
}
