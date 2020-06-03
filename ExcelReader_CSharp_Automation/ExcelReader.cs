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

        static ExcelReader()
        {
            _cache = new Dictionary<string, ExcelDocData>();

        }
        private static DataSet OpenExcel(string filePath, string pwd = null)
        {
            FileStream stream = new FileStream(filePath, FileMode.Open);
            IExcelDataReader reader;
            ExcelReaderConfiguration readerConf = null;
            if (!String.IsNullOrEmpty(pwd))
            {
                readerConf = new ExcelReaderConfiguration { Password = pwd };
                
            }
                reader = ExcelReaderFactory.CreateOpenXmlReader(stream, readerConf);


                DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    { UseHeaderRow = true }
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

        public static void Load(string dataFilePath, string pwd = null)
        {
            try
            {
                if (!_cache.ContainsKey(dataFilePath) && File.Exists(dataFilePath))
                {

                    DataSet result = OpenExcel(dataFilePath, pwd);
                    PopulateExcelDictionary(dataFilePath, result);
                }
            }
            catch (Exception ex)
            {

            }
        }

        public static int GetRowCount(string filepath, string sheetName, string pwd = null)
        {
            try
            {
                int count;
                //Load the file if it doesn't exist already
                if (!_cache.ContainsKey(filepath))
                {
                    if (File.Exists(filepath))
                    {
                        Load(filepath,pwd);
                        count = _cache[filepath].sheetData[sheetName].Rows.Count;
                        return count;
                    }
                    else
                    {
                        return 0;
                    }

                }
                else
                {
                    //Getting row count 
                    count = _cache[filepath].sheetData[sheetName].Rows.Count;
                    return count;
                }
            }
            catch (Exception ex)
            {
                return 0;
            }
        }




        public static string ReadData(string filepath, string sheetName, int rowNumber, string columnName, string pwd = null)
        {
            try
            {
                //Load the file if it doesn't exist already
                if (!_cache.ContainsKey(filepath))
                {
                    if (File.Exists(filepath))
                    {
                        Load(filepath, pwd);
                        //First data row corresponds to index 0, to get row 1, do rowNumber -1
                        //Retrieving Data 
                        string data = _cache[filepath].sheetData[sheetName].Rows[rowNumber - 1][columnName].ToString();

                        return data;
                    }
                    else
                    { return null; }

                }
                else
                {

                    //First data row corresponds to index 0, to get row 1, do rowNumber -1
                    //Retrieving Data 
                    string data = _cache[filepath].sheetData[sheetName].Rows[rowNumber - 1][columnName].ToString();

                    return data;
                }
            }
            catch (Exception ex)
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
