using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace EasyExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            // easiest read from any Excel file to DataSet object
            var xlsxDataset = ReadFromAllExcelFormats(@".\Files\EasyRead.xlsx");
            var xlsDataset = ReadFromAllExcelFormats(@".\Files\EasyRead.xls");
            var xlsmDataset = ReadFromAllExcelFormats(@".\Files\EasyRead.xlsm");
            // Dear .XLSB file is still stubborn and does not want to work
            //var xlsbDataset = ReadFromAllExcelFormats(@".\Files\EasyRead.xlsb");
            var csvDataset = ReadFromAllExcelFormats(@".\Files\EasyRead.csv");

            // now print prettyfied values just to be readable in console
            var datasets = new List<DataSet>
            {
                xlsxDataset,
                xlsDataset,
                xlsmDataset,
                //xlsbDataset,
                csvDataset
            };
            foreach (var item in datasets)
            {
                var prettyfiedStr = TransformToPrettyString(item);
                Console.WriteLine(prettyfiedStr);
            }
        }

        #region The actual work -- read from any Excel file and return DataSet object
        private static DataSet ReadFromAllExcelFormats(string path)
        {
            if (!File.Exists(path))
            {
                return null;
                //throw new IOException();
            }
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (var reader = CreateDataReader(path, stream))
                {
                    return reader.AsDataSet();
                }
            }
        }

        private static IExcelDataReader CreateDataReader(string path, FileStream stream)
        {
            var fileExtension = Path.GetExtension(path);
            var config = new ExcelReaderConfiguration
            {
                // configuration stuff...
            };

            if (fileExtension == ".csv")
            {
                return ExcelReaderFactory.CreateCsvReader(stream, config);
            }
            else if(fileExtension == ".xlsb")
            {
                return ExcelReaderFactory.CreateBinaryReader(stream, config);
            }
            else
            {
                return ExcelReaderFactory.CreateReader(stream, config);
            }
        }
        #endregion

        #region Transform to prettyfied string
        private static string TransformToPrettyString(DataSet dataSet)
        {
            var sb = new StringBuilder();
            foreach (var table in CreateList(dataSet.Tables))
            {
                sb.AppendLine("--" + table.TableName + "--");
                sb.AppendLine(string.Join(" | ", CreateList(table.Columns)));
                foreach (DataRow row in table.Rows)
                {
                    sb.AppendLine(string.Join(" | ", row.ItemArray));
                }
                sb.AppendLine();
            }
            return sb.ToString();
        }

        private static List<DataTable> CreateList(DataTableCollection collection)
        {
            var list = new List<DataTable>();
            foreach (var table in collection)
            {
                list.Add((DataTable)table);
            }
            return list;
        }

        private static List<DataColumn> CreateList(DataColumnCollection collection)
        {
            var list = new List<DataColumn>();
            foreach (var column in collection)
            {
                list.Add((DataColumn)column);
            }
            return list;
        }
        #endregion
    }
}
