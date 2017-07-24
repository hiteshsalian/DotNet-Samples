using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataReader.Core
{
    public static class DistributionOptions
    {
        public static bool ExportToExcel<T>(string fileName, string sheetName, List<T> entityList) where T:class
        {
            // Export to Excel -- Approach 2 Using Seperate Class and passing the List of Entities
            ExportToExcel<T, List<T>> exportToExcel = new ExportToExcel<T, List<T>>();
            exportToExcel.dataToPrint = entityList;
            exportToExcel.GenerateReport(fileName, sheetName);
            return true;
        }

        public static bool ExportToExcel<T>(string fileName, string sheetName, DataTable dt)
        {
            // Export to Excel -- Approach 1 Extension Method to DataTable
            dt.ExportToExcel(fileName, sheetName);
            return true;
        }

        public static bool ExportToJson<T>(string fileName, List<T> entityList)
        {
            // serialize JSON to a string and then write string to a file
            File.WriteAllText(fileName,JsonConvert.SerializeObject(entityList, Formatting.Indented));
            return true;
        }
    }
}
