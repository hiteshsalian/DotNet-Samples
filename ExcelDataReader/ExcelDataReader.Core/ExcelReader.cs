using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataReader.Core
{ 
    public static class ExcelReader
    {
        public static List<T> ReadExcelData<T>(string fileName, string sheetName, Func<DataRow,T> entityMapper)
        {
            string connectionString = string.Empty;
            string fileExtension = Path.GetExtension(fileName);
            switch (fileExtension.Trim().ToUpper())
            {
                case ".XLS":
                    connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties='Excel 8.0;HDR=YES;IMEX=2'", fileName);
                    break;
                case ".XLSX":
                    connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}; Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=2'", fileName);
                    break;
                default:
                    break;
            }

            var queryString = string.Format("SELECT * FROM [{0}$]", sheetName);
            var adapter = new OleDbDataAdapter(queryString, connectionString);

            var ds = new DataSet();
            adapter.Fill(ds, "ExcelData");

            var dataRowCollection = ds.Tables["ExcelData"].AsEnumerable();
            var query = dataRowCollection.Select(x => entityMapper(x));

            //DataTable dt = query.ToList<T>().ToDataTable<T>();

            //return dt;

            var entityList = query.ToList<T>();

            return entityList;
        }

        
    }
}
