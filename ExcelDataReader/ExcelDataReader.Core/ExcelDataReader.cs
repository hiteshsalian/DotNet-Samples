using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using ExcelDataReader.Core.Entities;
 
namespace ExcelDataReader.Core
{
    public class ExcelDataReader
    {
        public DataTable ReadExcelData(string fileName, string sheetName)
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

            return ds.Tables["ExcelData"];
        }

        public void ReadExcel(string fileName, string sheetName)
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
            adapter.Fill(ds, "RawDataTable");

           // DataTable dataTbl = ds.Tables["RawDataTable"];
            var dataTbl = ds.Tables["RawDataTable"].AsEnumerable();
            var query = dataTbl.Where(x => x.Field<bool>("Discontinued") == false).Select(x => 
            new Product
            {
                ProductID = Convert.ToInt32(x["ProductID"]),
                ProductName = x["ProductName"].ToString(),
                UnitPrice = Convert.ToDouble(x["UnitPrice"]),
                CategoryID = Convert.ToInt32(x["CategoryID"])
            });

            // Export to Excel -- Approach 1 Extension Method to DataTable
            DataTable dt = query.ToList<Product>().ToDataTable<Product>();
            dt.ExportToExcel(@"C:\Product.xlsx", "product");
            Console.WriteLine(dt.Rows.Count);

            // Export to Excel -- Approach 2 Using Seperate Class and passing the List of Products
            //List<Product> productList = query.ToList<Product>();
            //ExportToExcel<Product, List<Product>> exportToExcel = new ExportToExcel<Product, List<Product>>();
            //exportToExcel.dataToPrint = productList;
            //exportToExcel.GenerateReport(@"C:\testfile.xlsx", "finalData");
        }

        private Product MapDataRow(DataRow x)
        {
            return new Product {
                ProductID = x.Field<int>("ProductID"),
                ProductName = x.Field<string>("ProductName"),
                UnitPrice = x.Field<double>("UnitPrice"),
                CategoryID = x.Field<int>("CategoryID")
            };
           

            //Product p = new Product();

            //p.ProductID = Convert.ToInt32(dr["ProductID"]);
            //p.ProductName = dr["ProductName"].ToString();
            //p.UnitPrice = Convert.ToDouble(dr["UnitPrice"]);
            //p.CategoryID = Convert.ToInt32(dr["CategoryID"]);
            //return p;
            
            
        }
    }
}
