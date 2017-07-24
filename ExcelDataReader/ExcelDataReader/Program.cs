using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader.Core;
using ExcelDataReader.Core.Entities;
namespace ExcelDataReader
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello");

            // NON GENERIC WAY
            //  ExcelDataReader excelReader = new ExcelDataReader();
            // excelReader.ReadExcel(@"C:\Northwind.xls","Product");
            //  excelReader.ReadExcel(@"C:\Travelers Data- WMC.xlsx", "Data");  

            //GENERIC WAY
            var products = ExcelReader.ReadExcelData<Product>(@"C:\Northwind.xls", "Product", EntityMapper.MapProduct);
            var categories = ExcelReader.ReadExcelData<Category>(@"C:\Northwind.xls", "Categories", EntityMapper.MapCategory);
            var employees = ExcelReader.ReadExcelData<Employee>(@"C:\Northwind.xls", "Employees", EntityMapper.MapEmployee);

            DistributionOptions.ExportToExcel(@"C:\ProductOutput.xlsx", "ProductOutput",products);
            DistributionOptions.ExportToJson<Product>(@"C:\products.json", products);

            DistributionOptions.ExportToExcel(@"C:\CategoryOutput.xlsx", "CategoryOutput", categories);
            DistributionOptions.ExportToJson<Category>(@"C:\categories.json", categories);

            DistributionOptions.ExportToExcel(@"C:\EmployeesOutput.xlsx", "EmployeesOutput", employees);
            DistributionOptions.ExportToJson<Employee>(@"C:\employees.json", employees);
            Console.Read();
               
        }
    }
}
