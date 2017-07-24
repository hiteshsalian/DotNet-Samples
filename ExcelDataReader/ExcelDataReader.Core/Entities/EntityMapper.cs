using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataReader.Core.Entities
{
    public static class EntityMapper
    {
        public static Product MapProduct(DataRow x)
        {
            return new Product
            {
                ProductID = Convert.ToInt32(x["ProductID"]),
                ProductName = x["ProductName"].ToString(),
                UnitPrice = Convert.ToDouble(x["UnitPrice"]),
                CategoryID = Convert.ToInt32(x["CategoryID"])
            };
        }

        public static Category MapCategory(DataRow x)
        {
            return new Category
            {
                CategoryID = Convert.ToInt32(x["CategoryID"]),
                CategoryName = x["CategoryName"].ToString(),
                Description = x["Description"].ToString() 
            };
        }

        public static Employee MapEmployee(DataRow x)
        {
            return new Employee
            {
                EmployeeID = Convert.ToInt32(x["EmployeeID"]),
                FirstName = x["FirstName"].ToString(),
                LastName = x["LastName"].ToString(),
                Gender = x["TitleOfCourtesy"].ToString().ToUpper()  == "MR." ? "M" : (x["TitleOfCourtesy"].ToString().ToUpper() == "DR." ? "M" : "F"),
                BirthDate = Convert.ToDateTime(x["BirthDate"]),
                Age = DateTime.Now.Year - Convert.ToDateTime(x["BirthDate"]).Year,
                HireDate = Convert.ToDateTime(x["HireDate"]),
                YearsOfExperience = DateTime.Now.Year - Convert.ToDateTime(x["HireDate"]).Year,
                Title = x["Title"].ToString(),
                TitleOfCourtesy = x["TitleOfCourtesy"].ToString(),
                Address = x["Address"].ToString(),
                City = x["City"].ToString(),
                Country = x["Country"].ToString(),
                PostalCode = x["PostalCode"].ToString(),
                HomePhone = x["HomePhone"].ToString(),
                Extension = Convert.ToInt32(x["Extension"]),
                Region = x["Region"].ToString(),
                ReportsTo = x["ReportsTo"].ToString(),
                Notes = x["Notes"].ToString()
            };
        }
    }
}
