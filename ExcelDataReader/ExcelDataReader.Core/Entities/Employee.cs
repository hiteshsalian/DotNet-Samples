﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataReader.Core.Entities
{
    public class Employee
    {
        public int EmployeeID { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string Title { get; set; }
        public string TitleOfCourtesy { get; set; }
        public string Gender { get; set; }
        public DateTime BirthDate { get; set; }
        public int Age { get; set; }
        public DateTime HireDate { get; set; }
        public int YearsOfExperience { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Region { get; set; }
        public string PostalCode { get; set; }
        public string Country { get; set; }
        public string HomePhone { get; set; }
        public int Extension { get; set; }
        public string Notes { get; set; }
        public string ReportsTo { get; set; }
        

    }
}
