using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Configuration;
using LotusTimesheet.Class;

namespace LotusTimesheet
{
    class Program
    {
        static void Main(string[] args)
        {
            var siteUrl = ConfigurationManager.AppSettings["SP_Address_Live"];
            GetEmployeeFromEasyHR getEmployeeFromEasyHR = new GetEmployeeFromEasyHR();

            List<EmployeeModel> employeeModels = new List<EmployeeModel>();

            EmployeeModel emp = new EmployeeModel();
            employeeModels = getEmployeeFromEasyHR.GetEmployeAttendance();
            Console.WriteLine("Please Wait...");

            Console.WriteLine("Dont Stop program....");
            for (var i = 0; i < employeeModels.Count; i++)
            {
              emp = getEmployeeFromEasyHR.CheckNewEntry(employeeModels[i].empno,employeeModels[i].attendance_date, siteUrl);
                Console.WriteLine("Please Wait...");
                if (emp.count == 0)
                {
                    //Console.WriteLine("Insert " + employeeModels[i].empno +" "+ employeeModels[i].attendance_date);
                    getEmployeeFromEasyHR.InsertEntry(employeeModels[i], siteUrl);

                }
                else
                {
                    //Console.WriteLine("Update " + employeeModels[i].empno + " " + employeeModels[i].attendance_date);
                    getEmployeeFromEasyHR.UpdateEntry(employeeModels[i], siteUrl, emp.ID);

                }

            }





            getEmployeeFromEasyHR.CheckNewEntry("","", siteUrl);
        }
    }
}
