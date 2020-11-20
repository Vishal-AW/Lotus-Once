using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using LotusTimesheet.Class;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Security.Cryptography;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using MSC = Microsoft.SharePoint.Client;
using System.Configuration;
//using UserInformation;
using System.Text;
using System.Xml;
namespace LotusTimesheet.Class
{
    public class GetEmployeeFromEasyHR
    {
        //static UserOperation _UserOperation = new UserOperation();
        public static StreamWriter logFile;
        static byte[] bytes = ASCIIEncoding.ASCII.GetBytes("ZeroCool");
        public static string Decrypt(string cryptedString)
        {
            if (String.IsNullOrEmpty(cryptedString))
            {
                throw new ArgumentNullException("The string which needs to be decrypted can not be null.");
            }

            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            MemoryStream memoryStream = new MemoryStream(Convert.FromBase64String(cryptedString));
            CryptoStream cryptoStream = new CryptoStream(memoryStream, cryptoProvider.CreateDecryptor(bytes, bytes), CryptoStreamMode.Read);
            StreamReader reader = new StreamReader(cryptoStream);

            return reader.ReadToEnd();
        }
        public static MSC.ClientContext GetContext(string siteUrl)
        {
            try
            {
                AppConfiguration _AppConfiguration = GetSharepointCredentials(siteUrl);
                string pass = ConfigurationManager.AppSettings["SP_Password_Live"];  //"Lotus@123";
                var securePassword = new SecureString();
                foreach (char c in pass)
                {
                    securePassword.AppendChar(c);
                }
                var onlineCredentials = new MSC.SharePointOnlineCredentials(_AppConfiguration.ServiceUserName, securePassword);
                var context = new MSC.ClientContext(_AppConfiguration.ServiceSiteUrl);
                context.Credentials = onlineCredentials;
                return context;
            }
            catch (Exception ex)
            {
                WriteLog("Error in  CustomSharePointUtility.GetContext: " + ex.ToString());
                return null;
            }
        }
        public static void WriteLog(string logmsg)
        {
            // StreamWriter logFile;

            try
            {

                string LogString = DateTime.Now.ToString("dd/MM/yyyy HH:MM") + " " + logmsg.ToString();

                //  logFile.WriteLine(DateTime.Now);
                //  logFile.WriteLine(logmsg.ToString());
                //logFile.WriteLine(LogString);

                //logFile.Close();
            }
            catch (Exception ex)
            {
                WriteLog(ex.ToString());

            }

        }

        public static AppConfiguration GetSharepointCredentials(string siteUrl)
        {
            AppConfiguration _AppConfiguration = new AppConfiguration();

            _AppConfiguration.ServiceSiteUrl = siteUrl;// _UserOperation.ReadValue("SP_Address");
            _AppConfiguration.ServiceUserName = ConfigurationManager.AppSettings["SP_USER_ID_Live"];
            // _UserOperation.ReadValue("SP_USER_ID_Live");
            // _AppConfiguration.ServicePassword = Decrypt(_UserOperation.ReadValue("SP_Password_Live"));

            return _AppConfiguration;
        }
        public List<EmployeeModel> GetEmployeAttendance()
        {
            List<EmployeeModel> employeeModels = new List<EmployeeModel>();
            //var dataDateValue = DateTime.Today;
            int datecount = Convert.ToInt32(ConfigurationManager.AppSettings["DateCount"]);
            var dataDateValue = DateTime.Today.AddDays(-datecount);

            var endDateValue = DateTime.Today;
            //+ datecount;

            // var Url ="https://test.easyhrworld.com/api/v2/attendance/getEmployeeAttendanceByRange?start_date=2020-09-16&end_date=2020-09-16";
            // var Url = "https://lotustechnicals.easyhrworld.com/api/v2/attendance/getEmployeeAttendanceByRange?start_date=2020-10-13";
            var Url = "https://lotustechnicals.easyhrworld.com/api/v2/attendance/getEmployeeAttendanceByRange?start_date=" + dataDateValue.ToString("yyyy-MM-dd") + "&end_date=" + endDateValue.ToString("yyyy-MM-dd");

            HttpWebRequest request = HttpWebRequest.CreateHttp(Url);
            request.Accept = "application/json;odata=verbose";
            request.Headers.Add("X-API-KEY", "356a192b7913b04c54574d18c28d46e6395428ab");
            Stream webStream = request.GetResponse().GetResponseStream();
            StreamReader responseReader = new StreamReader(webStream);
            string response = responseReader.ReadToEnd();
            JObject jobj = JObject.Parse(response);
            JArray jarr = (JArray)jobj["data"];

            foreach (JObject j in jarr)
            {
                //string ass = j["ID"].ToString();
                //string Title = j["Title"].ToString();
                employeeModels.Add(new EmployeeModel
                {
                    empno = j["empno"].ToString(),
                    name = j["name"].ToString(),
                    office_email = j["office_email"].ToString(),
                    attendance_date = j["attendance_date"].ToString(),
                    checkin_time = j["checkin_time"].ToString(),
                    checkout_time = j["checkout_time"].ToString(),
                    value = j["value"].ToString(),
                    hours = j["hours"].ToString(),
                    comment = j["comment"].ToString(),
                });
            }

            return employeeModels;
        }

        public EmployeeModel CheckNewEntry(string EMPCode,string attendance_date, string siteUrl)
        {
            int returnVal = 0;
            EmployeeModel employeeModels = new EmployeeModel();
            using (MSC.ClientContext context = GetContext(siteUrl))
            {
                //var dataDateValue = DateTime.Today;

                string dateValue = "";
                try
                {
                    var dataDateValue = Convert.ToDateTime(attendance_date);
                    dateValue = dataDateValue.ToString("dd-MM-yyyy");
                }
                catch(Exception ex)
                {
                    dateValue = "";
                }

                MSC.List list = context.Web.Lists.GetByTitle("TIM_DailyAttendance");
                MSC.ListItemCollectionPosition itemPosition = null;
                var q = new CamlQuery() { ViewXml = "<View><Query><Where><And><Eq><FieldRef Name='EmpNo' /><Value Type='Text'>"+ EMPCode + "</Value></Eq><Eq><FieldRef Name='AttendanceDate' /><Value Type='Text'>" + dateValue+"</Value></Eq></And></Where></Query></View>" };
                MSC.ListItemCollection Items = list.GetItems(q);

                context.Load(Items);
                context.ExecuteQuery();
                itemPosition = Items.ListItemCollectionPosition;

                returnVal = Items.Count;
                if (returnVal > 0)
                {
                    employeeModels.ID = Items[0]["ID"].ToString();
                    employeeModels.count = Items.Count;
                }
                else
                {
                    employeeModels.count = Items.Count;
                }
                

            }


            return employeeModels;
        }

        public int InsertEntry(EmployeeModel emp, string siteUrl)
        {
            try{
                using (MSC.ClientContext context = GetContext(siteUrl))
                {
                    MSC.List list = context.Web.Lists.GetByTitle("TIM_DailyAttendance");

                    MSC.ListItem listItem = null;

                    MSC.ListItemCreationInformation itemCreateInfo = new MSC.ListItemCreationInformation();
                    listItem = list.AddItem(itemCreateInfo);

                    listItem["AttendanceDate"] = Convert.ToDateTime(emp.attendance_date).ToString("dd-MM-yyyy");
                    listItem["CheckinTime"] = emp.checkin_time;
                    listItem["CheckoutTime"] = emp.checkout_time;
                    listItem["Comment"] = emp.comment;
                    listItem["EmpNo"] = emp.empno;
                    listItem["Hours"] = emp.hours;
                    listItem["EmpName"] = emp.name;
                    listItem["EmpMail"] = emp.office_email;
                    listItem.Update();
                    context.ExecuteQuery();

                }
            }
            catch (Exception ex)
            {


            }

            return 0;
        }


        public int UpdateEntry(EmployeeModel emp, string siteUrl,string ID)
        {
            try
            {
                using (MSC.ClientContext context = GetContext(siteUrl))
                {
                    MSC.List list = context.Web.Lists.GetByTitle("TIM_DailyAttendance");

                    MSC.ListItem listItem = null;

                    MSC.ListItemCreationInformation itemCreateInfo = new MSC.ListItemCreationInformation();
                    listItem = list.GetItemById(Convert.ToInt32(ID));

                    listItem["AttendanceDate"] = Convert.ToDateTime(emp.attendance_date).ToString("dd-MM-yyyy");
                    listItem["CheckinTime"] = emp.checkin_time;
                    listItem["CheckoutTime"] = emp.checkout_time;
                    listItem["Comment"] = emp.comment;
                    listItem["EmpNo"] = emp.empno;
                    listItem["Hours"] = emp.hours;
                    listItem["EmpName"] = emp.name;
                    listItem["EmpMail"] = emp.office_email;
                    listItem.Update();
                    context.ExecuteQuery();

                }
            }
            catch (Exception ex)
            {


            }

            return 0;
        }


        public class AppConfiguration
        {
            public string ServiceSiteUrl;
            public string ServiceUserName;
            public string ServicePassword;
        }


    }
}
