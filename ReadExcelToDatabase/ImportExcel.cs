using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReadExcelToDatabase.Entity;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace ReadExcelToDatabase
{
    class ImportExcel
    {
        public string filePath { get; set; }
        public string fileName { get; set; }

        public Process[] oldProcess { get; set; }

        public ImportExcel()
        {

        }
        public Boolean ImportToDB()
        {

            Process[] oldProcess = Process.GetProcessesByName("Excel");
            var db = new AttendanceEntities();
            try
            {
                Application excelApp = new Application();
                Workbook excelWorkbook = excelApp.Workbooks.Open(filePath);
                Worksheet excelWorkSheet = excelWorkbook.Sheets[1];
                Range exRange = excelWorkSheet.UsedRange;
                int rowCount = exRange.Rows.Count;
                int colCount = exRange.Columns.Count;

                for (int i = 2; i <= rowCount; i++)
                {
                    string LGT_ID = exRange[i, 1].Value2;
                    string IDCard = exRange[i, 2].Value2;
                    string firstName = exRange[i, 3].Value2;
                    string lastName = exRange[i, 4].Value2;
                    string khmerName = exRange[i, 5].Value2;
                    string khmerSureName = exRange[i, 6].Value2;
                    string gender = exRange[i, 7].Value2;
                    DateTime dob = ConvertToDateTime(Convert.ToString(exRange[i, 8].Value2));
                    string email = exRange[i, 9].Value2;
                    string phoneNumber = exRange[i, 10].Value2;
                    string emergencyContact = exRange[i, 11].Value2;
                    string maritalStatus = exRange[i, 12].Value2;
                    int child = Convert.ToInt16(exRange[i, 13].Value2);
                    string presentAddress = exRange[i, 14].Value2;
                    string permanentAddress = exRange[i, 15].Value2;
                    string levelOfEducation = exRange[i, 16].Value2;
                    string fieldOfEducation = exRange[i, 17].Value2;
                    string spouseName = exRange[i, 18].Value2;
                    string spouseOccupation = exRange[i, 19].Value2;
                    string recruitmentBase = exRange[i, 20].Value2;
                    string workingStaus = exRange[i, 21].Value2;
                    double basicSalary = exRange[i, 22].Value2;
                    string workType = exRange[i, 23].Value2;
                    DateTime startDate = ConvertToDateTime(Convert.ToString(exRange[i, 24].Value2));
                    string position = exRange[i, 25].Value2;
                    string site = exRange[i, 26].Value2;
                    string bankAccount = exRange[i, 27].Value2;
                    double lgtFund = exRange[i, 28].Value2;


                    if (!string.IsNullOrEmpty(LGT_ID))
                    {
                        if(!db.EmployeeMasters.Any(x => x.EmployeeID.Equals(LGT_ID)))
                        {
                            EmployeeMaster emp = new EmployeeMaster();
                            emp.EmployeeID = LGT_ID;
                            emp.created_at = DateTime.Now;
                            emp.created_by = "System";
                            emp.status = true;
                            emp.FirstName = firstName;
                            emp.LastName = lastName;
                            emp.FirstNameKH = khmerName;
                            emp.LastNameKH = khmerSureName;
                            emp.gender = gender == "M" ? "Male" : "Female";
                            emp.dob = dob;
                            emp.Email = email;
                            emp.PhoneNumber = phoneNumber;
                            emp.EmergencyPhoneNumber = emergencyContact;
                            emp.MarritalStatus = string.IsNullOrEmpty(maritalStatus) ? "Single" : "Married";
                            emp.ChildNumber = child;
                            emp.PresentAddress = presentAddress;
                            emp.PermanetAddress = permanentAddress;
                            emp.FieldofEducation = fieldOfEducation;
                            emp.LevelEducation = levelOfEducation;
                            emp.SpouseGuardian = spouseName;
                            emp.SpouseOccupation = spouseOccupation;
                            emp.RecruitmentBase = 1;
                            emp.Working_Status = "Working";
                            emp.BankAccount = bankAccount;
                            emp.LGTFund = lgtFund.ToString();
                            db.EmployeeMasters.Add(emp);
                          //  db.SaveChanges();

                            Contract con = new Contract();
  
                        }
                    }
                }
                KillExcelProc();
                return true;

            }
            catch(Exception ex)
            {
                return false;
            }


        }

        private void KillExcelProc()
        {
            Process[] excelProcesses = System.Diagnostics.Process.GetProcessesByName("Excel");
            if (oldProcess != null && oldProcess.Count() > 0)
            {
                foreach (Process p in excelProcesses)
                {
                    foreach (Process o in oldProcess)
                    {
                        if (p.Id != o.Id) p.Kill();
                    }
                }
            }
            else
            {
                foreach (Process p in excelProcesses)
                {
                    p.Kill();
                }
            }
        }

        public ImportExcel(string filePath, string fileName)
        {
            this.filePath = filePath;
            this.fileName = fileName;
        }

        private DateTime ConvertToDateTime(string data)
        {
            double d = Convert.ToDouble(data);
            DateTime con = DateTime.FromOADate(d);
            return con;
        }

       
    }
}
