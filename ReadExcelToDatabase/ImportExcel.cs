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
                Worksheet excelWorkSheet = excelWorkbook.Sheets[4];
                Range exRange = excelWorkSheet.UsedRange;
                int rowCount = exRange.Rows.Count;
                int colCount = exRange.Columns.Count;

                for (int i = 2; i <= rowCount; i++)
                {
                    string siteLocation = exRange[i, 2].Value2;
                    string siteManagerName = exRange[i, 3].Value2;
                    decimal siteAllowance = Convert.ToDecimal(exRange[i, 7].Value2);
                    string project = exRange[i, 5].Value2;
                    string description = exRange[i, 6].Value2;


                    if (!string.IsNullOrEmpty(siteLocation))
                    {
                        var userName = siteManagerName.Split(' ')[1];

                        if (!db.Sites.Any(x => x.name.Equals(siteLocation)))
                        { 
                            Site obj = new Site();
                            obj.name = siteLocation;
                            obj.location = siteLocation;
                            obj.Province_FK = 1;
                            obj.siteAllowance = siteAllowance;
                            obj.description = description;
                            obj.Owner = db.AspNetUsers.Where(x => x.status == true && x.UserName.Equals(userName)).Select(x => x.Id).FirstOrDefault();
                            if (!string.IsNullOrEmpty(project))
                                obj.project_fk = db.tb_Project.FirstOrDefault(x => x.sonumber.Equals(project)).id;
                            obj.status = true;
                            db.Sites.Add(obj);
                            db.SaveChanges();
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
