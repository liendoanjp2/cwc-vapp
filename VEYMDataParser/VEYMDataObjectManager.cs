using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace VEYMDataParser
{
    class VEYMDataObjectManager
    {
        List<AllUsersDataObject.Value> allHTandFriends;
        readonly string[] doansInJP2 = new string[19] { "Kitô Vua - (Wheat Ridge, CO)", "Anrê Dũng Lạc - (Charlotte, NC)", "Anre Dung Lac - (Kansas City, MO)", "Anrê Dũng Lạc - (Lincoln, NE)", "Anrê Phú Yên - (Dayton, OH)", "Anrê Trần Anh Dũng - (Warren, MI)", "Chúa Cứu Thế - (Louisville, KY)", "Chúa Thăng Thiên - (St. Louis, MO)", "Đồng Hành - (Franklin, WI)", "Kito Vua - (Wichita, KS)", "Maria Nu Vuong - (Lincoln, NE)", "Nguồn Sống - (Glen Ellyn, IL)", "Phaolo Buong- (St. Paul, MN)", "Phaolô Hạnh - (Des Moines, IA)", "Phêrô - (Lansing, MI)", "Sao Biển - (Chicago, IL)", "Thánh Giuse - (Minneapolis, MN)", "Thánh Tâm - (Cincinnati, OH)", "Tôma Thiện - (Buffalo, NY)" };
        //Dictionary of (doan, HT email)
        IDictionary<string, List<string>> doansAndTheirHT;

        public VEYMDataObjectManager(List<AllUsersDataObject.RootObject> pages)
        {
            allHTandFriends = new List<AllUsersDataObject.Value>();
            doansAndTheirHT = new Dictionary<string, List<string>>();

            //initalize the "buckets" to sort out the HT
            foreach(string doan in doansInJP2)
            {
                doansAndTheirHT.Add(doan, new List<string>());
            }

            //Break it down
            foreach (AllUsersDataObject.RootObject page in pages)
            {
                foreach (AllUsersDataObject.Value HT in page.value)
                {
                    allHTandFriends.Add(HT);
                }
            }

            CreateExcel();
        }

        private void CreateExcel()
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("All of VEYM");
                excel.Workbook.Worksheets.Add("JPII");

                List<string[]> headerRow = new List<string[]>()
                {
                     new string[] { "Name", "First Name", "Rank/Title", "Email", "Phone", "Doan", "ID" }
                };

                // Determine the header range (e.g. A1:E1)
                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

                // Target a worksheet
                ExcelWorksheet allWorksheet = excel.Workbook.Worksheets["All of VEYM"];
                ExcelWorksheet jp2Worksheet = excel.Workbook.Worksheets["JPII"];

                List<ExcelWorksheet> allWorksheets = new List<ExcelWorksheet>();

                allWorksheets.Add(allWorksheet);
                allWorksheets.Add(jp2Worksheet);

                foreach(ExcelWorksheet myWorksheet in allWorksheets)
                {
                    //Style the worksheet
                    myWorksheet.Cells[headerRange].Style.Font.Bold = true;
                    myWorksheet.Cells[headerRange].Style.Font.Size = 24;
                    myWorksheet.Cells[headerRange].Style.Font.Color.SetColor(System.Drawing.Color.Blue);

                    // Populate header row data
                    myWorksheet.Cells[headerRange].LoadFromArrays(headerRow);
                }
                
                //master List of Arrays
                List<string[]> allDataOut = new List<string[]>();

                //Each Doan
                List<string[]> jp2DataOut = new List<string[]>();

                string[] htAsStringArray;
                List<string> htEmails;

                //Add in everything else
                foreach (AllUsersDataObject.Value HT in allHTandFriends)
                {
                    htAsStringArray = new string[7] { HT.displayName, HT.givenName, HT.jobTitle, HT.mail, HT.mobilePhone, HT.officeLocation, HT.id };

                    if (doansInJP2.Contains(HT.officeLocation))
                    {
                        //Add the HT to the jp2 colleciton
                        jp2DataOut.Add(htAsStringArray);

                        //Grab the value paired to the doan key
                        doansAndTheirHT.TryGetValue(HT.officeLocation, out htEmails);

                        //add the HT email specifically to the Email List
                        htEmails.Add(HT.mail);
                    }

                    allDataOut.Add(htAsStringArray);
                }

                allWorksheet.Cells[2, 1].LoadFromArrays(allDataOut);
                allWorksheet.Cells.AutoFitColumns();

                jp2Worksheet.Cells[2, 1].LoadFromArrays(jp2DataOut);
                jp2Worksheet.Cells.AutoFitColumns();

                FileInfo excelFile = new FileInfo(@"C:\Users\"+ Environment.UserName + @"\Desktop\VEYM_Dump.xlsx");
                excel.SaveAs(excelFile);

                //Folder containing all Doan's emials
                string doanEamilsFolder = @"C:\Users\" + Environment.UserName + @"\Desktop\Doan Emails\";
                if (!Directory.Exists(doanEamilsFolder))
                {
                    DirectoryInfo di = Directory.CreateDirectory(doanEamilsFolder);
                }

                StringBuilder lineOfAllHTEmails = new StringBuilder();
                foreach (string doan in doansAndTheirHT.Keys)
                {
                    using (System.IO.StreamWriter outfile = new System.IO.StreamWriter(doanEamilsFolder + @"\" + doan + ".txt"))
                    {
                        outfile.WriteLine(doan);

                        doansAndTheirHT.TryGetValue(doan, out htEmails);
                        
                        foreach (string htEmail in htEmails)
                        {
                            lineOfAllHTEmails.Append(htEmail);
                            lineOfAllHTEmails.Append(';');
                        }

                        outfile.WriteLine(lineOfAllHTEmails.ToString());
                        lineOfAllHTEmails.Clear();
                    }
                }

                MessageBox.Show("Data Scrape Finished");
            }

        }
    }
}
