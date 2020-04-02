using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;

namespace VEYMDataParser
{
    class VEYMDataObjectManager
    {
        //This List holds all of the VEYM users, populated through the crawl of the Get All Users Request
        List<AllUsersDataObject.Value> allHTandFriends;

        //This is in according to the VEYM List of Chapters. Incudes No Accent marks to try and get EVERYONE
        readonly string[] doansInJP2 = new string[23]
        {
         "Kitô Vua - (Wheat Ridge, CO)",
         "Anrê Dũng Lạc - (Des Moines, IA)",
         "Anrê Dũng Lạc - (Wyoming, MI)",
         "Anrê Dũng Lạc - (Kansas City, MO)",
         "Anre Dung Lac - (Kansas City, MO)",
         "Anrê Dũng Lạc - (Lincoln, NE)",
         "Anrê Phú Yên - (Dayton, OH)",
         "Anrê Trần Anh Dũng - (Warren, MI)",
         "Chúa Cứu Thế - (Louisville, KY)",
         "Chúa Thăng Thiên - (St. Louis, MO)",
         "Đồng Hành - (Franklin, WI)",
         "Kito Vua - (Wichita, KS)",
         "Kitô Vua - (Wichita, KS)",
         "Maria Nu Vuong - (Lincoln, NE)",
         "Maria Nữ Vương - (Lincoln, NE)",
         "Nguồn Sống - (Glen Ellyn, IL)",
         "Phaolo Buong- (St. Paul, MN)",
         "Phaolô Hạnh - (Des Moines, IA)",
         "Phêrô - (Lansing, MI)",
         "Sao Biển - (Chicago, IL)",
         "Thánh Giuse - (Minneapolis, MN)",
         "Thánh Tâm - (Cincinnati, OH)",
         "Tôma Thiện - (St. Paul, MN)"
        };

        //Dictionary of (doan, HT emails)
        IDictionary<string, List<string>> doansAndTheirHT;

        //Constructor requires all the pages of Get All Users Data for inital population
        public VEYMDataObjectManager(List<AllUsersDataObject.RootObject> pages)
        {
            allHTandFriends = new List<AllUsersDataObject.Value>();
            doansAndTheirHT = new Dictionary<string, List<string>>();

            //initalize the "buckets" to sort out the HT
            foreach (string doan in doansInJP2)
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
                //I am declaring variables as I go
                
                //This is a mapping of shortened Doan Names -> Full Doan Names (Excel cant handle wierd characters + super long names)
                Dictionary<string, string> doanShortDictionary = new Dictionary<string, string>();
                //Variables used to shorten names in a more "english" manner
                string stateAbbrev;
                string actualDoanName;
                string shortnedName;

                //Add the easy ones
                excel.Workbook.Worksheets.Add("All of VEYM");
                excel.Workbook.Worksheets.Add("JPII");

                //Create Worksheets for each Doan and populate the look up dictionary
                foreach (string doan in doansInJP2)
                {
                    //Need to Break down the Doan name. only 10 characters, no dashes. Sheets are Doan + State Abbrev
                    stateAbbrev = doan.Substring(doan.LastIndexOf(" ") + 1, 2);
                    actualDoanName = doan.Substring(0, doan.IndexOf("-") - 1);
                    shortnedName = actualDoanName + " " + stateAbbrev;
                    excel.Workbook.Worksheets.Add(shortnedName);

                    doanShortDictionary.Add(shortnedName, doan);
                }

                //This is the header that goes on top of each Worksheet
                List<string[]> headerRow = new List<string[]>()
                {
                     new string[] { "Name", "First Name", "Rank/Title", "Email", "Phone", "Doan", "ID" }
                };

                // Determine the header range (e.g. A1:E1), this is the "Size"
                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

                // Create Instances of each worksheet so we can assign/manipulate later
                ExcelWorksheet allWorksheet = excel.Workbook.Worksheets["All of VEYM"];
                ExcelWorksheet jp2Worksheet = excel.Workbook.Worksheets["JPII"];

                //Collection of all the worksheet instances
                List<ExcelWorksheet> allWorksheets = new List<ExcelWorksheet>();

                //Add the easy ones first
                allWorksheets.Add(allWorksheet);
                allWorksheets.Add(jp2Worksheet);

                //Create instances of the worksheets for each Doan and add them to allWorksheets
                foreach (string shortenedDoanName in doanShortDictionary.Keys)
                {
                    allWorksheets.Add(excel.Workbook.Worksheets[shortenedDoanName]);
                }                

                //Style Each worksheet
                foreach (ExcelWorksheet myWorksheet in allWorksheets)
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

                //All of Jp2
                List<string[]> jp2DataOut = new List<string[]>();
                //All the Doans Specifically
                Dictionary<string, List<string[]>> doansDataOutDictionary = new Dictionary<string, List<string[]>>();
                //For use of each doan
                List<string[]> doanDataOut = new List<string[]>();

                //Ititalize
                foreach (string doan in doansInJP2)
                {
                    doansDataOutDictionary.Add(doan, new List<string[]>());
                }

                //For use in loop to hold data
                string[] htAsStringArray;

                //Out Value for later
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
                        doansDataOutDictionary.TryGetValue(HT.officeLocation, out doanDataOut);
                        //add the HT specifically to the Doan page
                        doanDataOut.Add(htAsStringArray);

                        //Grab the value paired to the doan key
                        doansAndTheirHT.TryGetValue(HT.officeLocation, out htEmails);

                        //add the HT email specifically to the Email List
                        htEmails.Add(HT.mail);
                    }

                    allDataOut.Add(htAsStringArray);
                }

                //Load up all the data into the sheets
                allWorksheet.Cells[2, 1].LoadFromArrays(allDataOut);
                allWorksheet.Cells.AutoFitColumns();

                jp2Worksheet.Cells[2, 1].LoadFromArrays(jp2DataOut);
                jp2Worksheet.Cells.AutoFitColumns();

                string fullDoanName;
                foreach (ExcelWorksheet myWorksheet in allWorksheets)
                {
                    if (myWorksheet.Name != "All of VEYM" && myWorksheet.Name != "JPII")
                    {
                        //Grab the value paired to the doan key
                        //First find the full name
                        doanShortDictionary.TryGetValue(myWorksheet.Name, out fullDoanName);
                        doansDataOutDictionary.TryGetValue(fullDoanName, out doanDataOut);
                        myWorksheet.Cells[2, 1].LoadFromArrays(doanDataOut);
                        myWorksheet.Cells.AutoFitColumns();
                    }
                }

                //Save it off to the desktop!
                FileInfo excelFile = new FileInfo(@"C:\Users\" + Environment.UserName + @"\Desktop\VEYM_Dump.xlsx");
                excel.SaveAs(excelFile);

                //Folder containing all Doan's emails! Create it!
                string doanEamilsFolder = @"C:\Users\" + Environment.UserName + @"\Desktop\Doan Emails\";
                if (!Directory.Exists(doanEamilsFolder))
                {
                    DirectoryInfo di = Directory.CreateDirectory(doanEamilsFolder);
                }

                //Create Text Files of email chains for each Doan!
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