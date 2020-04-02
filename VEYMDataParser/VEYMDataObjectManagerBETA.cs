using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OfficeOpenXml;

namespace VEYMDataParser
{
    class VEYMDataObjectManagerBETA
    {
        //This List holds all of the VEYM users, populated through the crawl of the Get All Users Request
        List<AllUsersDataObjectBETA.Value> allUserValues;
        //This List holds all of the VEYM users, populated through the crawl of the Get All Users Request
        List<string[]> allUserData;
        //Dictionary of Key = Leauge Value, = User Value List
        Dictionary<string, List<string[]>> leaugesAndTheirUsers;
        //Dictionary of Key = Chapter Value, = User Value List
        Dictionary<string, List<string[]>> chaptersAndTheirUsers;
        //This is a mapping of shortened Doan Names -> Full Doan Names (Excel cant handle wierd characters + super long names)
        Dictionary<string, string> doanShortDictionary;
        //Collection of all the worksheet instances
        List<ExcelWorksheet> allWorksheets = new List<ExcelWorksheet>();
        //The excel sheet
        ExcelPackage excel;

        //Constructor requires all the pages of Get All Users Data for inital population
        public VEYMDataObjectManagerBETA(List<AllUsersDataObjectBETA.RootObject> pages)
        {
            allUserData = new List<string[]>();
            allUserValues = new List<AllUsersDataObjectBETA.Value>();
            leaugesAndTheirUsers = new Dictionary<string, List<string[]>>();
            chaptersAndTheirUsers = new Dictionary<string, List<string[]>>();
            doanShortDictionary = new Dictionary<string, string>();

            //Break it down, only get the values
            foreach (AllUsersDataObjectBETA.RootObject page in pages)
            {
                foreach (AllUsersDataObjectBETA.Value HT in page.value)
                {
                    allUserValues.Add(HT);
                }
            }

            SortUsers();
            CreateExcel();
        }

        private void SortUsers()
        {
            //Temp variable old output
            List<string[]> outUserValues;
            //Read the user and place it into the correct sheet
            //For use in loop to hold data
            string[] userAsStringArray;
            //Variables used to shorten names in a more "english" manner
            string stateAbbrev;
            string actualChapterName;
            string shortnedName;
            string patternNorm = @"^[ a-z0-9A-Z_ÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚĂĐĨŨƠàáâãèéêìíòóôõùúăđĩũơƯĂẠẢẤẦẨẪẬẮẰẲẴẶẸẺẼỀỀỂưăạảấầẩẫậắằẳẵặẹẻẽềềểỄỆỈỊỌỎỐỒỔỖỘỚỜỞỠỢỤỦỨỪễệỉịọỏốồổỗộớờởỡợụủứừỬỮỰỲỴÝỶỸửữựỳỵỷỹ]+[ ][-][ ][(][a-zA-Z]+[,][ ][a-zA-Z]+[)]$";
            string patternID = @"LD[0-9]{2}-[A-Z]{2}[0-9]+";
            string leaugeChapterID;
            Match match;
            int actualLengthOfChapter;
            string ldWorkName;

            foreach (AllUsersDataObjectBETA.Value user in allUserValues)
            {
                //Do not add wonky users
                if(!user.displayName.Contains('@'))
                {
                    leaugeChapterID = "";

                    if(!string.IsNullOrEmpty(user.chapter))
                    {
                        match = Regex.Match(user.chapter, patternID, RegexOptions.IgnoreCase);
                        if (match.Success)
                        {
                            leaugeChapterID = match.Value;
                        }
                    }

                    userAsStringArray = new string[9] { user.displayName, user.rank, user.memberID, user.mail, user?.otherMails?.FirstOrDefault()?.ToString() ?? "", user.mobilePhone, user.leauge, user.officeLocation, leaugeChapterID };
                    if (!string.IsNullOrEmpty(user.leauge))
                    {
                        ldWorkName = "LD " + user.leauge;

                        //check if key exists first, if not add one!
                        if (!leaugesAndTheirUsers.ContainsKey(ldWorkName))
                        {
                            leaugesAndTheirUsers.Add(ldWorkName, new List<string[]>());
                        }

                        //add them!
                        leaugesAndTheirUsers.TryGetValue(ldWorkName, out outUserValues);
                        outUserValues.Add(userAsStringArray);
                    }

                    if (!string.IsNullOrEmpty(user.officeLocation))
                    {
                        match = Regex.Match(user.officeLocation, patternNorm, RegexOptions.IgnoreCase);
                        //Use Regex to find if the names are "Valid"

                        ////There are chapters without a -

                        //Need to Break down the Doan name. only 10 characters, no dashes. Sheets are Doan + State Abbrev
                        //Trim there's bad data with spaces at the end! Handling no state

                        if (match.Success)
                        {
                            stateAbbrev = user.officeLocation.Substring(user.officeLocation.Trim().LastIndexOf(" ") + 1, 2);

                            actualLengthOfChapter = user.officeLocation.IndexOf("-") -1;

                            //Blanket check for super long chapter names
                            if(actualLengthOfChapter > 29)
                            {
                                //Give it some room to append!
                                actualLengthOfChapter = actualLengthOfChapter - (actualLengthOfChapter - 29);
                            }

                            actualChapterName = user.officeLocation.Substring(0, actualLengthOfChapter - 1);
                            shortnedName = actualChapterName + " " + stateAbbrev;

                            if (!chaptersAndTheirUsers.ContainsKey(shortnedName))
                            {
                                chaptersAndTheirUsers.Add(shortnedName, new List<string[]>());
                            }

                            chaptersAndTheirUsers.TryGetValue(shortnedName, out outUserValues);
                            outUserValues.Add(userAsStringArray);
                        }
                    }

                    allUserData.Add(userAsStringArray);
                }
            }
        }

        private void CreateExcel()
        {
            //This is the header that goes on top of each Worksheet
            List<string[]> headerRow = new List<string[]>()
                {
                     new string[] { "Name", "Rank/Title","Member ID", "VEYM Email", "Other Email" , "Phone", "Leauge", "Chapter", "Leauge-Chapter ID" }
                };

            // Determine the header range (e.g. A1:E1), this is the "Size"
            string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
            excel = new ExcelPackage();

            using (excel)
            {
                //Collection of all the worksheet instances
                allWorksheets = new List<ExcelWorksheet>();

                excel.Workbook.Worksheets.Add("All of VEYM");
                // Create Instances of each worksheet so we can assign/manipulate later
                ExcelWorksheet allWorksheet = excel.Workbook.Worksheets["All of VEYM"];

                //Add the easy ones first
                allWorksheets.Add(allWorksheet);

                //Create Worksheets for each leauge and populate the look up dictionary
                foreach (string leauge in leaugesAndTheirUsers.Keys)
                {
                    excel.Workbook.Worksheets.Add(leauge);
                    allWorksheets.Add(excel.Workbook.Worksheets[leauge]);
                }

                List<string> chaptersToUpper = new List<string>();
                string upperName;

                List<string> sortedChaptersAndTheirUsers = chaptersAndTheirUsers.Keys.ToList();
                sortedChaptersAndTheirUsers.Sort();

                //Create Worksheets for each chapter and populate the look up dictionary
                foreach (string shortenedChapter in sortedChaptersAndTheirUsers)
                {
                    //dominico savio AZ has ugly name of one upper and one lower, this checks and accounts for that... urgg
                    upperName = shortenedChapter.ToUpper();

                    if (!chaptersToUpper.Contains(upperName))
                    {
                        excel.Workbook.Worksheets.Add(shortenedChapter);
                        allWorksheets.Add(excel.Workbook.Worksheets[shortenedChapter]);
                        chaptersToUpper.Add(upperName);
                    }
                }

                //Style + fill Each worksheet
                List<string[]> populateData;

                foreach (ExcelWorksheet myWorksheet in allWorksheets)
                {
                    //There are failed worksheets! Name is too long I think
                    if (myWorksheet != null)
                    {
                        //Style the worksheet
                        myWorksheet.Cells[headerRange].Style.Font.Bold = true;
                        myWorksheet.Cells[headerRange].Style.Font.Size = 24;
                        myWorksheet.Cells[headerRange].Style.Font.Color.SetColor(System.Drawing.Color.Blue);

                        // Populate header row data
                        myWorksheet.Cells[headerRange].LoadFromArrays(headerRow);

                        //find which data array to populate
                        if (leaugesAndTheirUsers.Keys.Contains(myWorksheet.Name))
                        {
                            leaugesAndTheirUsers.TryGetValue(myWorksheet.Name, out populateData);
                        }
                        else if (chaptersAndTheirUsers.Keys.Contains(myWorksheet.Name))
                        {
                            chaptersAndTheirUsers.TryGetValue(myWorksheet.Name, out populateData);
                        }
                        else
                        {
                            populateData = allUserData;
                        }

                        myWorksheet.Cells[2, 1].LoadFromArrays(populateData);
                        myWorksheet.Cells.AutoFitColumns();

                    }
                }

                //Save it off to the desktop!
                FileInfo excelFile = new FileInfo(@"C:\Users\" + Environment.UserName + @"\Desktop\VEYM_Dump.xlsx");
                excel.SaveAs(excelFile);

                MessageBox.Show("DONE!");
            }
        }
    }
}
