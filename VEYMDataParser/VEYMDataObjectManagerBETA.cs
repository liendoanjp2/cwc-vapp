﻿using System;
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
            PopulateExcel();
        }

        private void SortUsers()
        {
            //Temp variable old output
            List<string[]> outUserValues;

            //Read the user and place it into the correct sheet

            //For use in loop to hold data
            string[] UserAsStringArray;

            //Variables used to shorten names in a more "english" manner
            string stateAbbrev;
            string actuaChapterName;
            string shortnedName;
            try
            {
                foreach (AllUsersDataObjectBETA.Value user in allUserValues)
                {
                    Console.WriteLine(user.mail);
                    UserAsStringArray = new string[7] { user.displayName, user.rank, user.mail, user?.otherMails?.FirstOrDefault()?.ToString() ?? "", user.mobilePhone, user.leauge, user.officeLocation };
                    if (!string.IsNullOrEmpty(user.leauge))
                    {
                        //check if key exists first, if not add one!
                        if (!leaugesAndTheirUsers.ContainsKey(user.leauge))
                        {
                            leaugesAndTheirUsers.Add(user.leauge, new List<string[]>());
                        }

                        //add them!
                        leaugesAndTheirUsers.TryGetValue(user.leauge, out outUserValues);
                        outUserValues.Add(UserAsStringArray);
                    }

                    if (!string.IsNullOrEmpty(user.officeLocation))
                    {
                        //There are chapters without a -
                        if (user.officeLocation.Contains("-") && user.officeLocation.Contains(" "))
                        {
                            //Need to Break down the Doan name. only 10 characters, no dashes. Sheets are Doan + State Abbrev
                            //Trim there's bad data with spaces at the end! Handling no state

                            if(user.officeLocation.Trim().LastIndexOf(" ") + 3 < user.officeLocation.Trim().Length)
                            {
                                stateAbbrev = user.officeLocation.Substring(user.officeLocation.Trim().LastIndexOf(" ") + 1, 2);
                            }
                            else
                            {
                                stateAbbrev = "";

                            }
                            
                            actuaChapterName = user.officeLocation.Substring(0, user.officeLocation.IndexOf("-") - 1);
                            shortnedName = actuaChapterName + " " + stateAbbrev;
                        }
                        else
                        {
                            shortnedName = user.officeLocation;
                        }

                        if (!chaptersAndTheirUsers.ContainsKey(shortnedName))
                        {
                            chaptersAndTheirUsers.Add(shortnedName, new List<string[]>());
                        }

                        chaptersAndTheirUsers.TryGetValue(shortnedName, out outUserValues);
                        outUserValues.Add(UserAsStringArray);
                    }

                    allUserData.Add(UserAsStringArray);
                }
            }
            catch(Exception ex)
            {

            }

        }

        private void CreateExcel()
        {
            //This is the header that goes on top of each Worksheet
            List<string[]> headerRow = new List<string[]>()
                {
                     new string[] { "Name", "Rank/Title", "VEYM Email", "Other Email" , "Phone", "Leauge", "Chapter" }
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

                //Create Worksheets for each chapter and populate the look up dictionary
                foreach (string shortenedChapter in chaptersAndTheirUsers.Keys)
                {
                    //rule out the super funky
                    if(!shortenedChapter.Contains("-"))
                    {
                        if(shortenedChapter != "Canada" || shortenedChapter != "dominico savio AZ")
                        {
                            excel.Workbook.Worksheets.Add(shortenedChapter);
                            allWorksheets.Add(excel.Workbook.Worksheets[shortenedChapter]);
                        }
                    }
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
            }
        }

        private void PopulateExcel()
        {
            List<string[]> populateData;

            foreach (ExcelWorksheet worksheet in allWorksheets)
            {

                //find which data array to populate
                if (leaugesAndTheirUsers.Keys.Contains(worksheet.Name))
                {
                    leaugesAndTheirUsers.TryGetValue(worksheet.Name, out populateData);
                }
                else if (chaptersAndTheirUsers.Keys.Contains(worksheet.Name))
                {
                    chaptersAndTheirUsers.TryGetValue(worksheet.Name, out populateData);
                }
                else
                {
                    populateData = allUserData;
                }

                worksheet.Cells[2, 1].LoadFromArrays(populateData);
                worksheet.Cells.AutoFitColumns();
            }

            //Save it off to the desktop!
            FileInfo excelFile = new FileInfo(@"C:\Users\" + Environment.UserName + @"\Desktop\VEYM_Dump.xlsx");
            excel.SaveAs(excelFile);
        }
    }
}
