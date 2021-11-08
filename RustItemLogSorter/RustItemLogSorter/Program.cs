using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace RustItemLogSorter
{
    class Program
    {
        #region vars
        public static List<ResourceData> resourceList = new List<ResourceData>();
        public static List<int> ItemList = new List<int>();
        private static Workbook excelFile;
        private static Application excelApp;
        private static string excelFileName;
        #endregion

        public static void Main(string[] args)
        {

            #region startFunctions
            ExcelFileCreator();
            ResourceDataSet();
            #endregion

            ReadLogs();
            bool finishChecking = true;
            #region CheckingResource
            /*
            while (finishChecking)
            {
                Console.WriteLine("What resource do you want to check out?");
                string info = Console.ReadLine();
                for (int i = 0; i < resourceList.Count; i++)
                {
                    if (resourceList[i].ResourceName.ToUpper() == info.ToUpper())
                    {
                        Console.WriteLine("The total amount of resource {0} gathered is {1}", resourceList[i].ResourceName, resourceList[i].storedAmount);
                    }
                }

                Console.WriteLine("Want to end?");
                string yn = Console.ReadLine();
                if (yn.ToUpper() == "YES" || yn.ToUpper() == "YEA")
                {
                    finishChecking = false;
                }
            }
            */
            SettingExcelFile();
            #endregion



            //
            Console.ReadLine();
        }

        //Loops through all rust resources in the list I have and it adds them to a resource data list.
        static void ResourceDataSet()
        {
            String txt;
            String[] resources;
            //grabs the file 
            StreamReader sr = new StreamReader("D:\\Programs\\CodingStuff\\RustItemLogSorter\\RustItemStorage\\ItemList.txt");
            //reads the entire txt file
            txt = sr.ReadToEnd();
            sr.Close();
            //splits each word and puts it into an array based on the , (comma)
            resources = txt.Split(',');
            int length = resources.Length / 2;
            //Console.WriteLine(resources.Length);
            //Console.WriteLine(length);
            String[] resourceName = new String[length];
            String[] resourceType = new String[length];
            string nameholder;
            string typeholder;
            for (int i = 0; i < resourceName.Length; i++)
            {
                nameholder = resources[i].Split('-')[0];
                typeholder = resources[i].Split('-')[1];
                resourceType[i] = typeholder;
                resourceName[i] = nameholder;

            }
            Console.WriteLine(resourceName.Length);
            //this for loop runs through the array and creats a new item in the resource data and makes sets the name value of it to the array postions name, thus putting all the txt from the file into a nicely sorted list.
            //sets a new array that has its length set to the length of the resource names array.
            ResourceData[] data = new ResourceData[resourceName.Length];
            for (int i = 0; i < resourceName.Length; i++)
            {
                data[i].ResourceName = resourceName[i];
                data[i].ResourceType = resourceType[i];
                data[i].storedAmount = 0;
            }
            //one its done setting the names it sets the resource to the array thus keeping it in a list, then it prints out the resources
            resourceList = data.ToList<ResourceData>();
            //Console.WriteLine("There are {0} resources in the list, keep in mind they are not all resources, just using this for most names.", resourceList.Count);
        }

        //this void runs a file an prints all the numbers from the resource values.
        static void ReadLogs()
        {
            foreach (var file in Directory.EnumerateFiles(@"D:\Programs\CodingStuff\RustItemLogSorter\LogStorage\", "*.txt"))
            {
                //this puts the resourcelist into an array.
                ResourceData[] data = resourceList.ToArray();
                String line;
                int number;
                //Pass the file path and file name to the StreamReader constructor
                StreamReader sr = new StreamReader(file);
                //Read the first line of text and searches through it for the x and then all before it untill the end of the number.
                line = sr.ReadLine();
                //Continue to read until you reach end of file while applying the same regex search script, this is for depositite
                while (line != null)
                {
                    //how the regex pattern is setup: First it gets the ending of the deposited or looted to get the index on the line, then it groups the space, next group is the numbers upto 10, next group is the x, next group is a space, next group it gets the resource name.
                    string pattern = @"(ed)(\s)(\d{1,10})(x)(\s)(\D{1,50}$)";
                    RegexOptions options = RegexOptions.Multiline;
                    foreach (Match m in Regex.Matches(line, pattern, options))
                    {
                        //grabs wether its deposited or not.
                        char action;
                        action = line.ElementAt(m.Groups[1].Index - 3);
                        //if it S then the player put resources into a container, if its o then the player took resources from a container.

                        //adds items to the list
                        if (action == 's')
                        {
                            //this loops through all the names and links it to a item in the resource list.
                            for (int i = 0; i < resourceList.Count; i++)
                            {
                                //this checks to see if it gets a match for the resource in the line on the log.
                                if (data[i].ResourceName.ToLower() == m.Groups[6].Value.ToLower())
                                {
                                    bool valSuccess = Int32.TryParse(m.Groups[3].Value, out number);
                                    if (valSuccess)
                                    {
                                        data[i].storedAmount += number;
                                    }
                                    //Console.WriteLine("Got resource {0} ", resourceList[i].ResourceName);
                                }
                            }
                            //Console.WriteLine("desposited!");
                        }
                        //removes items from the list.
                        else if (action == 'o')
                        {
                            //Console.WriteLine("looted!!");
                            for (int i = 0; i < resourceList.Count; i++)
                            {
                                //this checks to see if it gets a match for the resource in the line on the log.
                                if (data[i].ResourceName == m.Groups[6].Value)
                                {
                                    bool valSuccess = Int32.TryParse(m.Groups[3].Value, out number);
                                    if (valSuccess)
                                    {
                                        data[i].storedAmount -= number;
                                    }
                                    //Console.WriteLine("Got resource {0} ", resourceList[i].ResourceName);
                                }
                            }
                        }
                        //adding the resource to the max.

                    }

                    //Read the next line
                    line = sr.ReadLine();
                }
                //close the file
                sr.Close();
                //sets the resource list to the data array of resourceDatas created
                resourceList = data.ToList();
            }

        }

        static void ExcelFileCreator()
        {
            Console.WriteLine("Enter a WipeNumber ");
            excelFileName = Console.ReadLine();
            if (File.Exists(@"D:\Programs\CodingStuff\RustItemLogSorter\ExcelLogOutput\" + excelFileName + ".xlsx") == false)
            {
                //creating the excel file and deals with saving it. and makes sure to close and quit so we dont have problems!
                excelApp = new Application();
                if (excelApp == null)
                {
                    Console.WriteLine("Excel is not properly installed!!");
                    return;
                }
                excelFile = excelApp.Workbooks.Add();
                Worksheet excelWorksheet = (Worksheet)excelFile.Worksheets.get_Item(1);
                excelWorksheet.Cells[1, 1] = "Resource Type";
                excelWorksheet.Cells[1, 2] = "Resource Name";
                excelWorksheet.Cells[1, 2] = "Stored Amount";
                excelFile.SaveAs(@"D:\Programs\CodingStuff\RustItemLogSorter\ExcelLogOutput\" + excelFileName + ".xlsx");
            }
            else
            {
                File.Delete(@"D:\Programs\CodingStuff\RustItemLogSorter\ExcelLogOutput\" + excelFileName + ".xlsx");
                //creating the excel file after deletin the old and deals with saving it. and makes sure to close and quit so we dont have problems!
                excelApp = new Application();
                if (excelApp == null)
                {
                    Console.WriteLine("Excel is not properly installed!!");
                    return;
                }
                excelFile = excelApp.Workbooks.Add();
                Worksheet excelWorksheet = (Worksheet)excelFile.Worksheets.get_Item(1);
                excelWorksheet.Cells[1, 1] = "Resource Type";
                excelWorksheet.Cells[1, 2] = "Resource Name";
                excelWorksheet.Cells[1, 3] = "Stored Amount";
                excelFile.SaveAs(@"D:\Programs\CodingStuff\RustItemLogSorter\ExcelLogOutput\" + excelFileName + ".xlsx");
            }

        }

        static void SettingExcelFile()
        {
            //editing the premade excel file.
            if (excelApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }
            excelFile = excelApp.Workbooks.Open(@"D:\Programs\CodingStuff\RustItemLogSorter\ExcelLogOutput\" + excelFileName + ".xlsx");
            Worksheet excelWorksheet = (Worksheet)excelFile.Worksheets.get_Item(1);
            List<ResourceData> tempResourceList = resourceList;
            for (int z = 0; z < tempResourceList.Count; z++)
            {
                if (tempResourceList[z].storedAmount == 0)
                {
                    tempResourceList.RemoveAt(z);
                }
            }
            for (int i = 0; i < tempResourceList.Count; i++)
            {
                if (tempResourceList[i].storedAmount < 0)
                {
                    excelWorksheet.Cells[(i + 1), 1] = tempResourceList[i].ResourceType;
                    excelWorksheet.Cells[(i + 1), 2] = tempResourceList[i].ResourceName;
                    excelWorksheet.Cells[(i + 1), 3] = tempResourceList[i].storedAmount * -1;
                }
                else if (tempResourceList[i].storedAmount != 0)
                {
                    excelWorksheet.Cells[(i + 1), 1] = tempResourceList[i].ResourceType;
                    excelWorksheet.Cells[(i + 1), 2] = tempResourceList[i].ResourceName;
                    excelWorksheet.Cells[(i + 1), 3] = tempResourceList[i].storedAmount;
                }

            }
            Console.WriteLine("Compiling all the logs!");

            excelFile.SaveAs(@"D:\Programs\CodingStuff\RustItemLogSorter\ExcelLogOutput\" + excelFileName + ".xlsx");
            excelFile.Close(@"D:\Programs\CodingStuff\RustItemLogSorter\ExcelLogOutput\" + excelFileName + ".xlsx");
            excelApp.Quit();

            Console.WriteLine("This is done, you have now compiled all the logs! \nPress enter to end.");
        }
    }
    public struct ResourceData
    {
        public string ResourceType;
        public string ResourceName;
        public int storedAmount;
    }
}