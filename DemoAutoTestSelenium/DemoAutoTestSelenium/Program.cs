using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AutoTestSelenium
{
    [TestClass]
    public class Program
    {
        [TestMethod]
        public void Main()
        {
            string user, password, path;

            // クラウドサービスのリストを取得する
            List<string> listCloudService = new List<string>();

            int index = -1;

            string systemPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location).ToString();

        
            if (!(File.Exists(systemPath + "\\CloudServices_Information\\CloudServices.xlsx")))
            {
                //log.WriteLog("ファイルエクセルクラウドサービス情報が存在しません", null);
                Console.WriteLine("ファイルエクセルクラウドサービス情報が存在しません");          
                return;
            }

            // Excelファイル情報クラウドサービスへのパス
            path = systemPath + "\\CloudServices_Information\\CloudServices.xlsx";

            SpreadsheetDocument doc = SpreadsheetDocument.Open(path, false);

            WorkbookPart workbookPart = doc.WorkbookPart;
   
            Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById("rId1")).Worksheet;
      
            SheetData thesheetdata = theWorksheet.GetFirstChild<SheetData>();

            var lastRow = thesheetdata.Descendants<Row>().LastOrDefault();

            var lastRowToInt = Int32.Parse(lastRow.RowIndex);

            Cell theCellUser = (Cell)thesheetdata.ElementAt(1).ChildElements.ElementAt(4);                
            SharedStringItem textUser = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Int32.Parse(theCellUser.InnerText));
            user = textUser.Text.Text;

            Cell theCellPassword = (Cell)thesheetdata.ElementAt(1).ChildElements.ElementAt(5);
            SharedStringItem textPassword = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Int32.Parse(theCellPassword.InnerText));
            password = textPassword.Text.Text;
       
            // クラウドサービスのステータスのリストを取得する 
            for (int i = 1; i < lastRowToInt - 1; i++)
            {
                Cell theCellCheck = (Cell)thesheetdata.ElementAt(i).ChildElements.ElementAt(1);

                Int32.TryParse(theCellCheck.InnerText, out index);
                
                string textCheck = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(index).InnerText;

                if ((String.Compare("〇", textCheck)) == 0)
                {
                    Cell theCellCloudService = (Cell)thesheetdata.ElementAt(i).ChildElements.ElementAt(2);
                    SharedStringItem textCloudService = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(Int32.Parse(theCellCloudService.InnerText));
                    listCloudService.Add(textCloudService.Text.Text);
                }
            }

            var stateStoping = "stop";

            MyTestCaseTest test_case = new MyTestCaseTest();

            test_case.SetUp();

            test_case.StartAzure(user, password);

            for (int i = 0; i < listCloudService.Count; i++)
            {
                if (test_case.TestCases(listCloudService[i], stateStoping) == true)
                {
                    //log.WriteLog(stateStoping + " " + listCloudService[i], "OK");
                    Console.WriteLine(stateStoping + " " + listCloudService[i] + " " + "OK");
                }
                else
                {
                    //log.WriteLog(stateStoping + " " + listCloudService[i], "ERROR");
                    Console.WriteLine(stateStoping + " " + listCloudService[i] + " " + "ERROR");
                }
            }
            test_case.TearDown();
        }
    }
}

