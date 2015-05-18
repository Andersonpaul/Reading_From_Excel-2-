using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support;

namespace Read_From_Excel
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            //instantiating Excel Application
            Excel.Application appExcel = new Excel.Application();

            //instantiating Excel Workbook
            Excel.Workbook excelWorkbook = appExcel.Workbooks.Add("C:\\Users\\Udokoro\\Desktop\\Book1.xlsx");

            //instantiating Excel Sheets
            Excel.Sheets excelWorksheet = excelWorkbook.Sheets;

            //Getting the Active Sheet in currently opened Excel Application
            Excel.Worksheet activeSheet = excelWorksheet.get_Item("Sheet1");


            //code to get username from excel document
            Excel.Range username = (Excel.Range)activeSheet.get_Range("B1", "B1");
            string usernameID = username.Value.ToString();

            //code to get password from excel document
            Excel.Range password = (Excel.Range)activeSheet.get_Range("D1", "D1");
            string passwordId = password.ToString();

            IWebDriver driver = new ChromeDriver();
            driver.Url = "http://yahoomail.com/";
            driver.FindElement(By.Id("login-username")).SendKeys(usernameID);

            driver.FindElement(By.Id("login-passwd")).SendKeys(passwordId);

            driver.FindElement(By.Id("login-signin")).Click();


        }
    }
}
