using Microsoft.VisualStudio.TestTools.UnitTesting;
//using New1.TestPagesObjects;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Text.RegularExpressions;
using excel = Microsoft.Office.Interop.Excel;
using System.Net.Mail;

namespace New1
{
    [TestClass]
    public class UnitTest1
    {
        //Global variables declaration
        protected IWebDriver driver;

        private excel.Workbook x1WorkBook;

        
        string username = "";
        string password = "";

        [TestInitialize]
        public void Initialise()
        {
         
           ////Initialization

            string browserName = "Chrome";

            if (browserName.Equals("Chrome"))
            {
                driver = new ChromeDriver(@"C:\Workspace\Hippo\Frontend_Auto\Exe");
            }
            else if (browserName.Equals("IE"))
            {
                driver = new InternetExplorerDriver(@"C:\Workspace\Hippo\Frontend_Auto\IE Driver");
            }

            driver.Manage().Window.Maximize();

            driver.Navigate().GoToUrl("http://automationpractice.com/");

            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(20));

            excel.Application x1App = new excel.Application();
            x1WorkBook = x1App.Workbooks.Open(@"C:\Workspace\Hippo\Frontend_Auto\TestData\Data.xlsx"); 

            username = "isaiamohube@gmail.com";
            password = "Mohube@123";
        }

        [TestMethod]
        public void validatePageTitle()
        {
           
            excel._Worksheet x1WorkSheet = x1WorkBook.Sheets[2];
            excel.Range x1Range = x1WorkSheet.UsedRange;
            int xlRowCnt;

            
            for (xlRowCnt = 2; xlRowCnt <= x1Range.Rows.Count; xlRowCnt++)
            {
                string expectedPageTitle = (string)(x1Range.Cells[xlRowCnt, 1] as excel.Range).Value2;
                string actualPageTitle = driver.Title;

                Assert.AreEqual<string>(expectedPageTitle, actualPageTitle);

            }
           
        }

        [TestMethod]
        public void registerNewUser()
        {

            excel._Worksheet x1WorkSheet = x1WorkBook.Sheets[3];
            excel.Range x1Range = x1WorkSheet.UsedRange;
            int xlRowCnt;

            for (xlRowCnt = 2; xlRowCnt <= x1Range.Rows.Count; xlRowCnt++)
            {

            driver.FindElement(By.XPath("//*[@id='header']/div[2]/div/div/nav/div[1]")).Click();
            driver.FindElement(By.XPath("//input[@id='email_create']")).SendKeys((string)(x1Range.Cells[xlRowCnt, 1] as excel.Range).Value2);
            driver.FindElement(By.XPath("//button[@name='SubmitCreate']")).Click();
            string title = (string)(x1Range.Cells[xlRowCnt, 2] as excel.Range).Value2;

            System.Threading.Thread.Sleep(10000);

            if (title == "Mr")
            {
                driver.FindElement(By.XPath("//*[@id='id_gender1']")).Click();
            }
            else
            {
                driver.FindElement(By.XPath("//*[@id='id_gender2']")).Click();
            }
            driver.FindElement(By.XPath("//input[@id='customer_firstname']")).SendKeys((string)(x1Range.Cells[xlRowCnt, 3] as excel.Range).Value2);
            driver.FindElement(By.XPath("//input[@id='customer_lastname']")).SendKeys((string)(x1Range.Cells[xlRowCnt, 4] as excel.Range).Value2);

            driver.FindElement(By.XPath("//*[@id='passwd']")).SendKeys((string)(x1Range.Cells[xlRowCnt, 5] as excel.Range).Value2);

            SetMethods.Select(driver, "//select[@id='days']", (string)(x1Range.Cells[xlRowCnt, 6] as excel.Range).Value2, "XPath");
            SetMethods.Select(driver, "//select[@id='months']", (string)(x1Range.Cells[xlRowCnt, 7] as excel.Range).Value2, "XPath");
            SetMethods.Select(driver, "//select[@id='years']", (string)(x1Range.Cells[xlRowCnt, 8] as excel.Range).Value2, "XPath");

            driver.FindElement(By.XPath("//*[@id='newsletter']")).Click();
            driver.FindElement(By.XPath(" //input[@id='optin']")).Click();

            driver.FindElement(By.XPath("//*[@id='firstname']")).SendKeys((string)(x1Range.Cells[xlRowCnt, 9] as excel.Range).Value2);
            driver.FindElement(By.XPath("//*[@id='lastname']")).SendKeys((string)(x1Range.Cells[xlRowCnt, 10] as excel.Range).Value2);
            driver.FindElement(By.XPath("//*[@id='company']")).SendKeys((string)(x1Range.Cells[xlRowCnt, 10] as excel.Range).Value2);
            driver.FindElement(By.XPath("//*[@id='address1']")).SendKeys((string)(x1Range.Cells[xlRowCnt, 12] as excel.Range).Value2);
            driver.FindElement(By.XPath("//*[@id='address2']")).SendKeys((string)(x1Range.Cells[xlRowCnt, 13] as excel.Range).Value2);
            driver.FindElement(By.XPath("//*[@id='city']")).SendKeys((string)(x1Range.Cells[xlRowCnt, 14] as excel.Range).Value2);
            SetMethods.Select(driver, "//*[@id='id_country']", (string)(x1Range.Cells[xlRowCnt, 15] as excel.Range).Value2, "XPath");

            SetMethods.Select(driver, "//*[@id='id_state']", (string)(x1Range.Cells[xlRowCnt, 16] as excel.Range).Value2, "XPath");
            driver.FindElement(By.XPath("//*[@id='postcode']")).SendKeys((string)(x1Range.Cells[xlRowCnt, 17] as excel.Range).Value2);
            driver.FindElement(By.XPath("//*[@id='other']")).SendKeys((string)(x1Range.Cells[xlRowCnt, 18] as excel.Range).Value2);
            driver.FindElement(By.XPath("//*[@id='phone']")).SendKeys((string)(x1Range.Cells[xlRowCnt, 19] as excel.Range).Value2);
            driver.FindElement(By.XPath("//*[@id='phone_mobile']")).SendKeys((string)(x1Range.Cells[xlRowCnt, 20] as excel.Range).Value2);
            driver.FindElement(By.XPath("//*[@id='alias']")).SendKeys((string)(x1Range.Cells[xlRowCnt, 21] as excel.Range).Value2);

            driver.FindElement(By.XPath("//*[@id='submitAccount']")).Click();
        }
        }

        [TestMethod]
        public void addItemsToCart()
        {
            SetMethods.logIn(driver, username, password);

            driver.FindElement(By.XPath("//*[@id='center_column']/ul/li/a/span")).Click();
            
           

            Random rnd = new Random();
            int numOfItem = rnd.Next(1, 4);



            for(int i= 1; i<=numOfItem; i++)
            {


                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                var itemX = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='homefeatured']/li[" + i + "]/div/div[1]/div/a[1]/img")));

                

                Actions actions = new Actions(driver);
                actions.MoveToElement(itemX).Perform();

                driver.FindElement(By.XPath("//*[@id='homefeatured']/li[" + i + "]/div/div[2]/div[2]/a[1]/span")).Click();

                System.Threading.Thread.Sleep(10000);

                

                if(i == numOfItem)
                {
                    driver.FindElement(By.XPath("//*[@id='layer_cart']/div[1]/div[2]/div[4]/a/span")).Click();
                }
                else
                {
                    driver.FindElement(By.XPath("//*[@id='layer_cart']/div[1]/div[2]/div[4]/span/span")).Click();

                }

            }
          

            driver.FindElement(By.XPath("//*[@id='center_column']/p[2]/a[1]")).Click();

            driver.FindElement(By.XPath("//*[@id='center_column']/form/p/button/span")).Click();

            driver.FindElement(By.XPath("//input[@name='cgv']")).Click();

            driver.FindElement(By.XPath("//*[@id='form']/p/button/span")).Click();

            driver.FindElement(By.XPath("//*[@id='HOOK_PAYMENT']/div[1]/div/p/a")).Click();

            // driver.FindElement(By.XPath("//*[@id='HOOK_PAYMENT']/div[2]/div/p/a")).Click();
            driver.FindElement(By.XPath("//*[@id='cart_navigation']/button/span")).Click();

           
        }

        [TestMethod]
        public void viewOrderHistory()
        {
            SetMethods.logIn(driver, username, password);
          

            driver.FindElement(By.XPath("//*[@id='center_column']/div/div[1]/ul/li[1]/a/span")).Click();

            IWebElement tblOrders = driver.FindElement(By.XPath("//table[@id='order-list']"));
            int rowCnt = tblOrders.FindElements(By.TagName("tr")).Count;

           

            ///////////////send mail///////////////////////


            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com",587);
            mail.From = new MailAddress("isaiamohube@gmail.co.za");
            mail.To.Add("isaiam@hippo.co.za");
            mail.Subject = "Order details";
            mail.Body = "See attached order details";

            if (rowCnt > 1)
            {

                Random rnd = new Random();
                int orderNum = rnd.Next(1, rowCnt);

                driver.FindElement(By.XPath("//*[@id='order-list']/tbody/tr[" + orderNum + "]/td[1]/a")).Click();
                driver.FindElement(By.XPath("//*[@id='block-order-detail']/div[2]/p[3]/a")).Click();

                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment("C:\\Users\\mo5599\\Downloads\\Downloads\\New folder\\IN035480.pdf");
                mail.Attachments.Add(attachment);

            }
         

            
            SmtpServer.UseDefaultCredentials = false;
            SmtpServer.Credentials = new System.Net.NetworkCredential("isaiamohube@gmail.com", "Isaia@gmail1");
            SmtpServer.EnableSsl = true;

            SmtpServer.Send(mail);



        }

        [TestMethod]
        public void submitQuery()
        {

            SetMethods.logIn(driver, username, password);
            excel._Worksheet x1WorkSheet = x1WorkBook.Sheets[4];
            excel.Range x1Range = x1WorkSheet.UsedRange;
            int xlRowCnt;

            for (xlRowCnt = 2; xlRowCnt <= x1Range.Rows.Count; xlRowCnt++)
            {

                driver.FindElement(By.XPath("//*[@id='contact-link']/a")).Click();

                SetMethods.Select(driver, "//select[@name='id_contact']", (string)(x1Range.Cells[xlRowCnt, 1] as excel.Range).Value2, "XPath");

                driver.FindElement(By.XPath("//*[@id='message']")).SendKeys((string)(x1Range.Cells[xlRowCnt, 2] as excel.Range).Value2);

                driver.FindElement(By.XPath("//button[@id='submitMessage']")).Click();
            }
        }

      
       
    

        [TestCleanup]
        public void cleanUp()
        {
            driver.Close();
        }
    }
}