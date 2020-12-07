using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

namespace New1
{
    internal class SetMethods
    {
        

        public static void logIn(IWebDriver driver, string username, string password)
        {
            driver.FindElement(By.XPath("//*[@id='header']/div[2]/div/div/nav/div[1]")).Click();
            driver.FindElement(By.XPath("//*[@id='email']")).SendKeys(username);
            driver.FindElement(By.XPath("//*[@id='passwd']")).SendKeys(password);
            driver.FindElement(By.XPath("//button[@name='SubmitLogin']")).Click();
        }
        //Selecting a drop down
        public static void Select(IWebDriver driver, string Object, string value, string objectType)
        {
            if (objectType == "Id")
                new SelectElement(driver.FindElement(By.Id(Object))).SelectByText(value);
            if (objectType == "name")
                new SelectElement(driver.FindElement(By.Name(Object))).SelectByText(value);
            if (objectType == "XPath")
                new SelectElement(driver.FindElement(By.XPath(Object))).SelectByValue(value);
        }
    }
}