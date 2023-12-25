using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Automation
{
    public class Outlook
    {
        private IWebDriver _driver;

        public Outlook(IWebDriver driver)
        {
            _driver = driver;
            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);

        }

        public void Login()
        {
            _driver.Navigate().GoToUrl("https://login.live.com/login.srf");

            _driver.FindElement(By.XPath("//input[@type=\"email\"]"))
                .SendKeys("adityakc776@outlook.com");

            _driver.FindElement(By.XPath("//input[@type=\"submit\"]"))
                .Click();

            Thread.Sleep(1000);

            _driver.FindElement(By.XPath("//input[@type=\"password\"]"))
                .SendKeys("iloveceridian77");

            _driver.FindElement(By.XPath("//div[@class='inline-block button-item ext-button-item']"))
                .Click();

            Thread.Sleep(3000);

            _driver.FindElement(By.XPath("//input[@id='idBtn_Back']/.."))
                .Click();

            _driver.Navigate().GoToUrl("https://outlook.live.com/mail/0/");

        }

        public void NewMail()
        {
            Thread.Sleep(10000);

            _driver.FindElement(By.XPath("//button[@aria-label='New mail']"))
                .Click();

        }

        public void SendMail(string unique_string)
        {
            Thread.Sleep(3000);
            _driver.FindElement(By.XPath("//div[@aria-label='To']"))
                .SendKeys("adityakc776@outlook.com");

            _driver.FindElement(By.XPath("//input[@aria-label='Add a subject']"))
                .SendKeys($"Hello, World - {unique_string}");

            _driver.FindElement(By.XPath("//div[@aria-label='Message body, press Alt+F10 to exit']"))
                .SendKeys($"Body - {unique_string}");

            _driver.FindElement(By.XPath("//button[@aria-label='Send']"))
                .Click();
        }

        public void OpenLastMail()
        {
            _driver.FindElement(By.XPath("(//div[@id='MailList']/div/div/div/div/div/div/div/div)[2]"))
                .Click();
        }

        public Email GetLastMailBody()
        {
            string subject = _driver.FindElement(
                    By.XPath("//div[contains(@class, 'allowTextSelection')]/span")
                    ).Text;

            string body = _driver.FindElement(
                    By.XPath("(//div[@dir='ltr'])[2]/div")
                    ).Text;

            return new Email()
            {
                Subject = subject,
                Body = body
            };

        }
    }
}
