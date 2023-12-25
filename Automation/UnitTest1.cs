using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Security.Cryptography;
using System.Xml.Linq;

namespace Automation
{
    public class Tests
    {
        private IWebDriver driver;

        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
        }

        [Test]
        public void SendMailTest()
        {
            var outlook = new Outlook(driver);

            var unique_string = get_unique_string(10);

            outlook.Login();
            outlook.NewMail();
            outlook.SendMail(unique_string);

            // Wait 10 second to receive mail
            Thread.Sleep(10000);

            outlook.OpenLastMail();
            Email email = outlook.GetLastMailBody();

            Assert.IsNotNull(email);
            Assert.AreEqual($"Hello, World - {unique_string}", email.Subject);
            Assert.AreEqual($"Body - {unique_string}", email.Body);

        }

        string get_unique_string(int string_length)
        {
            using (var rng = new RNGCryptoServiceProvider())
            {
                var bit_count = (string_length * 6);
                var byte_count = ((bit_count + 7) / 8); // rounded up
                var bytes = new byte[byte_count];
                rng.GetBytes(bytes);
                return Convert.ToBase64String(bytes);
            }
        }
    }
}