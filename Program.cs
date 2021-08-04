using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Whatsapp.Marketing.Entities;

namespace WhatsappMarketing
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("|--------------------------------------------------|");
            Console.WriteLine("                WHATSAPP MARKETING                  ");
            Console.WriteLine("|--------------------------------------------------|");
            Console.WriteLine("");
            Console.WriteLine("");

            Console.WriteLine(">> LENDO MENSAGEM");
            ReadMessage();
            Console.WriteLine("     [Ok] Mensagem lida!");
            Console.WriteLine("");
            Console.WriteLine("");

            Console.WriteLine(">> BUSCANDO CONTATOS");
            GetContacts();
            Console.WriteLine("     [Ok] Contatos guardados!");
            Console.WriteLine("");
            Console.WriteLine("");

            Console.WriteLine(">> ABRINDO O WHATSAPP, AGUARDADO AUTENTICAÇÃO");
            OpenChromeDriverAndWaitForAuthentication();
            Console.WriteLine("     [Ok] Autenticado!");
            Console.WriteLine("");
            Console.WriteLine("");

            Console.WriteLine(">> ENVIANDO MENSAGENS");
            foreach (var contact in Contacts)
            {
                try
                {
                    NavigateToContactUrl(contact.Number);
                    if (IsElementPresent(By.XPath("/html/body/div[1]/div/span[2]/div/span/div/div/div/div/div/div[1]")))
                    {
                        while (!IsElementPresent(By.XPath("/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[2]/div/div[2]")) && !IsElementPresent(By.XPath("/html/body/div[1]/div/span[2]/div/span/div/div/div/div/div/div[1]")))
                            Thread.Sleep(2000);

                        if (ChromeDriver.FindElement(By.XPath("/html/body/div[1]/div/span[2]/div/span/div/div/div/div/div/div[1]")).Text.Contains("inválido") || ChromeDriver.FindElement(By.XPath("/html/body/div[1]/div/span[2]/div/span/div/div/div/div/div/div[1]")).Text.Contains("invalid"))
                        {
                            if (!string.IsNullOrEmpty(contact.SecondNumber))
                                NavigateToContactUrl(contact.SecondNumber);

                            if (IsElementPresent(By.XPath("/html/body/div[1]/div/span[2]/div/span/div/div/div/div/div/div[1]")))
                            {
                                Thread.Sleep(2000);
                                if (ChromeDriver.FindElement(By.XPath("/html/body/div[1]/div/span[2]/div/span/div/div/div/div/div/div[1]")).Text.Contains("inválido") || ChromeDriver.FindElement(By.XPath("/html/body/div[1]/div/span[2]/div/span/div/div/div/div/div/div[1]")).Text.Contains("invalid"))
                                {
                                    ColorErrorRow(contact.Row);
                                    Console.WriteLine($"     [ERRO] {contact.Name} - {contact.Number} / {contact.SecondNumber} (Número inválido)");
                                    Log($"[ERRO] {contact.Name} - {contact.Number} / {contact.SecondNumber} (Número inválido)\n");
                                    continue;
                                }
                            }

                        }
                        Thread.Sleep(5000);
                    }

                    SendMessage(contact);

                    Console.WriteLine($"     [Mensagem enviada] {contact.Name} - {contact.Number}");

                    if (IsElementPresent(By.XPath("/html/body/div[1]/div/div/div[4]/div/div[3]/div/div/div[3]/div[20]/div/div/div/div[2]/div/div")))
                    {
                        while (ChromeDriver.FindElementsByXPath("/html/body/div[1]/div/div/div[4]/div/div[3]/div/div/div[3]/div[20]/div/div/div/div[2]/div/div").Last().FindElement(By.TagName("span")).GetAttribute("aria-label").Contains("Pendente"))
                            Thread.Sleep(1000);
                    }
                    Thread.Sleep(5000);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"     [ERRO] {contact.Name} - {contact.Number} ({e.Message})");
                    Log($"[ERRO] {contact.Name} - {contact.Number} / {contact.SecondNumber} ({e.Message})\n");
                    continue;
                }
            }

            ChromeDriver.Close();
            ChromeDriver.Dispose();
        }
        private static bool IsElementPresent(By by)
        {
            try
            {
                ChromeDriver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        private static bool IsAlertPresent()
        {
            try
            {
                ChromeDriver.SwitchTo().Alert();
                return true;
            }
            catch (NoAlertPresentException)
            {
                return false;
            }
        }
        private static void Log(string logMessage) => File.AppendAllText($@"{CurrentDirectoryPath}\Log.txt", logMessage);
        private static void ReadMessage()
        {
            var messageFile = Directory.GetFiles(CurrentDirectoryPath, "message.txt").First();
            Message = File.ReadAllLines(messageFile);
        }
        private static void GetContacts()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var spreadsheetFilePath = Directory.GetFiles(CurrentDirectoryPath, "*.*").First(f => f.EndsWith(".xlsx") || f.EndsWith(".xls"));
            using (var package = new ExcelPackage(new FileInfo(spreadsheetFilePath.Split('/').Last())))
            {
                var worksheet = package.Workbook.Worksheets.First();
                var lastRow = worksheet.Dimension.End.Row + 1;
                for (int i = 2; i <= lastRow; i++)
                {
                    Contacts.Add(new Contact(worksheet.Cells[$"A{i}"].Text, worksheet.Cells[$"B{i}"].Text, worksheet.Cells[$"C{i}"].Text, i));
                }
            }
        }
        private static void OpenChromeDriverAndWaitForAuthentication()
        {
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArgument("--start-maximized");
            chromeOptions.AddArgument("user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36");
            ChromeDriver = new ChromeDriver(CurrentDirectoryPath, chromeOptions);
            ChromeDriver.Navigate().GoToUrl("https://web.whatsapp.com/");
            try
            {
                while (IsElementPresent(By.ClassName("landing-title")))
                    Thread.Sleep(2000);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
        }
        private static void NavigateToContactUrl(string number)
        {
            var treatedNumber = Regex.Replace(number, "[^0-9]", "");
            treatedNumber = treatedNumber.TrimStart('0');
            if (treatedNumber.Substring(0, 2) != "55")
                treatedNumber = "55" + treatedNumber;

            if (treatedNumber.Length > 13)
                treatedNumber = treatedNumber.Substring(0, 13);

            ChromeDriver.Navigate().GoToUrl($"http://web.whatsapp.com/send?phone={treatedNumber}");
            Thread.Sleep(2000);

            while (!IsElementPresent(By.ClassName("copyable-text")))
                Thread.Sleep(1500);

            if (IsAlertPresent())
                ChromeDriver.SwitchTo().Alert().Accept();

            while (!IsElementPresent(By.ClassName("copyable-text")))
                Thread.Sleep(1500);
        }
        private static void SendMessage(Contact contact)
        {
            var personalMessage = Message.Select(line => line.Replace("#nome", contact.Name)).ToList();
            var messageField = ChromeDriver.FindElementByXPath("/html/body/div/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div[2]/div/div[1]/div/div[2]");
            var sendMessageButton = ChromeDriver.FindElementByXPath("/html/body/div/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div[2]/div/div[2]");
            Thread.Sleep(800);

            foreach (var line in personalMessage.Where(line => line != ""))
            {
                messageField.Click();
                messageField.SendKeys(line);
                Thread.Sleep(800);
                sendMessageButton.Click();

                while (IsElementPresent(By.XPath("span[@aria-label='Pendente']")))
                    Thread.Sleep(500);
            }
            ColorSuccessfullRow(contact.Row);
            Thread.Sleep(1000);
        }
        private static void ColorSuccessfullRow(int row)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var spreadsheetFilePath = Directory.GetFiles(CurrentDirectoryPath, "*.*").First(f => f.EndsWith(".xlsx") || f.EndsWith(".xls"));
            var package = new ExcelPackage(new FileInfo(spreadsheetFilePath.Split('/').Last()));
            var worksheet = package.Workbook.Worksheets.First();
            worksheet.Cells[$"A{row}"].Style.Fill.SetBackground(Color.Green, ExcelFillStyle.Solid);

            package.Save();
            package.Dispose();
        }
        private static void ColorErrorRow(int row)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var spreadsheetFilePath = Directory.GetFiles(CurrentDirectoryPath, "*.*").First(f => f.EndsWith(".xlsx") || f.EndsWith(".xls"));
            var package = new ExcelPackage(new FileInfo(spreadsheetFilePath.Split('/').Last()));
            var worksheet = package.Workbook.Worksheets.First();
            worksheet.Cells[$"A{row}"].Style.Fill.SetBackground(Color.Red, ExcelFillStyle.Solid);

            package.Save();
            package.Dispose();
        }

        public static string CurrentDirectoryPath { get; private set; } = Directory.GetCurrentDirectory();
        public static List<Contact> Contacts { get; private set; } = new List<Contact>();
        public static IEnumerable<string> Message { get; private set; }
        public static ChromeDriver ChromeDriver { get; private set; }
    }
}
