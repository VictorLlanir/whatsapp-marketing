using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Whatsapp.Marketing.Entities;

namespace WhatsappMarketing
{
    class Program
    {
        static void Main(string[] args)
        {
            RowsToRemove = new List<int>();
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
                    var couldNavigate = NavigateToContactUrl(contact.Number);
                    if (!couldNavigate)
                    {
                        if (!string.IsNullOrEmpty(contact.SecondNumber))
                            couldNavigate = NavigateToContactUrl(contact.SecondNumber);

                        if (!couldNavigate)
                        {
                            Console.WriteLine($"     [ERRO] {contact.Name} - {contact.Number} / {contact.SecondNumber} (Número inválido)");
                            Log($"[ERRO] {contact.Name} - {contact.Number} / {contact.SecondNumber} (Número inválido)\n");
                            continue;
                        }
                    }

                    SendMessage(contact);
                    Console.WriteLine($"     [Mensagem enviada] {contact.Name} - {contact.Number}");

                    //if (IsElementPresent(By.ClassName("_1qPwk")))
                    //{
                    //    if (ChromeDriver.FindElementsByClassName("_1qPwk").Last().FindElement(By.TagName("span")).GetAttribute("aria-label") == " Pendente ")
                    //        Thread.Sleep(1000);
                    //}
                    Thread.Sleep(10000);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"     [ERRO] {contact.Name} - {contact.Number} ({e.Message})");
                    continue;
                }
            }

            ChromeDriver.Close();
            ChromeDriver.Dispose();
            Console.ReadLine();
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
                    Contacts.Add(new Contact(worksheet.Cells[$"A{i}"].Text, worksheet.Cells[$"B{i}"].Text, worksheet.Cells[$"C{i}"].Text));
                    RowsToRemove.Add(i);
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
        private static bool NavigateToContactUrl(string number)
        {
            var treatedNumber = Regex.Replace(number, "[^0-9]", "");
            treatedNumber = treatedNumber.TrimStart('0');            if (treatedNumber.Substring(0, 2) != "55")
                treatedNumber = "55" + treatedNumber;

            ChromeDriver.Navigate().GoToUrl($"http://web.whatsapp.com/send?phone={treatedNumber}");
            Thread.Sleep(2000);

            while (!IsElementPresent(By.ClassName("copyable-text")))
                Thread.Sleep(1500);

            if (IsAlertPresent())
                ChromeDriver.SwitchTo().Alert().Accept();

            while (!IsElementPresent(By.ClassName("copyable-text")))
                Thread.Sleep(1500);

            if (IsElementPresent(By.ClassName("_9a59P")))
            {
                if (ChromeDriver.FindElement(By.ClassName("_9a59P")).Text.Contains("O número de telefone compartilhado através de url é inválido."))
                    return false;
            }

            return true;
        }
        private static void SendMessage(Contact contact)
        {
            if (ChromeDriver.FindElementsByClassName("copyable-text").Count < 2)
            {
                Console.WriteLine($"     [ERRO] {contact.Name} - {contact.Number} (Não foi possível enviar, tente novamente)");
                Log($"[ERRO] {contact.Name} - {contact.Number} (Não foi possível enviar, tente novamente)\n");

                return;
            }

            var personalMessage = Message.Select(line => line.Replace("#nome", contact.Name)).ToList();
            var messageField = ChromeDriver.FindElementsByClassName("copyable-text").Last();

            foreach (var line in personalMessage.Where(line => line != ""))
            {
                messageField.SendKeys(line);
                while (!IsElementPresent(By.XPath("/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[3]/button/span")))
                    Thread.Sleep(500);
                ChromeDriver.FindElementByXPath("/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[3]/button/span").Click();
                Thread.Sleep(600);
            }
        }
        private static void RemoveRowsFromWorksheet()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var spreadsheetFilePath = Directory.GetFiles(CurrentDirectoryPath, "*.*").First(f => f.EndsWith(".xlsx") || f.EndsWith(".xls"));
            var package = new ExcelPackage(new FileInfo(spreadsheetFilePath.Split('/').Last()));
            var worksheet = package.Workbook.Worksheets.First();
            foreach (var row in RowsToRemove)
                worksheet.DeleteRow(row);

            package.Save();
            package.Dispose();
        }

        public static string CurrentDirectoryPath { get; private set; } = Directory.GetCurrentDirectory();
        public static List<Contact> Contacts { get; private set; } = new List<Contact>();
        public static IEnumerable<string> Message { get; private set; }
        public static ChromeDriver ChromeDriver { get; private set; }
        public static List<int> RowsToRemove { get; private set; }
    }
}
