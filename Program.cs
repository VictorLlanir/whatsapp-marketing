using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Whatsapp.Marketing.Entities;

namespace WhatsappMarketing
{
    class Program
    {
        private static readonly string CurrentDirectoryPath = Directory.GetCurrentDirectory();
        static void Main(string[] args)
        {

            Console.WriteLine("|--------------------------------------------------|");
            Console.WriteLine("                WHATSAPP MARKETING                  ");
            Console.WriteLine("|--------------------------------------------------|");
            Console.WriteLine("");
            Console.WriteLine("");

            Console.WriteLine(">> LENDO MENSAGEM");
            var messageFile = Directory.GetFiles(CurrentDirectoryPath, "message.txt").First();
            var message = File.ReadAllLines(messageFile);
            Console.WriteLine("     [Ok] Mensagem lida!");
            Console.WriteLine("");
            Console.WriteLine("");

            Console.WriteLine(">> BUSCANDO CONTATOS");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var contacts = new List<Contact>();
            var spreadsheetFilePath = Directory.GetFiles(CurrentDirectoryPath, "*.*").First(f => f.EndsWith(".xlsx") || f.EndsWith(".xls"));
            using (var package = new ExcelPackage(new FileInfo(spreadsheetFilePath.Split('/').Last())))
            {
                var worksheet = package.Workbook.Worksheets.First();
                var lastRow = worksheet.Dimension.End.Row + 1;
                for (int i = 2; i < lastRow; i++)
                    contacts.Add(new Contact(worksheet.Cells[$"A{i}"].Text, worksheet.Cells[$"B{i}"].Text));
            }
            Console.WriteLine("     [Ok] Contatos guardados!");
            Console.WriteLine("");
            Console.WriteLine("");

            Console.WriteLine(">> ABRINDO O WHATSAPP, AGUARDADO AUTENTICAÇÃO");
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArgument("--start-maximized");
            chromeOptions.AddArgument("user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36");
            var chromeDriver = new ChromeDriver(CurrentDirectoryPath, chromeOptions);
            chromeDriver.Navigate().GoToUrl("https://web.whatsapp.com/");
            try
            {
                while (IsElementPresent(chromeDriver, By.ClassName("landing-title")))
                    Thread.Sleep(2000);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
            Console.WriteLine("     [Ok] Autenticado!");
            Console.WriteLine("");
            Console.WriteLine("");

            Console.WriteLine(">> ENVIANDO MENSAGENS");
            foreach (var contact in contacts)
            {
                chromeDriver.Navigate().GoToUrl($"http://web.whatsapp.com/send?phone={contact.Number}");
                var personalMessage = message.Select(line => line.Replace("#nome", contact.Name)).ToList();

                while (!IsElementPresent(chromeDriver, By.ClassName("copyable-text")))
                    Thread.Sleep(1500);

                if (IsAlertPresent(chromeDriver))
                    chromeDriver.SwitchTo().Alert().Accept();

                if (IsElementPresent(chromeDriver, By.XPath("/html/body/div[1]/div/span[2]/div/span/div/div/div/div/div/div[1]")))
                {
                    if (chromeDriver
                        .FindElement(By.XPath("/html/body/div[1]/div/span[2]/div/span/div/div/div/div/div/div[1]")).Text
                        .Contains("O número de telefone compartilhado através de url é inválido."))
                    {
                        Console.WriteLine($"     [ERRO] {contact.Name} - {contact.Number} (Número inválido)");
                        Log($"[ERRO] {contact.Name} - {contact.Number} (Número inválido)\n");
                        chromeDriver
                            .FindElement(
                                By.XPath("/html/body/div[1]/div/span[2]/div/span/div/div/div/div/div/div[2]/div"))
                            .Click();
                        continue;
                    }

                }

                while (!IsElementPresent(chromeDriver, By.ClassName("copyable-text")))
                    Thread.Sleep(1000);

                var messageField = chromeDriver.FindElementsByClassName("copyable-text").Last();
                foreach (var line in personalMessage.Where(line => line != ""))
                {
                    messageField.SendKeys(line);
                    while (!IsElementPresent(chromeDriver,
                        By.XPath("/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[3]/button/span")))
                        Thread.Sleep(1000);
                    chromeDriver.FindElementByXPath("/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[3]/button/span").Click();
                }

                Console.WriteLine($"     [Mensagem enviada] {contact.Name} - {contact.Number}");
                Thread.Sleep(2000);
            }
        }
        private static bool IsElementPresent(ChromeDriver driver, By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch
            {
                return false;
            }
        }
        private static bool IsAlertPresent(ChromeDriver driver)
        {
            try
            {
                driver.SwitchTo().Alert();
                return true;
            }
            catch
            {
                return false;
            }
        }
        private static void Log(string logMessage) => File.AppendAllText($@"{CurrentDirectoryPath}\Log.txt", logMessage);
    }
}
