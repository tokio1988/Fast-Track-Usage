using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using OpenQA.Selenium.Interactions;
using Actions = OpenQA.Selenium.Interactions.Actions;
using System.IO;
using System.Diagnostics;
using System.Security.AccessControl;
using Microsoft.Office.Interop.Excel;
using System.Reflection.Emit;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PoC_Sharepoint
{
    class FastTrack
    {
        // CONFIGURATION: All configuration is done here and only here :)

        private static int GeneralTimeInMinutes;
        //  private static int GeneralTimeInDays;
        private enum GenerationType
        {
            Generate,
            Sustain
        }

        // Parameters
        private static GenerationType generationType;
        private static string userFilePath;
        private static string sharepointUrl;

        // Excel variables
        private static Excel.Application xlApp = new Excel.Application();
        private static Range xlRange;
        private static int lastUsedRow;


        // Logger variables
        private static string SuccessSharepoint;
        private static string FailedSharepoint;
        private static string SuccessTeams;
        private static string FailedTeams;
        private static string SuccessYammer;
        private static string FailedYammer;
        private static string SuccessOutlook;
        private static string FailedOutlook;
        private static string FailedExceptions;
        private static string time;

        // User Count Variables
        private static int UsersCount = 0;

        // Script Duration variables
        private static int generalTime;
        // private static int generalTimeInDays;
        private static DateTime scriptEndTime;
        private static DateTime startHour;
        private static DateTime endHour;
        private static int randomStartHour;
        private static int randomEndHour;
        private static int randomStartMinute;
        private static int randomEndMinute;
        //  private static int addDays;

        private static bool isUpdateDay = true;
        static void Main(string[] args)
        {
            InitializeParameters(args);
            ChooseGenerationType(Convert.ToString(generationType));

        }

        private static void ChooseGenerationType(string GenerationType)
        {
            if (generationType == FastTrack.GenerationType.Generate)
            {
                Console.WriteLine("Wprowadź ścieżke do pliku Excel : ");
                userFilePath = Console.ReadLine();
                Console.WriteLine("Wprowadź ścieżke do dokumentu w sharepoint : ");
                sharepointUrl = Console.ReadLine();
                Console.WriteLine("Wprowadź czas działania skryptu w minutach : ");
                GeneralTimeInMinutes = Convert.ToInt32(Console.ReadLine());

                InitializeExcel(userFilePath);
                InitializeLogger();
                InitializeScriptDuration();

                for (int i = 4; i <= lastUsedRow; i++)
                {
                    string login = GetUserLogin(i);
                    string password = GetUserPassword(i);
                    int sleepAfterUser = CalculateUserSleepTime();

                    LogUserStart(login);

                    // BEING Main Processing Section

                    Yammer(login, password);
                    SharePoint(sharepointUrl, login, password);
                    Teams(login, password);
                    Outlook(login, password);

                    // END Main Processing Section
                    RecalculateUsers();
                    LogUserEnd(login, sleepAfterUser);
                    SleepAfterUser(sleepAfterUser);
                }
            }
            else if (generationType == FastTrack.GenerationType.Sustain)
            {
                Console.WriteLine("Wprowadź ścieżke do pliku Excel : ");
                userFilePath = Console.ReadLine();
                InitializeExcel(userFilePath);
                InitializeLogger();
                InitializeScriptDuration();

                int processedToday = 0;
                UsersCounter();
                SetStartEndToday();
                int dailyUsers = DailyUsers();
                while (true)
                {
                    DateTime now = DateTime.Now;
                    DateTime tenMinutesToMidnight = new DateTime(now.Year, now.Month, now.Day, 23, 50, 0);
                    DateTime lastMinuteInDay = new DateTime(now.Year, now.Month, now.Day, 23, 59, 59);
                    DateTime firstMinuteInDay = new DateTime(now.Year, now.Month, now.Day, 00, 00, 00);
                    DateTime tenMinutesAfterMidnight = new DateTime(now.Year, now.Month, now.Day, 00, 10, 00);

                    if (DateTime.Now >= tenMinutesToMidnight && DateTime.Now <= lastMinuteInDay)
                    {
                        isUpdateDay = true;
                    }

                    if (DateTime.Now >= firstMinuteInDay && DateTime.Now <= tenMinutesAfterMidnight && isUpdateDay == true)
                    {
                        dailyUsers = DailyUsers();
                        SetStartEndNextDay();
                        processedToday = 0;
                        isUpdateDay = false;
                    }
                    int randomUser = UserRandomizer();
                    //  DateTime today = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);
                    if (DateTime.Now > startHour && DateTime.Now < endHour && processedToday < dailyUsers)
                    {
                        try
                        {
                            Process(randomUser, ref processedToday, dailyUsers);
                        }
                        catch
                        {
                            Console.WriteLine();
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Problem with realization actions for user ");
                            Console.ForegroundColor = ConsoleColor.White;
                        }

                    }
                    else
                    {
                        // SetStartEndNextDay();
                        //processedToday = 0;

                    }
                }
            }
            else
            {
                Console.WriteLine("Wprowadzono niepoprawny typ generowania");
            }
        }
        private static string SetSharepointUrl(string email)
        {
            //później zrefakoyzowac żeby bylo mapowanie czyli np
            //   wataha - https://testowy.sharepoint.com/sites/Weryfikacjauprawnie2/Shared%20Documents/Forms/AllItems.aspx
            //itd
            if (Regex.Match(email, @"testowy.onmicrosoft.com").Success)
            {
                sharepointUrl = "https://testowy.sharepoint.com/sites/Weryfikacjauprawnie2/Shared%20Documents/Forms/AllItems.aspx";
            }
            else if (Regex.Match(email, @"testowy.onmicrosoft.com").Success)
            {
                sharepointUrl = "https://testowy.sharepoint.com/sites/Weryfikacja/Shared%20Documents/Forms/AllItems.aspx";
            }
            else if (Regex.Match(email, @"testowy.onmicrosoft.com").Success)
            {
                sharepointUrl = "https://testowy.sharepoint.com/sites/Weryfikacjauprawnie/Shared%20Documents/Forms/AllItems.aspx";
            }
            else if (Regex.Match(email, @"testowy.onmicrosoft.com").Success)
            {
                sharepointUrl = "https://testowy.sharepoint.com/sites/Weryfikacjauprawnie/Shared%20Documents/Forms/AllItems.aspx";
            }
            else if (Regex.Match(email, @"testowy.onmicrosoft.com").Success)
            {
                sharepointUrl = "https://testowy.sharepoint.com/sites/Weryfikacjauprawnie/Shared%20Documents/Forms/AllItems.aspx";
            }
            else if (Regex.Match(email, @"testowy.pl").Success)
            {
                sharepointUrl = "https://testowy.sharepoint.com/sites/Weryfikacjauprawnie/Shared%20Documents/Forms/AllItems.aspx";
            }
            return sharepointUrl;
        }
        private static void Process(int randomUser, ref int processedToday, int dailyUsers)
        {
            string login;
            string password;
            try
            {
                login = GetUserLogin(randomUser);
                password = GetUserPassword(randomUser);
            }
            catch
            {
                System.Threading.Thread.Sleep(getRandom());
                login = GetUserLogin(randomUser);
                password = GetUserPassword(randomUser);
            }
            int currentUserIndex = GetUserIndex(processedToday);

            int sleepAfterUser = CalculateUserSleepTime();
            string sharepoint = SetSharepointUrl(login);
            LogUserStart(login);
            ConfirmationStart(login);
            // BEING Main Processing Section
            Yammer(login, password);
            Teams(login, password);
            SharePoint(sharepoint, login, password);

            Outlook(login, password);
            // END Main Processing Section
            ConfirmationEnd(login, currentUserIndex, dailyUsers);
            LogUserEnd(login, sleepAfterUser);
            processedToday++;
        }
        private static int UserRandomizer()
        {
            int randomUser = new Random().Next(4, lastUsedRow);
            return randomUser;
        }
        private static int GetUserIndex(int index)
        {
            int currentUserIndex = index + 1;
            return currentUserIndex;
        }
        private static int DailyUsers()
        {
            double UsersCountConverted = Convert.ToDouble(UsersCount);
            double dailyUsers = (UsersCountConverted / 22.0);
            double dailyUsersPercent = (((new Random().Next(-15, 15)) * dailyUsers) / 100);
            double dailyUsersRandomizer = dailyUsers + dailyUsersPercent;

            double dailyUsersRandom = Math.Ceiling(dailyUsersRandomizer);


            if ((DateTime.Now.DayOfWeek == DayOfWeek.Saturday))
            {
                dailyUsersRandom = Math.Ceiling(dailyUsersRandomizer * 0.15);
            }
            else if (DateTime.Now.DayOfWeek == DayOfWeek.Sunday)
            {
                dailyUsersRandom = Math.Ceiling(dailyUsersRandomizer * 0.05);
            }

            int dailyUsersRand = Convert.ToInt32(dailyUsersRandom);

            return dailyUsersRand;
        }
        private static bool isHoliday(DateTime now)
        {
            List<string> hoidaysString = new List<string>() {"01/01/2020","06/01/2020",
                "12/04/2020","13/04/2020","01/05/2020","03/05/2020","31/05/2020","11/06/2020","15/08/2020",
            "01/11/2020","11/11/2020","25/12/2020","26/12/2020"};
            List<DateTime> holidays = hoidaysString.Select(date => DateTime.Parse(date)).ToList();
            return holidays.Contains(now);
        }
        private static void SleepAfterUser(int sleepAfterUser)
        {
            System.Threading.Thread.Sleep(sleepAfterUser);
        }
        private static void SetStartEndToday()
        {
            SetStartEndHour();
        }

        private static void SetStartEndNextDay()
        {
            SetStartEndHour();
        }

        private static void SetStartEndHour()
        {
            if ((DateTime.Now.DayOfWeek != DayOfWeek.Saturday) && (DateTime.Now.DayOfWeek != DayOfWeek.Sunday))
            {
                randomStartHour = new Random().Next(7, 8);
                randomEndHour = new Random().Next(17, 18);
            }
            else if ((DateTime.Now.DayOfWeek == DayOfWeek.Saturday) || (DateTime.Now.DayOfWeek == DayOfWeek.Sunday))
            {
                randomStartHour = new Random().Next(8, 10);
                randomEndHour = new Random().Next(12, 14);
            }
            randomStartMinute = RandMinute();
            randomEndMinute = RandMinute();
            //startHour = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, randomStartHour, RandMinute(), 0).AddDays(addDays);
            //endHour = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, randomEndHour, RandMinute(), 0).AddDays(addDays);

            startHour = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, randomStartHour, RandMinute(), 0);
            endHour = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, randomEndHour, RandMinute(), 0);
        }

        private static int RandMinute()
        {
            int randomMinute = new Random().Next(0, 59);
            return randomMinute;
        }
        private static void RecalculateUsers()
        {
            UsersCount--;
        }

        private static int CalculateUserSleepTime()
        {
            if (UsersCount == 0)
            {
                UsersCount = lastUsedRow - 3;
            }

            double oneUserTime = 5000;
            if (scriptEndTime > DateTime.Now)
            {
                oneUserTime = (scriptEndTime - DateTime.Now).TotalMilliseconds / UsersCount;
            }

            double smallestUserTime = oneUserTime / 2;
            double biggestUserTime = oneUserTime * 1.5;
            Random random = new Random();

            return random.Next(Convert.ToInt32(smallestUserTime), Convert.ToInt32(biggestUserTime));
        }
        private static int UsersCounter()
        {
            UsersCount = lastUsedRow - 3;
            return UsersCount;
        }
        private static dynamic GetUserPassword(int i)
        {
            return xlRange.Cells[i, 3].Value;
        }

        private static dynamic GetUserLogin(int i)
        {
            return xlRange.Cells[i, 2].Value2;
        }

        private static void LogUserEnd(string login, int sleepAfterUser)
        {
            using (StreamWriter sw = File.AppendText(time))
            {
                sw.WriteLine(login + " end at " + date() + " sleep time " + sleepAfterUser);
            }
        }

        private static void LogUserStart(string login)
        {
            using (StreamWriter sw = File.AppendText(time))
            {
                sw.WriteLine(login + " start " + date());
            }
        }
        private static void ConfirmationStart(string login)
        {
            Console.WriteLine(login + " start " + date());
        }
        private static void ConfirmationEnd(string login, int currentUserIndex, int dailyUsers)
        {
            Console.WriteLine(login + " ends " + date());
            Console.WriteLine();
            Console.WriteLine("Processed : " + currentUserIndex + " of " + dailyUsers);
            Console.WriteLine("-------------------------");
            Console.WriteLine();
        }
        private static void InitializeScriptDuration()
        {
            if (generationType ==
               FastTrack.GenerationType.Generate)
            {
                generalTime = FastTrack.GeneralTimeInMinutes * 60 * 1000;
                scriptEndTime = DateTime.Now.AddMilliseconds(generalTime);
            }
            //else if (generationType ==
            //   FastTrack.GenerationType.Sustain)
            //{
            //    generalTimeInDays = FastTrack.GeneralTimeInDays * 24 * 60 * 60 * 1000;
            //    scriptEndTime = DateTime.Now.AddMilliseconds(generalTimeInDays);
            //}

        }

        private static void InitializeParameters(string[] args)
        {
            if (args.Length != 1)
            {
                Console.WriteLine("Podaj parametr - rodzaj generowania : 'Generate' lub 'Sustain'");
                throw new InvalidProgramException("Podaj parametr - rodzaj generowania : 'Generate' lub 'Sustain'");
            }
            generationType = (GenerationType)Enum.Parse(typeof(GenerationType), args[0]);
        }

        private static void InitializeExcel(string excelFilePath)
        {
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(excelFilePath);
            Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;
            lastUsedRow = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            Console.WriteLine(lastUsedRow);
        }

        private static void InitializeLogger()
        {
            long ticks = DateTime.Now.Ticks;
            FastTrack.SuccessSharepoint = String.Format(@"C:\Repo\SuccessSharepoint{0}.txt", ticks);
            FastTrack.FailedSharepoint = String.Format(@"C:\Repo\FailedSharepoint{0}.txt", ticks);
            FastTrack.SuccessTeams = String.Format(@"C:\Repo\SuccsessTeams{0}.txt", ticks);
            FastTrack.FailedTeams = String.Format(@"C:\Repo\FailedTeams{0}.txt", ticks);
            FastTrack.SuccessYammer = String.Format(@"C:\Repo\SuccessYammer{0}.txt", ticks);
            FastTrack.FailedYammer = String.Format(@"C:\Repo\FailedYammmer{0}.txt", ticks);
            FastTrack.SuccessOutlook = String.Format(@"C:\Repo\SuccsessOutlook{0}.txt", ticks);
            FastTrack.FailedOutlook = String.Format(@"C:\Repo\FailedOutlook{0}.txt", ticks);
            FastTrack.FailedExceptions = String.Format(@"C:\Repo\FailedExceptions{0}.txt", ticks);
            FastTrack.time = String.Format(@"C:\Repo\time{0}.txt", ticks);
        }

        private static void Yammer(string login, string password)
        {
            var yammerdriver = new ChromeDriver();

            try
            {
                yammerdriver.Navigate().GoToUrl("https://www.yammer.com/login?locale=pl-PL&locale_type=standard");
                System.Threading.Thread.Sleep(getRandom());
                LoginToYammer(login, password, yammerdriver);
                try
                {
                    yammerdriver.FindElementByClassName("login-back").FindElement(By.TagName("a")).Click();
                    System.Threading.Thread.Sleep(getRandom());
                    //LoginToYammer(login, password, yammerdriver, false);
                    LoginToYammer(login, password, yammerdriver);
                }
                catch
                {

                }
                System.Threading.Thread.Sleep(getRandom());
                try
                {
                    yammerdriver.FindElementByClassName("publisher-placeholder--text").Click();
                }
                catch
                {
                    System.Threading.Thread.Sleep(getRandom());
                    LoginToYammer(login, password, yammerdriver, false);
                    System.Threading.Thread.Sleep(getRandom());
                    yammerdriver.FindElementByClassName("publisher-placeholder--text").Click();
                }
                System.Threading.Thread.Sleep(getRandom());
                if (generationType == FastTrack.GenerationType.Generate)
                {

                    yammerdriver.FindElementByClassName("notranslate").SendKeys(getPost());
                    System.Threading.Thread.Sleep(getRandom());
                    yammerdriver.FindElementByXPath("//*[@aria-label='Add a group and/or people']").Click();
                    System.Threading.Thread.Sleep(getRandom());
                    yammerdriver.FindElementByXPath("//*[@aria-label='Add a group and/or people']").SendKeys("Weryfikacja Uprawnień");
                    System.Threading.Thread.Sleep(10000);
                    yammerdriver.FindElementByXPath("//*[@aria-label='Add a group and/or people']").SendKeys(Keys.Enter);
                    System.Threading.Thread.Sleep(getRandom());
                    repeatAction(yammerdriver.FindElementByClassName("publisher-submit"));
                    System.Threading.Thread.Sleep(getRandom());
                }
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(login + " completed Yammer");
                Console.ForegroundColor = ConsoleColor.White;
                using (StreamWriter sw = File.AppendText(SuccessYammer))
                {
                    sw.WriteLine(login);
                }
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(login + " failed Yammer");
                Console.WriteLine("Reason : " + e.Message);
                Console.ForegroundColor = ConsoleColor.White;
                using (StreamWriter sw = File.AppendText(FailedYammer))
                {
                    sw.WriteLine(login);
                }
                using (StreamWriter sw = File.AppendText(FailedExceptions))
                {
                    sw.WriteLine(e.Message);
                }
            }
            finally
            {
                try
                {
                    yammerdriver.Close();
                }
                catch (Exception e)
                {
                    using (StreamWriter sw = File.AppendText(FailedExceptions))
                    {
                        sw.WriteLine(e.Message);
                    }
                }

            }
        }

        private static void LoginToYammer(string login, string password, ChromeDriver yammerdriver, bool requirePass = true)
        {
            yammerdriver.FindElement(By.Id("login")).SendKeys(login);
            System.Threading.Thread.Sleep(getRandom());
            yammerdriver.FindElementById("password").SendKeys(password);
            System.Threading.Thread.Sleep(getRandom());
            System.Threading.Thread.Sleep(10000);

            if (requirePass)
            {
                yammerdriver.FindElementById("i0116").SendKeys(login + Keys.Enter);
                System.Threading.Thread.Sleep(getRandom());
                yammerdriver.FindElementById("i0118").SendKeys(password + Keys.Enter);
                System.Threading.Thread.Sleep(getRandom());

                try
                {
                    yammerdriver.FindElement(By.Id("idSIButton9")).Click();
                    System.Threading.Thread.Sleep(getRandom());
                    try
                    {
                        yammerdriver.FindElement(By.Id("idSIButton9")).Click();
                        System.Threading.Thread.Sleep(getRandom());
                    }
                    catch
                    {

                    }
                }
                catch
                {
                    yammerdriver.FindElement(By.ClassName("form-group")).FindElement(By.TagName("a")).Click();
                    System.Threading.Thread.Sleep(getRandom());
                }

                System.Threading.Thread.Sleep(getRandom());

                try
                {
                    yammerdriver.FindElement(By.ClassName("form-group")).FindElement(By.TagName("a")).Click();
                    System.Threading.Thread.Sleep(getRandom());
                }
                catch
                {

                }
            }

            System.Threading.Thread.Sleep(15000);
        }
        private static void LoginToTeams(IWebDriver teamsdriver, string login, string password)
        {
            teamsdriver.Navigate().GoToUrl("https://teams.microsoft.com");
            System.Threading.Thread.Sleep(getRandom());
            teamsdriver.FindElement(By.Id("i0116")).SendKeys(login + Keys.Enter);
            System.Threading.Thread.Sleep(getRandom());
            try
            {
                teamsdriver.FindElement(By.Id("i0118")).SendKeys(password + Keys.Enter);
            }
            catch
            {
                System.Threading.Thread.Sleep(getRandom());
                teamsdriver.FindElement(By.Id("i0118")).SendKeys(password + Keys.Enter);
            }

            System.Threading.Thread.Sleep(getRandom());
        }
        private static void LoginToSharepoint(IWebDriver driver, string sharepoint, string login, string password)
        {
            System.Threading.Thread.Sleep(getRandom());
            driver.Navigate().GoToUrl(sharepoint);
            System.Threading.Thread.Sleep(getRandom());
            driver.FindElement(By.Id("i0116")).SendKeys(login + Keys.Enter);
            System.Threading.Thread.Sleep(getRandom());
            driver.FindElement(By.Id("i0118")).SendKeys(password + Keys.Enter);
            System.Threading.Thread.Sleep(getRandom());
            try
            {
                driver.FindElement(By.Id("idSIButton9")).Click();
                System.Threading.Thread.Sleep(getRandom());
                try
                {
                    driver.FindElement(By.Id("idSIButton9")).Click();
                    System.Threading.Thread.Sleep(getRandom());
                }
                catch
                {

                }
            }
            catch
            {
                driver.FindElement(By.ClassName("form-group")).FindElement(By.TagName("a")).Click();
                System.Threading.Thread.Sleep(getRandom());
            }
        }

        private static void Teams(string login, string password)
        {
            IWebDriver teamsdriver = new ChromeDriver();
            Actions teamsactions = new Actions(teamsdriver);
            //TEAMS
            try
            {
                try
                {
                    LoginToTeams(teamsdriver, login, password);
                }
                catch
                {
                    System.Threading.Thread.Sleep(10000);
                    if (teamsdriver != null)
                    {
                        teamsdriver.Close();
                        System.Threading.Thread.Sleep(10000);
                        LoginToTeams(teamsdriver, login, password);
                    }

                }

                try
                {
                    teamsdriver.FindElement(By.Id("idSIButton9")).Click();
                    System.Threading.Thread.Sleep(getRandom());
                    try
                    {
                        teamsdriver.FindElement(By.Id("idSIButton9")).Click();
                        System.Threading.Thread.Sleep(getRandom());
                    }
                    catch
                    {

                    }
                }
                catch
                {
                    teamsdriver.FindElement(By.ClassName("form-group")).FindElement(By.TagName("a")).Click();
                    System.Threading.Thread.Sleep(getRandom());
                }
                System.Threading.Thread.Sleep(10000);
                teamsdriver.FindElement(By.ClassName("use-app-lnk")).Click();
                System.Threading.Thread.Sleep(getRandom());
                try
                {
                    System.Threading.Thread.Sleep(10000);
                    if ((teamsdriver.FindElement(By.ClassName("prompt-cards-holder"))) != null)
                    {
                        System.Threading.Thread.Sleep(5000);
                        teamsdriver.FindElement(By.XPath("//*[@aria-label='Cancel']")).Click();
                    }
                }
                catch (Exception e)
                {

                }
                System.Threading.Thread.Sleep(getRandom());
                try
                {
                    var TeamsWeryfikacjaUprawnien = ((teamsdriver.FindElement(By.XPath("//*[@data-tid='team-Weryfikacja uprawnień-li']"))) != null);
                }
                catch (Exception e)
                {
                    System.Threading.Thread.Sleep(7000);
                    teamsdriver.Navigate().GoToUrl("https://teams.microsoft.com/_#/discover");
                    System.Threading.Thread.Sleep(7000);
                    try
                    {
                        repeatAction(teamsdriver.FindElement(By.Id("create_join_team_text")));
                    }
                    catch
                    {
                        teamsdriver.Navigate().Refresh();
                        System.Threading.Thread.Sleep(5000);
                        repeatAction(teamsdriver.FindElement(By.Id("create_join_team_text")));
                    }
                    System.Threading.Thread.Sleep(getRandom());
                    repeatAction(teamsdriver.FindElement(By.ClassName("team-info")));
                    System.Threading.Thread.Sleep(getRandom());
                    teamsdriver.FindElement(By.XPath("//*[@track-summary='Join Team']")).Click();
                }
                System.Threading.Thread.Sleep(5000);
                try
                {
                    if ((teamsdriver.FindElement(By.Id("engagement-surface-dialog"))) != null)
                    {
                        System.Threading.Thread.Sleep(5000);
                        teamsdriver.FindElement(By.XPath("//*[@title='Close']")).Click();
                        System.Threading.Thread.Sleep(getRandom());
                    }
                }
                catch
                {

                }
                teamsdriver.FindElement(By.XPath("//*[@data-tid='team-Weryfikacja uprawnień-li']")).Click();
                System.Threading.Thread.Sleep(getRandom());
                teamsdriver.FindElement(By.XPath("//*[@data-tid='team-Weryfikacja uprawnień-li']")).Click();
                System.Threading.Thread.Sleep(getRandom());
                if (generationType ==
                FastTrack.GenerationType.Generate)
                {

                    teamsdriver.FindElement(By.XPath("//*[@aria-label='Start a new conversation. Type @ to mention someone.']")).SendKeys(getPost() + Keys.Enter);
                    System.Threading.Thread.Sleep(10000);
                    repeatAction(teamsdriver.FindElement(By.ClassName("user-information-button")));
                    System.Threading.Thread.Sleep(getRandom());
                    repeatAction(teamsdriver.FindElement(By.XPath("//*[@data-tid='logout-button']")));
                }
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(login + " completed Teams");
                Console.ForegroundColor = ConsoleColor.White;
                using (StreamWriter sw = File.AppendText(SuccessTeams))
                {
                    sw.WriteLine(login);
                }
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(login + " failed Teams");
                Console.WriteLine("Reason : " + e.Message);
                Console.ForegroundColor = ConsoleColor.White;
                using (StreamWriter sw = File.AppendText(FailedTeams))
                {
                    sw.WriteLine(login);
                }
                using (StreamWriter sw = File.AppendText(FailedExceptions))
                {
                    sw.WriteLine(e.Message);
                }
            }
            finally
            {
                try
                {
                    teamsdriver.Close();
                }
                catch
                {

                }
            }
        }
        private static void SharePoint(string sharepoint, string login, string password)
        {
            IWebDriver driver = new ChromeDriver();
            Actions actions = new Actions(driver);

            try
            {
                try
                {
                    LoginToSharepoint(driver, sharepoint, login, password);
                }
                catch
                {
                    if (driver != null)
                    {
                        System.Threading.Thread.Sleep(getRandom());
                        driver.Close();
                    }
                    System.Threading.Thread.Sleep(getRandom());
                    LoginToSharepoint(driver, sharepoint, login, password);
                }

                try
                {
                    System.Threading.Thread.Sleep(getRandom());
                    System.Threading.Thread.Sleep(5000);
                    if (driver.FindElement(By.ClassName("FirstRunDialog-main")) != null)
                    {
                        System.Threading.Thread.Sleep(getRandom());
                        driver.FindElement(By.XPath("//*[@aria-label='Close dialog']")).Click();
                        System.Threading.Thread.Sleep(getRandom());
                    }
                }
                catch (Exception e)
                {

                }
                System.Threading.Thread.Sleep(getRandom());

                if (generationType == FastTrack.GenerationType.Generate)

                {
                    try
                    {
                        System.Threading.Thread.Sleep(getRandom());
                        IWebElement elementLocator = driver.FindElement(By.ClassName("ms-FocusZone"));
                        actions.DoubleClick(elementLocator).Perform();
                    }
                    catch
                    {
                        driver.Navigate().GoToUrl(sharepoint);
                        System.Threading.Thread.Sleep(getRandom());
                        IWebElement elementLocator = driver.FindElement(By.ClassName("ms-FocusZone"));
                        actions.DoubleClick(elementLocator).Perform();

                    }

                    System.Threading.Thread.Sleep(getRandom());
                    var tabs = driver.WindowHandles;
                    if (tabs.Count > 1)
                    {
                        driver.SwitchTo().Window(tabs[1]);
                        driver.Close();
                        driver.SwitchTo().Window(tabs[0]);
                    }
                    System.Threading.Thread.Sleep(15000);
                    driver.Navigate().GoToUrl(sharepoint);
                    repeatAction(driver.FindElement(By.ClassName("ms-SelectionZone")));
                    System.Threading.Thread.Sleep(getRandom());
                    driver.FindElement((By.XPath("//button[@aria-label='Download']"))).Click();
                    System.Threading.Thread.Sleep(getRandom());
                }
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(login + " completed Sharepoint");
                Console.ForegroundColor = ConsoleColor.White;
                using (StreamWriter sw = File.AppendText(SuccessSharepoint))
                {
                    sw.WriteLine(login);
                }
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(login + " failed Sharepoint");
                Console.WriteLine("Reason : " + e.Message);
                using (StreamWriter sw = File.AppendText(FailedSharepoint))
                {
                    sw.WriteLine(login);
                }
                using (StreamWriter sw = File.AppendText(FailedExceptions))
                {
                    sw.WriteLine(e.Message);
                }
            }
            finally
            {
                try
                {
                    driver.Close();
                }
                catch
                {

                }
            }
        }
        private static void LoginToOutlook(IWebDriver outlookdriver, string login, string password)
        {

            System.Threading.Thread.Sleep(10000);

            outlookdriver.Navigate().GoToUrl("https://portal.office.com/");
            System.Threading.Thread.Sleep(getRandom());
            outlookdriver.FindElement(By.Id("i0116")).SendKeys(login + Keys.Enter);
            System.Threading.Thread.Sleep(getRandom());
            outlookdriver.FindElement(By.Id("i0118")).SendKeys(password + Keys.Enter);
            System.Threading.Thread.Sleep(getRandom());
            outlookdriver.FindElement(By.Id("idSIButton9")).Click();
            System.Threading.Thread.Sleep(getRandom());


        }
        private static void Outlook(string login, string password)
        {
            IWebDriver outlookdriver = new ChromeDriver();
            Actions outlookactions = new Actions(outlookdriver);
            try
            {
                try
                {

                    LoginToOutlook(outlookdriver, login, password);
                }
                catch
                {
                    System.Threading.Thread.Sleep(getRandom());
                    if (outlookdriver != null)
                    {
                        outlookdriver.Close();
                    }

                    System.Threading.Thread.Sleep(getRandom());
                    LoginToOutlook(outlookdriver, login, password);
                }

                try
                {
                    if (outlookdriver.FindElement(By.Id("slide-container")) != null)
                    {
                        System.Threading.Thread.Sleep(getRandom());
                        (outlookdriver.FindElement(By.Id("close-welcome-overlay"))).Click();
                        System.Threading.Thread.Sleep(getRandom());
                    }
                }
                catch
                {

                }
                try
                {
                    outlookdriver.FindElement(By.Id("O365_MainLink_NavMenu")).Click();
                }
                catch (Exception ex)
                {
                    outlookdriver.Navigate().GoToUrl("https://portal.office.com/");
                    System.Threading.Thread.Sleep(getRandom());
                    outlookdriver.FindElement(By.Id("O365_MainLink_NavMenu")).Click();

                }
                System.Threading.Thread.Sleep(getRandom());
                outlookdriver.FindElement(By.Id("O365_AppTile_Mail")).Click();
                System.Threading.Thread.Sleep(getRandom());
                System.Threading.Thread.Sleep(getRandom());
                try
                {
                    System.Threading.Thread.Sleep(getRandom());
                    var OutlookPopup = ((outlookdriver.FindElement(By.ClassName("ms-Dialog-main"))) != null);
                    System.Threading.Thread.Sleep(10000);
                    outlookdriver.FindElement(By.ClassName("_1EsTEQqnWolfpJqbg_aysc")).Click();

                }
                catch (Exception ex)
                {

                }
                System.Threading.Thread.Sleep(5000);
                try
                {
                    outlookdriver.FindElement(By.ClassName("ms-Modal-scrollableContent")).Click();
                    System.Threading.Thread.Sleep(getRandom());
                    outlookdriver.FindElement(By.ClassName("ms-Dialog-button--close")).Click();

                }
                catch
                {

                }
                try
                {
                    outlookdriver.FindElement(By.ClassName("ms-Button-label")).Click();
                }
                catch
                {
                    outlookdriver.Navigate().Refresh();
                    System.Threading.Thread.Sleep(5000);
                    outlookdriver.FindElement(By.ClassName("ms-Button-label")).Click();
                }


                System.Threading.Thread.Sleep(getRandom());
                outlookdriver.FindElement(By.XPath("//*[@aria-label = 'To']")).SendKeys(login + Keys.Enter);
                System.Threading.Thread.Sleep(getRandom());
                outlookdriver.FindElement(By.XPath("//*[@aria-label = 'Send']")).Click();
                System.Threading.Thread.Sleep(getRandom());
                outlookdriver.FindElement(By.Id("ok-1")).Click();
                System.Threading.Thread.Sleep(getRandom());
                if (generationType == FastTrack.GenerationType.Sustain)
                {
                    System.Threading.Thread.Sleep(getRandom());
                    outlookdriver.FindElement(By.XPath("//*[@title='Inbox']")).Click();
                    System.Threading.Thread.Sleep(getRandom());
                    outlookdriver.FindElement(By.ClassName("ms-Check")).Click();
                    System.Threading.Thread.Sleep(getRandom());
                    try
                    {
                        outlookdriver.FindElement(By.XPath("//*[@name='Empty Focused']")).Click();
                        System.Threading.Thread.Sleep(getRandom());
                    }
                    catch
                    {
                        outlookdriver.FindElement(By.XPath("//*[@name='Empty Folder']")).Click();
                        System.Threading.Thread.Sleep(getRandom());
                    }
                    outlookdriver.FindElement(By.Id("ok-7")).Click();
                    System.Threading.Thread.Sleep(getRandom());
                    outlookdriver.FindElement(By.XPath("//*[@title='Deleted Items']")).Click();
                    System.Threading.Thread.Sleep(getRandom());
                    outlookdriver.FindElement(By.ClassName("ms-Check")).Click();
                    System.Threading.Thread.Sleep(getRandom());
                    outlookdriver.FindElement(By.XPath("//*[@name='Empty folder']")).Click();
                    System.Threading.Thread.Sleep(getRandom());
                    outlookdriver.FindElement(By.Id("ok-13")).Click();

                }
                outlookdriver.FindElement(By.Id("O365_MainLink_MePhoto")).Click();
                System.Threading.Thread.Sleep(getRandom());
                outlookdriver.FindElement(By.Id("meControlSignoutLink")).Click();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(login + " completed Outlook");
                Console.ForegroundColor = ConsoleColor.White;
                using (StreamWriter sw = File.AppendText(SuccessOutlook))
                {
                    sw.WriteLine(login);
                }
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(login + " failed Outlook");
                Console.WriteLine("Reason : " + e.Message);
                Console.ForegroundColor = ConsoleColor.White;
                using (StreamWriter sw = File.AppendText(FailedOutlook))
                {
                    sw.WriteLine(login);
                }
                using (StreamWriter sw = File.AppendText(FailedExceptions))
                {
                    sw.WriteLine(e.Message);
                }

            }
            finally
            {
                try
                {
                    outlookdriver.Close();
                }
                catch (Exception e)
                {

                }

            }
        }
        private static void repeatAction(IWebElement element)
        {
            int retry = 10;

            while (retry > 0)
            {
                try
                {
                    element.Click();
                    break;
                }
                catch
                {
                    retry--;
                    if (retry == 0)
                    {
                        throw;
                    }
                    System.Threading.Thread.Sleep(5000);
                }
            }
        }
        private static string getPost()
        {
            string[] posts = { "hej", "czesc", "Cześć", "ok", "test", "działa", "witam", "Dzień Dobry" };
            Random rand = new Random();
            int index = rand.Next(1, posts.Length);
            var post = posts[index];
            return post;
        }
        private static int getRandom()
        {
            Random random = new Random();
            var mseconds = random.Next(4, 6) * 1000;
            return mseconds;
        }
        private static DateTime date()
        {
            return DateTime.Now;
        }
    }
}
