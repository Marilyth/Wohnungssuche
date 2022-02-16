//http://chromedriver.storage.googleapis.com/index.html?path=&sort=desc
//Download matching chromedriver.exe here, and put it into the root of the repo

using OpenQA.Selenium.Chrome;
using Newtonsoft.Json;

var RowChosen = -1;
var apartments = new List<Apartment>();
loadApartments();

var options = new ChromeOptions();
options.SetLoggingPreference(OpenQA.Selenium.LogType.Browser, OpenQA.Selenium.LogLevel.Off);
options.SetLoggingPreference(OpenQA.Selenium.LogType.Client, OpenQA.Selenium.LogLevel.Off);
options.SetLoggingPreference(OpenQA.Selenium.LogType.Driver, OpenQA.Selenium.LogLevel.Off);
options.SetLoggingPreference(OpenQA.Selenium.LogType.Profiler, OpenQA.Selenium.LogLevel.Off);
options.SetLoggingPreference(OpenQA.Selenium.LogType.Server, OpenQA.Selenium.LogLevel.Off);

var browser = new ChromeDriver(options);
//browser.Manage().Window.Position = new System.Drawing.Point(-9000, -9000);
browser.Manage().Window.Minimize();

DrawApartmentTable();
Task.Run(() => UserInput());

while (true)
{
    //await handleImmowelt(0);
    await handleImmobilienscout24(0);
    await Task.Delay(60000);
}

async Task handleImmowelt(int tab=0){
    try
    {
        //Switch to correct tab
        while(browser.WindowHandles.Count - 1 < tab){
            browser.ExecuteScript("window.open()");
        }
        browser.SwitchTo().Window(browser.WindowHandles[tab]);

        //Go to website
        if(browser.Url != "https://www.immowelt.de/liste/52355/wohnungen/mieten?ami=65&d=true&pma=750&r=20&sd=DESC&sf=TIMESTAMP&sp=1")
            browser.Navigate().GoToUrl("https://www.immowelt.de/liste/52355/wohnungen/mieten?ami=65&d=true&pma=750&r=20&sd=DESC&sf=TIMESTAMP&sp=1");
        else
            browser.Navigate().Refresh();

        await Task.Delay(10000);
        var test = browser.PageSource;
        try{browser.FindElementByClassName("bJrwSQ").Click();} catch(Exception e){Console.WriteLine(e);}
        browser.Manage().Window.Minimize();

        //Solve captcha
        if (browser.Title.Contains("Ich bin kein Roboter"))
        {
            Console.WriteLine("Captcha, trying to solve...");
            try
            {
                var captcha = browser.FindElementById("captcha-box");
                captcha.Click();
                await Task.Delay(5000);
            }
            catch { }
            if (browser.Title.Contains("Ich bin kein Roboter"))
            {
                Console.WriteLine("Couldn't solve captcha, please help\n");
                browser.Manage().Window.Maximize();
            }
            while (browser.Title.Contains("Ich bin kein Roboter"))
                await Task.Delay(1000);
            
            browser.Manage().Window.Minimize();
        }

        //Parse listings
        if (browser.Title.Contains("Wohnung mieten in"))
        {
            var results = browser.FindElementsByClassName("EstateItem-1c115").ToList();
            try{
                var exclude = browser.FindElementByClassName("recommendationsList").FindElements(OpenQA.Selenium.By.ClassName("EstateItem-1c115"));
                results = results.Except(exclude).ToList();
            } catch {}
            foreach (var result in results)
            {
                var debug = result.Text;
                var texts = result.Text.Replace("NEU", "").Replace("Grundriss", "").Replace("\r", "").Split("\n").ToList();
                var offset = texts[0] == "Neu" ? 0 : -1;
                var cost = double.Parse(texts[2 + offset].Split(" ").First().Replace(",", ".").Replace("~",""));
                var area = double.Parse(texts[3 + offset].Split(" ").First().Replace(",", "."));
                var rooms = double.Parse(texts[4 + offset].Split(" ").First().Replace(",", "."));
                var contact = texts[1 + offset];
                var location = texts[7 + offset];
                var description = texts[5 + offset];
                
                if (apartments.FirstOrDefault(x => x.Description == description && x.Location == location) is null)
                {
                    apartments.Add(new Apartment()
                    {
                        Area = area,
                        Contact = contact,
                        Cost = cost,
                        Url = ((OpenQA.Selenium.IWebElement)result.FindElement(OpenQA.Selenium.By.TagName("a"))).GetAttribute("href"),
                        Description = description,
                        isMale = contact.Contains("Herr"),
                        Location = location,
                        RoomCount = rooms
                    });

                    Console.WriteLine($"\nNew Apartment found!\n{apartments.Last().ToString()}\n");
                    Console.Beep();
                    RowChosen = apartments.Count - 1;
                    DrawApartmentTable();
                    saveApartments();
                }
            }
        }
    }
    catch(Exception e) {
        Console.WriteLine(e);
    }
}

async Task handleImmobilienscout24(int tab=0)
{
    try
    {
        //Switch to correct tab
        //while(browser.WindowHandles.Count - 1 < tab){
        //    browser.ExecuteScript("window.open()");
        //}
        //browser.SwitchTo().Window(browser.WindowHandles[tab]);

        //Go to website
        if(browser.Url != "https://www.immobilienscout24.de/Suche/de/nordrhein-westfalen/dueren-kreis/wohnung-mit-garage-mieten?price=-750.0&livingspace=65.0-&pricetype=calculatedtotalrent&sorting=2")
            browser.Navigate().GoToUrl("https://www.immobilienscout24.de/Suche/de/nordrhein-westfalen/dueren-kreis/wohnung-mit-garage-mieten?price=-750.0&livingspace=65.0-&pricetype=calculatedtotalrent&sorting=2");
        else
            browser.Navigate().Refresh();

        browser.Manage().Window.Minimize();

        await Task.Delay(10000);

        //Solve captcha
        if (browser.Title.Contains("Ich bin kein Roboter"))
        {
            Console.WriteLine("Captcha, trying to solve...");
            try
            {
                var captcha = browser.FindElementById("captcha-box");
                captcha.Click();
                await Task.Delay(5000);
            }
            catch { }
            if (browser.Title.Contains("Ich bin kein Roboter"))
            {
                Console.WriteLine("Couldn't solve captcha, please help\n");
                browser.Manage().Window.Maximize();
            }
            while (browser.Title.Contains("Ich bin kein Roboter"))
                await Task.Delay(1000);
            
            browser.Manage().Window.Minimize();
        }

        //Parse listings
        if (browser.Title.Contains("Wohnung mit Garage"))
        {
            var results = browser.FindElementsByClassName("result-list__listing");
            var exclude = browser.FindElementByClassName("recommendation-list").FindElements(OpenQA.Selenium.By.ClassName("result-list__listing"));
            foreach (var result in results.Except(exclude))
            {
                if (result.GetAttribute("data-id") is not null)
                {
                    var debug = result.Text;
                    var texts = result.Text.Replace("NEU", "").Replace("Grundriss", "").Replace("\r", "").Split("\n").ToList();
                    var cost = double.Parse(texts[texts.IndexOf("Warmmiete") - 1].Split(" ").First().Replace(",", ".").Replace("~",""));
                    var area = double.Parse(texts[texts.IndexOf("Wohnfläche") - 1].Split(" ").First().Replace(",", "."));
                    var rooms = double.Parse(texts[texts.IndexOf("Zi.") - 1].Split(" ").First().Replace(",", "."));
                    var contact = "";
                    try{
                        contact = texts[texts.IndexOf("Zi.") + 2];
                    } catch{}
                    var location = texts[texts.IndexOf("Warmmiete") - 2];
                    var description = texts[texts.IndexOf("Warmmiete") - 3];
                    if (apartments.FirstOrDefault(x => x.Description == description && x.Location == location) is null)
                    {
                        apartments.Add(new Apartment()
                        {
                            Area = area,
                            Contact = contact,
                            Cost = cost,
                            Url = $"https://www.immobilienscout24.de/expose/{ulong.Parse(result.GetAttribute("data-id"))}",
                            Description = description,
                            isMale = contact.Contains("Herr"),
                            Location = location,
                            RoomCount = rooms
                        });

                        Console.WriteLine($"\nNew Apartment found!\n{apartments.Last().ToString()}\n");
                        Console.Beep();
                        RowChosen = apartments.Count - 1;
                        DrawApartmentTable();
                        saveApartments();
                    }
                }
            }
        }
    }
    catch { }
}

void saveApartments()
{
    File.WriteAllText("apartments.json", JsonConvert.SerializeObject(apartments));
}

void loadApartments()
{
    if (File.Exists("apartments.json"))
        apartments = JsonConvert.DeserializeObject<List<Apartment>>(File.ReadAllText("apartments.json"));
}

async Task UserInput()
{
    //#if !DEBUG
    RowChosen = apartments.Count - 1;
    while (true)
    {
        try
        {
            while (!Console.KeyAvailable)
            {
                await Task.Delay(200);
            }
            var key = Console.ReadKey(true);

            switch (key.KeyChar)
            {
                case 'w':
                    if (RowChosen > 0) RowChosen--;
                    else RowChosen = apartments.Count - 1;
                    break;
                case 's':
                    if (RowChosen < apartments.Count - 1) RowChosen++;
                    else RowChosen = 0;
                    break;
                case 'm':
                    System.Diagnostics.Process.Start("cmd", $"/C start https://www.google.com/maps/dir/Valencienner+Str.+269,+52355+Düren/{apartments.OrderByDescending(x => x.FoundAt).ToList()[RowChosen].Location.Replace(" ", "+").Replace(",", "")}");
                    continue;
                case (char)System.ConsoleKey.Enter:
                    System.Diagnostics.Process.Start("cmd", $"/C start {apartments.OrderByDescending(x => x.FoundAt).ToList()[RowChosen].Url}");
                    continue;
                case (char)System.ConsoleKey.Backspace:
                    apartments.Remove(apartments.OrderByDescending(x => x.FoundAt).ToList()[RowChosen]);
                    saveApartments();
                    break;
                case '-':
                    apartments.OrderByDescending(x => x.FoundAt).ToList()[RowChosen].CurrentStatus = Apartment.Status.Ignored;
                    saveApartments();
                    break;
                case '+':
                    apartments.OrderByDescending(x => x.FoundAt).ToList()[RowChosen].CurrentStatus = Apartment.Status.Applied;
                    saveApartments();
                    break;
                default:
                    Console.WriteLine((int)key.KeyChar);
                    RowChosen = -1;
                    break;
            }
            DrawApartmentTable();
        }
        catch(Exception e)
        {
            Console.WriteLine(e);
        }
    }
    //#endif
}

void DrawApartmentTable(bool onlyMoved=false)
{
    //#if !DEBUG
    Console.Clear();
    //#endif
    Console.WriteLine(String.Format("Index | {0, -7} | {1, -7} | {2, -27} | {3, -100}", "Cost", "Area", "Date", "Location"));
    Console.WriteLine(new String('-', 100));

    var sortedApartments = apartments.OrderByDescending(x => x.FoundAt).ToList();
    for (int i = Math.Max(0, RowChosen - 9); i < Math.Min(sortedApartments.Count, Math.Max(10, RowChosen + 1)); i++)
    {
        if (sortedApartments[i].CurrentStatus == Apartment.Status.Applied)
            Console.ForegroundColor = ConsoleColor.Green;
        else if (sortedApartments[i].CurrentStatus == Apartment.Status.Ignored)
            Console.ForegroundColor = ConsoleColor.Red;

        if (i == RowChosen)
            Console.BackgroundColor = ConsoleColor.DarkGray;

        Console.WriteLine($"#{i + 1,-4} | {sortedApartments[i].Cost,-7} | {sortedApartments[i].Area,-7} | {sortedApartments[i].FoundAt,-23} UTC | {sortedApartments[i].Location}");

        Console.ResetColor();
    }
    Console.WriteLine(new String('-', 100));

    if (RowChosen > -1 && RowChosen < apartments.Count)
    {
        var chosen = sortedApartments[RowChosen];
        Console.WriteLine(chosen.ToString());
    }

    Console.WriteLine("\n[W] [S] to navigate between rows, [Enter] to visit posting, [M] to look at map.\n[-] To set ignored, [+] to set applied, [Backspace] to delete.");
}

public class Apartment
{
    public DateTime FoundAt = DateTime.UtcNow;
    public double Cost, RoomCount, Area;
    public string Contact, Description, Location, Url;
    public bool isMale;
    public Status CurrentStatus = Status.Undecided;

    public String CreateBewerbung(string source = "ImmobilienScout24")
    {
        string bewerbung = File.ReadAllText("Wohnungsbewerbung.txt");
        bewerbung = bewerbung.Replace("[CONTACT]", Contact.Contains("Herr") || Contact.Contains("Frau") ? $"Sehr geehrte{(isMale ? "r" : "")} {Contact}" : "Sehr geehrte Damen und Herren");
        bewerbung = bewerbung.Replace("[ROOMCOUNT]", RoomCount.ToString());
        bewerbung = bewerbung.Replace("[SOURCE]", source);

        return bewerbung;
    }

    public new string ToString()
    {
        return $"{new String('-', 100)}\nLocation: {Location}\nCost: {Cost}\nArea: {Area}\nRooms: {RoomCount}\nDescription: " +
               $"{Description}\nContact: {Contact}\nPossible application:\n{CreateBewerbung()}\n{new String('-', 100)}";
    }

    public enum Status{
        Undecided,
        Applied,
        Ignored
    }
}