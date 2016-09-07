using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using HtmlAgilityPack;

namespace OtoDom
{
    class Program
    {
        static string adres = @"http://dom.gratka.pl/";
        static List<string> links = new List<string>();
        static List<Ogloszenie> ogloszenia = new List<Ogloszenie>();


        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        static void Main(string[] args)
        {
            //var url = "https://otodom.pl/sprzedaz/mieszkanie/wroclaw/apartamentowiec--blok--plomba--loft--szeregowiec/?search%5Bfilter_float_price%3Ato%5D=370000&search%5Bfilter_float_m%3Afrom%5D=35&search%5Bfilter_enum_rooms_num%5D%5B0%5D=2&search%5Bfilter_enum_rooms_num%5D%5B1%5D=3&search%5Bfilter_enum_floor_no%5D%5B0%5D=floor_2&search%5Bfilter_enum_floor_no%5D%5B1%5D=floor_3&search%5Bfilter_enum_floor_no%5D%5B2%5D=floor_4&search%5Bfilter_float_building_floors_num%3Ato%5D=6&search%5Bdescription%5D=1&search%5Bdist%5D=0&search%5Bdistrict_id%5D=1&nrAdsPerPage=72";
            var url = "https://otodom.pl/sprzedaz/mieszkanie/wroclaw/apartamentowiec--blok--dom-wolnostojacy--plomba--loft--szeregowiec/?search%5Bfilter_float_price%3Ato%5D=370000&search%5Bfilter_float_m%3Afrom%5D=35&search%5Bfilter_enum_rooms_num%5D%5B0%5D=2&search%5Bfilter_enum_rooms_num%5D%5B1%5D=3&search%5Bfilter_enum_floor_no%5D%5B0%5D=floor_2&search%5Bfilter_enum_floor_no%5D%5B1%5D=floor_3&search%5Bfilter_enum_floor_no%5D%5B2%5D=floor_4&search%5Bfilter_float_building_floors_num%3Ato%5D=6&search%5Bdescription%5D=1&search%5Bdist%5D=0&search%5Bdistrict_id%5D=2&nrAdsPerPage=72";
            while (url != "")
            {
                var hw = new HtmlWeb();
                var htmlPage = hw.Load(url);
                PobierzOgloszenia(htmlPage);
                url = PobierzKolejnaStrone(htmlPage);
                Console.WriteLine("Zczytano " + links.Count);
            }

            foreach (var link in links)
            {
                var ogloszenie = PobierzInformacje(link);
                ogloszenia.Add(ogloszenie);
                Console.WriteLine("Linki " + ogloszenia.Count);
            }

            var wlasciwosciKluczy = ogloszenia.SelectMany(o => o.Properties.Keys.ToArray()).Distinct();
            ZapiszDoExcela(wlasciwosciKluczy);
        }

        private static void ZapiszDoExcela(IEnumerable<string> wlasciwosciKluczy)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            int x, y = 1;
            x = y;
            xlWorkSheet.Cells[y, x] = "URL";
            var lista = wlasciwosciKluczy.ToList();

            foreach (var wlasciwosc in wlasciwosciKluczy)
            {
                x++;
                xlWorkSheet.Cells[y, x] = wlasciwosc;
            }

            foreach (var ogloszenie in ogloszenia)
            {
                y++;
                xlWorkSheet.Cells[y, 1] = ogloszenie.Url + " ";
                foreach (var wlasciwosci in ogloszenie.Properties)
                {
                    x = lista.IndexOf(wlasciwosci.Key) + 2;
                    xlWorkSheet.Cells[y, x] = wlasciwosci.Value;
                }
            }


            xlWorkBook.SaveAs("d:\\SrodmiescieOtoDom.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

        }

        private static Ogloszenie PobierzInformacje(string url)
        {
            var ogloszenie = new Ogloszenie();
            ogloszenie.Url = url;
            var hw = new HtmlWeb();
            var htmlPage = hw.Load(url);
            try
            {
                PobierzMainInfo(htmlPage, ogloszenie);
            }
            catch
            { }
            try
            {
                PobierzSubListe(htmlPage, ogloszenie);
            }
            catch
            {
            }
            try
            {
                PobierzWlasciwosci(htmlPage, ogloszenie);
            }
            catch
            {
            }

            return ogloszenie;
        }

        private static void PobierzWlasciwosci(HtmlDocument htmlPage, Ogloszenie ogloszenie)
        {
            var mieszkanieInfo = htmlPage.DocumentNode.SelectNodes("//ul[@class='" + "params-list" + "']")[0];
            var info = mieszkanieInfo.ChildNodes.Where(c => c.Name == "li").Skip(1);
            foreach (var informacja in info)
            {
                var tekst = informacja.InnerText.Split('\n');
                var wartosci = tekst.Select(te => te.Trim()).Where(te => te != "").ToList();
                var klucz = wartosci[0];
                ogloszenie.Properties.Add(klucz, wartosci.Skip(1).Aggregate((i, j) => i + ", " + j)); 
            }
        }

        private static void PobierzSubListe(HtmlDocument htmlPage, Ogloszenie ogloszenie)
        {
            var mieszkanieInfo = htmlPage.DocumentNode.SelectNodes("//ul[@class='" + "sub-list" + "']")[0];
            var nodes = mieszkanieInfo.ChildNodes.Where(c => c.Name == "li");
            foreach (var node in nodes)
            {
                var info  = node.InnerText.Split(':');
                ogloszenie.Properties.Add(info[0].Trim(), info[1].Trim());
            }
        }

        private static void PobierzMainInfo(HtmlDocument htmlPage, Ogloszenie ogloszenie)
        {
            var mieszkanieInfo = htmlPage.DocumentNode.SelectNodes("//ul[@class='" + "main-list" + "']")[0];
            var info = mieszkanieInfo.InnerHtml.Split(new string[] { "<li>" }, StringSplitOptions.None);
            var dzielCena = info[1].Split(new string[] { "<span><strong>" }, StringSplitOptions.None);
            var dzielCenaWZl = dzielCena[1].Split('z');
            var dzielCenaZaMetr = dzielCenaWZl[1].Split(new string[] { "</span>" }, StringSplitOptions.None);
            ogloszenie.Properties.Add("Cena", dzielCenaWZl[0].Replace(" ", "").Trim());
            ogloszenie.Properties.Add("Średnia", dzielCenaZaMetr[1].Replace(" ", "").Replace(",", ".").Trim());

            var dzielPowierzchnia = info[2].Split(new string[] { "<span><strong>" }, StringSplitOptions.None);
            var dzielPowierzchniaMetry = dzielPowierzchnia[1].Split('m');
            ogloszenie.Properties.Add("Powierzchnia", dzielPowierzchniaMetry[0].Replace(" ", "").Replace(",", ".").Trim());

            var dzielLiczbaPokoi = info[3].Split(new string[] { "<span><strong>" }, StringSplitOptions.None);
            var liczbaPokoi = dzielLiczbaPokoi[1].Split('<');
            ogloszenie.Properties.Add("Liczba pokoi", liczbaPokoi[0].Replace(" ", "").Trim());

            var dzielPietro = info[4].Split(new string[] { "<span><strong>" }, StringSplitOptions.None);
            var Pietro = dzielPietro[1].Split('<');
            ogloszenie.Properties.Add("Piętro", Pietro[0].Replace(" ", "").Trim());

            var wysokoscBudynku = info[4].Split(new string[] { "<span><strong>" }, StringSplitOptions.None);
            var budynek = wysokoscBudynku[1].Split('<');
            var temp = budynek[1].Split(new string[] { ("(z ") }, StringSplitOptions.None);

            ogloszenie.Properties.Add("Liczba pieter", temp[1].Replace(")", "").Trim());
            
                


        }

        private static void PobierzWlasciwosci(HtmlDocument htmlPage, string klasa, Ogloszenie ogloszenie)
        {
            var mieszkanieInfo = htmlPage.DocumentNode.SelectNodes("//div[@class='" + klasa + "']")[0];
            var wlasnosci = mieszkanieInfo.ChildNodes[3].ChildNodes.Where(node => node.Name == "li");
            foreach (var wlasnosc in wlasnosci)
            {
                var paraWlasnosci = wlasnosc.InnerText.Split('\n');
                var klucz = paraWlasnosci[1].Trim();
                var wartoscKlucza = paraWlasnosci[2].Trim();
                if (klucz == "Cena")
                {
                    var ceny = wartoscKlucza.Split('z').ToList();
                    ceny[0] = ceny[0].Replace(" ", "");
                    ogloszenie.Properties.Add(klucz, ceny[0].Trim());

                    var nawias = ceny[1].IndexOf("(");
                    var srednia = ceny[1].Substring(nawias + 1).Trim().Replace(" ", "");
                    ogloszenie.Properties.Add("Średnia", srednia);
                }
                else if (klucz == "Powierzchnia")
                {
                    var powierzchnia = wartoscKlucza.Split('m').ToList();
                    powierzchnia[0] = powierzchnia[0].Replace(',', '.');
                    ogloszenie.Properties.Add(klucz, powierzchnia[0].Trim());
                }
                else
                    ogloszenie.Properties.Add(klucz, wartoscKlucza);
            }
        }

        public static void PobierzOgloszenia(HtmlDocument htmlPage)
        {
            var nodes = htmlPage.DocumentNode.SelectNodes("//span[starts-with(@class, 'offer-item-title')]");
            foreach (var node in nodes)
            {
                ;
                var getLink = node.ParentNode.ParentNode.OuterHtml.Split('"');
                var link = getLink[1];
                links.Add(link);
            }
        }

        public static string PobierzKolejnaStrone(HtmlDocument htmlPage)
        {
            var nastepnaStrona = htmlPage.DocumentNode.SelectNodes("//a[@data-dir='next']");
            if (nastepnaStrona == null)
                return "";
            var kolejnaStrony = nastepnaStrona[0].OuterHtml.Split('"');
            var nastepna = kolejnaStrony[1];
            return nastepna;
        }
    }
}
