using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ParserMEBEL
{
    public static class klenmarket
    {
        public delegate void ProgressCallBack();
        public delegate void ProgressMax(int maxValue);
        public static string ExcelName, name, artikul, zapchasti, cena, opisanie, kategor, image, soptov, proizvoditel, kod, strana, gabarit, vec, moch, napr, osnxar, analog;
        public static int ExcelStr=0;
        public static CookieContainer cookies;
        public static void GetObj(string put, ProgressCallBack incCallBack, ProgressMax maximum, int sdelano)
        {
            HttpWebRequest request1 = (HttpWebRequest)WebRequest.Create("https://www.klenmarket.ru/shop/");
            request1.CookieContainer = cookies;
            request1.Headers["Upgrade-Insecure-Requests"] = "1";
            request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
            request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
            HttpWebResponse response1 = (HttpWebResponse)request1.GetResponse();
            StreamReader streamReader = new StreamReader(response1.GetResponseStream());
            string result = streamReader.ReadToEnd();
            result = Regex.Match(result, "(?<=<i>По категориям</i>)[\\w\\W]*?(?=<b>Услуги</b>)").Value;
            string lichnee = Regex.Match(result, "(?<=<i>По типам предприятий</i>)[\\w\\W]*?(?=<b>Посуда и столовые приборы</b>)").Value;
            result = result.Replace(lichnee, "");
            Regex regex = new Regex("(?<=a href=\")[\\w\\W]*?(?=\">)");
            int k = 0;
            foreach (Match match in regex.Matches(result))
            {
                k++;
            }
            string[] allkatalog = new string[k];
            int i = 0;
            foreach (Match match in regex.Matches(result))
            {
                allkatalog[i] = match.Value;
                i++;
            }
            maximum?.Invoke(allkatalog.Length);
            for (i = sdelano; i < allkatalog.Length; i++)
            {
                incCallBack?.Invoke();
                request1 = (HttpWebRequest)WebRequest.Create("https://www.klenmarket.ru"+allkatalog[i]);
                request1.CookieContainer = cookies;
                request1.Headers["Upgrade-Insecure-Requests"] = "1";
                request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                response1 = (HttpWebResponse)request1.GetResponse();
                streamReader = new StreamReader(response1.GetResponseStream());
                result = streamReader.ReadToEnd();
                if (result.Contains("class=\"shop-eq-list__cats-item"))
                {
                    regex = new Regex("(?<=list__cats-item\" href=\").*?(?=\">)");
                    k = 0;
                    foreach (Match match in regex.Matches(result))
                    {
                        k++;
                    }
                    string[] allgrupp = new string[k];
                    k = 0;
                    foreach (Match match in regex.Matches(result))
                    {
                        allgrupp[k] = match.Value;
                        k++;
                    }
                    for (k = 0; k < allgrupp.Length; k++)
                    {
                        try
                        {
                            for (int page = 1; ; page++)
                            {
                                request1 = (HttpWebRequest)WebRequest.Create("https://www.klenmarket.ru" + allgrupp[k] + "page-" + page);
                                request1.CookieContainer = cookies;
                                request1.Headers["Upgrade-Insecure-Requests"] = "1";
                                request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                                request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                                response1 = (HttpWebResponse)request1.GetResponse();
                                streamReader = new StreamReader(response1.GetResponseStream());
                                result = streamReader.ReadToEnd();
                                regex = new Regex("(?<=<a href=\"/shop).*?(?=\">\\s*<img)");
                                int kolobj = 0, c = 0;
                                foreach (Match match in regex.Matches(result))
                                {
                                    kolobj++;
                                }
                                string[] allobj = new string[kolobj];
                                foreach (Match match in regex.Matches(result))
                                {
                                    allobj[c] = match.Value;
                                    c++;
                                }
                                for (int z = 0; z < allobj.Length; z++)
                                {
                                    parsobj(put, allobj[z], k, page, z);
                                }
                            }
                        }
                        catch (WebException)
                        {
                            continue;
                        }
                    }
                }
                else
                {
                    try
                    {
                        for (int page = 1; ; page++)
                        {
                            request1 = (HttpWebRequest)WebRequest.Create("https://www.klenmarket.ru" + allkatalog[i] + "page-" + page);
                            request1.CookieContainer = cookies;
                            request1.Headers["Upgrade-Insecure-Requests"] = "1";
                            request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                            request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                            response1 = (HttpWebResponse)request1.GetResponse();
                            streamReader = new StreamReader(response1.GetResponseStream());
                            result = streamReader.ReadToEnd();
                            regex = new Regex("(?<=<a href=\"/shop).*?(?=\">\\s*<img)");
                            int kolobj = 0, c = 0;
                            foreach (Match match in regex.Matches(result))
                            {
                                kolobj++;
                            }
                            string[] allobj = new string[kolobj];
                            foreach (Match match in regex.Matches(result))
                            {
                                allobj[c] = match.Value;
                                c++;
                            }
                            for (int z = 0; z < allobj.Length; z++)
                            {
                                parsobj(put, allobj[z], k,page, z);
                            }
                        }
                    }
                    catch (WebException)
                    {
                        continue;
                    }
                }
            }
        }
        public static void Excel(string silka, string name, string artikul, string cena, string opisanie, string kategor, string gruppa, string image, string allsoptov, string proizvoditel, string kod, string strana, string gabarit, string ves, string moch, string napr, string osnxar, string allzapchasti, string allanalogi)
        {
            ExcelPackage package = new ExcelPackage(new FileInfo(ExcelName));
            ExcelWorksheet sheet = package.Workbook.Worksheets[1];
            sheet.Cells[ExcelStr + 2, 1].Value = silka;
            sheet.Cells[ExcelStr + 2, 2].Value = name;
            sheet.Cells[ExcelStr + 2, 3].Value = artikul;
            sheet.Cells[ExcelStr + 2, 4].Value = cena;
            //sheet.Cells[ExcelStr + 2, 5].Value = nalichie;
            sheet.Cells[ExcelStr + 2, 6].Value = opisanie;
            sheet.Cells[ExcelStr + 2, 7].Value = kategor;
            sheet.Cells[ExcelStr + 2, 8].Value = gruppa;
            sheet.Cells[ExcelStr + 2, 9].Value = image;
            sheet.Cells[ExcelStr + 2, 10].Value = allsoptov;
            sheet.Cells[ExcelStr + 2, 11].Value = proizvoditel;
            sheet.Cells[ExcelStr + 2, 12].Value = artikul;
            sheet.Cells[ExcelStr + 2, 13].Value = strana;
            sheet.Cells[ExcelStr + 2, 14].Value = gabarit;
            sheet.Cells[ExcelStr + 2, 15].Value = ves;
            sheet.Cells[ExcelStr + 2, 16].Value = moch;
            sheet.Cells[ExcelStr + 2, 17].Value = napr;
            //sheet.Cells[ExcelStr + 2, 18].Value = dopopis;
            sheet.Cells[ExcelStr + 2, 19].Value = osnxar;
            sheet.Cells[ExcelStr + 2, 20].Value = allzapchasti;
            //sheet.Cells[ExcelStr + 2, 21].Value = inst;
            sheet.Cells[ExcelStr + 2, 22].Value = allanalogi;
            package.Save();
            ExcelStr++;
        }
        public static void getexc(string put)
        {
            if (ExcelName == null)
            {
                ExcelName = $"{put}\\{DateTime.Now.ToString("dd.MM.yyyy hh.mm.ss")}.xlsx";
            }
            ExcelPackage package = new ExcelPackage(new FileInfo(ExcelName));
            if (package.Workbook.Worksheets.Count == 0)
            {
                package.Workbook.Worksheets.Add("klenmarket");
            }
            ExcelWorksheet sheet = package.Workbook.Worksheets[1];
            sheet.Cells[1, 1].Value = "Ссылка на страницу";
            sheet.Cells[1, 2].Value = "Название";
            sheet.Cells[1, 3].Value = "Артикул";
            sheet.Cells[1, 4].Value = "Цена";
            sheet.Cells[1, 5].Value = "Наличие";
            sheet.Cells[1, 6].Value = "Описание";
            sheet.Cells[1, 7].Value = "Категория";
            sheet.Cells[1, 8].Value = "Группа";
            sheet.Cells[1, 9].Value = "Изображение";
            sheet.Cells[1, 10].Value = "Сопутствующие товары";
            sheet.Cells[1, 11].Value = "Производитель";
            sheet.Cells[1, 12].Value = "Код товара";
            sheet.Cells[1, 13].Value = "Страна";
            sheet.Cells[1, 14].Value = "Габариты (нетто)";
            sheet.Cells[1, 15].Value = "Вес (брутто)";
            sheet.Cells[1, 16].Value = "Мощность";
            sheet.Cells[1, 17].Value = "Напряжение ";
            sheet.Cells[1, 18].Value = "Дополнительное описание";
            sheet.Cells[1, 19].Value = "Основные характеристики";
            sheet.Cells[1, 20].Value = "Запчасти и комплектующие";
            sheet.Cells[1, 21].Value = "Инструкция";
            sheet.Cells[1, 22].Value = "Аналоги";
            package.Save();
        }
        public static void parsobj(string put,string sil,int k,int page,int z)
        {
            string putimag = "", gruppa = "";
            string silkaobj = "https://www.klenmarket.ru/shop" + sil;
            HttpWebRequest request1 = (HttpWebRequest)WebRequest.Create("https://www.klenmarket.ru/shop" + sil);
            request1.CookieContainer = cookies;
            request1.Headers["Upgrade-Insecure-Requests"] = "1";
            request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
            request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
            HttpWebResponse response1 = (HttpWebResponse)request1.GetResponse();
            StreamReader streamReader = new StreamReader(response1.GetResponseStream());
            string result = streamReader.ReadToEnd();
            string result1 = result;
            name = Regex.Match(result, "(?<=title=\").*?(?=\" alt=)").Value;
            artikul = Regex.Match(result, "(?<=Артикул:)[\\w\\W]*?(?=</span>)").Value;
            artikul = Regex.Replace(artikul, "<.*?>", string.Empty).Replace("\n", "").Replace(" ", "");
            cena = Regex.Match(result, "(?<=\"price\" content=\")[\\w\\W]*?(?=\">)").Value;
            opisanie = Regex.Match(result, "(?<=<h2>Описание.*?</h2>)[\\w\\W]*?(?=<div class=\"shop)").Value;
            opisanie = Regex.Replace(opisanie, "<.*?>", string.Empty).Replace("&nbsp;", " ").Replace("&laquo;", "(").Replace("&raquo;", ")").Replace("&deg;", "").Replace("&mdash;", "-").Replace("\r\n", " ");
            kategor = Regex.Match(result, "(?<=name\">).*?(?=<meta)").Value;
            kategor = Regex.Replace(kategor, "<.*?>", string.Empty).Replace("  ", "");
            Regex regex = new Regex("(?<=name\">).*?(?=</span></a>)");
            int gr = 0;
            foreach (Match match in regex.Matches(result))
            {
                gr++;
            }
            string[] grupmas = new string[gr];
            gr = 0;
            foreach (Match match in regex.Matches(result))
            {
                grupmas[gr] = match.Value;
                gr++;
            }
            for (int x = 2; x < gr; x++)
            {
                gruppa = String.Concat(gruppa, "\\" + grupmas[x]);
            }
            DirectoryInfo dirInfo = new DirectoryInfo(put + "\\" + kategor + gruppa);
            if (!dirInfo.Exists)
            {
                dirInfo.Create();
            }
            image = Regex.Match(result, "(?<=href=\").*?(?=\" rel)").Value;
            if (image != "")
            {
                request1 = (HttpWebRequest)WebRequest.Create("https://www.klenmarket.ru" + image);
                request1.CookieContainer = cookies;
                request1.Headers["Upgrade-Insecure-Requests"] = "1";
                request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                response1 = (HttpWebResponse)request1.GetResponse();
                putimag = put + "\\" + kategor + gruppa + "\\" + k + page + z + ".jpg";
                using (var image1 = File.OpenWrite(putimag))
                {
                    response1.GetResponseStream().CopyTo(image1);
                }
            }
            proizvoditel = Regex.Match(result, "(?<=Производитель:)[\\w\\W]*?(?=</a>)").Value;
            proizvoditel = Regex.Replace(proizvoditel, "<.*?>", string.Empty).Replace("  ", "").Replace("\n", "");
            kod = artikul;
            strana = Regex.Match(result, "(?<=country\">)[\\w\\W]*?(?=</span>)").Value;
            strana = Regex.Replace(strana, "<.*?>", string.Empty).Replace("  ", "").Replace("\n", "");
            gabarit = Regex.Match(result, "(?<=Габаритные размеры:</span>)[\\w\\W]*?(?=</span>)").Value;
            gabarit = Regex.Replace(gabarit, "<.*?>", string.Empty).Replace("  ", "").Replace("\n", "");
            vec = Regex.Match(result, "(?<=Вес:</span>)[\\w\\W]*?(?=</span>)").Value;
            vec = Regex.Replace(vec, "<.*?>", string.Empty).Replace("  ", "").Replace("\n", "");
            if (vec == "")
            {
                vec = Regex.Match(result, "(?<=Масса:</span>)[\\w\\W]*?(?=</span>)").Value;
                vec = Regex.Replace(vec, "<.*?>", string.Empty).Replace("  ", "").Replace("\n", "");
            }
            moch = Regex.Match(result, "(?<=ощность:</span>)[\\w\\W]*?(?=</span>)").Value;
            moch = Regex.Replace(moch, "<.*?>", string.Empty).Replace("  ", "").Replace("\n", "");
            napr = Regex.Match(result, "(?<=апряжение:</span>)[\\w\\W]*?(?=</span>)").Value;
            napr = Regex.Replace(napr, "<.*?>", string.Empty).Replace("  ", "").Replace("\n", "");
            osnxar = Regex.Match(result, "(?<=text__specs\">)[\\w\\W]*?(?=<div class=\"text__collapsed\">)").Value;
            osnxar = Regex.Replace(osnxar, "<.*?>", string.Empty).Replace("  ", "").Replace("\n", " ");
            analog = Regex.Match(result1, "(?<=Похожие товары)[\\w\\W]*?(?=<div class=\"modals-list\">)").Value;
            regex = new Regex("(?<=item-title\">\\s*?<a href=\")[\\w\\W]*?(?=\">)");
            int chetanal = 0;
            foreach (Match match in regex.Matches(analog))
            {
                chetanal++;
            }
            string[] grupanal = new string[chetanal];
            chetanal = 0;
            foreach (Match match in regex.Matches(analog))
            {
                grupanal[chetanal] = match.Value;
                chetanal++;
            }
            analog = "";
            for (chetanal = 0; chetanal < grupanal.Length; chetanal++)
            {
                request1 = (HttpWebRequest)WebRequest.Create("https://www.klenmarket.ru" + grupanal[chetanal]);
                request1.CookieContainer = cookies;
                request1.Headers["Upgrade-Insecure-Requests"] = "1";
                request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                response1 = (HttpWebResponse)request1.GetResponse();
                streamReader = new StreamReader(response1.GetResponseStream());
                result = streamReader.ReadToEnd();
                string artikulzap = Regex.Match(result, "(?<=Артикул:)[\\w\\W]*?(?=</span>)").Value;
                artikulzap = Regex.Replace(artikulzap, "<.*?>", string.Empty).Replace("\n", "").Replace(" ", "");
                analog = String.Concat(analog, artikulzap, ",");
            }
            zapchasti = Regex.Match(result1, "(?<=data-title=\"Запчасти\">)[\\w\\W]*?(?=<div class=\"qa__form\">)").Value;
            regex = new Regex("(?<=item-title\">\\s*?<a href=\")[\\w\\W]*?(?=\">)");
            int chetzap = 0;
            foreach (Match match in regex.Matches(zapchasti))
            {
                chetzap++;
            }
            string[] grupzap = new string[chetzap];
            chetzap = 0;
            foreach (Match match in regex.Matches(zapchasti))
            {
                grupzap[chetzap] = match.Value;
                chetzap++;
            }
            zapchasti = "";
            for (chetzap=0;chetzap<grupzap.Length;chetzap++)
            {
                request1 = (HttpWebRequest)WebRequest.Create("https://www.klenmarket.ru" + grupzap[chetzap]);
                request1.CookieContainer = cookies;
                request1.Headers["Upgrade-Insecure-Requests"] = "1";
                request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                response1 = (HttpWebResponse)request1.GetResponse();
                streamReader = new StreamReader(response1.GetResponseStream());
                result = streamReader.ReadToEnd();
                string artikulzap = Regex.Match(result, "(?<=Артикул:)[\\w\\W]*?(?=</span>)").Value;
                artikulzap = Regex.Replace(artikulzap, "<.*?>", string.Empty).Replace("\n", "").Replace(" ", "");
                zapchasti = String.Concat(zapchasti,artikulzap,",");
            }
            soptov = Regex.Match(result1, "(?<=Необходимые аксессуары)[\\w\\W]*?(?=Описание)").Value;
            regex = new Regex("(?<=item-title\">\\s*?<a href=\"/)[\\w\\W]*?(?=\">)");
            int chetsop = 0;
            foreach (Match match in regex.Matches(soptov))
            {
                chetsop++;
            }
            string[] grupsop = new string[chetsop];
            chetsop = 0;
            foreach (Match match in regex.Matches(soptov))
            {
                grupsop[chetsop] = match.Value;
                chetsop++;
            }
            soptov = "";
            for (chetsop = 0; chetsop < grupsop.Length; chetsop++)
            {
                request1 = (HttpWebRequest)WebRequest.Create("https://www.klenmarket.ru/" + grupsop[chetsop]);
                request1.CookieContainer = cookies;
                request1.Headers["Upgrade-Insecure-Requests"] = "1";
                request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                response1 = (HttpWebResponse)request1.GetResponse();
                streamReader = new StreamReader(response1.GetResponseStream());
                result = streamReader.ReadToEnd();
                string artikulsop = Regex.Match(result, "(?<=Артикул:)[\\w\\W]*?(?=</span>)").Value;
                artikulsop = Regex.Replace(artikulsop, "<.*?>", string.Empty).Replace("\n", "").Replace(" ", "");
                soptov = String.Concat(soptov, artikulsop, ",");
            }
            Excel(silkaobj,name,artikul,cena,opisanie,kategor,gruppa,putimag,zapchasti,proizvoditel,kod,strana,gabarit,vec,moch,napr,osnxar,soptov,analog);
        }
    }
}
