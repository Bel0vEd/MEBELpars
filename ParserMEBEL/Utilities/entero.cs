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
    public static class entero
    {
        public delegate void ProgressCallBack();
        public delegate void ProgressMax(int maxValue);
        public static string ExcelName, garanti, name, artikul,nalichie, cena, opisanie, kategor, image, proizvoditel, kod, strana, gabarit, vec, moch, napr, osnxar, gruppa, inst;
        public static int ExcelStr = 0;
        public static void GetObj(string put,ProgressCallBack incCallBack, ProgressMax maximum,int sdelano, int stroka)
        {
            ExcelStr = stroka;
            string putimag = "", putinst = "";
            CookieContainer cookies = new CookieContainer();
            HttpWebRequest request1 = (HttpWebRequest)WebRequest.Create("http://www.entero.ru/vendors/");
            request1.CookieContainer = cookies;
            request1.Headers["Upgrade-Insecure-Requests"] = "1";
            request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
            request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
            HttpWebResponse response1 = (HttpWebResponse)request1.GetResponse();
            StreamReader streamReader = new StreamReader(response1.GetResponseStream());
            string result = streamReader.ReadToEnd();
            Regex regex = new Regex("(?<=<a href=/vendors)[\\w\\W]*?(?=>)");
            int k = 0;
            foreach (Match match in regex.Matches(result))
            {
                k++;
            }
            string[] allvend = new string[k];
            k = 0;
            foreach (Match match in regex.Matches(result))
            {
                allvend[k]=match.Value;
                k++;
            }
            maximum?.Invoke(allvend.Length);
            for (int chetvend = sdelano;chetvend <=allvend.Length;chetvend++)
            {
                request1 = (HttpWebRequest)WebRequest.Create("http://www.entero.ru/vendors"+allvend[chetvend]);
                request1.CookieContainer = cookies;
                request1.Headers["Upgrade-Insecure-Requests"] = "1";
                request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                response1 = (HttpWebResponse)request1.GetResponse();
                streamReader = new StreamReader(response1.GetResponseStream());
                result = streamReader.ReadToEnd();
                regex = new Regex("(?<=<a href=/list).*?(?=><img)");
                k = 0;
                foreach (Match match in regex.Matches(result))
                {
                    k++;
                }
                string [] allkatal = new string[k];
                k = 0;
                foreach (Match match in regex.Matches(result))
                {
                    allkatal[k] = match.Value;
                    k++;
                }
                for(int katalchet = 0;katalchet<allkatal.Length;katalchet++)
                {
                    request1 = (HttpWebRequest)WebRequest.Create("http://www.entero.ru/list" + allkatal[katalchet] + "&p=1");
                    request1.CookieContainer = cookies;
                    request1.Headers["Upgrade-Insecure-Requests"] = "1";
                    request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                    request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                    response1 = (HttpWebResponse)request1.GetResponse();
                    streamReader = new StreamReader(response1.GetResponseStream());
                    result = streamReader.ReadToEnd();
                    regex = new Regex("(?<=<span>)\\d.*?(?=</span>)");
                    int kolstr = 1;
                    foreach (Match match in regex.Matches(result))
                    {
                        kolstr++;
                    }
                    for (int strchet = 1; strchet <= kolstr; strchet++)
                    {
                        request1 = (HttpWebRequest)WebRequest.Create("http://www.entero.ru/list" + allkatal[katalchet] + "&p=" + strchet);
                        request1.CookieContainer = cookies;
                        request1.Headers["Upgrade-Insecure-Requests"] = "1";
                        request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                        request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                        response1 = (HttpWebResponse)request1.GetResponse();
                        streamReader = new StreamReader(response1.GetResponseStream());
                        result = streamReader.ReadToEnd();
                        regex = new Regex("(?<=m><a href=).*?(?= title=)");
                        k = 0;
                        foreach (Match match in regex.Matches(result))
                        {
                            k++;
                        }
                        string[] allobj = new string[k];
                        k = 0;
                        foreach (Match match in regex.Matches(result))
                        {
                            allobj[k] = match.Value;
                            k++;
                        }
                        for (int objchet = 0; objchet < allobj.Length; objchet++)
                        {
                            string silkobj = "http://www.entero.ru" + allobj[objchet];
                            request1 = (HttpWebRequest)WebRequest.Create(silkobj);
                            request1.CookieContainer = cookies;
                            request1.Headers["Upgrade-Insecure-Requests"] = "1";
                            request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                            request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                            response1 = (HttpWebResponse)request1.GetResponse();
                            streamReader = new StreamReader(response1.GetResponseStream());
                            result = streamReader.ReadToEnd();
                            name = Regex.Match(result, "(?<=<title>).*?(?=</title>)").Value;
                            nalichie = Regex.Match(result, "(?<=lowercase;'>).*?(?=</span>)").Value;
                            artikul = Regex.Match(result, "(?<=Код товара: <b>).*?(?=</b>)").Value;
                            cena = Regex.Match(result, "(?<=price\">).*?(?=</div>)").Value.Replace("&thinsp;"," ").Replace("&nbsp;", " ");
                            opisanie = Regex.Match(result, "(?<=>Описание)[\\w\\W]*?(?=</div></div>)").Value;
                            opisanie = Regex.Replace(opisanie, "<.*?>", string.Empty).Replace("  ", "").Replace("\n", "").Replace("&quot;", "\"").Replace("&deg;", "").Replace("\t", " ").Replace("&nbsp;", " ");
                            kategor = Regex.Match(result, "(?<=href=/categories/\\d*?>).*?(?=</a>)").Value;
                            regex = new Regex("(?<=href=/categories/\\d*?>).*?(?=</a>)");
                            k = 0;
                            foreach (Match match in regex.Matches(result))
                            {
                                k++;
                            }
                            string[] allgru = new string[k];
                            k = 0;
                            foreach (Match match in regex.Matches(result))
                            {
                                allgru[k] = match.Value;
                                k++;
                            }
                            string inst = Regex.Match(result, "(?<=pdf'><a href=).*(?=>Инструкция)").Value;
                            gruppa = allgru[k-1];
                            image = Regex.Match(result, "(?<=image src=//).*?(?= alt)").Value;
                            proizvoditel = Regex.Match(result, "(?<=Все товары ).*?(?=</a>)").Value;
                            kod = artikul;
                            strana = Regex.Match(result, "(?<=Страна-производитель<span></td><td class=value>).*?(?=</td>)").Value;
                            string shir = Regex.Match(result, "(?<=Ширина)<.*?(?=<td class=name>)").Value;
                            shir = Regex.Replace(shir, "<.*?>", string.Empty).Replace("&nbsp;", " ");
                            string vis = Regex.Match(result, "(?<=Высота)<.*?(?=<td class=name>)").Value;
                            vis = Regex.Replace(vis, "<.*?>", string.Empty).Replace("&nbsp;", " ");
                            if (vis != "" && shir != "")
                                gabarit = String.Concat(shir, " x ", vis);
                            else if (vis != "" && shir == "")
                                gabarit = vis;
                            else if (vis == "" && shir != "")
                                gabarit = shir;
                            string glub = Regex.Match(result, "(?<=Глубина)<.*?(?=<td class=name>)").Value;
                            glub = Regex.Replace(glub, "<.*?>", string.Empty).Replace("&nbsp;", " ");
                            if (gabarit != "" && glub != "")
                                gabarit = String.Concat(gabarit, " x ", glub);
                            else if (gabarit == "" && glub != "")
                                gabarit = glub;
                            vec = Regex.Match(result, "(?<=Вес ).*?(?=<td class=name>)").Value;
                            vec = Regex.Replace(vec, "<.*?>", string.Empty).Replace("&nbsp;", " ");
                            moch = Regex.Match(result, "(?<=Мощность).*?(?=<td class=name>)").Value;
                            moch = Regex.Replace(moch, "<.*?>", string.Empty).Replace("&nbsp;", " ");
                            napr = Regex.Match(result, "(?<=Напряжение).*?(?=<td class=name>)").Value;
                            if (napr == "")
                                napr = Regex.Match(result, "(?<=Питание).*?(?=<td class=name>)").Value;
                            if(napr == "")
                                napr = Regex.Match(result, "(?<=Подключение).*?(?=<td class=name>)").Value;
                            napr = Regex.Replace(napr, "<.*?>", string.Empty).Replace("&nbsp;", " ");
                            osnxar = Regex.Match(result, "(?<=<td class=name>).*?(?=</table>)").Value;
                            osnxar = Regex.Replace(osnxar, "<.*?>", " ").Replace("  ", " ").Replace("&nbsp;", "");
                            garanti = Regex.Match(result, "(?<=Гарантия:).*?(?=.</span>)").Value;
                            garanti = Regex.Replace(garanti, "<.*?>", "");
                            DirectoryInfo dirInfo = new DirectoryInfo(put + "\\" + kategor + "\\" +gruppa);
                            if (!dirInfo.Exists)
                            {
                                dirInfo.Create();
                            }
                            if (image != "")
                            {
                                request1 = (HttpWebRequest)WebRequest.Create("http://" + image);
                                request1.CookieContainer = cookies;
                                request1.Headers["Upgrade-Insecure-Requests"] = "1";
                                request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                                request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                                response1 = (HttpWebResponse)request1.GetResponse();
                                putimag = put + "\\" + kategor + "\\" + gruppa + "\\" + katalchet + strchet + objchet + ".jpg";
                                using (var image1 = File.OpenWrite(putimag))
                                {
                                    response1.GetResponseStream().CopyTo(image1);
                                }
                            }
                            if (inst != "")
                            {
                                request1 = (HttpWebRequest)WebRequest.Create("http://www.entero.ru" + inst);
                                request1.CookieContainer = cookies;
                                request1.Headers["Upgrade-Insecure-Requests"] = "1";
                                request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                                request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                                response1 = (HttpWebResponse)request1.GetResponse();
                                putinst = put + "\\" + kategor + "\\" + gruppa + "\\" + katalchet + strchet + objchet + ".pdf";
                                using (var image1 = File.OpenWrite(putinst))
                                {
                                    response1.GetResponseStream().CopyTo(image1);
                                }
                            }
                            Excel(chetvend, silkobj,name,artikul,cena,nalichie,opisanie,kategor,gruppa,putimag,proizvoditel,kod,strana,gabarit,vec,moch,napr,osnxar,putinst,garanti);
                        }
                    }
                }
                incCallBack?.Invoke();
            }
        }
        public static void Excel(int chetvend, string silka, string name, string artikul, string cena, string nalichie, string opisanie, string kategor, string gruppa, string image, string proizvoditel, string kod, string strana, string gabarit, string ves, string moch, string napr, string osnxar, string inst, string garanti)
        {
            ExcelPackage package = new ExcelPackage(new FileInfo(ExcelName));
            ExcelWorksheet sheet = package.Workbook.Worksheets[1];
            sheet.Cells[ExcelStr + 2, 1].Value = silka;
            sheet.Cells[ExcelStr + 2, 2].Value = name;
            sheet.Cells[ExcelStr + 2, 3].Value = artikul;
            sheet.Cells[ExcelStr + 2, 4].Value = cena;
            sheet.Cells[ExcelStr + 2, 5].Value = nalichie;
            sheet.Cells[ExcelStr + 2, 6].Value = opisanie;
            sheet.Cells[ExcelStr + 2, 7].Value = kategor;
            sheet.Cells[ExcelStr + 2, 8].Value = gruppa;
            sheet.Cells[ExcelStr + 2, 9].Value = image;
            //sheet.Cells[ExcelStr + 2, 10].Value = allsoptov;
            sheet.Cells[ExcelStr + 2, 11].Value = proizvoditel;
            sheet.Cells[ExcelStr + 2, 12].Value = artikul;
            sheet.Cells[ExcelStr + 2, 13].Value = strana;
            sheet.Cells[ExcelStr + 2, 14].Value = gabarit;
            sheet.Cells[ExcelStr + 2, 15].Value = ves;
            sheet.Cells[ExcelStr + 2, 16].Value = moch;
            sheet.Cells[ExcelStr + 2, 17].Value = napr;
            //sheet.Cells[ExcelStr + 2, 18].Value = dopopis;
            sheet.Cells[ExcelStr + 2, 19].Value = osnxar;
            //sheet.Cells[ExcelStr + 2, 20].Value = allzapchasti;
            sheet.Cells[ExcelStr + 2, 21].Value = inst;
            //sheet.Cells[ExcelStr + 2, 22].Value = allanalogi;
            sheet.Cells[ExcelStr + 2, 23].Value = garanti;
            package.Save();
            ExcelStr++;
            if (ExcelStr % 100 == 0)
            {
                string ExcelName2 = ExcelName.Replace(".xlsx", chetvend + ".xlsx");
                FileInfo package2 = new FileInfo(ExcelName);
                package2.CopyTo(ExcelName2, true);
            }
        }
        public static void getexc(string put,string backup)
        {
            if (ExcelName == null)
            {
                ExcelName = $"{put}\\{DateTime.Now.ToString("dd.MM.yyyy hh.mm.ss")}.xlsx";
            }
            if (backup != "")
            {
                FileInfo package2 = new FileInfo(backup);
                package2.CopyTo(ExcelName, true);
            }
            ExcelPackage package = new ExcelPackage(new FileInfo(ExcelName));
            if (package.Workbook.Worksheets.Count == 0)
            {
                package.Workbook.Worksheets.Add("entero");
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
            sheet.Cells[1, 23].Value = "Гарантия";
            package.Save();
        }
    }
}
