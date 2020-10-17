using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace ParserMEBEL
{
    public static class rp
    {
        public delegate void ProgressCallBack();
        public delegate void ProgressMax(int maxValue);
        public static string ExcelName;
        public static int ExcelStr=0;
        public static void GetObj(string put, ProgressCallBack incCallBack, ProgressMax maximum, int sdelano, string prox, int stroka)
        {
            ExcelStr = stroka;
            string result2 = "";
            int inet = 0;
            int i = 0;
            int proxchet1 = 0;
            int objectchet = 0;
            string putimag = "", putinst = "";
            int allstr = 2, str = 1;
            StreamReader objReader1 = new StreamReader(prox);
            string sLine1 = "";
            ArrayList proxy = new ArrayList();
            while (sLine1 != null)
            {
                sLine1 = objReader1.ReadLine();
                if (sLine1 != null)
                    proxy.Add(sLine1);
            }
            objReader1.Close();
            CookieContainer cookies = new CookieContainer();
            for (int first=0; first<proxy.Count; first++)
            {
                try
                {
                    HttpWebRequest request2 = (HttpWebRequest)WebRequest.Create("http://www.rp.ru/shop/?set_filter=Y&query=&q=&manufacturer=0&type=0&count=100&XML_ID=&section=0&sort=0&show_type=2&instore=0&PAGEN_1=" + str);
                    request2.CookieContainer = cookies;
                    request2.Proxy = new WebProxy(proxy[first].ToString());
                    request2.Headers["Upgrade-Insecure-Requests"] = "1";
                    request2.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                    request2.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                    HttpWebResponse response2 = (HttpWebResponse)request2.GetResponse();
                    StreamReader streamReader2 = new StreamReader(response2.GetResponseStream());
                    result2 = streamReader2.ReadToEnd();
                    break;
                }
                catch(WebException)
                {
                    continue;
                }
            }
            Regex regex = new Regex("(?<=PAGEN_1=).*(?=\">...</a>)");
            Match match1 = regex.Match(result2);
            while (match1.Success)
            {
                allstr = Int32.Parse(match1.Value);
                match1 = match1.NextMatch();

            }
            maximum?.Invoke(allstr);
            for (str = sdelano+1; str < allstr; str++)
            {
                if (inet > 400)
                    break;
                for (int proxchet = proxchet1; proxchet < proxy.Count;proxchet++)
                {
                    try
                    {
                        HttpWebRequest request1 = (HttpWebRequest)WebRequest.Create("http://www.rp.ru/shop/?set_filter=Y&query=&q=&manufacturer=0&type=0&count=100&XML_ID=&section=0&sort=0&show_type=2&instore=0&PAGEN_1=" + str);
                        request1.CookieContainer = cookies;
                        request1.Proxy = new WebProxy(proxy[proxchet].ToString());
                        request1.Headers["Upgrade-Insecure-Requests"] = "1";
                        request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                        request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                        HttpWebResponse response1 = (HttpWebResponse)request1.GetResponse();
                        StreamReader streamReader = new StreamReader(response1.GetResponseStream());
                        string result = streamReader.ReadToEnd();
                        regex = new Regex("(?<=<td><a href=\").*(?=\" >)");
                        int k = 0;
                        foreach (Match match in regex.Matches(result))
                        {
                            k++;
                        }
                        string[] obiekt = new string[k];
                        regex = new Regex("(?<=<td><a href=\").*(?=\" >)");
                        i = 0;
                        foreach (Match match in regex.Matches(result))
                        {
                            obiekt[i] = match.Value;
                            i++;
                        }
                        if (i == 0)
                        {
                            proxchet1++;
                            continue;
                        }
                        for (i = objectchet; i < obiekt.Length; i++)
                        {
                            string allanalogi = "";
                            string allzapchasti = "";
                            string allsoptov = "";
                            string silkaobj = "http://www.rp.ru" + obiekt[i];
                            request1 = (HttpWebRequest)WebRequest.Create(silkaobj);
                            request1.CookieContainer = cookies;
                            request1.Proxy = new WebProxy(proxy[proxchet].ToString());
                            request1.Headers["Upgrade-Insecure-Requests"] = "1";
                            request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                            request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                            response1 = (HttpWebResponse)request1.GetResponse();
                            streamReader = new StreamReader(response1.GetResponseStream());
                            result = streamReader.ReadToEnd();
                            string name = Regex.Match(result, "(?<=<title>).*(?=</title>)").Value;
                            string artikul = Regex.Match(result, "(?<=/images/).*(?=_o.jpg\" rel)").Value;
                            if (artikul=="")
                                artikul = Regex.Match(silkaobj, "(?<=\\d/).*?(?=/)").Value;
                            string cena = Regex.Match(result, "(?<=<br /> ).*(?=<br/>)").Value;
                            string nalichie = Regex.Match(result, "(?<=\\s).*?(?=</b></span>)").Value;
                            nalichie = nalichie.Replace(" ", "");
                            string opisanie = Regex.Match(result, "(?<=товара</b>)[\\w\\W]*?(?=<a href=)").Value;
                            if (opisanie == "")
                            {
                                opisanie = Regex.Match(result, "(?<=Описание</b></span>)[\\w\\W]*?(?=<img)").Value;
                                if (opisanie == "")
                                {
                                    opisanie = Regex.Match(result, "(?<=justify\">)[\\w\\W]*?(?=<iframe)").Value;
                                    if (opisanie == "")
                                    {
                                        opisanie = Regex.Match(result, "(?<=tooltip\\(\\);)[\\w\\W]*?(?=<span class=\"techlabel\")").Value;
                                        if (opisanie == "")
                                        {
                                            opisanie = Regex.Match(result, "(?<=text-align: justify;\">)[\\w\\W]*?(?=<p style=)").Value;
                                        }
                                    }
                                }
                            }
                            opisanie = opisanie.Replace("&gt;", " ").Replace("\r", "").Replace("\n", "").Replace("\t", "").Replace("</b>", "").Replace("&nbsp;", "").Replace("<span style=\"color:#ff0000;\"", "").Replace("<b>", "").Replace("</span", "").Replace("</font>", "").Replace("</div>", "").Replace("<br />", "").Replace("});", "").Replace("   ", "").Replace("&mdash;", "");
                            /*while (opisanie.Contains("<"))
                            {
                                opisanie = opisanie.Remove(opisanie.IndexOf("<"), opisanie.IndexOf(">", opisanie.IndexOf("<")) - opisanie.IndexOf("<"));
                            }*/
                            opisanie = Regex.Replace(opisanie, "<.*?>", string.Empty);
                            string kategor = Regex.Match(result, "(?<=<a href=\"/shop/\\?set_filter=Y&level1=).*?(?=&level2=)").Value;
                            string gruppa = Regex.Match(result, "(?<=<a href=\"/shop/\\?set_filter=Y&level1=" + kategor + "&level2=).*?(?=&level3=)").Value;
                            string image = Regex.Match(result, "(?<=<a href=\").*?(?=.jpg)").Value;
                            string soptovari = Regex.Match(result, "(?<=techlabel\">Сопутствующие товары:)[\\w\\W]*?(?=\"techlabel\">)").Value;
                            if(soptovari=="")
                                soptovari = Regex.Match(result, "(?<=techlabel\">Сопутствующие товары:)[\\w\\W]*?(?=</table>)").Value;
                            regex = new Regex("(?<=\\d/).*?(?=/\">)");
                            foreach (Match match in regex.Matches(soptovari))
                            {
                                allsoptov = String.Concat(allsoptov, match.Value + ",");
                            }
                            string proizvoditel = Regex.Match(result, "(?<=/'>).*?(?=</A>)").Value;
                            string strana = Regex.Match(result, "(?<=Страна:</div></td>)[\\w\\W]*?(?=</td>)").Value;
                            strana = strana.Replace(" ", "").Replace("<td>", "").Replace("\n", "").Replace("\t", "");
                            string gabarit = Regex.Match(result, "(?<=\\(нетто\\):</div></td>)[\\w\\W]*?(?=</td>)").Value;
                            gabarit = gabarit.Replace(" ", "").Replace("<td>", "").Replace("\n", "").Replace("\t", "");
                            string ves = Regex.Match(result, "(?<=\\(брутто\\):</div></td>)[\\w\\W]*?(?=</td>)").Value;
                            ves = ves.Replace(" ", "").Replace("<td>", "").Replace("\n", "").Replace("\t", "");
                            string moch = Regex.Match(result, "(?<=ость:</div></td>)[\\w\\W]*?(?=</td>)").Value;
                            moch = moch.Replace(" ", "").Replace("<td>", "").Replace("\n", "").Replace("\t", "");
                            string napr = Regex.Match(result, "(?<=жение:</div></td>)[\\w\\W]*?(?=</td>)").Value;
                            napr = napr.Replace(" ", "").Replace("<td>", "").Replace("\n", "").Replace("\t", "");
                            string dopopisanie = Regex.Match(result, "(?<=Описание:</div></td>)[\\w\\W]*?(?=</td>)").Value;
                            if (dopopisanie == "")
                            {
                                dopopisanie = Regex.Match(result, "(?<=Описание:</div>)[\\w\\W]*?(?=</tr>)").Value;
                            }
                            dopopisanie = dopopisanie.Replace("<td>", "").Replace("\n", "").Replace("\t", "").Replace("&nbsp;", "").Replace("&#40", "").Replace("&#41", "");
                            string osnxar = Regex.Match(result, "()").Value;
                            string zapchasti = Regex.Match(result, "(?<=\"techlabel\">Запчасти и комплектующие:)[\\w\\W]*?(?=</table>)").Value;
                            if(zapchasti=="")
                                zapchasti = Regex.Match(result, "(?<=\"techlabel\">Этот товар Запчасть и комплектующие для:)[\\w\\W]*?(?=</table>)").Value;
                            regex = new Regex("(?<=\\d/).*?(?=/\">)");
                            foreach (Match match in regex.Matches(zapchasti))
                            {
                                allzapchasti = String.Concat(allzapchasti, match.Value + ",");
                            }
                            string allsopdlya = "";
                            string sopdlya = Regex.Match(result, "(?<=Этот товар сопутствующий для:)[\\w\\W]*?(?=</table>)").Value;
                            regex = new Regex("(?<=\\d/).*?(?=/\">)");
                            foreach (Match match in regex.Matches(sopdlya))
                            {
                                allsopdlya = String.Concat(allsopdlya, match.Value + ",");
                            }
                            string instruction = Regex.Match(result, "(?<=href=\").*?(?=_inst.pdf)").Value;
                            string analogi = Regex.Match(result, "(?<=\"techlabel\">Этот товар Аналог)[\\w\\W]*?(?=</table>)").Value;
                            if(analogi=="")
                                analogi = Regex.Match(result, "(?<=\"techlabel\">Аналоги:)[\\w\\W]*?(?=</table>)").Value;
                            regex = new Regex("(?<=\\d/).*?(?=/\">)");
                            foreach (Match match in regex.Matches(analogi))
                            {
                                allanalogi = String.Concat(allanalogi, match.Value + ",");
                            }
                            DirectoryInfo dirInfo = new DirectoryInfo(put + "\\" + kategor + "\\" + gruppa + "\\");
                            if (!dirInfo.Exists)
                            {
                                dirInfo.Create();
                            }
                            if (instruction != "")
                            {
                                request1 = (HttpWebRequest)WebRequest.Create("http://www.rp.ru" + instruction + "_inst.pdf");
                                request1.CookieContainer = cookies;
                                request1.Proxy = new WebProxy(proxy[proxchet].ToString());
                                request1.Headers["Upgrade-Insecure-Requests"] = "1";
                                request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                                request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                                response1 = (HttpWebResponse)request1.GetResponse();
                                putinst = put + "\\" + kategor + "\\" + gruppa + "\\" + str + i + ".pdf";
                                using (var inst = File.OpenWrite(putinst))
                                {
                                    response1.GetResponseStream().CopyTo(inst);
                                }
                            }
                            if (image != "")
                            {
                                request1 = (HttpWebRequest)WebRequest.Create("http://www.rp.ru" + image + ".jpg");
                                request1.CookieContainer = cookies;
                                request1.Proxy = new WebProxy(proxy[proxchet].ToString());
                                request1.Headers["Upgrade-Insecure-Requests"] = "1";
                                request1.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8";
                                request1.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36";
                                response1 = (HttpWebResponse)request1.GetResponse();
                                putimag = put + "\\" + kategor + "\\" + gruppa + "\\" + str + i + ".jpg";
                                using (var image1 = File.OpenWrite(putimag))
                                {
                                    response1.GetResponseStream().CopyTo(image1);
                                }
                            }
                            Excel(str, put, silkaobj, name, artikul, cena, nalichie, opisanie, kategor, gruppa, putimag, allsoptov, proizvoditel, artikul, strana, gabarit, ves, moch, napr, dopopisanie, osnxar, allzapchasti, putinst, allanalogi, allsopdlya);
                        }
                        incCallBack?.Invoke();
                        inet = 0;
                        if (proxchet == proxy.Count-1)
                            proxchet1 = 0;
                        objectchet = 0;
                        proxchet1++;
                        break;
                    }
                    catch (WebException)
                    {
                        if (proxchet == proxy.Count-1)
                            proxchet1 = 0;
                        objectchet = i;
                        proxchet1++;
                        inet++;
                        continue;
                    }
                    catch (IOException)
                    {
                        if (proxchet == proxy.Count - 1)
                            proxchet1 = 0;
                        objectchet = i;
                        proxchet1++;
                        inet++;
                        continue;
                    }
                }
            }
        }
        public static void Excel(int str, string put,string silka,string name,string artikul,string cena,string nalichie,string opisanie,string kategor,string gruppa,string image,string allsoptov,string proizvoditel,string kod, string strana,string gabarit, string ves, string moch,string napr,string dopopis, string osnxar,string allzapchasti,string inst,string allanalogi,string allsopdlya)
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
            sheet.Cells[ExcelStr + 2, 10].Value = allsoptov;
            sheet.Cells[ExcelStr + 2, 11].Value = proizvoditel;
            sheet.Cells[ExcelStr + 2, 12].Value = artikul;
            sheet.Cells[ExcelStr + 2, 13].Value = strana;
            sheet.Cells[ExcelStr + 2, 14].Value = gabarit;
            sheet.Cells[ExcelStr + 2, 15].Value = ves;
            sheet.Cells[ExcelStr + 2, 16].Value = moch;
            sheet.Cells[ExcelStr + 2, 17].Value = napr;
            sheet.Cells[ExcelStr + 2, 18].Value = dopopis;
            sheet.Cells[ExcelStr + 2, 19].Value = osnxar;
            sheet.Cells[ExcelStr + 2, 20].Value = allzapchasti;
            sheet.Cells[ExcelStr + 2, 21].Value = inst;
            sheet.Cells[ExcelStr + 2, 22].Value = allanalogi;
            sheet.Cells[ExcelStr + 2, 23].Value = allsopdlya;
            package.Save();
            ExcelStr++;
            if (ExcelStr % 100 == 0)
            {
                string ExcelName2 = ExcelName.Replace(".xlsx", str + ".xlsx");
                FileInfo package2 = new FileInfo(ExcelName);
                package2.CopyTo(ExcelName2, true);
            }
        }
        public static void getexc(string put, string backup)
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
                package.Workbook.Worksheets.Add("rp");
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
            sheet.Cells[1, 23].Value = "Этот товар сопутствующий для";
            package.Save();
        }
    }
}