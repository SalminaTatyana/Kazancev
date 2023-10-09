using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Pdf;
using Spire.Pdf.General.Find;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace CheckPageCount
{
    internal class Program
    {
        static void Main(string[] args)
        {

            try
            {
                if (args.Length > 0)
                {
                    if (args.Length != 2)
                    {
                        Console.WriteLine("Неверное количество файлов. Проверьте наличие строго 2 файлов в bat файле");
                        Console.ReadKey();
                    }
                    else
                    {
                        Document doc = OpenDoc(args[0]);
                        FileManager.GetFileNameResult(args[1]);
                        int count = ListCount(doc);
                        string allCount = FindAllPagesInStartPage(doc);
                        Console.WriteLine(count);
                        Console.WriteLine(allCount);
                        if (!string.IsNullOrEmpty(allCount))
                        {
                            int page = 0;
                            string number = string.Empty;
                            if (allCount != null && allCount.Length > 0)
                            {
                                for (int i = 0; i < allCount.Length; i++)
                                {
                                    if (Char.IsDigit(allCount[i]))
                                        number += allCount[i];
                                }

                                if (number.Length > 0)
                                    page = int.Parse(number);
                            }

                            if (count != page)
                            {
                                FileManager.WriteFile($"Количество страниц не совпадает! Посчитано: {count}, на странице реферата: {page}");
                                Console.WriteLine($"Количество страниц не совпадает! Посчитано: {count}, на странице реферата: {page}");
                            }
                        }
                        else
                        {
                            FileManager.WriteFile("Нет количества страниц на странице реферата!");
                            Console.WriteLine("Нет количества страниц на странице реферата!");
                        }
                        Console.ReadKey();

                    }
                }
                else
                {
                    Console.WriteLine("Не заданы файлы. Проверьте наличие строго 2 файлов в bat файле");
                    Console.ReadKey();
                    //Document doc = OpenDoc("1.docx");
                    //FileManager.GetFileNameResult("1");
                    //int count = ListCount(doc);
                    //string allCount = FindAllPagesInStartPage(doc);
                    //Console.WriteLine(count);
                    //Console.WriteLine(allCount);
                    //if (!string.IsNullOrEmpty(allCount))
                    //{
                    //    int page = 0;
                    //    string number = string.Empty;
                    //    if (allCount != null && allCount.Length > 0)
                    //    {
                    //        for (int i = 0; i < allCount.Length; i++)
                    //        {
                    //            if (Char.IsDigit(allCount[i]))
                    //                number += allCount[i];
                    //        }

                    //        if (number.Length > 0)
                    //            page = int.Parse(number);
                    //    }

                    //    if (count != page)
                    //    {
                    //        FileManager.WriteFile($"Количество страниц не совпадает! Посчитано: {count}, на странице реферата: {page}");
                    //        Console.WriteLine($"Количество страниц не совпадает! Посчитано: {count}, на странице реферата: {page}");
                    //    }
                    //}
                    //else
                    //{
                    //    FileManager.WriteFile("Нет количества страниц на странице реферата!");
                    //    Console.WriteLine("Нет количества страниц на странице реферата!");
                    //}
                    //Console.ReadKey();
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Что-то пошло не так!");
                Console.ReadKey();
            }


            Document OpenDoc(string path)
            {
                if (File.Exists(path))
                {
                    Document document = new Document(path.Trim());
                    return document;
                }
                else
                {
                    FileManager.WriteFile("Файла не существует");
                    Console.WriteLine("Файла не существует");
                    Console.ReadKey();
                    return null;
                }
            }

            int ListCount(Document doc)
            {
                if (doc != null)
                {
                    doc.SaveToFile("ToPDF.pdf", Spire.Doc.FileFormat.PDF);
                    PdfDocument pd = new PdfDocument("ToPDF.pdf");
                    int pageStart = FindDocPart(pd, "Введение");
                    if (pageStart == 0)
                    {
                        pageStart = FindDocPart(pd, "ВВЕДЕНИЕ");
                        if (pageStart == 0)
                        {
                            pageStart = FindDocPart(pd, "введение");
                        }
                        if (pageStart == 0)
                        {
                            FileManager.WriteFile("Не обнаружено введение в содержании.");
                            Console.WriteLine("Не обнаружено введение в содержании.");
                            return 0;
                        }
                    }
                    int pageStop = doc.PageCount;
                    return pageStop - pageStart;
                }
                else
                {
                    return 0;
                }
            }

            int FindDocPart(PdfDocument doc, string part)
            {
                if (doc != null)
                {
                    int page = 0;
                    PdfTextFind[] result = null;
                    foreach (PdfPageBase pageItem in doc.Pages)
                    {
                        result = pageItem.FindText(part, TextFindParameter.None).Finds;
                        string b = string.Empty;
                        if (result != null && result.Length > 0)
                        {
                            for (int i = 0; i < result[0].LineText.Length; i++)
                            {
                                if (Char.IsDigit(result[0].LineText[i]))
                                    b += result[0].LineText[i];
                            }

                            if (b.Length > 0)
                                page = int.Parse(b);
                            if (result[0].LineText.Length > 0)
                            {
                                break;
                            }
                        }

                    }
                    return page;
                }
                else
                {
                    return 0;
                }
            }

            string FindAllPagesInStartPage(Document doc)
            {
                Regex regex = new Regex(@"стр\..{0,1}[0-9]{1,3}");
                TextSelection text = doc.FindPattern(regex);
                return text?.SelectedText.ToString() ?? "";
            }
        }



    }
}
