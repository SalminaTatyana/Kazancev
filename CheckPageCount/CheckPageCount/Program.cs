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
            string fileNameResult = "";

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
                        string[] name = args[1].Split('.');

                        string expansion = $".{name[name.Length - 1]}";
                        if (expansion == ".txt")
                        {
                            fileNameResult = args[1];
                        }
                        else
                        {
                            fileNameResult = $"{args[1]}.txt";
                        }
                        if (!File.Exists(fileNameResult))
                        {
                            File.Create(fileNameResult);
                        }
                        WriteFile($"Проверяемый файл {args[0]}, выходной файл {args[1]}", fileNameResult);

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
                                WriteFile($"Количество страниц не совпадает! Посчитано: {count}, на странице реферата: {page}", fileNameResult);
                                Console.WriteLine($"Количество страниц не совпадает! Посчитано: {count}, на странице реферата: {page}");
                            }
                        }
                        else
                        {
                            WriteFile("Нет количества страниц на странице реферата!", fileNameResult);
                            Console.WriteLine("Нет количества страниц на странице реферата!");
                        }
                        
                        Console.ReadKey();

                    }
                }
                else
                {
                    Console.WriteLine("Не заданы файлы. Проверьте наличие строго 2 файлов в bat файле");
                    Console.ReadKey();
                    //Document doc = OpenDoc("3.docx");
                    //string[] name = "1".Split('.');

                    //string expansion = $".{name[name.Length - 1]}";
                    //if (expansion == ".txt")
                    //{
                    //    fileNameResult = "1";
                    //}
                    //else
                    //{
                    //    fileNameResult = $"1.txt";
                    //}
                    //if (!File.Exists(fileNameResult))
                    //{
                    //    File.Create(fileNameResult);
                    //}
                    //WriteFile($"Проверяемый файл 3.docx, выходной файл 1.txt", fileNameResult);

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
                    //        WriteFile($"Количество страниц не совпадает! Посчитано: {count}, на странице реферата: {page}", fileNameResult);
                    //        Console.WriteLine($"Количество страниц не совпадает! Посчитано: {count}, на странице реферата: {page}");
                    //    }
                    //    else
                    //    {
                    //        WriteFile("Проверка прошла успешно!", fileNameResult);
                    //    }
                    //}
                    //else
                    //{
                    //    WriteFile("Нет количества страниц на странице реферата!", fileNameResult);
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

                    WriteFile("Файла не существует", fileNameResult);
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
                    int pageStart = FindDocPart(pd, "Реферат");
                    if (pageStart == 0)
                    {

                        WriteFile("Не обнаружено реферата.", fileNameResult);
                        Console.WriteLine("Не обнаружено реферата.");
                        return 0;
                    }
                    int pageStop = doc.PageCount;
                    return pageStop+1 - (pageStart-2); //Включается сама старница реферата (по ворду она 4, на деле 3 поэтому -2) +1 - включается сама паследняя страница
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
                        result = pageItem.FindText(part, TextFindParameter.IgnoreCase).Finds;
                        foreach (PdfTextFind find in result)
                        {
                            //get the page number of the text
                            page = find.SearchPageIndex + 1;
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

        static public void WriteFile(string text,string fileNameResult)
        {
            using (StreamWriter writer = new StreamWriter(fileNameResult, true, System.Text.Encoding.Default))
            {
                writer.WriteLine(text);
            }

        }

    }
}
