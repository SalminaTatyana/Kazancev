using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CheckPageCount
{
    internal class FileManager
    {
        static string fileNameResult { get; set; }

        // название выходного файла
        static internal void GetFileNameResult(string line)
        {
            string[] name = line.Split('.');
            string expansion = $".{name[name.Length - 1]}";
            if (expansion == ".txt")
            {
                fileNameResult = line;
            }
            else
            {
                fileNameResult = $"{line}.txt";
            }
            File.Delete(fileNameResult);
        }

        // подставляет расширения
        static internal string ReadFile(string fileName = "Тест")
        {
            string[] name = fileName.Split('.');
            string expansion = $".{name[name.Length - 1]}";
            string filePath;
            string[] expansions = new string[3] { ".doc", ".docx", ".docm" };
            if (expansion == expansions[0] || expansion == expansions[1] || expansion == expansions[2])
            {
                filePath =  $"\\{fileName}";
                if (File.Exists(filePath))
                {
                    return filePath;
                }
                else
                {
                    return "ERROR";
                }
            }
            else
            {
                filePath = $"\\{fileName}{expansions[0]}";
                if (File.Exists(filePath))
                {
                    return filePath;
                }
                else
                {
                    filePath =  $"\\{fileName}{expansions[1]}";
                }

                if (File.Exists(filePath))
                {
                    return filePath;
                }
                else
                {
                    filePath = $"\\{fileName}{expansions[2]}";
                }

                if (File.Exists(filePath))
                {
                    return filePath;
                }
                else
                {
                    return "ERROR";
                }
            }
        }
        static internal void WriteFile(string text)
        {
            using (StreamWriter writer = File.AppendText(fileNameResult))
                writer.WriteLine(text);
        }
    }
}
