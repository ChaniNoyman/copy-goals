using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using HtmlAgilityPack;
using System.Text;
using System.Linq;

public class CsvCopyAndSplit
{
    public static void Main(string[] args)
    {
        string sourceFilePath = @"C:\Users\Chanif\Desktop\C#\Copy Excel File\data.csv";
        string destinationFilePath = @"C:\Users\Chanif\Desktop\C#\Copy Excel File\fixData.csv";
        string logFilePath = @"C:\Users\Chanif\Desktop\C#\Copy Excel File\LOG.csv"; // נתיב לקובץ הלוג

        var processedRecords = new List<string[]>();

        using (var reader = new StreamReader(sourceFilePath, Encoding.UTF8))
        using (var writer = new StreamWriter(destinationFilePath, false, Encoding.UTF8, 1024) { NewLine = "\n" }) // כתיבה עם LF
        using (var logWriter = new StreamWriter(logFilePath, false, Encoding.UTF8, 1024) { NewLine = "\n" }) // כתיבה עם LF
        {
            string line;
            while ((line = ReadLineLFOnly(reader)) != null)
            {
                string[] columns = Regex.Split(line, @"\t");

                if (columns.Length > 0 && Regex.IsMatch(columns[0], @"\d")) // בדיקה אם עמודה A מכילה לפחות מספר אחד
                {
                    string colA = columns[0];
                    List<string> outputValues = new List<string> { colA }; // אתחול outputValues עם עמודה A
                    bool splitOccurred = false;
                    List<string> splitParts = new List<string>();

                    if (columns.Length > 1)
                    {
                        string colB = columns[1];
                        if (columns.Length > 2)
                        {
                            colB = string.Join("", columns.Skip(1));
                        }
                        colB = ReplaceStrongTagsAndKeepHebrew(colB);
                        colB = truncateAtFirstMatch(colB);
                        colB = Regex.Replace(colB, "שבטייצירת", "שבט יצירת");
                        if (colB.EndsWith("\""))
                        {
                            colB = colB.Substring(0, colB.Length - 1);
                        }

                        string textWithoutHtml = RemoveHtmlTags(colB); // חילוץ טקסט ללא HTML

                        // פיצול על סמך מספר ואחריו נקודה או סוגריים ואז רווח אופציונלי (יכול להופיע בכל מקום)
                        if (Regex.IsMatch(textWithoutHtml, @"\d+[\.\)]\s"))
                        {
                            string[] parts = Regex.Split(textWithoutHtml, @"\s*\d+[\.\)]\s*");
                            foreach (string part in parts)
                            {
                                string trimmedPart = part.Trim();
                                if (!string.IsNullOrEmpty(trimmedPart))
                                {
                                    trimmedPart = Regex.Replace(trimmedPart, "שבטייצירת", "שבט יצירת");
                                    splitParts.Add(RemoveNonHebrewPrefix(trimmedPart));
                                }
                            }
                            if (splitParts.Any())
                            {
                                outputValues.AddRange(splitParts);
                                splitOccurred = true;
                            }
                        }
                        else if (Regex.IsMatch(colB, @"<\s*\w+.*?>")) // אם אין מספור, בודקים אם יש תגיות HTML
                        {
                            var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                            htmlDoc.LoadHtml(colB);
                            var textNodes = htmlDoc.DocumentNode.SelectNodes("//text()[normalize-space()]");
                            if (textNodes != null)
                            {
                                foreach (var node in textNodes)
                                {
                                    string trimmedText = node.InnerText.Trim();
                                    trimmedText = Regex.Replace(trimmedText, "שבטייצירת", "שבט יצירת");
                                    splitParts.Add(RemoveNonHebrewPrefix(RemoveHtmlTags(trimmedText)));
                                }
                                if (splitParts.Any())
                                {
                                    outputValues.AddRange(splitParts);
                                    splitOccurred = true;
                                }
                            }
                            else
                            {
                                outputValues.Add(RemoveNonHebrewPrefix(RemoveHtmlTags(colB)));
                            }
                        }
                        else
                        {
                            outputValues.Add(RemoveNonHebrewPrefix(RemoveHtmlTags(colB)));
                        }
                    }
                    else if (columns.Length == 1) // אם יש רק עמודה אחת
                    {
                        processedRecords.Add(new[] { colA });
                        continue;
                    }

                    processedRecords.Add(outputValues.ToArray());
                    if (!splitOccurred && columns.Length > 1 && !string.IsNullOrWhiteSpace(columns[1]))
                    {
                        logWriter.WriteLine($"{colA}\t{columns[1]}");
                    }
                }
            }

            var shiftedRecords = new List<string[]>();
            foreach (var record in processedRecords)
            {
                shiftedRecords.Add(ShiftLeft(record));
            }

            foreach (var record in shiftedRecords)
            {
                writer.WriteLine(string.Join("\t", record));
            }
        }

        Console.WriteLine("הפעולה הסתיימה בהצלחה!");
    }

    private static string ReadLineLFOnly(StreamReader reader)
    {
        StringBuilder sb = new StringBuilder();
        int charCode;
        char? previousChar = null;

        while ((charCode = reader.Read()) != -1)
        {
            char currentChar = (char)charCode;

            if (currentChar == '\n' && previousChar != '\r')
            {
                return sb.ToString();
            }
            else if (currentChar == '\r')
            {
                previousChar = currentChar;
            }
            else
            {
                sb.Append(currentChar);
                previousChar = currentChar;
            }
        }
        return sb.Length > 0 ? sb.ToString() : null;
    }

    private static string[] ShiftLeft(string[] row)
    {
        var nonNullValues = row.Where(s => !string.IsNullOrEmpty(s)).ToList();
        var newRow = new string[row.Length];
        for (int i = 0; i < newRow.Length; i++)
        {
            newRow[i] = i < nonNullValues.Count ? nonNullValues[i] : "";
        }
        return newRow;
    }

    private static string RemoveNonHebrewPrefix(string input)
    {
        return Regex.Replace(input, @"^\P{IsHebrew}+", "");
    }

    private static string ReplaceStrongTagsAndKeepHebrew(string input)
    {
        string inputWithoutB = Regex.Replace(input, @"<[/]?b>", "");
        string inputWithoutU = Regex.Replace(inputWithoutB, @"<[/]?u>", "");
        string inputWithoutSpan = Regex.Replace(inputWithoutU, @"<span style=""font-size: 1rem;\""\s*/>", "");
        string inputWithoutMStrong = Regex.Replace(inputWithoutSpan, @"<strong>\s*מטרות\s*<\/strong>", "", RegexOptions.IgnoreCase);
        string inputWithoutMB = Regex.Replace(inputWithoutMStrong, @"<b>\s*מטרות\s*<\/b>", "", RegexOptions.IgnoreCase);
        string inputWithoutMRegular = Regex.Replace(inputWithoutMB, @"\s*מטרות\s*", "", RegexOptions.IgnoreCase);
        string inputWithoutNbsp = inputWithoutMRegular.Replace("&nbsp;", " ");
        string inputWithoutEM = inputWithoutNbsp.Replace("<em>", " ");

        string pattern = @"<strong[^>]*>(.*?)<\/strong>";
        return Regex.Replace(inputWithoutEM, pattern, match =>
        {
            string content = match.Groups[1].Value;
            return Regex.IsMatch(content, @"\p{IsHebrew}") ? content : "";
        });
    }

    private static string truncateAtFirstMatch(string text)
    {
        string[] keywords = { "ציוד", "מקורות להרחבה", "העשרה למדריכים", "אביזרים", "עזרים", "שימו לב,", "הערה כללית", "כובעים", "מדריכים יקרים"
                                    , "הערה למדריכים" , "בקבוק", "חשבי", "רעיון ליישום", "להרחבה", "המלצות ללימוד", "כמה רעיונות", "מקורות מידע", "בנו טקס", "טיפים",
                                    "הקדמה:", "שים לב פעולה", "למדריכים", "כפיסי קפלה","חומרי עזר","מאמרי הראי\"\"ה"};
        int firstIndex = -1;

        foreach (string keyword in keywords)
        {
            int index = text.IndexOf(keyword);
            if (index != -1 && (firstIndex == -1 || index < firstIndex))
            {
                firstIndex = index;
            }
        }

        if (firstIndex != -1)
        {
            return text.Substring(0, firstIndex).Trim();
        }

        return text;
    }

    private static string RemoveHtmlTags(string input)
    {
        if (string.IsNullOrEmpty(input))
        {
            return input;
        }
        HtmlDocument doc = new HtmlDocument();
        doc.LoadHtml(input);
        return HtmlEntity.DeEntitize(doc.DocumentNode.InnerText);
    }
}