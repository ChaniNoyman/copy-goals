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
        string tabNewlineRegex = @"\t{10,}"; // ביטוי רגולרי שמחפש 10 טאבים או יותר

        using (var reader = new StreamReader(sourceFilePath, Encoding.UTF8))
        using (var writer = new StreamWriter(destinationFilePath, false, Encoding.UTF8))
        using (var logWriter = new StreamWriter(logFilePath, false, Encoding.UTF8)) // פתיחת StreamWriter עבור הלוג
        {
            string fileContent = reader.ReadToEnd();
            string[] lines = Regex.Split(fileContent, tabNewlineRegex);

            foreach (string line in lines)
            {
                string[] columns = Regex.Split(line, @"\t");
                string delimiter = "!#!#";
                int delimiterIndex = -1;

                if (columns.Length > 1)
                {
                    for (int i = 0; i < columns.Length; i++)
                    {
                        if (columns[i].Contains(delimiter))
                        {
                            delimiterIndex = i;
                            break;
                        }
                    }

                    string colA = columns[0];
                    if (colA == "\r\n10495")
                    {
                        colA = "10495";
                    }
                    string colB = string.Join("", columns.Skip(1));

                    // בדיקה אם columns מכיל יותר משני תאים
                    if (delimiterIndex != -1)
                    {
                        colB = string.Join("\t", columns.Skip(1).Take(delimiterIndex - 1));
                        if (string.IsNullOrEmpty(colB))
                        {
                            colB = string.Join("\t", columns.Skip(delimiterIndex));
                            string searchTerm = "מטרות";
                            colB = FindAllSiblingsLiAfterText(colB, searchTerm);
                        }
                    }

                    if (!string.IsNullOrEmpty(colB))
                    {
                        colB = ReplaceStrongTagsAndKeepHebrew(colB);
                        colB = truncateAtFirstMatch(colB);
                        colB = Regex.Replace(colB, "שבטייצירת", "שבט יצירת");

                        List<string> outputValues = new List<string> { colA };
                        bool splitOccurred = false;

                        // ********** טיפול בפיצול לפי מספור והסרת מספור ושמירה בשורה אחת **********

                        if (Regex.IsMatch(colB, @"\d+[\.\)]"))
                        {
                            // פיצול לפי מספור ומחיקת המספור
                            string[] parts = Regex.Split(colB, @"\s*\d+[\.\)]\s*");
                            foreach (string part in parts)
                            {
                                string trimmedPart = part.Trim();
                                if (!string.IsNullOrEmpty(trimmedPart))
                                {
                                    trimmedPart = Regex.Replace(trimmedPart, "שבטייצירת", "שבט יצירת"); // החלפה גם פה
                                    outputValues.Add(RemoveNonHebrewPrefix(trimmedPart));
                                }
                            }
                            splitOccurred = true;
                        }

                        // ********** פיצול לפי HTML אם לא פוצל לפי מספור **********
                        else if (Regex.IsMatch(colB, @"<[^>]+>"))
                        {
                            var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                            htmlDoc.LoadHtml(colB);

                            var leafNodes = htmlDoc.DocumentNode.Descendants()
                                .Where(n => n.NodeType == HtmlNodeType.Element &&
                                            !n.ChildNodes.Any(c => c.NodeType == HtmlNodeType.Element) &&
                                            !string.IsNullOrWhiteSpace(n.InnerText));

                            foreach (var node in leafNodes)
                            {
                                string trimmedText = node.InnerText.Trim();
                                trimmedText = Regex.Replace(trimmedText, "שבטייצירת", "שבט יצירת");
                                trimmedText = Regex.Replace(trimmedText, @"\r\n|\r|\n", " "); // החלפת ירידות שורה ברווח
                                outputValues.Add(RemoveNonHebrewPrefix(trimmedText));
                            }
                            if (leafNodes.Any())
                            {
                                splitOccurred = true;
                            }
                        }
                        // ********** אם לא פוצל באף אחת מהשיטות **********
                        else
                        {
                            string value = colB.Trim();
                            value = Regex.Replace(value, @"\r\n|\r|\n", " "); // החלפת ירידות שורה ברווח
                            outputValues.Add(RemoveNonHebrewPrefix(value));
                        }

                        if (!string.IsNullOrWhiteSpace(colA) || outputValues.Count > 2 || !string.IsNullOrWhiteSpace(colB))
                        {
                            processedRecords.Add(outputValues.ToArray());
                        }
                        else if (outputValues.Count <= 2 && !string.IsNullOrWhiteSpace(colB))
                        {
                            logWriter.WriteLine($"{colA}\t{colB}");
                        }
                    }
                    else if (!string.IsNullOrWhiteSpace(colA))
                    {
                        processedRecords.Add(new string[] { colA, "" });
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
        // קוד למחיקת שורות ריקות מקובץ היעד
        var existingLines = File.ReadAllLines(destinationFilePath).ToList();
        var nonBlankLines = existingLines.Where(line => !string.IsNullOrWhiteSpace(line));
        File.WriteAllLines(destinationFilePath, nonBlankLines);

        Console.WriteLine("הפעולה הסתיימה בהצלחה, ושורות ריקות הוסרו מקובץ היעד.");
    } // סיום בלוק ה-using של reader ו-logWriter


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
                                    "הקדמה:", "שים לב פעולה", "למדריכים", "כפיסי קפלה"};
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

    public static string FindAllSiblingsLiAfterText(string html, string searchText, string separator = "")
    {
        HtmlDocument doc = new HtmlDocument();
        doc.LoadHtml(html);

        var nodesContainingText = doc.DocumentNode.SelectNodes($"//*[contains(text(), '{searchText}')]");

        if (nodesContainingText == null || !nodesContainingText.Any())
        {
            return null;
        }

        foreach (var textNode in nodesContainingText)
        {
            HtmlNode currentNode = textNode;
            HtmlNode firstFoundNode = null;
            string foundNodeType = "";

            // חפש את תגית ה-<li> או <p> הראשונה אחרי הטקסט
            while ((currentNode = GetNextNode(currentNode)) != null)
            {
                if (currentNode.Name.Equals("li", StringComparison.OrdinalIgnoreCase))
                {
                    firstFoundNode = currentNode;
                    foundNodeType = "li";
                    break; // מצאנו <li>, יוצאים מהלולאה
                }
                if (currentNode.Name.Equals("p", StringComparison.OrdinalIgnoreCase))
                {
                    firstFoundNode = currentNode;
                    foundNodeType = "p";
                    break; // מצאנו <p>, יוצאים מהלולאה
                }
            }

            if (firstFoundNode != null)
            {
                if (foundNodeType == "li" && firstFoundNode.ParentNode != null)
                {
                    // קבל את כל הילדים של ההורה של ה-<li> הראשון
                    var siblingNodes = firstFoundNode.ParentNode.ChildNodes;
                    List<string> siblingLis = new List<string>();

                    foreach (var sibling in siblingNodes)
                    {
                        if (sibling.Name.Equals("li", StringComparison.OrdinalIgnoreCase))
                        {
                            siblingLis.Add(sibling.OuterHtml);
                        }
                    }

                    if (siblingLis.Any())
                    {
                        return string.Join(separator, siblingLis);
                    }
                }
                else if (foundNodeType == "p")
                {
                    return firstFoundNode.OuterHtml; // החזר את ה-HTML של ה-<p> הראשון שנמצא
                }
            }
        }

        return null;
    }

    private static HtmlNode GetNextNode(HtmlNode node)
    {
        if (node == null) return null;

        if (node.FirstChild != null)
            return node.FirstChild;

        HtmlNode current = node;
        while (current != null)
        {
            if (current.NextSibling != null)
                return current.NextSibling;

            current = current.ParentNode;
        }

        return null;
    }
}