using Syncfusion;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Text.Json;
using System.Net.Http.Json;
using System.Reflection.Metadata;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Linq;
using System.Text.Json.Serialization.Metadata;
using System.Text.Encodings.Web;

namespace WordToJson
{

    internal class Program
    {
        private static FileStream filestreamPath = new FileStream("C:\\Users\\User\\Downloads\\Работа\\Sherpa\\MyDemoProject\\Copy_of_!Регламент_разработки_и_пересомтр_норм_расхода_топлива_ред.docx"
                , FileMode.Open, FileAccess.Read, FileShare.ReadWrite); // Вот это довести до ума надо
        static void Main(string[] args)
        {
            BuildJSON();
        }

        private static void BuildJSON()
        {
            using (WordDocument doc = new WordDocument(filestreamPath, FormatType.Automatic))
            {

                foreach (WSection section in doc.Sections)
                {
                    IterateTextBody(section.Body, doc);
                }

                doc.Close();
            }
        }
        private static void IterateTextBody(WTextBody body, in WordDocument doc)
        {
            List<JSONobject> processedWord = new List<JSONobject>();
            List<string> paragraphs = [], tables = [];

            string docTitle = GetFileName(filestreamPath.Name);
            string? sectionTitle = null;
            JSONobject jSONobject = new JSONobject(docTitle, sectionTitle); 
            foreach (IEntity entity in body.ChildEntities)
            {

                switch (entity.EntityType)
                {
                    case EntityType.Paragraph:
                        WParagraph paragraph = (WParagraph)entity;

                        if (IsHeading(paragraph))
                        {
                            jSONobject.content = [.. paragraphs];
                            jSONobject.tables = [.. tables];

                            paragraphs = [];
                            tables = [];

                            if (jSONobject.content.Length > 0)
                            {
                                processedWord.Add(jSONobject);
                            }

                            sectionTitle = paragraph.Text.Trim();

                            jSONobject = new JSONobject(docTitle, sectionTitle);

                            continue;
                        }
                        else
                        {
                            if (paragraph.Text != "")
                            {
                                paragraphs.Add(paragraph.Text.Trim());
                            }
                        }
                        break;

                    case EntityType.Table:
                        tables.Add(IterateTable((entity as WTable)!));
                        break;

                }
            }

            jSONobject.content = [.. paragraphs];
            jSONobject.tables = [.. tables];

            SerializeObjects(processedWord);
        }

        private static string IterateTable(WTable table)
        {
            List<string> result = [];
            foreach (WTableRow row in table.Rows)
            {
                List<string> row_data = [];
                foreach (WTableCell cell in row.Cells)
                {
                    foreach (WParagraph paragraph in cell.Paragraphs)
                    {
                        row_data.Add(paragraph.Text.Replace("\n", " ").Trim());
                    }
                }
                result.Add(String.Join(";",row_data));
            }
            return String.Join("\n",result);
        }

        private static bool IsHeading(WParagraph paragraph)
        {
            if (paragraph.StyleName.Contains("Heading"))
            {
                return true;
            }
            return false;
        }

        private static string GetFileName(string str)
        {
            Regex rgx = new Regex(@"[^\\]*$", RegexOptions.None);
            return new string(rgx.Match(str).Value);
        }

        private static void SerializeObjects(List<JSONobject> list)
        {
            var options = new JsonSerializerOptions { WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

            string json = JsonSerializer.Serialize<List<JSONobject>>(list, options);

            File.WriteAllText(@"processed_document.json", json);
        }
    }
}
