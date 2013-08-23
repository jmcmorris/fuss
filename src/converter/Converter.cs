using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Threading;
using OfficeOpenXml;

namespace converter
{
    public enum ConverterType {
        TXT,
        XLSX
    }

    public class Converter
    {
        private ConverterType type;
        private string[] files;
        private string output;
        private ConversionCompletedCallback onComplete;
        private Dispatcher uiDispatcher;

        public Converter(ConverterType type, string[] files, string output, ConversionCompletedCallback callback, Dispatcher uiDispatcher)
        {
            this.type = type;
            this.files = files;
            this.output = output;
            onComplete = callback;
            this.uiDispatcher = uiDispatcher;
        }

        public void perform()
        {
            foreach (var file in files)
            {
                //Attempt to convert the file
                try
                {
                    convertFile(file, output);
                }
                catch (Exception ex)
                {
                    
                }
            }
            object[] param = { };
            uiDispatcher.BeginInvoke(onComplete, param);
        }

        private void convertFile(string file, string output)
        {
            var chapters = getChapters(file);

            //Now that we have all the chapters as lists of questions lets write them out to the new format
            //One file per chapter
            if (this.type == ConverterType.TXT)
            {
                writeTxt(file, output, chapters);
            }
            else if (this.type == ConverterType.XLSX)
            {
                writeXlsx(new FileInfo(file), chapters);
            }
        }

        private Dictionary<string, List<Question>> getChapters(string file)
        {
            FlowDocument document = new FlowDocument();

            //Read the file stream to a Byte array 'data'
            TextRange txtRange = null;

            using (MemoryStream stream = new MemoryStream(File.ReadAllBytes(file)))
            {
                // create a TextRange around the entire document
                txtRange = new TextRange(document.ContentStart, document.ContentEnd);
                txtRange.Load(stream, DataFormats.Rtf);
            }

            //First open up the file and pull out all of the questions from it
            List<string> lines = new List<string>();
            foreach (string line in Regex.Split(txtRange.Text, @"\r\n|\r|\n"))
            {
                string trimmedLine = line.Trim();
                if (trimmedLine != string.Empty)
                {
                    lines.Add(trimmedLine);
                }
            }

            Dictionary<string, List<Question>> chapters = new Dictionary<string, List<Question>>();
            int position = 0;
            while (position < lines.Count)
            {
                Question question = Question.ReadFromQG(ref lines, ref position);
                if (question == null)
                {
                    break;
                }
                if (!chapters.ContainsKey(question.Chapter))
                {
                    chapters.Add(question.Chapter, new List<Question>());
                }
                chapters[question.Chapter].Add(question);
                ++position;
            }
            return chapters;
        }

        private void writeTxt(string file, string output, Dictionary<string, List<Question>> chapters)
        {
            string directory = output + Path.DirectorySeparatorChar + Path.GetFileNameWithoutExtension(file);
            foreach (char invalid in Path.GetInvalidPathChars())
            {
                directory = directory.Replace(invalid.ToString(), "");
            }
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
            foreach (var chapter in chapters)
            {
                string filename = chapter.Key;
                foreach (char invalid in Path.GetInvalidFileNameChars())
                {
                    filename = filename.Replace(invalid.ToString(), "");
                }
                string filePath = directory + Path.DirectorySeparatorChar + filename + ".txt";
                using (StreamWriter stream = File.CreateText(filePath))
                {
                    foreach (Question question in chapter.Value)
                    {
                        question.Write(stream);
                    }
                }
            }
        }

        private void writeXlsx(FileInfo file, Dictionary<string, List<Question>> chapters)
        {
            string filename = file.DirectoryName + @"\" + Path.GetFileNameWithoutExtension(file.Name) + ".xlsx";
            FileInfo newFile = new FileInfo(filename);
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(filename);
            }
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                worksheet.Cells[1, 1].Value = "Total Questions";
                worksheet.Cells[1, 2].Value = "Chapter";
                worksheet.Cells[1, 3].Value = "Chapter Number";
                worksheet.Cells[1, 4].Value = "Question Number (per chapter)";
                worksheet.Cells[1, 5].Value = "Question";
                worksheet.Cells[1, 6].Value = "Correct Answer";
                worksheet.Cells[1, 7].Value = "Incorrect Answer";
                worksheet.Cells[1, 8].Value = "Incorrect Answer";
                worksheet.Cells[1, 9].Value = "Incorrect Answer";
                worksheet.Cells[1, 10].Value = "Reference";
                worksheet.Column(1).AutoFit();
                worksheet.Column(3).AutoFit();
                worksheet.Column(4).AutoFit();
                worksheet.Column(2).Width = 25.0;
                worksheet.Column(10).Width = 25.0;
                worksheet.Column(6).Width = 50.0;
                worksheet.Column(7).Width = 50.0;
                worksheet.Column(8).Width = 50.0;
                worksheet.Column(9).Width = 50.0;
                worksheet.Column(5).Width = 80.0;
                int totalQuestions = 1;
                foreach (var chapter in chapters)
                {
                    int chapterQuestions = 1;
                    var match = Regex.Match(chapter.Key, @"Ch ([0-9]+) - (.*)");
                    int chapterNumber = Convert.ToInt32(match.Groups[1].Value);
                    string chapterName = match.Groups[2].Value;
                    foreach (var question in chapter.Value)
                    {
                        int row = totalQuestions + 1;
                        worksheet.Row(row).Style.WrapText = true;
                        worksheet.Cells[row, 1].Value = totalQuestions.ToString();
                        worksheet.Cells[row, 2].Value = chapterName;
                        worksheet.Cells[row, 3].Value = chapterNumber.ToString();
                        worksheet.Cells[row, 4].Value = chapterQuestions.ToString();
                        question.WriteXlsx(worksheet, row);
                        ++totalQuestions;
                        ++chapterQuestions;
                    }
                }
                package.Workbook.Properties.Title = Path.GetFileNameWithoutExtension(file.Name);
                package.Workbook.Properties.Author = "Fire Instructor Testing Software, LLC";
                package.Save();
            }
        }
    }
}
