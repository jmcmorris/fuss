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

namespace converter
{
    public class Converter
    {
        private string[] files;
        private string output;
        private ConversionCompletedCallback onComplete;
        private Dispatcher uiDispatcher;

        public Converter(string[] files, string output, ConversionCompletedCallback callback, Dispatcher uiDispatcher)
        {
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

            //Now that we have all the chapters as lists of questions lets write them out to the new format
            //One file per chapter
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
    }
}
