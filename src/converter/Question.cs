using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;

namespace converter
{
    class Question
    {
        private string text;
        private Dictionary<string, string> answers;
        private string correctAnswer;
        private string reference;
        private string chapter;

        public string Chapter
        {
            get { return chapter; }
            private set { chapter = value; }
        }

        public Question(string text, Dictionary<string, string> answers, string correct, string reference, string chapter)
        {
            this.text = text;
            this.answers = answers;
            this.correctAnswer = correct;
            this.reference = reference;
            Chapter = chapter;
        }

        public void Write(StreamWriter stream)
        {
            string questionText = text;
            StringBuilder answerBuilder = new StringBuilder("{");
            foreach (var answer in answers)
            {
                if (correctAnswer == answer.Key)
                {
                    answerBuilder.Append("=");
                }
                else
                {
                    answerBuilder.Append("~");
                }
                answerBuilder.Append(answer.Value);
                answerBuilder.Append(" ");
            }
            //Remove the extra space at the end
            answerBuilder.Remove(answerBuilder.Length - 1, 1);
            answerBuilder.Append("}");

            if (Regex.IsMatch(text, @"\b_{4,}\b"))
            {
                questionText = Regex.Replace(text, @"\b_{4,}\b", answerBuilder.ToString());
            } else
            {
                questionText += " " + answerBuilder.ToString();
            }

            stream.WriteLine("::" + reference);
            stream.WriteLine("::" + questionText);
            stream.WriteLine();
        }

        /// <summary>
        /// Reads a question out from a stream.
        /// </summary>
        /// <exception cref=""></exception>
        /// <param name="stream">The stream to read a question from.</param>
        public static Question ReadFromQG(ref List<string> lines, ref int position)
        {
            string text;
            Dictionary<string, string> answers = new Dictionary<string,string>();
            string correct;
            string reference;
            string chapter;

            //Determine the type of content this line has
            while (!Regex.IsMatch(lines[position], @"^[0-9]+\.\t(.*)"))
            {
                if (++position >= lines.Count) { return null; }
            }
            var match = Regex.Match(lines[position], @"^[0-9]+\.\t(.*)");
            if (match.Groups.Count < 2)
            {
                return null;
            }
            text = match.Groups[1].Value;
            if (++position >= lines.Count) { return null; }

            while (!Regex.IsMatch(lines[position], @"^[A-Z]\.\t"))
            {
                text += lines[position];
                if (++position >= lines.Count) { return null; }
            }

            while (Regex.IsMatch(lines[position], @"^[A-Z]\.\t"))
            {
                match = Regex.Match(lines[position], @"^([A-Z])\.\s*(.*)");
                if (match.Groups.Count < 3)
                {
                    return null;
                }
                string answer = match.Groups[2].Value;

                if (position + 3 >= lines.Count) { return null; }

                while (!Regex.IsMatch(lines[position + 1], @"^[A-Z]\.\t") && !Regex.IsMatch(lines[position + 3], @"^Answer\:"))
                {
                    if (++position >= lines.Count) { return null; }
                    answer += " " + lines[position];
                }
                try
                {
                    answers.Add(match.Groups[1].Value, answer);
                }
                catch (Exception ex)
                {
                    return null;
                }
                if (++position >= lines.Count) { return null; }
            }
            chapter = lines[position];

            if (++position >= lines.Count) { return null; }
            reference = lines[position];

            if (++position >= lines.Count) { return null; }
            match = Regex.Match(lines[position], @"^Answer\:\s*(.*)");
            if (match.Groups.Count < 2)
            {
                return null;
            }
            correct = match.Groups[1].Value;

            return new Question(text, answers, correct, reference, chapter);
        }
    }
}
