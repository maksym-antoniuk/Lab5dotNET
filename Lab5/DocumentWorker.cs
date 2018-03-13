using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace Lab5
{
    public class DocumentWorker
    {
        public static object template = "C:\\QuestionTemplate.dot";
        Word.Application application;
        Word.Document document;
        Object missingObj = Missing.Value;
        Object trueObj = true;
        Object falseObj = false;
        StreamReader sr;

        public DocumentWorker(string path)
        {
            sr = File.OpenText(path);
            application = new Word.Application();
            object temp = template;
            
        }

        public void Init()
        {
            try
            {
                document = application.Documents.Add(ref template, ref missingObj, ref missingObj, ref missingObj);
            }
            catch (Exception error)
            {
                Console.WriteLine(error.ToString());
                document.Close(ref falseObj, ref missingObj, ref missingObj);
                application.Quit(ref missingObj, ref missingObj, ref missingObj);
                document = null;
                application = null;
                throw error;
            }
            application.Visible = true;
        }

        public void SwapText(params string[] marks)
        {
            bool rangeFound;
            Word.Range wordRange;
            foreach (string mark in marks)
            {
                rangeFound = false;
                for (int i = 1; i <= document.Sections.Count; i++)
                {
                    wordRange = document.Sections[i].Range;
                    Word.Find wordFindObj = wordRange.Find;
                    object[] wordFindParameters = new object[15]
                    { mark, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj,
                    missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj };
                    rangeFound = (bool)wordFindObj.GetType().InvokeMember
                        ("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
                    if (rangeFound)
                    {
                        wordRange.Text = sr.ReadLine();
                        break;
                    }
                }
            }
        }

        public void Save(string path, string fileName)
        {
            try
            {
                document.SaveAs(FileName: path + "\\" + fileName + ".docx");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

    }
}
