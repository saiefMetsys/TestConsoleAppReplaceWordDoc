using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
//using System.Windows.Forms;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using System.Xml.Linq;

namespace TestConsoleAppReplaceWord
{
    internal class Program
    {
        private void FindAndReplace(Word.Application wordApp, object TextToFind, object replaceWithText)
        {
            //this.Application.Documents.Add(@"C:\Test\SampleTemplate.dotx");
            //object matchCase = true;
            //object matchWholeWord = true;
            // object matchWildCards = false;

            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = true;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            wordApp.Selection.Find.Execute(ref TextToFind, ref matchCase,
                ref matchWholeWord, ref matchWildCards, ref matchSoundsLike,
                ref matchAllWordForms, ref forward, ref wrap,
                ref format, ref replaceWithText, ref replace, ref matchKashida, ref matchDiacritics,
                ref matchAlefHamza, ref matchControl);
            
        }

        public void CreateDocument(object filename, object SaveAs)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document myWordDoc = null;
            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly, ref missing,
                   ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();
                //find and replace : 

                this.FindAndReplace(wordApp, "<nom>", "ABBASSI");
                this.FindAndReplace(wordApp, "<email>", "saief.abbassi@metsys.fr");
                this.FindAndReplace(wordApp, "<phoneNumber>", "07*******");
                this.FindAndReplace(wordApp, "<prenom>", "saief");
                
                // finding and replacung image content : 
                Word.Range rng = myWordDoc.Content;
                Word.Find wdFind = rng.Find;
                wdFind.Text = "<photo>";
                bool found = wdFind.Execute();
                if (found)
                {
                    rng.InsertAfter("\n");
                    rng.MoveStart(Word.WdUnits.wdParagraph, 1);
                    Word.InlineShape ils = rng.InlineShapes.AddPicture(@"C:\Users\saief.abbassi\OneDrive - METSYS\Desktop\Saief\Projets\c#\Capture.jpg", false, true, rng);
                    Console.WriteLine("<photo> was found and the  photo added !! ");
                }
            }
            else
            {
                Console.WriteLine("<photo> was not found !! ");
            }

            //SAveAs 
            myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing);
            myWordDoc.Close();
            wordApp.Quit();
            // myWordDoc.Quit();
            Console.WriteLine("Word document is updated and saved succefully !");
        }

        static void Main(string[] args)
        {
            var pr = new Program();
            Console.WriteLine("Hello we are going to find & replace words in word doc !!");
            pr.CreateDocument(@"C:\Users\saief.abbassi\OneDrive - METSYS\Desktop\Saief\Projets\c#\temp.docx", @"C:\Users\saief.abbassi\OneDrive - METSYS\Desktop\Saief\Projets\c#\Output.docx");
            Console.WriteLine("New word file created and saved, See u next time !");
            Console.ReadLine();
        }
    }
}
