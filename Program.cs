using Microsoft.Office.Interop.Word;
using System;
using System.IO;

namespace ConsoleApp.Interop.Word.AddPicture
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            string imagePath = @"C:\Temp\test.png";
            string docPath = @"c:\temp\test.docx";
            string newDocPath = @"c:\temp\test-new.docx";

            Application wordApp = new Application();

            if(!File.Exists(docPath))
            {
                Console.WriteLine("Source Word document doesn't exist, please put a document in the path: " + docPath);
                return;
            }

            if (!File.Exists(imagePath))
            {
                Console.WriteLine("Image file doesn't exist, please put a document in the path: " + imagePath);
                return;
            }

            Document wordDoc = wordApp.Documents.Open(docPath);

            object saveWithDocument = true;
            object missing = Type.Missing;

            // Get pages count
            WdStatistic PagesCountStat = WdStatistic.wdStatisticPages;
            int PagesCount = wordDoc.ComputeStatistics(PagesCountStat, ref missing);

            //Get pages
            object What = WdGoToItem.wdGoToPage;
            object Which = WdGoToDirection.wdGoToAbsolute;
            object Start;
            object End;
            object CurrentPageNumber;
            object NextPageNumber;

            for (int Index = 1; Index < PagesCount + 1; Index++)
            {
                CurrentPageNumber = (Convert.ToInt32(Index.ToString()));
                NextPageNumber = (Convert.ToInt32((Index + 1).ToString()));

                // Get start position of current page
                Start = wordApp.Selection.GoTo(ref What, ref Which, ref CurrentPageNumber, ref missing).Start;

                // Get end position of current page                                
                End = wordApp.Selection.GoTo(ref What, ref Which, ref NextPageNumber, ref missing).End;

                // Get text
                object oRange;
                if (Convert.ToInt32(Start.ToString()) != Convert.ToInt32(End.ToString()))
                {
                    //Pages.Add(wordDoc.Range(ref Start, ref End).Text);

                    oRange = wordDoc.Range(ref Start, ref End);
                }
                else
                {
                    //Pages.Add(wordDoc.Range(ref Start).Text);

                    oRange = wordDoc.Range(ref Start);
                }

                InlineShape pic = wordDoc.InlineShapes.AddPicture(imagePath, ref missing, ref saveWithDocument, ref oRange);
                pic.Width = 595;
                pic.Height = 842;

                Shape shapePic = pic.ConvertToShape();
                shapePic.Left = -50;
                shapePic.Top = -30;
                shapePic.WrapFormat.Type = WdWrapType.wdWrapFront;
            }

            wordDoc.SaveAs2(newDocPath);
            wordApp.Quit();

            Console.WriteLine("Please check the new file: " + newDocPath);
        }
    }
}


