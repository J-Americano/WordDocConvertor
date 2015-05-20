using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace HelpNoteWorker
{
    class Program
    {
        public static int i = 0;
        public static int total = 0;
        static void Main(string[] args)
        {
            Console.WriteLine("Ensure all Word documents are saved and closed. Press enter to continue.");
            Console.ReadLine();
            string path = Directory.GetCurrentDirectory();
            string s = null;
            bool success = processDirectory(path);
            if (success)
            {
                Process[] process = Process.GetProcesses();
                foreach (Process proc in process)
                {
                    s = proc.ProcessName.ToLower();
                    if (s.CompareTo("winword") == 0)
                    {
                        proc.Kill();
                    }
                }

                Console.WriteLine("Finished reformatting help documents.");
            }
            else
                Console.WriteLine("No files found.");
        }

        //Create document method
        private static void ProcessFile(string path)
        {
            int output = ++i;
            Microsoft.Office.Interop.Word.Application app;
            Microsoft.Office.Interop.Word.Documents doc;
            Microsoft.Office.Interop.Word.Document document;
            Microsoft.Office.Interop.Word.Range rng;
            Microsoft.Office.Interop.Word.Range rng2;

            object saveOption = Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges;
            object missing = System.Reflection.Missing.Value;
            object styleHeading1 = "Heading 1";
            object oEndOfDoc = "\\startofdoc";

            string title;
            try
            {
                app = new Microsoft.Office.Interop.Word.Application();
                

                app.Visible = false;
                app.DisplayAlerts = WdAlertLevel.wdAlertsNone;

                doc = app.Documents;
                
                document = doc.Open(path);
                //adding text to document
                
                rng = document.Bookmarks[ref oEndOfDoc].Range;
                rng2 = document.Bookmarks[ref oEndOfDoc].Range;

                if (rng.Tables.Count >= 1)
                {
                    if (rng.Tables[1].Range.Start == rng2.Start)
                    {
                        rng.Tables[1].Rows.Add(rng.Tables[1].Rows[1]);
                        rng.Tables[1].Rows[1].Cells.Merge();
                    }
                }

                rng2.set_Style(document.Styles[styleHeading1]);

                title = path.Substring(path.LastIndexOf("\\") + 1);

                title = title.Substring(0, title.LastIndexOf("."));

                if (rng.Tables.Count >= 1)
                {
                    if (rng.Tables[1].Range.Start == rng2.Start)
                    {
                        rng2.Text = title;
                        rng2.Tables[1].Rows[1].ConvertToText();
                    }
                }
                else
                {
                    rng2.InsertBefore(title);
                }

                app.NormalTemplate.Saved = true;

                //Save the document
                document.SaveAs2(path);
                if (document != null)
                {
                    ((_Document)document).Close(ref saveOption, ref missing, ref missing);
                    while(Marshal.FinalReleaseComObject(document) > 0 );
                    document = null;
                }

                if (doc != null)
                {
                    while(Marshal.FinalReleaseComObject(doc) > 0);
                    doc = null;
                    
                }

                if (doc != null)
                {
                    ((_Application)app).Quit(ref saveOption, ref missing, ref missing);
                    while(Marshal.FinalReleaseComObject(app) > 0);
                    app = null;
                }

                while(Marshal.FinalReleaseComObject(rng) > 0);
                while (Marshal.FinalReleaseComObject(rng2) > 0) ;


                Console.WriteLine(output + "/" + total + " files written");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
                
            }
        }


        private static bool processDirectory(string path)
        {
            string[] fileEntries = Directory.GetFiles(path, "*.doc");
            total = fileEntries.Length;
            if (!path.Contains("docx") && !path.Contains("?") && !path.Contains("~") && fileEntries.Length > 0)
            {
                Console.WriteLine("Starting write process.");
                //fileEntries.AsParallel().ForAll(ProcessFile);
                Parallel.ForEach(fileEntries, fileName => ProcessFile(fileName));
                /*foreach (string fileName in fileEntries)
                {
                    ProcessFile(fileName);
                    i++;
                    Console.WriteLine(i + "/" + fileEntries.Length + " files written");
                }*/
            }

            if (i >= 1)
                return true;
            else
                return false;
        }

    }
}
