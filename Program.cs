using System;
using System.IO;
using NPOI.POIFS.FileSystem;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.Model;
using NPOI.XWPF.Extractor;
using System.Collections.Generic;

namespace npoi.test
{
    class Program
    {

        // define the "template fields" - nothing fancy, just pattern matching
        static HashSet<string> tokens = new HashSet<string> { "<tem-firstname>", "<tem-lastname>", "<tem-startdate>" };

        static void Main(string[] args)
        {
            var here = AppContext.BaseDirectory;
            using (var input = new FileStream(@"docxtemplate.docx", FileMode.Open))
            {
                var doc = new XWPFDocument(input);

                //for standard paragraphs of text
                foreach (XWPFParagraph p in doc.Paragraphs)
                {
                    var runs = p.Runs;
                    if (runs != null)
                    {
                        foreach (XWPFRun r in runs)
                        {
                            r.SetText(ReplaceTokens(r.Text));
                        }
                    }
                }

                //for text within tables
                foreach (XWPFTable tbl in doc.Tables)
                {
                    foreach (XWPFTableRow row in tbl.Rows)
                    {
                        foreach (XWPFTableCell cell in row.GetTableCells())
                        {
                            foreach (XWPFParagraph p in cell.Paragraphs)
                            {
                                foreach (XWPFRun r in p.Runs)
                                {
                                    r.SetText(ReplaceTokens(r.Text));
                                }
                            }
                        }
                    }
                }

                // write the new document
                using (var fs = new FileStream(Path.Combine(here, "output2.docx"), FileMode.Create, FileAccess.Write))
                {
                    doc.Write(fs);
                }
            }
        }

        //horrible, but for the purposes of the demo, it works ok ish
        static string ReplaceTokens(string text)
        {
            foreach (var token in tokens)
            {
                switch (token)
                {
                    case "<tem-firstname>":
                        {
                            text = text.Replace(token, "JAYNE");
                            break;
                        }
                    case "<tem-lastname>":
                        {
                            text = text.Replace(token, "DOE");
                            break;
                        }
                    case "<tem-startdate>":
                        {
                            text = text.Replace(token, "1st April 2019");
                            break;
                        }
                }
            }
            return text;
        }
    }
}