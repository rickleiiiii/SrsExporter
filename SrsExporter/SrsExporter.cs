using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using TfsExporter;

using Microsoft.Office.Interop.Word;
using System.IO;

namespace SrsExporter
{
    internal class SrsExporter
    {
        private readonly QueryExecutor qe;

        public SrsExporter() {
            this.qe = new QueryExecutor("appsettings.json");
        }

        public async System.Threading.Tasks.Task SaveSrs()
        {
            string docPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SRS.docx");

            Application wordApp = new Application();
            wordApp.Visible = true;

            Document doc = wordApp.Documents.Open(docPath);

            // build a list of the fields we want to see
            var fields = new[] { "System.Id", "System.Title", "System.Description" };

            var epics = await this.qe.QueryEpics(fields).ConfigureAwait(false);

            Console.WriteLine("Query Results: {0} items found", epics.Count);

            // loop though work items and write to console
            foreach (var workItem in epics)
            {
                string placeholderTag = "<<EpicTitle>>";

                string title = workItem.Fields["System.Title"].ToString();

                Console.WriteLine(title);

                FindReplace(doc, placeholderTag, title);

                doc.Application.Selection.EndKey(WdUnits.wdStory);
            }

            doc.Save();

            doc.Close();
            wordApp.Quit();

            ReleaseComObject(doc);
            ReleaseComObject(wordApp);
        }

        private void ReleaseComObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void FindReplace(Document doc, string findText, string replaceText)
        {
            Find find = doc.Content.Find;

            find.Text = findText;

            bool found = find.Execute();

            if (found)
            {
                Range range = find.Parent;

                range.InsertBefore(replaceText);
            }
        }

        public async System.Threading.Tasks.Task ListEpics()
        {
            // build a list of the fields we want to see
            var fields = new[] { "System.Id", "System.Title", "System.Description", "Microsoft.VSTS.Common.StackRank" };

            var epics = await this.qe.QueryEpics(fields).ConfigureAwait(false);

            Console.WriteLine("Query Results: {0} items found", epics.Count);

            // loop though work items and write to console
            foreach (var workItem in epics)
            {
                var stackRankString = "9999999999";

                if (workItem.Fields.TryGetValue("Microsoft.VSTS.Common.StackRank", out var stackRankField))
                {
                    if (!string.IsNullOrEmpty(stackRankField?.ToString()))
                    {
                        stackRankString = stackRankField.ToString();
                    }
                }
                Console.WriteLine(
                    "{0}\t{1}\t{2}\t{3}",
                    workItem.Id,
                    workItem.Fields["System.Title"],
                    "", //workItem.Fields["System.Description"],
                    stackRankString
                    );
            }
        }
    }
}
