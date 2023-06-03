using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TfsExporter;

namespace SrsExporter
{
    internal class SrsExporter
    {
        private readonly QueryExecutor qe;

        public SrsExporter() {
            this.qe = new QueryExecutor("appsettings.json");
        }

        public async Task ListEpics()
        {
            var workItems = await this.qe.QueryEpics().ConfigureAwait(false);

            Console.WriteLine("Query Results: {0} items found", workItems.Count);

            // loop though work items and write to console
            foreach (var workItem in workItems)
            {
                Console.WriteLine(
                    "{0}\t{1}\t{2}\t{3}",
                    workItem.Id,
                    workItem.Fields["System.Title"],
                    workItem.Fields["System.Description"],
                    workItem.Fields["Microsoft.VSTS.Common.StackRank"]);
            }
        }
    }
}
