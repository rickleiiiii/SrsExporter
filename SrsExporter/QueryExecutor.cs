using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;

using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;

using Microsoft.Extensions.Configuration;
using System.IO;

namespace TfsExporter
{
    class QueryExecutor
    {
        private readonly Uri uri;
        private readonly string project;
        private readonly string username;
        private readonly string password;

        /// <summary>
        ///     Initializes a new instance of the <see cref="QueryExecutor" /> class.
        /// </summary>
        /// <param name="orgName">
        ///     An organization in Azure DevOps Services. If you don't have one, you can create one for free:
        ///     <see href="https://go.microsoft.com/fwlink/?LinkId=307137" />.
        /// </param>
        /// <param name="personalAccessToken">
        ///     A Personal Access Token, find out how to create one:
        ///     <see href="/azure/devops/organizations/accounts/use-personal-access-tokens-to-authenticate?view=azure-devops" />.
        /// </param>
        public QueryExecutor(string configFile)
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile(configFile)
                .Build();

            var username = config["Username"];
            var password = config["Password"];
            var uri = config["Uri"];
            var collection = config["Collection"];
            var project = config["Project"];

            if (uri.EndsWith("/"))
            {
                this.uri = new Uri(String.Format("{0}{1}", uri, collection));
            }
            else
            {
                this.uri = new Uri(String.Format("{0}/{1}", uri, collection));
            }
            this.project = project;
            this.username = username;
            this.password = password;

        }

        /// <summary>
        ///     Execute a WIQL (Work Item Query Language) query to return a list of open bugs.
        /// </summary>
        /// <param name="project">The name of your project within your organization.</param>
        /// <returns>A list of <see cref="WorkItem"/> objects representing all the open bugs.</returns>
        public async Task<IList<WorkItem>> QueryOpenBugs()
        {
            NetworkCredential netCre = new NetworkCredential(this.username, this.password);
            WindowsCredential winCre = new WindowsCredential(netCre);

            // create a wiql object and build our query
            var wiql = new Wiql()
            {
                // NOTE: Even if other columns are specified, only the ID & URL are available in the WorkItemReference
                Query = "Select [Id] " +
                        "From WorkItems " +
                        "Where [Work Item Type] = 'Epic' " +
                        "And [System.TeamProject] = '" + this.project + "' " +
                        "Order By [State] Asc, [Changed Date] Desc",
            };

            // create instance of work item tracking http client
            using (var httpClient = new WorkItemTrackingHttpClient(this.uri, winCre))
            {
                // execute the query to get the list of work items in the results
                var result = await httpClient.QueryByWiqlAsync(wiql).ConfigureAwait(false);
                var ids = result.WorkItems.Select(item => item.Id).ToArray();

                // some error handling
                if (ids.Length == 0)
                {
                    return Array.Empty<WorkItem>();
                }

                if (ids.Length >= 10)
                {
                    ids = ids.Take<int>(10).ToArray();
                }

                // build a list of the fields we want to see
                var fields = new[] { "System.Id", "System.Title", "System.State" };

                // get work items for the ids found in query
                return await httpClient.GetWorkItemsAsync(ids, fields, result.AsOf).ConfigureAwait(false);
            }
        }

        /// <summary>
        ///     Execute a WIQL (Work Item Query Language) query to print a list of open bugs.
        /// </summary>
        /// <param name="project">The name of your project within your organization.</param>
        /// <returns>An async task.</returns>
        public async Task PrintOpenBugsAsync()
        {
            var workItems = await this.QueryOpenBugs().ConfigureAwait(false);

            Console.WriteLine("Query Results: {0} items found", workItems.Count);

            // loop though work items and write to console
            foreach (var workItem in workItems)
            {
                Console.WriteLine(
                    "{0}\t{1}\t{2}",
                    workItem.Id,
                    workItem.Fields["System.Title"],
                    workItem.Fields["System.State"]);
            }
        }

        public async Task<IList<WorkItem>> QueryEpics()
        {
            NetworkCredential netCre = new NetworkCredential(this.username, this.password);
            WindowsCredential winCre = new WindowsCredential(netCre);

            // create a wiql object and build our query
            var wiql = new Wiql()
            {
                // NOTE: Even if other columns are specified, only the ID & URL are available in the WorkItemReference
                Query = "Select [Id] " +
                        "From WorkItems " +
                        "Where [Work Item Type] = 'Epic' " +
                        "And [System.TeamProject] = '" + this.project + "' " +
                        "Order By [Microsoft.VSTS.Common.StackRank]",
            };

            // create instance of work item tracking http client
            using (var httpClient = new WorkItemTrackingHttpClient(this.uri, winCre))
            {
                // execute the query to get the list of work items in the results
                var result = await httpClient.QueryByWiqlAsync(wiql).ConfigureAwait(false);
                var ids = result.WorkItems.Select(item => item.Id).ToArray();

                // some error handling
                if (ids.Length == 0)
                {
                    return Array.Empty<WorkItem>();
                }

                if (ids.Length >= 10)
                {
                    ids = ids.Take<int>(10).ToArray();
                }

                // build a list of the fields we want to see
                var fields = new[] { "System.Id", "System.Title", "System.Description", "Microsoft.VSTS.Common.StackRank" };

                // get work items for the ids found in query
                return await httpClient.GetWorkItemsAsync(ids, fields, result.AsOf).ConfigureAwait(false);
            }
        }
    }
}
