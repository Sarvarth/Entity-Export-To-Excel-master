using Octokit;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.IO;
using System.Threading.Tasks;
using System.Globalization;

namespace MainApp
{
    class Program
    {
        public static readonly string GitHubIdentity = Assembly
            .GetEntryAssembly()
            .GetCustomAttribute<AssemblyProductAttribute>()
            .Product;

        static DateTime StartDate { get; set; } = DateTime.Now.AddDays(-10);
        static DateTime EndDate { get; set; } = DateTime.Now;

        static async Task Main(string[] args)
        {
            var prodHeader = new ProductHeaderValue(GitHubIdentity);
            var credentials = new Credentials("7f4d674999cf17062ca99a2b9c2fdbce56222839");
            var enterpriseUrl = "https://github-rd.carefusion.com/vanguard";
            var client = new GitHubClient(prodHeader, new Uri(enterpriseUrl))
            {
                Credentials = credentials
            };

            var currentDirectory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location).Replace("MainApp\\bin\\Debug", "");
            var repoNameFile = "RepositoryNames.txt";
            var repoNamePath = currentDirectory + repoNameFile;

            var allRepoNames = File.ReadAllLines(repoNamePath);

            var prr = new PullRequestRequest
            {
                State = ItemStateFilter.Closed
            };

            var str = "https://github-rd.carefusion.com/vanguard/logistics-storagespaceItem-orchestrator/pulls";
            var abcd = str.Split('/');

            try
            {
                Console.WriteLine("Please mention the start date (dd/MM/yyyy)");
                StartDate = DateTime.Parse(Console.ReadLine());

                Console.WriteLine("Please mention the end date (dd/MM/yyyy)");
                EndDate = DateTime.Parse(Console.ReadLine());
            }
            catch
            {
                Console.WriteLine("You were wrong. Try again!!");
                Environment.Exit(0);
            }

            var excelWorkbookName = "PRDetails";
            var excelWorkSheetName = "PRsList " + StartDate.ToShortDateString() + " to " + EndDate.ToShortDateString();
            var excelWorkbookPath = currentDirectory + excelWorkbookName;

            var excelExport = new ExcelExport(excelWorkbookPath, excelWorkSheetName);

            var rc = new RepositoryCollection();
            foreach (var repoName in allRepoNames)
            {
                rc.Add($"vanguard/{repoName}");
            }

            var abc = new SearchIssuesRequest
            {
                Merged = new DateRange(StartDate, EndDate),
                Type = IssueTypeQualifier.PullRequest,
                Repos = rc
            };

            //var prs = await client.PullRequest.GetAllForRepository("vanguard", repoName, prr);
            //var filteredPrs = prs.Where(pr => pr.Merged && pr.MergedAt.Value <= EndDate && pr.MergedAt.Value >= StartDate).ToList();

            var filteredPrs = await client.Search.SearchIssues(abc);

            var prsToBeAdded = new List<Issue>();

            foreach (var filteredPr in filteredPrs.Items)
            {
                var reviewCommentsCount = (await client.PullRequest.ReviewComment.GetAll("vanguard", filteredPr.GetName(), filteredPr.Number)).Count;
                if (reviewCommentsCount > 0)
                {
                    prsToBeAdded.Add(filteredPr);
                }
            }
            if (prsToBeAdded.Any())
            {
                excelExport.GenerateExcel(ConvertToDataTable<Issue>(prsToBeAdded));
                excelExport.SaveAndCloseExcel();
            }
        }
        
        // T : Generic Class
        static DataTable ConvertToDataTable<T>(List<Issue> models)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);

            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            // Loop through all the properties            
            // Adding Column to our datatable
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names  
                dataTable.Columns.Add(prop.Name);
            }
            dataTable.Columns.Add("Repository Name");

            // Adding Row
            foreach (Issue item in models)
            {
                var values = new object[Props.Length + 1];
                int i;
                for (i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows  
                    values[i] = Props[i].GetValue(item, null);
                }

                values[i] = item.GetName();

                // Finally add value to datatable  
                dataTable.Rows.Add(values);
            }

            dataTable.Columns.Remove("CommentsUrl");
            dataTable.Columns.Remove("EventsUrl");
            dataTable.Columns.Remove("ClosedBy");
            dataTable.Columns.Remove("User");
            dataTable.Columns.Remove("Labels");
            dataTable.Columns.Remove("Assignee");
            dataTable.Columns.Remove("Assignees");
            dataTable.Columns.Remove("Milestone");
            dataTable.Columns.Remove("Comments");
            dataTable.Columns.Remove("PullRequest");
            dataTable.Columns.Remove("ClosedAt");
            dataTable.Columns.Remove("CreatedAt");
            dataTable.Columns.Remove("UpdatedAt");
            dataTable.Columns.Remove("Locked");
            dataTable.Columns.Remove("Repository");
            dataTable.Columns.Remove("Reactions");

            return dataTable;
        }
    }
}
