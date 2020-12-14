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

            var prr = new PullRequestRequest { 
                State = ItemStateFilter.Closed
            };
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
                Console.ReadKey();
            }
            
            var excelWorkbookName = "PRsList " + StartDate.ToShortDateString() + " to " + EndDate.ToShortDateString();
            var excelWorkbookPath = currentDirectory + excelWorkbookName;


            foreach (var repoName in allRepoNames)
            {
                Console.WriteLine($"Checking PRs for {repoName}");
                var prs = await client.PullRequest.GetAllForRepository("vanguard", repoName, prr);
                var filteredPrs = prs.Where(pr => pr.Merged && pr.MergedAt.Value <= EndDate && pr.MergedAt.Value >= StartDate).ToList();
                var prsToBeAdded = new List<PullRequest>();
                foreach(var filteredPr in filteredPrs)
                {
                    var reviewCommentsCount = (await client.PullRequest.ReviewComment.GetAll("vanguard", repoName, filteredPr.Number)).Count;
                    if(reviewCommentsCount > 0)
                    {
                        prsToBeAdded.Add(filteredPr);
                    }
                }
                if (prsToBeAdded.Any())
                {
                    Console.WriteLine($"Found Repo {repoName} with Review Comments");
                    ExcelExport.GenerateExcel(ConvertToDataTable(prsToBeAdded, repoName.Substring(0, 30)), excelWorkbookPath);
                }
            }
         
            ExcelExport.SaveAndCloseExcel();
        }


        // T : Generic Class
        static DataTable ConvertToDataTable<T>(List<T> models, string sheetName)
        {
            DataTable dataTable = new DataTable(sheetName);

            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            // Loop through all the properties            
            // Adding Column to our datatable
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names  
                dataTable.Columns.Add(prop.Name);
            }

            // Adding Row
            foreach (T item in models)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows  
                    values[i] = Props[i].GetValue(item, null);
                }
                // Finally add value to datatable  
                dataTable.Rows.Add(values);
            }
            dataTable.Columns.Remove("NodeId");
            dataTable.Columns.Remove("HtmlUrl");
            dataTable.Columns.Remove("DiffUrl");
            dataTable.Columns.Remove("PatchUrl");
            dataTable.Columns.Remove("IssueUrl");
            dataTable.Columns.Remove("StatusesUrl");
            dataTable.Columns.Remove("Head");
            dataTable.Columns.Remove("Base");
            dataTable.Columns.Remove("User");
            dataTable.Columns.Remove("Assignee");
            dataTable.Columns.Remove("Assignees");
            dataTable.Columns.Remove("Milestone");
            dataTable.Columns.Remove("Draft");
            dataTable.Columns.Remove("Merged");
            dataTable.Columns.Remove("Mergeable");
            dataTable.Columns.Remove("MergeableState");
            dataTable.Columns.Remove("MergedBy");
            dataTable.Columns.Remove("MergeCommitSha");
            dataTable.Columns.Remove("Comments");
            dataTable.Columns.Remove("Commits");
            dataTable.Columns.Remove("Additions");
            dataTable.Columns.Remove("Deletions");
            dataTable.Columns.Remove("ChangedFiles");
            dataTable.Columns.Remove("Locked");
            dataTable.Columns.Remove("MaintainerCanModify");
            dataTable.Columns.Remove("RequestedReviewers");
            dataTable.Columns.Remove("RequestedTeams");
            dataTable.Columns.Remove("Labels");
            return dataTable;
        }
    }
}
