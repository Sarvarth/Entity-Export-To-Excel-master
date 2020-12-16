using Octokit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

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

            var currentDirectory = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location).Replace("MainApp\\bin\\Debug", "");
            var credsFilePath = currentDirectory + "Creds.txt";
            string username = string.Empty;
            string password = string.Empty;

            Console.WriteLine("Welcome to the Github Repo PR Extractor\n");

            if (!File.Exists(credsFilePath))
            {
                Console.Write("Would you like us to save your username or password? \nYou won't have to enter it next time (y/n):- ");
                var saveCredentials = Console.ReadLine();

                Console.Write("Please enter your BD email address:- ");
                username = Console.ReadLine();

                Console.Write("Please enter your password:- ");
                password = Console.ReadLine();

                if (saveCredentials.ToLower() == "y")
                {
                    CreateCredsFile(username, password, credsFilePath);
                    Console.WriteLine("\n\nALERT!! A Creds.txt file is created with your credentials. Keep them safe");
                }
            }
            else
            {
                Console.WriteLine("Logging you in!!");
                var creds = File.ReadAllLines(credsFilePath);
                username = creds[0];
                password = creds[1];

                Console.WriteLine($"Welcome {username}\n");
            }

            var prodHeader = new ProductHeaderValue(GitHubIdentity);
            var credentials = new Credentials(username, password);
            var enterpriseUrl = "https://github-rd.carefusion.com/vanguard";
            var client = new GitHubClient(prodHeader, new Uri(enterpriseUrl))
            {
                Credentials = credentials
            };

            var repoNameFile = "RepositoryNames.txt";
            var repoNamePath = currentDirectory + repoNameFile;
            var allRepoNames = File.ReadAllLines(repoNamePath);

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
                Environment.Exit(0);
            }

            var excelWorkbookName = "PRsList " + StartDate.ToShortDateString() + " to " + EndDate.ToShortDateString();
            var excelWorkSheetName= "PRDetails";
            var excelWorkbookPath = currentDirectory + excelWorkbookName;

            var excelExport = new ExcelExport(excelWorkbookPath, excelWorkSheetName);

            var rc = new RepositoryCollection();
            foreach (var repoName in allRepoNames)
            {
                rc.Add($"vanguard/{repoName}");
            }

            var searchIssuesRequest = new SearchIssuesRequest
            {
                Merged = new DateRange(StartDate, EndDate),
                Type = IssueTypeQualifier.PullRequest,
                Repos = rc,
                SortField = IssueSearchSort.Merged,
                Page = 1,
                PerPage = 100
            };
            
            var filteredPrs = await client.Search.SearchIssues(searchIssuesRequest);
            var totalNumberOfPrs = filteredPrs.TotalCount;

            Console.WriteLine($"Found {totalNumberOfPrs} of PRs within our range\n");

            var totalFilteredPRs = (List<Issue>)filteredPrs.Items;
            
            // Max page size is 100
            while ((totalNumberOfPrs / (searchIssuesRequest.Page * 100)) >= 1)
            {
                // 403 = 1 * 100, 2 * 100, 3 * 100, 4 * 100, 5 * 3
                searchIssuesRequest.PerPage = ((totalNumberOfPrs - searchIssuesRequest.Page * 100) < 100) ? (totalNumberOfPrs - searchIssuesRequest.Page * 100) : 100;
                searchIssuesRequest.Page += 1;
                filteredPrs = await client.Search.SearchIssues(searchIssuesRequest);
                totalFilteredPRs.AddRange(filteredPrs.Items);
            }

            var prsToBeAdded = new List<Issue>();

            Console.WriteLine("Checking if any PRs have any review comments...");

            foreach (var filteredPr in totalFilteredPRs)
            {
                var reviewCommentsCount = (await client.PullRequest.ReviewComment.GetAll("vanguard", filteredPr.GetName(), filteredPr.Number)).Count;
                if (reviewCommentsCount > 0)
                {
                    prsToBeAdded.Add(filteredPr);
                }
            }
            if (prsToBeAdded.Any())
            {
                Console.WriteLine($"Found {prsToBeAdded.Count} PRs with CR Comments");
                excelExport.GenerateExcel<Issue>(prsToBeAdded);
                excelExport.SaveAndCloseExcel();
            }
            else
            {
                Console.WriteLine("No PR found in our range with CR comments.\nHave a Nice Day!!");
                Console.ReadKey();
            }
        }

        private static void CreateCredsFile(string username, string password, string filePath)
        {
            using (StreamWriter sw = File.CreateText(filePath))
            {
                sw.WriteLine(username);
                sw.WriteLine(password);
            }
        }
    }
}
