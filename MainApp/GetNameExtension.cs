using Octokit;

namespace MainApp
{
    public static class GetNameExtension
    {
        public static string GetName(this Issue issue) => issue.Url.Split('/')[7];
        
    }
}
