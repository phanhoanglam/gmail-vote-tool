using Google.Apis.Gmail.v1;
using Tool;

class Program
{
    static readonly string LogPath = "C:\\Users\\lamph\\Desktop\\Tools\\Tool\\Logs.txt";
    static readonly string LogGetTokenPath = "C:\\Users\\lamph\\Desktop\\Tools\\Tool\\GetTokenLogs.txt";

    private List<string> ZohoMail = new List<string>
    {
        "phanhoanglam12061998112233447"
    };

    private List<string> Gmail = new List<string>
    {
        "hieuruoixanhxalo",
        "phanhoanglam120619981122334444",
        "phanhoanglam120619981122334455",
        "phanhoanglam120619981122334466"
    };

    Dictionary<string, GmailService> AddMailToDictionary()
    {
        var dictionary = new Dictionary<string, GmailService>();
        using (StreamWriter writer = new StreamWriter(LogGetTokenPath))
        {
            foreach (var email in Gmail)
            {
                try
                {
                    var service = GmailUtils.GetService($"${email}@gmail.com");
                    dictionary.Add(email, service);
                    writer.WriteLine($"=================> Successed Get Token for email: {email}@gmail.com <=================");
                    writer.Flush();
                }
                catch (Exception)
                {
                    writer.WriteLine($"=================> Failed Get Token for email: {email}@gmail.com <=================");
                    writer.Flush();
                }
            }
        }

        return dictionary;

    }

    static void Main(string[] args)
    {
        var mailDict = new Program().AddMailToDictionary();

        var indexSuccessed = 0;
        var indexFailed = 0;
        using (StreamWriter writer = new StreamWriter(LogPath))
        {
            foreach (var item in mailDict)
            {
                List<string> gmails = new List<string>();
                string gmailName = item.Key;
                string gmailExtensions = "@gmail.com";
                gmails.Add(gmailName + gmailExtensions);
                for (int i = 1; i < gmailName.Length; i++)
                {
                    var aliasGmail = gmailName.Substring(0, i) + "." + gmailName.Substring(i);
                    gmails.Add(aliasGmail + gmailExtensions);
                }

                foreach (var gmail in gmails)
                {
                    writer.WriteLine($"===============> Start email: {gmail} ===============");
                    try
                    {
                        SeleniumUtils.RunAutomation(item.Value, gmail);
                        indexSuccessed++;
                    }
                    catch (Exception ex)
                    {
                        indexFailed++;
                        writer.WriteLine($"----------------- {ex} -----------------");
                        writer.WriteLine($"===============> End email: {gmail} ===============");
                        writer.WriteLine($"============================================================");
                        writer.Flush();
                        continue;
                    }
                    writer.WriteLine($"===============> End email: {gmail} ===============");
                    writer.WriteLine($"============================================================");
                }
            }
            writer.WriteLine($"=================> Successfully vote {indexSuccessed} emails <=================");
            writer.WriteLine($"=================> Failed vote {indexFailed} emails <=================");
            writer.Flush();
        }
    }
}
