using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util;
using Google.Apis.Util.Store;

public static class GmailUtils
{
    static readonly string TokenStorePath = "C:\\Users\\lamph\\Desktop\\Tools\\Tool\\TokenStore.json";
    static readonly string ClientSecretPath = "C:\\Users\\lamph\\Desktop\\Tools\\Tool\\credential.json";
    static string[] Scopes = { GmailService.Scope.GmailReadonly, GmailService.Scope.GmailModify };
    static string ApplicationName = "Gmail API";

    public static GmailService GetService(string user)
    {
        UserCredential credential;

        // Đường dẫn đến tệp refresh token
        using (FileStream stream = new FileStream(ClientSecretPath, FileMode.Open, FileAccess.Read))
        {
            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.Load(stream).Secrets,
                Scopes,
                user,
                CancellationToken.None,
                new FileDataStore(TokenStorePath, true)).Result;
        }

        // Kiểm tra xem AccessToken còn hợp lệ không
        if (!credential.Token.IsExpired(SystemClock.Default))
        {
            // AccessToken còn hợp lệ, trả về dịch vụ Gmail
            GmailService service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }
        else
        {
            // AccessToken đã hết hạn, lấy lại AccessToken từ RefreshToken
            if (credential.RefreshTokenAsync(CancellationToken.None).Result)
            {
                // Lấy lại AccessToken thành công, trả về dịch vụ Gmail
                GmailService service = new GmailService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                return service;
            }
            else
            {
                // Lỗi khi lấy lại AccessToken từ RefreshToken
                throw new Exception("Không thể lấy lại AccessToken từ RefreshToken.");
            }
        }
    }

    public static (string, string) GetCodeInMail(GmailService service)
    {
        UsersResource.MessagesResource.ListRequest ListRequest = service.Users.Messages.List("me");
        ListRequest.LabelIds = "INBOX";
        ListRequest.IncludeSpamTrash = false;
        ListRequest.Q = "from:info@vnba.com.vn is:unread"; //ONLY FOR UNDREAD EMAIL'S...

        //GET ALL EMAILS
        ListMessagesResponse ListResponse = ListRequest.Execute();
        var index = 0;
        while (ListResponse.Messages == null)
        {
            if (index == 10)
            {
                return (null, null);
            }
            System.Threading.Thread.Sleep(1000);
            index++;
            ListResponse = ListRequest.Execute();
        }
        var msgID = ListResponse.Messages.FirstOrDefault().Id;

        Message message = service.Users.Messages.Get("me", msgID).Execute();
        var valueHeader = message.Payload.Headers.FirstOrDefault(x => x.Name == "Subject").Value;
        var code = valueHeader.Split(":").Last().Trim();

        return (msgID, code);
    }

    public static void MsgMarkAsRead(GmailService service, string MsgId)
    {
        //MESSAGE MARKS AS READ AFTER READING MESSAGE
        ModifyMessageRequest mods = new ModifyMessageRequest();
        mods.AddLabelIds = null;
        mods.RemoveLabelIds = new List<string> { "UNREAD" };
        service.Users.Messages.Modify(mods, "me", MsgId).Execute();
    }
}
