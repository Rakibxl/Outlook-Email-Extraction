using Microsoft.Graph;
using System;
using System.Net.Http.Headers;

namespace OutlookEmailExtraction
{
    class Program
    {
        static void Main(string[] args)
        {
            var client = new GraphServiceClient(new DelegateAuthenticationProvider(async request => {
                request.Headers.Authorization = new AuthenticationHeaderValue("bearer", "eyJ0eXAiOiJKV1QiLCJub25jZSI6IlFqMm1YR1VLRTNXbkZLVzF3aUhEYTdlYVc4ZG9aNHNIZ1FRaHUtVXJONXciLCJhbGciOiJSUzI1NiIsIng1dCI6IlNzWnNCTmhaY0YzUTlTNHRycFFCVEJ5TlJSSSIsImtpZCI6IlNzWnNCTmhaY0YzUTlTNHRycFFCVEJ5TlJSSSJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8xNDcxZmExMS04MjhlLTQ5ZGUtOWViNC1hMjFhYTFlZjhiYTYvIiwiaWF0IjoxNTkzMjg2MDA3LCJuYmYiOjE1OTMyODYwMDcsImV4cCI6MTU5MzI4OTkwNywiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFTUUEyLzhRQUFBQUJQN0Q2cWlUSmNadW5yTGtqWDNJd3pUV2h1Qzl3K3BoMEppcm1pdlM5Smc9IiwiYW1yIjpbInB3ZCJdLCJhcHBfZGlzcGxheW5hbWUiOiJHcmFwaCBleHBsb3JlciIsImFwcGlkIjoiZGU4YmM4YjUtZDlmOS00OGIxLWE4YWQtYjc0OGRhNzI1MDY0IiwiYXBwaWRhY3IiOiIwIiwiZmFtaWx5X25hbWUiOiJBSE1FRCIsImdpdmVuX25hbWUiOiJSYWtpYiIsImlwYWRkciI6IjE1MS40OC4xNzEuNDMiLCJuYW1lIjoiQUhNRUQsIFJha2liIiwib2lkIjoiNmFlNGY3OWEtODY1MS00ZWJiLTk1MjItMTEyZGQzNjIzZGFmIiwib25wcmVtX3NpZCI6IlMtMS01LTIxLTE3NDY0OTExMDMtOTczNTA4OTc1LTM0NzMxNDY2MjAtMTI2NjYzIiwicGxhdGYiOiIzIiwicHVpZCI6IjEwMDMzRkZGQTczNjlCN0UiLCJzY3AiOiJDYWxlbmRhcnMuUmVhZFdyaXRlIENvbnRhY3RzLlJlYWRXcml0ZSBEZXZpY2VNYW5hZ2VtZW50QXBwcy5SZWFkLkFsbCBEZXZpY2VNYW5hZ2VtZW50QXBwcy5SZWFkV3JpdGUuQWxsIERldmljZU1hbmFnZW1lbnRDb25maWd1cmF0aW9uLlJlYWQuQWxsIERldmljZU1hbmFnZW1lbnRDb25maWd1cmF0aW9uLlJlYWRXcml0ZS5BbGwgRGV2aWNlTWFuYWdlbWVudE1hbmFnZWREZXZpY2VzLlByaXZpbGVnZWRPcGVyYXRpb25zLkFsbCBEZXZpY2VNYW5hZ2VtZW50TWFuYWdlZERldmljZXMuUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudE1hbmFnZWREZXZpY2VzLlJlYWRXcml0ZS5BbGwgRGV2aWNlTWFuYWdlbWVudFJCQUMuUmVhZC5BbGwgRGV2aWNlTWFuYWdlbWVudFJCQUMuUmVhZFdyaXRlLkFsbCBEZXZpY2VNYW5hZ2VtZW50U2VydmljZUNvbmZpZy5SZWFkLkFsbCBEZXZpY2VNYW5hZ2VtZW50U2VydmljZUNvbmZpZy5SZWFkV3JpdGUuQWxsIEZpbGVzLlJlYWRXcml0ZS5BbGwgTWFpbC5SZWFkIE1haWwuUmVhZEJhc2ljIE1haWwuUmVhZFdyaXRlIE5vdGVzLlJlYWRXcml0ZS5BbGwgb3BlbmlkIFBlb3BsZS5SZWFkIHByb2ZpbGUgU2l0ZXMuUmVhZFdyaXRlLkFsbCBUYXNrcy5SZWFkV3JpdGUgVXNlci5SZWFkIFVzZXIuUmVhZEJhc2ljLkFsbCBVc2VyLlJlYWRXcml0ZSBlbWFpbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6IlJSS3loUVhIQmdBMWVPYUtIQ2g2UF9KM0dWUTE4ajQ5bVdUbE9FSnFzREUiLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiRVUiLCJ0aWQiOiIxNDcxZmExMS04MjhlLTQ5ZGUtOWViNC1hMjFhYTFlZjhiYTYiLCJ1bmlxdWVfbmFtZSI6IlJha2liLkFITUVEQGVjb25vY29tLmNvbSIsInVwbiI6IlJha2liLkFITUVEQGVjb25vY29tLmNvbSIsInV0aSI6Ik5qeHhfbXg4blVxbEtwelhxa2RLQVEiLCJ2ZXIiOiIxLjAiLCJ4bXNfc3QiOnsic3ViIjoiQWs2eUsza3d2SkpNZ3UyRGRpelk4WmxWb1JRVjVsa0JUMGkzSXdRX3NvNCJ9LCJ4bXNfdGNkdCI6MTMxODkyMjE1NH0.ovfyMc561MBa-VYKZrMeqZLKowhkIxiosIPpVoKHgPqbJNYHWhzZ36QNg2HnoRfcpCEKd9QSkdlebCYWdEwW-1nABPHLYbU5ASH3D0YHqQQ6Hpkv-ROVrImtF-VB3PKPBtePJUgiwlSYxlKoMqYnPjl_rPd88bXuPG9XlnPA8VxykWNl9dIY_AhnOqrfB1rztJwyxCAfsXvsiREDwmn7xz6zZYJg-MH0ewJSSMxK8Pzgnmr7bqwRkV09ABGHAz_rdJNK3sOsx3uO8jbLqyCVKGTfQy8gibKQJLD9z8KNdKIfnOkcFR-xZ8qctJhfiwBdYzuJqgA3MgP6kYce4k9sRg");
            }));

            //var messages =  client.Me.Messages.Request().Top(10).GetAsync().Result;

            // foreach (var item in messages)
            // {
            //     Console.WriteLine(item.Subject + " " + item.BodyPreview + " " + item.Sender + " " + item.SentDateTime);
            // }

            // Console.Read();


            // below example is only for getting email of those are sent by samsung

            string filter = "sender/emailaddress/address eq 'sales.op5@samsung.com'";

            var messages = client.Me.MailFolders.Inbox.Messages.Request().Filter(filter).Top(10).GetAsync().Result;

            foreach (var item in messages)
            {
                Console.WriteLine(item.Subject + " " + item.BodyPreview + " " + item.Sender + " " + item.SentDateTime);
            }

            Console.Read();
        }
    }
}
