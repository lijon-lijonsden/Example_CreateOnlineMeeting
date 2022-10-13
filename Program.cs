using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System.Globalization;
using System.Net.Http.Headers;
using System.Net.Http.Json;

namespace ConsoleApp2
{
    internal class Program
    {
        private static string BaseUrl = "https://graph.microsoft.com/v1.0";
        private static string[] Scopes = new string[] { "https://graph.microsoft.com/.default" };

        static async Task Main(string[] args)
        {
            /// Fields required
            string applicationId = "";
            string applicationSecretValue = "";
            string tenantIdOrName = "";
            string OrganizerAzureId = "";
            ///

            var instance = "https://login.microsoftonline.com/{0}";
            var authority = String.Format(CultureInfo.InvariantCulture, instance, tenantIdOrName);

            var meetingProperties = new Microsoft.Graph.OnlineMeeting()
            {
                StartDateTime = DateTime.UtcNow.AddHours(1),
                EndDateTime = DateTime.UtcNow.AddHours(2),
                Subject = "Hello World",
                AllowMeetingChat = Microsoft.Graph.MeetingChatMode.Limited,
                RecordAutomatically = true,
                AllowAttendeeToEnableCamera = false
            };

            var token = await CollectAccessToken(clientId: applicationId, clientSecret: applicationSecretValue, tenant: tenantIdOrName, authority: authority);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Token generated");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(token);
            Console.ResetColor();
            Console.WriteLine();
            Console.WriteLine("Creating meeting....");
            await CreateOnlineMeeting(token, OrganizerAzureId, meetingProperties);
            Console.ReadKey();
        }

        public static async Task CreateOnlineMeeting(string accessToken, string creatorAzureId, Microsoft.Graph.OnlineMeeting properties)
        {

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var url = $"{BaseUrl}/users/{creatorAzureId}/onlineMeetings";
                HttpResponseMessage response = await client.PostAsJsonAsync<Microsoft.Graph.OnlineMeeting>(@$"{url}", properties);

                if (response.IsSuccessStatusCode)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    var json = JsonConvert.DeserializeObject<Microsoft.Graph.OnlineMeeting>(content);
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine();
                    Console.WriteLine("New Meeting Created");
                    Console.ResetColor();
                    Console.WriteLine($"Id: {json.Id}");
                    Console.WriteLine($"Subject: {json.Subject}");
                    Console.WriteLine($"StartDate: {json.StartDateTime}");
                    Console.WriteLine($"EndDate: {json.EndDateTime}");
                    Console.WriteLine($"Meeting chat: {json.AllowMeetingChat}");
                    Console.WriteLine($"Record automatically: {json.RecordAutomatically}");
                    Console.WriteLine($"URL to join: {json.JoinWebUrl}");
                    Console.WriteLine($"Allow camera: {json.AllowAttendeeToEnableCamera}");
                    Console.WriteLine($"Created: {json.CreationDateTime}");
                    Console.WriteLine($"Organizer: {json.Participants.Organizer.Upn}");

                    for (int i = 0; i < json.Participants.Attendees.Count(); i++)
                    {
                        Console.WriteLine($"Attendee{i}: {json.Participants.Attendees?.ElementAt(i)?.Upn}");
                    }
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Error {response.StatusCode} - {response.ReasonPhrase}");
                    Console.ResetColor();

                    var content = await response.Content?.ReadAsStringAsync();

                    if (content != null)
                        Console.WriteLine($"Content: {content}");
                }
            }
        }


        public static async Task<string> CollectAccessToken(string clientId, string clientSecret, string authority, string tenant, string[] scopes = null)
        {
            if (scopes != null)
                Scopes = scopes;

            IConfidentialClientApplication app;

            app = ConfidentialClientApplicationBuilder.Create(clientId)
            .WithClientSecret(clientSecret)
            .WithAuthority(new Uri(authority))
            .Build();

            AuthenticationResult result = null;
            result = await app.AcquireTokenForClient(Scopes).ExecuteAsync();

            return result?.AccessToken;
        }
    }
}