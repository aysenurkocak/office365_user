using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace Office365
{
   
    public class MailIntegration : IMailIntegration
    {
        //  @author: Ayşe Nur Koçak
        //  @Date:   31.03.2020

        public Result res = new Result();
        //string Token = "6LhPtGED4Qhv2pTp84ZExWziyHSSPGX6GX4TF.AeTuC4VZDTf51q0Toz1F2JIx84Ltg21Md9EuWktEwKGFzbyb5jscYspvyhzauq5HfF72LPXTDo.DSNfOlbJ18Jhct4Nwk51iRxPEbfziWu5POhyQqq6SLo4344";

        /// <summary>
        /// Office 365 kullanıcısını mail bilgilerini getirir.
        /// </summary>
        /// <param name="mailNickname">@sample.com.tr olmadan mail nickname yazılır.</param>
        /// <returns></returns>
        public User GetUser(string mailNickname)
        {
            
            string AzureToken = getToken();
            var client = new RestClient("https://graph.microsoft.com/v1.0/users/" + mailNickname + "@sample.com.tr");
            client.Timeout = -1;
            var request = new RestRequest(Method.GET);
            request.AddHeader("SdkVersion", "postman-graph/v1.0");
            request.AddHeader("Authorization", "Bearer " + AzureToken);
            IRestResponse response = client.Execute(request);
             User userList = new User();
            JObject result = JObject.Parse(response.Content);
            try
            {
                if (result != null)
                {
                    foreach (var item in result)
                        if (item.Key == "givenName")
                            userList.name = item.Value.ToString();
                        else if (item.Key == "surname")
                            userList.surname = item.Value.ToString();
                        else if (item.Key == "userPrincipalName")
                            userList.mail = item.Value.ToString();

                }
                log("GetUser", mailNickname, "succes");
            }
            catch(Exception exc)
            {
                log("GetUser", exc.Message, "error");
            }
            
            return userList;


        }


        /// <summary>
        /// Office 365 için lisanssız yeni kullanıcı oluşur.
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public Result UserCreate(User user)
        {
            string AzureToken = getToken();
            Result res = new Result();
            try
            {
                MailUser newUser = new MailUser();
                newUser.accountEnabled = true;
                newUser.userPrincipalName = user.mailNickname + "@sample.com.tr";
                newUser.givenName = user.name;
                newUser.surname = user.surname;
                newUser.displayName = newUser.givenName + " " + newUser.surname;
                newUser.passwordPolicies = "DisablePasswordExpiration";
                newUser.passwordProfile = new passwordProfile
                {
                    forceChangePasswordNextSignIn = false,
                    password = RandomPassword()
                };
                newUser.mailNickname = user.mailNickname;
                newUser.country = "Türkiye";
                string json = JsonConvert.SerializeObject(newUser, Formatting.Indented);
                var client = new RestClient("https://graph.microsoft.com/v1.0/users");
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);
                request.AddHeader("SdkVersion", "postman-graph/v1.0");
                request.AddHeader("Authorization", "Bearer " + AzureToken);
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("application/json", json, ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                res.status = 1;
                res.result = "Success";
                res.mail = newUser.userPrincipalName;
                res.password = newUser.passwordProfile.password;
                log("UserCreate", res.mail +"(" + res.password + ")", "succes");

            }
            catch (Exception exc)
            {
                res.status = 0;
                res.result = exc.Message;
                log("UserCreate", exc.Message, "error");
            }

            return res;
        }

        /// <summary>
        /// Mail adresine göre silme yapar
        /// </summary>
        /// <param name="mail">@sample.com.tr olmadan mail nickname yazılır.</param>
        /// <returns></returns>
        public Result UserDelete(string mail)
        {
            string AzureToken = getToken();
            res = new Result();
            try
            {
                var client = new RestClient("https://graph.microsoft.com/v1.0/users/" + mail +"@sample.com.tr");
                client.Timeout = -1;
                var request = new RestRequest(Method.DELETE);
                request.AddHeader("SdkVersion", "postman-graph/v1.0");
                request.AddHeader("Authorization", "Bearer " + AzureToken);
                IRestResponse response = client.Execute(request);
                Console.WriteLine(response.Content);
                res.status = 1;
                res.result = "Success";
                log("UserDelete", mail + "@sample.com.tr", "succes");
            }
            catch (Exception exc)
            {
                res.status = 0;
                res.result = "Error, " + exc.Message;
                log("UserDelete", exc.Message, "error");
            }
            return res;
        }

        public string RandomPassword(int size = 0)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append(RandomString(4, true));
            builder.Append(RandomNumber(1000, 9999));
            builder.Append(RandomString(2, false));
            return builder.ToString();
        }
        public int RandomNumber(int min, int max)
        {
            Random random = new Random();
            return random.Next(min, max);
        }
        public string RandomString(int size, bool lowerCase)
        {
            StringBuilder builder = new StringBuilder();
            Random random = new Random();
            char ch;
            for (int i = 0; i < size; i++)
            {
                ch = Convert.ToChar(Convert.ToInt32(Math.Floor(26 * random.NextDouble() + 65)));
                builder.Append(ch);
            }
            if (lowerCase)
                return builder.ToString().ToLower();
            return builder.ToString();
        }
        public string getToken()
        {
            var client = new RestClient("https://login.microsoftonline.com/6ddf0661-bc96-4c34-bda2-ecc4ff23d80d/oauth2/v2.0/token");
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            request.AddHeader("SdkVersion", "postman-graph/v1.0");
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            request.AddHeader("Cookie", "x-ms-gateway-slice=prod; stsservicecookie=ests; fpc=ApwJcDYy5T1Og2-YA_RJQKs");
            request.AddParameter("grant_type", "password");
            request.AddParameter("client_id", ConfigurationManager.AppSettings["client_id"]);
            request.AddParameter("client_secret", ConfigurationManager.AppSettings["client_secret"]);
            request.AddParameter("scope", "https://graph.microsoft.com/.default");
            request.AddParameter("userName", ConfigurationManager.AppSettings["username"]);
            request.AddParameter("password", ConfigurationManager.AppSettings["password"]);
            IRestResponse response = client.Execute(request);
            var parsed = JsonConvert.DeserializeObject<Dictionary<string, string>>(response.Content);
            string token = parsed["access_token"];
            return token;
        }

        public void log(string func, string message, string type)
        {

            StreamWriter Log = new StreamWriter(ConfigurationManager.AppSettings["logURL"], true);
            Log.WriteLine("-----------------------------------------");
            Log.WriteLine("Date : " + DateTime.Now.ToString());
            Log.WriteLine(type + " => " + func + " : " + message);
            Log.WriteLine("-----------------------------------------");
            Log.Close();
        }

    }
}
