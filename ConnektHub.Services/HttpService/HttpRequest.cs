using System;
using System.Net;

namespace Prospecta.ConnektHub.Services.HttpService
{
    public class HttpRequest : IHttpRequest
    {
        private static string baseURL = string.Empty;
        public string BaseURL
        {
            get { return baseURL; }
            set { baseURL = value; }
        }

        public virtual string HttpGet(string url)
        {
            url = baseURL + url;
            using (var request = new WebClient())
            {
                try
                {
                    var URI = new Uri(url);
                    request.Headers["User-Agent"] = "XMLHTTP/1.0";
                    var contents = request.DownloadString(URI);
                    return contents;
                }
#pragma warning disable 0168
                catch (Exception ex) { }
#pragma warning disable 0168

                return null;
            }
        }

        public virtual string HttpPost(string url, string jsonString, bool isEncoded)
        {
            string contents = string.Empty;
            url = baseURL + url;
            using (var request = new WebClient())
            {
                try
                {
                    var URI = new Uri(url);
                    request.Headers["User-Agent"] = "XMLHTTP/1.0";
                    if (isEncoded)
                        request.Headers["Content-Type"] = "application/x-www-form-urlencoded";
                    else
                        request.Headers["Content-Type"] = "application/json";
                    contents = request.UploadString(URI, jsonString);
                    return contents;
                }
#pragma warning disable 0168
                catch (Exception ex) { }
#pragma warning disable 0168
            }
            return null;
        }
    }
}
