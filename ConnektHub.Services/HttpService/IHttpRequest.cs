namespace Prospecta.ConnektHub.Services.HttpService
{
    public interface IHttpRequest
    {
        string BaseURL { get; set; }
        string HttpGet(string url);
        string HttpPost(string url, string jsonString, bool isEncoded);
    }
}
