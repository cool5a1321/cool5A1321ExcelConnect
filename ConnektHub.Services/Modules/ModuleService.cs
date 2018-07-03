using Prospecta.ConnektHub.Services.HttpService;

namespace Prospecta.ConnektHub.Services.Modules
{
    public class ModuleService : IModuleService
    {
        private IHttpRequest _httpRequest;

        public ModuleService(IHttpRequest httpRequest)
        {
            _httpRequest = httpRequest;
        }

        public string GetModuleList(string userId)
        {
            var url = "restObjectList/getModuleList?userId=" + userId + "&source=excel";
            return _httpRequest.HttpGet(url);
        }
    }
}
