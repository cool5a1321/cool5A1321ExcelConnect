using Prospecta.ConnektHub.Services.Modules;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace Prospecta.ConnektHub.Controllers
{
    public class ModuleController
    {
        private IModuleService _moduleService;

        public ModuleController(IModuleService moduleService)
        {
            _moduleService = moduleService;
        }

        public Dictionary<string, string> GetModulesList(string userId)
        {
            var dictModules = new Dictionary<string, string>();
            var jsonString = _moduleService.GetModuleList(userId);
            dictModules = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonString);

            return dictModules;
        }
    }
}