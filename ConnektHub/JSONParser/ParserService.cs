using Prospecta.ConnektHub.Core;
using Newtonsoft.Json;

namespace Prospecta.ConnektHub.JSONParser
{
    public class ParserService
    {
        public static UserDetails Authenticate(string contents)
        {
            var user = JsonConvert.DeserializeObject<UserDetails>(contents);
            return user;
        }
    }
}
