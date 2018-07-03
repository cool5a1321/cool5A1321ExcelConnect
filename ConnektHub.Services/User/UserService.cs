using Prospecta.ConnektHub.Core;
using Prospecta.ConnektHub.Services.HttpService;
using Prospecta.ConnektHub.SQLiteHelper.User;

namespace Prospecta.ConnektHub.Services.User
{
    public class UserService : IUserService
    {
        private IHttpRequest _httpRequest;

        public UserService(IHttpRequest httpRequest)
        {
            _httpRequest = httpRequest;
        }

        public string Authenticate(string userName, string password)
        {
            var url = "restUserValidation/userValidation?userId=" + userName + "&password=" + password;
            
            return _httpRequest.HttpGet(url);
        }

        public bool AddUserDetails(UserDetails userDetails)
        {
            return SQLiteUser.AddUserDetails(userDetails);
        }
    }
}
