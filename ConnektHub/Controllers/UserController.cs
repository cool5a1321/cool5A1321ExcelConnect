using Newtonsoft.Json;
using Prospecta.ConnektHub.Core;
using Prospecta.ConnektHub.Services.User;

namespace Prospecta.ConnektHub.Controllers
{
    public class UserController
    {
        private IUserService _login;
        #region Constructors
        public UserController(IUserService login)
        {
            _login = login;
        }
        #endregion
        /// <summary>
        /// This method is used to check if the login status is success or failure
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        /// <param name="fullName"></param>
        /// <returns></returns>
        public int UserLogin(string userName, string password, out string fullName)
        {
            fullName = "";
            var jsonString = _login.Authenticate(userName, password);
            if (!string.IsNullOrEmpty(jsonString))
            {
                var user = JsonConvert.DeserializeObject<UserDetails>(jsonString);
                if (user.status.Equals("valid"))
                {
                    fullName = user.fName + " " + user.lname;
                    return 1;
                }
            }
            else
            { return 2; }
            return 0;
        }
        /// <summary>
        /// This method is used to add the user details to sqlite database
        /// </summary>
        /// <param name="userDetails"></param>
        /// <returns></returns>
        public bool AddUserDetails(UserDetails userDetails)
        {
            return _login.AddUserDetails(userDetails);
        }
    }
}