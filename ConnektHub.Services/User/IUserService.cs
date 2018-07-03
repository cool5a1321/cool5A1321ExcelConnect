using Prospecta.ConnektHub.Core;

namespace Prospecta.ConnektHub.Services.User
{
    public interface IUserService
    {
        string Authenticate(string userName, string password);

        bool AddUserDetails(UserDetails userDetails);
    }
}
