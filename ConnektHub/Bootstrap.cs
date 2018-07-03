using Autofac;
using Prospecta.ConnektHub.Services.HttpService;
using Prospecta.ConnektHub.Services.Modules;
using Prospecta.ConnektHub.Services.RibbonModule;
using Prospecta.ConnektHub.Services.Translation;
using Prospecta.ConnektHub.Services.User;

namespace Prospecta.ConnektHub
{
    public class Bootstrap
    {
        public static void RegisterTypes(ref ContainerBuilder builder)
        {
            builder.RegisterType<UserService>().As<IUserService>(); 

            builder.RegisterType<HttpRequest>().As<IHttpRequest>();
            builder.RegisterType<RibbonService>().As<IRibbonService>();
            builder.RegisterType<ModuleService>().As<IModuleService>();
            builder.RegisterType<TranslationService>().As<ITranslationService>();
        }
    }
}