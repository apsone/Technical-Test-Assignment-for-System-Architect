using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(TechnicalTest.Startup))]
namespace TechnicalTest
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
