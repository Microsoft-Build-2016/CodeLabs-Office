using System;
using System.Threading.Tasks;
using Microsoft.Owin;
using Owin;

[assembly: OwinStartup(typeof(MailCRM.Startup))]

namespace MailCRM
{
    public class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            //setup SignalR
            app.MapSignalR();
        }
    }
}
