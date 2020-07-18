using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Hosting;
using System.Diagnostics;
using System.IO;
using Microsoft.AspNetCore.Server.Kestrel.Core;
using System;
using Microsoft.Extensions.Logging;
using NLog.Web;
using Microsoft.AspNetCore;

namespace WebApplication1
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //var pathToExe = Process.GetCurrentProcess().MainModule.FileName;

            //var pathToContentRoot = Path.GetDirectoryName(pathToExe);

            //var host = Host.CreateDefaultBuilder(args)
            //    .UseContentRoot(pathToContentRoot)
            //    .ConfigureWebHostDefaults(webBuilder =>
            //    {
            //        webBuilder.UseStartup<Startup>();
            //    })
            //    .Build();

            //host.Run();

            var logger = NLog.Web.NLogBuilder.ConfigureNLog("nlog.config").GetCurrentClassLogger();

            try
            {
                logger.Info("Init main.");
                BuildWebHost(args).Run();
                //CreateHostBuilder(args).Build().Run();
            }
            catch (Exception e)
            {
                logger.Error(e, "Program was stoped.");
                throw;
            }
            finally
            {
                NLog.LogManager.Shutdown();
            }
        }

        public static IWebHost BuildWebHost(string[] args) =>     
            WebHost.CreateDefaultBuilder(args)
                .UseStartup<Startup>()
                .UseUrls(urls: "http://localhost:8080")
                .UseNLog()
                .Build();
        

        //public static IHostBuilder CreateHostBuilder(string[] args) =>
        //    Host.CreateDefaultBuilder(args)
        //        .ConfigureWebHostDefaults(webBuilder =>
        //        {
        //            webBuilder.UseStartup<Startup>();
        //        })
        //        .ConfigureLogging(logging =>
        //        {
        //            logging.ClearProviders();
        //            logging.SetMinimumLevel(Microsoft.Extensions.Logging.LogLevel.Trace);
        //        })
        //        .UseNLog();
    }
}
