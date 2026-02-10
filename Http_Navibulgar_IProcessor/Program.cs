using AuditCodeCoreMessage;
using LeS.FileStore;
using System.Reflection;
using AuditCodeCoreMessage;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace Http_Navibulgar_IProcessor
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                AuditCodeCoreMessage.AuditMessageData.LogText = Environment.NewLine + "-------Processor Started!---------";
                using IHost host = _createHostBuilder(args).Build();
                var wrapper = host.Services.GetRequiredService<Navibulgar_Routine>();
                wrapper.StartingProcess();
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                string n = Assembly.GetEntryAssembly()?.Location;
                string ProcessorName = Path.GetFileNameWithoutExtension(n);
                AuditCodeCoreMessage.AuditMessageData.LogText = "Exception in Main : " + ex.GetBaseException().ToString();
                AuditCodeCoreMessage.AuditMessageData.CreateAuditFile("", ProcessorName, "", AuditMessageData.UNEXPECTED_ERROR, "", "", "", " Unexpected exception in Main : " + ex.Message);
                AuditCodeCoreMessage.AuditMessageData.LogText = "Processor completed!";
                AuditCodeCoreMessage.AuditMessageData.LogText = "----------------------------------------------------------";
            }
            static IHostBuilder _createHostBuilder(string[] args) =>
                Host.CreateDefaultBuilder(args)
                    .ConfigureServices((context, services) =>
                    {
                        try
                        {
                            services.AddSingleton<LeSFileStore>(sp =>
                            {
                                var fs = new LeSFileStore();
                                fs.GetApplicationSettings();
                                return fs;
                            });
                            services.AddScoped<Navibulgar_Routine>();
                        }
                        catch (Exception ex)
                        {
                            AuditCodeCoreMessage.AuditMessageData.LogText = $"Error initializing LeSFileStore: {ex.Message}";
                            throw;
                        }
                    });
        }
    }
}
