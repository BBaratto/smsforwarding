using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Topshelf;

namespace SMSForwarding
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // 1. Initialiseer de Exchange Service EENMALIG (met authentificatie!)
            // Je moet hier je EWS URL, credentials of OAuth/Modern Auth opzetten.
            var exchangeService = new ExchangeService(ExchangeVersion.Exchange2013_SP1)
            {
                Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx"),
                // Voorbeeld: Gebruik van basis authenticatie (niet aanbevolen, gebruik Modern Auth!)
                Credentials = new WebCredentials("bruno.baratto@outlook.com", "password")
            };

            // Zoek automatisch de EWS URI op
            ExchangeService.AutodiscoverUrl("jouw.emailadres@domein.com", RedirectionUrlValidationCallback);

            // 2. Configureer Topshelf
            var exitCode = HostFactory.Run(x =>
            {
                // Vertel Topshelf welke klasse je service is, en hoe hij moet starten
                x.Service<MailboxMonitorService>(s =>
                {
                    // De constructor wordt aangeroepen en de exchangeService wordt doorgegeven
                    s.ConstructUsing(name => new MailboxMonitorService(exchangeService));

                    // Welke methoden aanroepen bij Start/Stop
                    s.WhenStarted(tc => tc.Start());
                    s.WhenStopped(tc => tc.Stop());
                });

                // Service Metadata
                x.RunAsLocalSystem(); // Draai als lokale systeem gebruiker
                x.SetServiceName("EwsMailboxMonitor");
                x.SetDisplayName("EWS Mailbox Monitoring Service");
                x.SetDescription("Monitors a mailbox for new emails using EWS Push Notifications.");

                // Logging (optioneel, maar aanbevolen)
                x.UseNLog();
            });

            return (int)exitCode;
        }
    }
}
