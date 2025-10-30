using Microsoft.Exchange.WebServices.Data;
using System;
using System.Threading.Tasks;
using Topshelf.Logging;

public class MailboxMonitorService
{
    // Topshelf heeft hier een logger
    private readonly LogWriter _log = HostLogger.Get<MailboxMonitorService>();

    // Jouw Exchange Service object
    private readonly ExchangeService _service;

    // Timer voor de periodieke controle (elke 5 minuten)
    private System.Timers.Timer _reestablishTimer;

    // Het huidige abonnement ID
    private PushSubscription _currentSubscription;
    private const string CallbackUrl = "https://jouwserver.com/api/ews-callback";

    // Constructor (hier zet je de ExchangeService op en authenticeer je)
    public MailboxMonitorService(ExchangeService service)
    {
        _service = service;
    }

    // Wordt aangeroepen wanneer de service start
    public bool Start()
    {
        _log.Info("MailboxMonitorService start aan het initialiseren...");

        // 1. Maak direct een abonnement aan bij de start
        // Gebruik Task.Run omdat Start synchroon is in Topshelf
        System.Threading.Tasks.Task.Run(() => SetupSubscriptionAsync());

        // 2. Start de periodieke timer
        _reestablishTimer = new System.Timers.Timer(TimeSpan.FromMinutes(5).TotalMilliseconds);
        _reestablishTimer.Elapsed += (sender, e) => ReestablishSubscriptionAsync().Wait();
        _reestablishTimer.Start();

        _log.Info("MailboxMonitorService is gestart. Timer voor hernieuwing actief.");
        return true;
    }

    // Wordt aangeroepen wanneer de service stopt
    public bool Stop()
    {
        _log.Info("MailboxMonitorService aan het stoppen...");

        // Stop de timer
        _reestablishTimer?.Stop();
        _reestablishTimer?.Dispose();

        // Probeer het huidige abonnement op te ruimen
        if (_currentSubscription != null)
        {
            try
            {
                _log.Info($"Abonnement {_currentSubscription.Id} aan het opruimen.");
                // Let op: Unsubscribe is asynchroon in EWS, gebruik Wait() of Task.Run.Wait()
                _currentSubscription.Unsubscribe().Wait();
            }
            catch (Exception ex)
            {
                _log.Error($"Fout bij het opruimen van abonnement: {ex.Message}");
            }
        }

        return true;
    }

    // --- EWS Logica ---

    private async System.Threading.Tasks.Task SetupSubscriptionAsync()
    {
        _log.Info("Aanmaken of hernieuwen van het EWS abonnement...");

        // Optioneel: Unsubscribe van oude abonnementen indien nodig (zie vorige antwoord)
        // ... Logica om oud abonnement op te ruimen ...

        try
        {
            // Fix 1: Use FolderId instead of WellKnownFolderName
            var folderIds = new[] { new FolderId(WellKnownFolderName.Inbox) };

            // Fix 2: Convert CallbackUrl string to Uri
            var callbackUri = new Uri(CallbackUrl);

            // Fix 3: Pass EventType as array
            _currentSubscription = await _service.SubscribeToPushNotifications(
                folderIds,
                callbackUri,
                5,
                "",
                new[] { EventType.NewMail }
            );
            _log.Info($"Nieuw Push Abonnement aangemaakt. ID: {_currentSubscription.Id}");
        }
        catch (Exception ex)
        {
            _log.Error($"FATALE FOUT bij het aanmaken van het abonnement: {ex.Message}");
            // Hier zou je logica moeten toevoegen om te blijven proberen
        }
    }

    private async System.Threading.Tasks.Task ReestablishSubscriptionAsync()
    {
        _log.Info("5-minuten controle: Vernieuwen van het EWS abonnement.");
        // De 5-minuten strategie: we maken proactief een nieuwe aan om stabiliteit te garanderen.
        await SetupSubscriptionAsync();
    }
}