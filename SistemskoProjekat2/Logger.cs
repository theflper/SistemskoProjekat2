using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SistemskoProjekat2
{
    internal class Logger
    {
        // Thread-safe bafer za poruke
        private static readonly BlockingCollection<string> _queue = new BlockingCollection<string>();
        static Logger()
        {
            // Pokrećemo LongRunning Task koji će raditi u pozadini dok god server živi
            Task.Factory.StartNew(() =>
            {
                // GetConsumingEnumerable() automatski blokira (uspava) ovu nit ako je red prazan.
                // Čim neka nit ubaci poruku preko .Add(), ova petlja se budi, uzima je i ispisuje.
                foreach (var logMessage in _queue.GetConsumingEnumerable())
                {
                    Console.WriteLine(logMessage);
                }
            }, TaskCreationOptions.LongRunning);//služi za kreiranje zadataka koji će se izvršavati dugo vremena,
                                                //poput beskonačnih petlji ili pozadinskih radnika.
                                                //Ova opcija obično uzrokuje da se zadatak izvršava na zasebnoj niti,
                                                //a ne na ThreadPool niti,
                                                //što je idealno za naš scenario logovanja koji traje dok god server radi.
        }
        // Metoda koju pozivaju radne niti - samo ubace poruku u bafer i idu dalje
        public static void Log(string message)
        {
            _queue.Add($"[{DateTime.Now:HH:mm:ss}] {message}");
        }
    }
}
/*
Koristio sam static klasu jer naš loger nema potrebu za stanjem objekta,
ne treba da implementira nikakve interfejse (poput ILogger),
niti nam je potrebno fleksibilno upravljanje njegovim životnim vekom.
static klasa u .NET-u nam garantuje da će se statički konstruktor izvršiti bezbedno
za niti (thread-safe) pri prvom pristupu i obezbeđuje globalno dostu
*/
/*
 "Razmišljao sam o korišćenju eksplicitne Thread klase,
ali sam se odlučio za Task sa opcijom TaskCreationOptions.LongRunning. .NET Framework u 
tom slučaju prepoznaje da se radi o dugovečnom poslu i eksplicitno kreira novu namensku nit van ThreadPool-a,
čime postižemo isti efekat, ali zadržavamo prednosti moderne TPL apstrakcije i lakše
upravljanje životnim vekom zadatka."

"Što se tiče zaključavanja, ručni lock nije potreban jer BlockingCollection<string>
unutar sebe već implementira thread-safe mehanizme i optimizovanu signalizaciju.
Ona uspešno rešava konkurentnost između više proizvođača (radnih niti) 
i jednog potrošača (loger niti), pritom automatski uspavljujući potrošača kada je red prazan,
što oslobađa procesorske resurse."
 */