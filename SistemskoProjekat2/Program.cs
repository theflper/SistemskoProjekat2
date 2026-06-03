using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;
using System.Threading.Tasks;
using ExcelDataReader;

namespace SistemskoProjekat2
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            //Kreiramo izvor žetona za otkazivanje
            using (CancellationTokenSource cts = new CancellationTokenSource())
            {
                Server s = new Server();

                //Pokrećemo server u pozadini preko Task.Run i prosleđujemo token
                Task serverTask = Task.Run(() => s.Start(cts.Token));

                Logger.Log("Server uspešno pokrenut i radi u pozadini.");
                Console.WriteLine("==================================================================");
                Console.WriteLine("--> Pritisnite [ENTER] u bilo kom trenutku za gašenje servera. <--");
                Console.WriteLine("==================================================================\n");

                // ASINHRONO ČEKANJE NA ENTER: 
                // ova linija oslobađa nit, program čeka unos, ali troši 0% procesorskog vremena.
                await Console.In.ReadLineAsync();

                // kada korisnik pritisne Enter, kod stiže ovde i signalizira tokenu da ugasi server
                Logger.Log("Detektovan pritisak na Enter. Pokrećem bezbedno gašenje...");
                cts.Cancel();
                try
                {
                    //Čekamo da server završi trenutne zahteve i potpuno oslobodi HttpListener
                    await serverTask;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[MAIN ERROR] Greška prilikom zaustavljanja: {ex.Message}");
                }
            }
        }
    }
}
