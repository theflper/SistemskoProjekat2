using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelDataReader;

namespace SistemskoProjekat1
{
    internal class Server
    {
        private class Request//lakse cuvanje zahteva koji su stigli na server i ne vidi se spolja
        {
            public HttpListenerContext Context { get; set; }//kontekst zahteva, sadrzi sve informacije o zahtevu i odgovoru
            public string FileName { get; set; }//ime fajla koji je trazen u zahtevu
        }
        //keš za cuvanje sadrzaja fajlova koji su vec ucitani, da se ne bi svaki put ucitavali sa diska
        //string je ime fajla, a byte[] je sadrzaj fajla
        private LRUCache cache = new LRUCache(40);//novi LRU cache velicine 40
        private bool isRunning = false;
        //semafor za ogranicavanje broja istovremenih obrada zahteva, da se ne bi preopteretio server
        private SemaphoreSlim semaphore = new SemaphoreSlim(10);
        //dozvoljava samo 3 radne niti da istovremeno obradjuju zahteve, ostale ce cekati dok se semafor ne oslobodi
        private ProcessingTracker tracker = new ProcessingTracker();
        //tracker prati sta se trenutno obradjuje
        public void Start()
        {
            HttpListener listener = new HttpListener();
            listener.Prefixes.Add("http://localhost:5050/");
            try
            {
                listener.Start();
                isRunning = true;
                Console.WriteLine("Server uspešno pokrenut na portu 5050...");
            }
            catch (HttpListenerException ex)
            {
                // Specifične greške sistema (Port zauzet ili Access Denied)
                if (ex.ErrorCode == 5) // Access Denied
                {
                    Console.WriteLine("Greška: Pokreni VS/Terminal kao administrator.");
                }
                else
                {
                    Console.WriteLine($"Sistemska greška pri pokretanju: {ex.Message}");
                }
            }
            catch (Exception ex)              // Sve ostale nepredviđene greške
            {
                Console.WriteLine($"Neočekivana greška: {ex.Message}");
            }
            finally
            {
                // Ako server NIJE uspeo da se pokrene, ovde proveravamo da li treba ugasiti listener
                if (!listener.IsListening)
                {
                    Console.WriteLine("Server nije pokrenut. Proverite konfiguraciju.");
                    listener.Close();
                }
            }
            // worker niti obradjuju zahteve iz reda, a glavna nit prihvata nove zahteve i stavlja ih u red
            ThreadPool.GetMaxThreads(out int maxWorkerThreads, out int maxCompletionPortThreads);
            //broj radnih niti koje ThreadPool moze da koristi i broj niti koje moze da koristi za IO operacije
            //10ak niti stavi
            int poolSize = Math.Min(10, maxWorkerThreads);
            // prihvatanje zahteva dok listener radi, i dodavanje zahteva u red da ih radne niti obrade
            while (listener.IsListening && isRunning)
            {
                try
                {
                    //listener.GetContext() blokira glavnu nit dok ne stigne novi zahtev, i onda vraca kontekst zahteva
                    // Koristimo await - nit se oslobađa dok klijent ne pošalje zahtev
                    var context = listener.GetContext();
                    //kada stigne zahtev, izvucemo ime fajla iz URL-a zahteva
                    string file = context.Request.Url.AbsolutePath;
                    // uzmi samo ime fajla (odbaci sve putanje)
                    file = Path.GetFileName(file);
                    var req = new Request { Context = context, FileName = file };
                    ThreadPool.QueueUserWorkItem(_ => ProcessRequest(req));
                }
                catch (HttpListenerException ex)
                {
                    // Dešava se ako se listener ugasi ili port postane nedostupan
                    if (!listener.IsListening) break;
                    Console.WriteLine($"Mrežna greška: {ex.Message}");
                }
                catch (ObjectDisposedException)
                {
                    Console.WriteLine("Listener je zatvoren iz drugog thread-a.");
                    break;
                }
                catch (Exception ex)
                {
                    // Za sve ostale nepredviđene greške
                    Console.WriteLine($"Greška prilikom obrade: {ex.Message}");
                }
            }
            if (!listener.IsListening)
            {
                Stop();//ako listener nije više aktivan, zaustavi server i oslobodi resurse
                Console.WriteLine("Server je zaustavljen.");
            }
        }
        public void Stop()//pri gasenju servera postavljamo isRunning na false da bi prestali sa while petljom koja osluskuje
        {
            isRunning = false;
        }
        private void ProcessRequest(Request req)
        {
            try
            {
                if (req == null || req.Context == null)
                    return;
                if (string.IsNullOrEmpty(req.FileName))
                {
                    SendText(req, "Missing file", 400);
                    return;
                }
                if (!req.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    SendText(req, "Only .csv files are allowed", 400);
                    return;
                }
                //kad smo izvukli zahtev iz reda sad mozemo da oslobodimo lock na redu
                //i da obradjujemo zahtev bez blokiranja drugih radnih niti koje mogu da prihvate nove zahteve
                semaphore.Wait();
                //iako nemamo catch ovo sve garantuje da se semafor oslobodi
                //i da ne ostane zauzet ako dodje do greske u obradi zahteva
                try
                {
                    HandleRequest(req);//obrada zahteva
                }
                finally
                {
                    semaphore.Release();//oslobadjamo semafor nakon obrade zahteva
                    // da bi druge radne niti mogle da obradjuju nove zahteve
                }
            }
            catch (Exception ex)
            {
                SendText(req, "Error: " + ex.Message, 500);
            }
        }
        private void HandleRequest(Request req)
        {
            try
            {
                if (string.IsNullOrEmpty(req.FileName))
                {
                    SendText(req, "Missing file parameter", 400);
                    return;
                }
                byte[] excelData = GetFromCacheOrProcess(req);
                if (excelData == null)
                {
                    if (req.Context.Response.StatusCode == 500) return;
                    SendText(req, "File not found", 404);
                    return;
                }
                //podesi tip odgovora na Excel format i posalji podatke nazad klijentu
                req.Context.Response.ContentType =
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                // dodajemo zaglavlje za ime fajla
                // "attachment" forsira download, a "filename" definiše ime pod kojim se čuva
                string downloadName = req.FileName.Replace(".csv", ".xlsx"); // Primer promene ekstenzije
                req.Context.Response.AddHeader("Content-Disposition", $"attachment; filename=\"{downloadName}\"");
                //Ako želiš da se fajl samo otvori u browseru(ako browser to podržava),
                //umesto attachment koristi reč inline.
                //podesi duzinu odgovora da bi klijent znao koliko podataka ce dobiti
                req.Context.Response.ContentLength64 = excelData.Length;
                //u body odgovora upisujemo sadrzaj excel fajla koji smo dobili iz cache-a ili obradom csv fajla
                req.Context.Response.OutputStream.Write(excelData, 0, excelData.Length);
                req.Context.Response.OutputStream.Close();
                //podesili smo sta treba da se posalje klijentu, sad zatvaramo output stream da bi se odgovor poslao
                req.Context.Response.Close();
                //zatvaramo celu HTTP vezu da oslobodimo resurse i da bi klijent znao da je odgovor kompletan
            }
            catch (Exception e)
            {
                SendText(req, "Error: " + e.Message);
                //ako je doslo do greske u obradi zahteva, posaljemo klijentu poruku o gresci
            }
            finally
            {
                try { req.Context.Response.Close(); } catch { }
                //ako je doslo do greske u obradi zahteva, pokusavamo da zatvorimo HTTP vezu da oslobodimo resurse
                //ali ako klijent vise nije tu, onda ce se desiti greska prilikom zatvaranja veze, pa zato imamo try-catch
            }
        }
        private byte[] GetFromCacheOrProcess(Request req)
        {
            string fileName = req.FileName;
            try
            {
                // probaj da ga pribavis iz kesa
                var cached = cache.Get(fileName);
                //ako ne vrati null cache hit
                if (cached != null)
                    return cached;
                //pokušaj da obradis fajl ako se vec ne obradjuje
                bool iAmWorker = tracker.WaitOrTake(fileName);
                if (!iAmWorker)
                {
                    // neko drugi je obradio i sad fajl mora biti u cache-u
                    return cache.Get(fileName);
                }
                byte[] result = null;
                try
                {
                    //losa putanja vrati null
                    if (fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                        return null;
                    //ako ne izdvoji samo ime fajla i spoji sa data folderom koji je u bin debug projekta
                    string path = Path.Combine("data", Path.GetFileName(fileName));
                    //ako putanja ne postoji vrati null
                    if (!File.Exists(path))
                        return null;
                    //samo nit koja obradjuje dodje do ovde
                    result = ConvertCsvToExcel(path);
                }
                finally
                {
                    if (result != null)
                        cache.Put(fileName, result);
                    tracker.Done(fileName);
                }
                return result;
            }
            catch (Exception e)
            {
                SendText(req, "Error: " + e.Message, 500);
                return null;
            }
        }
        private byte[] ConvertCsvToExcel(string path)
        {
            //jednostavna simulacija konverzije ne pravimo txt fajl vec TSV fajl koji browser cita kao .xlsx
            //TSV (Tab-Separated Values) je format koji koristi tab karakter kao separator između vrednosti
            //dok CSV (Comma-Separated Values) koristi zarez. TSV moze browser da prepozna fajl kao Excel format.
            // (ExcelDataReader je zapravo za čitanje, ali profesor često traži simbolički)
            // Koristimo FileStream sa FileShare.Read da bismo izbegli "File in use" greške
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new StreamReader(fs))//za čitanje CSV fajla
            using (var ms = new MemoryStream())//privremeni prosotor u RAM-u gde ćemo upisivati konvertovani sadržaj
            using (var writer = new StreamWriter(ms, System.Text.Encoding.UTF8))//za pisanje konvertovanog sadržaja u memorijski stream
            //kreiramo StreamWriter koji će nam olakšati upisivanje teksta u memorijski stream
            {
                string line;//čitanje CSV fajla liniju po liniju
                while ((line = reader.ReadLine()) != null)//dok ima linija u CSV fajlu
                {
                    // Menjamo zarez tabom za TSV format
                    writer.WriteLine(line.Replace(",", "\t"));
                }
                //obavezno flush-ujemo StreamWriter da bi se sav tekst upisao u MemoryStream pre nego što ga konvertujemo u byte[]
                writer.Flush();
                return ms.ToArray();
            }
        }
        private void SendText(Request req, string text, int statusCode = 200)
        {
            try
            {
                //Podesi status kod (npr. 404, 500, 400...)
                req.Context.Response.StatusCode = statusCode;
                //podesi tip odgovora na plain text i posalji poruku nazad klijentu
                req.Context.Response.ContentType = "text/plain; charset=utf-8";
                //pretvaramo tekst poruke u byte[] da bi ga mogli poslati klijentu
                byte[] data = System.Text.Encoding.UTF8.GetBytes(text);
                //podesi duzinu odgovora da bi klijent znao koliko podataka ce dobiti
                req.Context.Response.ContentLength64 = data.Length;
                req.Context.Response.OutputStream.Write(data, 0, data.Length);
            }
            catch
            {
                // Ako klijent više nije tu, samo ignorišemo grešku
            }
            finally
            {
                try { req.Context.Response.Close(); } catch { }
            }
        }
    }
}
