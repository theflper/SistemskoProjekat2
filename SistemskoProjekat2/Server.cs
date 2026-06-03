using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SistemskoProjekat2
{
    internal class Server
    {
        private class Request//lakse cuvanje zahteva koji su stigli na server i ne vidi se spolja
        {
            public HttpListenerContext Context { get; set; }//kontekst zahteva, sadrzi sve informacije o zahtevu i odgovoru
            public string FileName { get; set; }//ime fajla koji je trazen u zahtevu
        }
        // Dodaj na nivo klase (pored cache-a) jedan običan lock objekat
        private static readonly object cacheLock = new object();
        //keš za cuvanje sadrzaja fajlova koji su vec ucitani, da se ne bi svaki put ucitavali sa diska
        //string je ime fajla, a byte[] je sadrzaj fajla
        private LRUCache cache = new LRUCache(40);//novi LRU cache velicine 40
        //semafor za ogranicavanje broja istovremenih obrada zahteva, da se ne bi preopteretio server
        private SemaphoreSlim semaphore = new SemaphoreSlim(10);
        //dozvoljava samo 3 radne niti da istovremeno obradjuju zahteve, ostale ce cekati dok se semafor ne oslobodi
        private ProcessingTracker tracker = new ProcessingTracker();
        //tracker prati sta se trenutno obradjuje
        public void Start(CancellationToken token)
        {
            HttpListener listener = new HttpListener();
            listener.Prefixes.Add("http://localhost:5050/");
            try
            {
                listener.Start();
                Logger.Log("Server uspešno pokrenut na portu 5050...");
            }
            catch (HttpListenerException ex)
            {
                // Specifične greške sistema (Port zauzet ili Access Denied)
                if (ex.ErrorCode == 5) // Access Denied
                {
                    Logger.Log("Greška: Pokreni VS/Terminal kao administrator.");
                }
                else
                {
                    Logger.Log($"Sistemska greška pri pokretanju: {ex.Message}");
                }
            }
            catch (Exception ex)              // Sve ostale nepredviđene greške
            {
                Logger.Log($"Neočekivana greška: {ex.Message}");
            }
            finally
            {
                // Ako server NIJE uspeo da se pokrene, ovde proveravamo da li treba ugasiti listener
                if (!listener.IsListening)
                {
                    Logger.Log("Server nije pokrenut. Proverite konfiguraciju.");
                    listener.Close();
                }
            }
            // Kada se okine spoljašnje gašenje preko tokena, gasimo listener da prekine GetContext()
            token.Register(() =>
            {
                try { 
                    listener.Stop();//prekini da slusas
                    listener.Close();//oslobodi port 5050 koji server koristi
                } catch { }
            });
            //dodata je provera da li je token otkazan u while petlji da stanemo sa radom
            while (listener.IsListening && !token.IsCancellationRequested)
            {
                try
                {
                    // listener.GetContext() blokira glavnu nit dok ne stigne novi zahtev, i onda vraca kontekst zahteva
                    // Koristimo await - nit se oslobađa dok klijent ne pošalje zahtev
                    // Kada se listener.Stop() pozove unutar token.Register, 
                    // GetContext() će baciti HttpListenerException i tako prekinuti blokadu!
                    var context = listener.GetContext();
                    //kada stigne zahtev, izvucemo ime fajla iz URL-a zahteva
                    string file = context.Request.Url.AbsolutePath;
                    // uzmi samo ime fajla (odbaci sve putanje)
                    file = Path.GetFileName(file);
                    var req = new Request { Context = context, FileName = file };
                    var processingTask = Task.Run(async () => await ProcessRequest(req));
                    processingTask.ContinueWith(t =>
                    {
                        if (t.IsFaulted)
                        {
                            Logger.Log($"ERROR: Zahtev za fajl '{req.FileName}' je pretrpeo kritičnu grešku: {t.Exception.GetBaseException().Message}");
                        }
                        else
                        {
                            Logger.Log($"INFO: Zahtev za fajl '{req.FileName}' je uspešno procesuiran.");
                        }
                    }, TaskContinuationOptions.ExecuteSynchronously);
                }
                catch (HttpListenerException ex)
                {
                    // Dešava se ako se listener ugasi ili port postane nedostupan
                    if (!listener.IsListening || token.IsCancellationRequested) break;
                    Logger.Log($"Mrežna greška: {ex.Message}");
                }
                catch (ObjectDisposedException)
                {
                    Logger.Log("Listener je zatvoren iz drugog thread-a.");
                    break;
                }
                catch (Exception ex)
                {
                    // Ako nas sistem gasi, ignoriši greške pri prekinutom prihvatanju novih veza
                    if (token.IsCancellationRequested) break;
                    // Za sve ostale nepredviđene greške
                    Logger.Log($"Greška prilikom obrade: {ex.Message}");
                }
            }
            if (!listener.IsListening)
            {
                Logger.Log("Server je zaustavljen.");
            }
        }
        private async Task ProcessRequest(Request req)
        {
            try
            {
                if (req == null || req.Context == null)
                    return;
                if (string.IsNullOrEmpty(req.FileName))
                {
                    await SendText(req, "Missing file", 400);
                    return;
                }
                if (!req.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    await SendText(req, "Only .csv files are allowed", 400);
                    return;
                }
                // Koristimo WaitAsync() umesto sinhronog Wait()
                // Ako je server pun, ova nit se oslobađa dok se ne oslobodi mesto u semaforu!
                await HandleRequestAsync(req);//obrada zahteva
            }
            catch (Exception ex)
            {
                await SendText(req, "Error: " + ex.Message, 500);
            }
        }
        private async Task HandleRequestAsync(Request req)
        {
            try
            {
                if (string.IsNullOrEmpty(req.FileName))
                {
                    await SendText(req, "Missing file parameter", 400);
                    return;
                }
                byte[] excelData = await GetFromCacheOrProcessAsync(req);
                if (excelData == null)
                {
                    if (req.Context.Response.StatusCode == 500) return;
                    await SendText(req, "File not found", 404);
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
                // u body odgovora upisujemo sadrzaj excel fajla koji smo dobili iz cache-a ili obradom csv fajla
                // optimizacija u odnosu na sinhrono jer je asinhrono slanje podataka. 
                // dok se fajl šalje kroz mrežu spovom klijentu, nit se oslobađa i može da usluži nekog drugog!
                await req.Context.Response.OutputStream.WriteAsync(excelData, 0, excelData.Length);
                req.Context.Response.OutputStream.Close();
                //podesili smo sta treba da se posalje klijentu, sad zatvaramo output stream da bi se odgovor poslao
                req.Context.Response.Close();
                //zatvaramo celu HTTP vezu da oslobodimo resurse i da bi klijent znao da je odgovor kompletan
                Logger.Log("Zahtev za fajl '" + req.FileName + "' je uspešno obrađen i poslat klijentu.");
            }
            catch (Exception e)
            {
                await SendText(req, "Error: " + e.Message);
                //ako je doslo do greske u obradi zahteva, posaljemo klijentu poruku o gresci
            }
            finally
            {
                try { req.Context.Response.Close(); } catch { }
                //ako je doslo do greske u obradi zahteva, pokusavamo da zatvorimo HTTP vezu da oslobodimo resurse
                //ali ako klijent vise nije tu, onda ce se desiti greska prilikom zatvaranja veze, pa zato imamo try-catch
            }
        }
        private async Task<byte[]> GetFromCacheOrProcessAsync(Request req)
        {
            string fileName = req.FileName;
            try
            {
                //brza provera keša
                lock (cacheLock)
                {
                    var cached = cache.Get(fileName);
                    if (cached != null) return cached;
                }
                // ako nije u kešu, tek tada se borimo oko tracker-a
                bool iAmWorker = await tracker.WaitOrTakeAsync(fileName);
                if (!iAmWorker)
                {
                    //ekskluzivno pristupamo kešu jer svi krenu odjednom kad je gotovo
                    lock (cacheLock)
                    {
                        var postProcessedCached = cache.Get(fileName);
                        if (postProcessedCached != null) return postProcessedCached;
                    }
                    throw new Exception("Obrada fajla od strane druge niti nije uspela.");
                }
                // (jedna jedina nit za taj fajl) prolazi kroz semafor
                // Ovo štiti tvoj procesor i disk ako stigne 50 različitih fajlova odjednom
                await semaphore.WaitAsync();
                byte[] result = null;
                try
                {
                    if (fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) return null;
                    string path = Path.Combine("data", Path.GetFileName(fileName));
                    if (!File.Exists(path)) return null;
                    // Pokrećemo tvoju sinhronu metodu u pozadinskom Tasku
                    // conversionTask postaje tipa Task<byte[]> iako je metoda obična sinhrona
                    var conversionTask = Task.Run(() => ConvertCsvToExcel(path));
                    //Pošto je to sada Task, na njega MOŽEMO zakačiti ContinueWith osigurač
                    _=conversionTask.ContinueWith(t =>
                    {
                        Logger.Log($"CRITICAL: Konverzija za {fileName} je fatalno otkazala! Automatski čistim tracker.");
                        tracker.Done(fileName);
                    }, TaskContinuationOptions.OnlyOnFaulted);
                    //Pošto smo unutar async metode, radimo await nad tim taskom da dobijemo byte[]
                    result = await conversionTask;
                }
                finally
                {
                    // Prvo punimo keš
                    if (result != null)
                    {
                        lock (cacheLock)
                        {
                            cache.Put(fileName, result);
                        }
                    }
                    // Oslobađamo semafor za druge fajlove
                    semaphore.Release();
                    // Budimo ostale niti koje su čekale na ovaj konkretan fajl
                    tracker.Done(fileName);
                }
                return result;
            }
            catch (Exception e)
            {
                await SendText(req, "Error: " + e.Message, 500);
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
        private async Task SendText(Request req, string text, int statusCode = 200)
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
                await req.Context.Response.OutputStream.WriteAsync(data, 0, data.Length);
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
