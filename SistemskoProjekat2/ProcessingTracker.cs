using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelDataReader;

namespace SistemskoProjekat1
{
    internal class ProcessingTracker
    {
        //objekat za zakljucavanje koji ce nam pomoci da sinhronizujemo pristup map-i
        private class Entry
        {
            public bool IsProcessing = true;
        }
        private readonly Dictionary<string, Entry> map = new Dictionary<string, Entry>();
        private static readonly object globalLock = new object();
        // Ako neko već obrađuje key onda čekamo da taj neko obradi.
        // Ako ne nit koja je skontala da niko ne radi obradi zahtev.
        public bool WaitOrTake(string key)
        {
            Entry entry;
            lock (globalLock)
            {
                if (!map.TryGetValue(key, out entry))
                {
                    // nema ga da se trenutno obradjuje onda ga mi obrađujemo
                    entry = new Entry();
                    map[key] = entry;
                    return true; // ja sam worker
                }
            }
            //ako već postoji, čekaj NA TOM entry-ju
            //ovim blokiramo niti dok neka druga nit ne obradi fajl
            //to definise ovaj isProcessing
            lock (entry)
            {
                while (entry.IsProcessing)
                {
                    Monitor.Wait(entry);
                }
            }
            //neko drugi zavrsio mi cemo samo iz kesa da procitamo
            return false;
        }
        // Signalizira da je obrada gotova
        public void Done(string key)
        {
            Entry entry;
            //provera da ne dodje do greske za svaki slucaj
            lock (globalLock)
            {
                //nanbavi iz dictionary
                //ako ga vec nema nista
                if (!map.TryGetValue(key, out entry))
                    return;
                map.Remove(key);//ako je u dictionary ukloni key
            }
            lock (entry)
            {
                //promeni isProcessing da ne cekaju ostali u while
                entry.IsProcessing = false;
                //probudi niti da prodju kroz 
                Monitor.PulseAll(entry);
            }
        }
    }
}
