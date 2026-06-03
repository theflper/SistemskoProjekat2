using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SistemskoProjekat2
{
    internal class ProcessingTracker
    {
        private class Entry
        {
            // TaskCompletionSource nam omogućava da ručno signaliziramo kada je Task gotov
            public TaskCompletionSource<bool> Tcs { get; } = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        }
        //Dictionary gde čuvamo šta se trenutno obrađuje
        private readonly Dictionary<string, Entry> map = new Dictionary<string, Entry>();
        private static readonly object globalLock = new object();
        // Metoda je sada ASINHRONA i vraća Task<bool> nije obična funkcija kao u prvom projektu
        public async Task<bool> WaitOrTakeAsync(string key)
        {
            Entry entry;
            bool isWorker = false;
            lock (globalLock)
            {
                if (!map.TryGetValue(key, out entry))
                {
                    entry = new Entry();
                    map[key] = entry;
                    isWorker = true;
                }
            }
            if (isWorker)
            {
                return true;
            }
            // ako neko već obrađuje nema blokiranja niti 
            // nit se oslobađa, a runtime čeka da Tcs dobije rezultat.
            await entry.Tcs.Task;
            return false; // Neko drugi je završio obradu, ja samo čitam iz keša
        }
        // Signalizira da je obrada gotova
        public void Done(string key)
        {
            Entry entry;
            lock (globalLock)
            {
                if (!map.TryGetValue(key, out entry))
                    return;
                map.Remove(key);
            }
            // signaliziramo svim nitima koje su uradile 'await entry.Tcs.Task' 
            // da je posao gotov sve one se sada bude asinhrono.
            entry.Tcs.TrySetResult(true);
        }
    }
}