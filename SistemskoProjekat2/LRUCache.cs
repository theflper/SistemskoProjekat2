using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using System.Threading;

namespace SistemskoProjekat2
{
    internal class LRUCache
    {
        //kapacitet cache-a
        private readonly int capacity;
        // TTL trajanje (koliko dugo podatak važi)
        private readonly TimeSpan ttl = TimeSpan.FromMinutes(10); // podaci važe 10 minuta
        //pamtimo ime fajla kao key i value koji je byte[] je fajl
        private readonly Dictionary<string, LinkedListNode<(string key, byte[] value)>> map;
        //lancana lista ko je zadnji koriscen onaj ko je na kraju je prvi koriscen
        //na pocetku je ko je najskorije koriscen
        private readonly LinkedList<(string key, byte[] value)> lruList;
        // TTL lista drži samo ključ i vreme isteka (nema teških bajtova ovde!)
        private readonly LinkedList<(string key, DateTime expireAt)> ttlList;
        //objekat za pristup kesu posto je deljeni resurs
        private static object lockObj = new object();
        private static object ttlObj = new object();
        //konstruktor cache-a
        public LRUCache(int capacity)
        {
            this.capacity = capacity;
            map = new Dictionary<string, LinkedListNode<(string, byte[])>>();
            lruList = new LinkedList<(string, byte[])>();
            ttlList = new LinkedList<(string, DateTime)>();
            Thread t= new Thread(() => cleanup());
            t.IsBackground = true;
            t.Start();
        }
        private void cleanup()
        {
            while (true)
            {
                DateTime end=DateTime.MinValue;
                string key = null;
                bool hasElement = false;
                //Uzimamo podatke sa TTL liste pod JEDNIM lock-om
                lock (ttlObj)
                {
                    if (ttlList.Count > 0)
                    {
                        end = ttlList.First.Value.expireAt;
                        key = ttlList.First.Value.key;
                        ttlList.RemoveFirst(); // Skidamo ga sa trake za čekanje
                        hasElement = true;
                    }
                }
                if (hasElement)
                {
                    //Ako treba, spavaj VAN lock-a
                    if (end - DateTime.UtcNow > TimeSpan.Zero)
                    {
                        Thread.Sleep((int)(end - DateTime.UtcNow).TotalMilliseconds);
                    }
                    //Bezbedno brisanje iz glavnog keša
                    lock (lockObj)
                    {
                        // TryGetValue sprečava pucanje ako je fajl u međuvremenu ručno obrisan/izmenjen
                        if (map.TryGetValue(key, out var node))
                        {
                            lruList.Remove(node);
                            map.Remove(key);
                        }
                    }
                }
                else
                {
                    // Ako je lista bila prazna, odmori sekundu
                    Thread.Sleep(1000);
                }
            }
        }
        public byte[] Get(string key)
        {
            lock (lockObj)
            {
                if (!map.TryGetValue(key, out var node))
                    return null;
                // pomeri na početak (najskorije korišćen)
                lruList.Remove(node);
                lruList.AddFirst(node);
                return node.Value.value;//jer su stvari u cache-u key,value pa je key ime fajla a value fajl
            }
        }
        public void Put(string key, byte[] value)
        {
            lock (lockObj)
            {
                if (map.TryGetValue(key, out var existing))
                {
                    lruList.Remove(existing);
                }
                else if (map.Count >= capacity)
                {
                    // Izbacujemo LRU fajl
                    var lru = lruList.Last;
                    if (lru != null)
                    {
                        string lruKey = lru.Value.key; // Zapamtimo ključ starog fajla koji leti van
                        map.Remove(lruKey);// Uklanjamo iz mape
                        lruList.RemoveLast();// Uklanjamo iz LRU liste
                        lock (ttlObj)
                        {
                            //Nadjemo i uklanjamo stari ključ iz TTL liste
                            var current = ttlList.First;
                            while (current != null)
                            {
                                if (current.Value.key == lruKey)
                                {
                                    ttlList.Remove(current);
                                    break;
                                }
                                current = current.Next;
                            }
                        }
                    }
                }
                var node = new LinkedListNode<(string, byte[])>((key, value));
                lruList.AddFirst(node);
                map[key] = node;
                //racunamo vreme isteka podatka
                lock (ttlObj)
                {
                    DateTime expirationTime = DateTime.UtcNow.Add(ttl);
                    ttlList.AddLast((key, expirationTime));
                }
            }
        }
        public void Remove(string key)
        {
            //Čistimo glavni keš resurs
            lock (lockObj)
            {
                if (map.TryGetValue(key, out var node))
                {
                    lruList.Remove(node);
                    map.Remove(key);
                }
            }
            //Čistimo TTL evidenciju
            lock (ttlObj)
            {
                // Pošto nemamo direktan pokazivač na čvor u ttlList, moramo proći kroz nju i ukloniti ključ.
                var current = ttlList.First;
                while (current != null)
                {
                    var next = current.Next;
                    if (current.Value.key == key)
                    {
                        ttlList.Remove(current); // Uklanjamo stari TTL čvor
                        break;
                    }
                    current = next;
                }
            }
        }
    }
}
