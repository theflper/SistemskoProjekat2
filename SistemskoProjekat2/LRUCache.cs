using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;

namespace SistemskoProjekat1
{
    internal class LRUCache
    {
        //kapacitet cache-a
        private readonly int capacity;
        //pamtimo ime fajla kao key i value koji je byte[] je fajl
        private readonly Dictionary<string, LinkedListNode<(string key, byte[] value)>> map;
        //lancana lista ko je zadnji koriscen onaj ko je na kraju je prvi koriscen
        //na pocetku je ko je najskorije koriscen
        private readonly LinkedList<(string key, byte[] value)> lruList;
        //objekat za pristup kesu posto je deljeni resurs
        private object lockObj = new object();
        //konstruktor cache-a
        public LRUCache(int capacity)
        {
            this.capacity = capacity;
            map = new Dictionary<string, LinkedListNode<(string, byte[])>>();
            lruList = new LinkedList<(string, byte[])>();
        }
        public byte[] Get(string key)
        {
            //provera da ne dodje do greske za svaki slucaj
            if (lockObj == null)
                lockObj = new object();
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
            //provera da ne dodje do greske za svaki slucaj
            if (lockObj == null)
                lockObj = new object();
            lock (lockObj)
            {
                if (map.TryGetValue(key, out var existing))
                {
                    lruList.Remove(existing);
                }
                else if (map.Count >= capacity)
                {
                    // izbaci najmanje korišćen (zadnji)
                    var lru = lruList.Last;
                    if (lru != null)
                    {
                        map.Remove(lru.Value.key);
                        lruList.RemoveLast();
                    }
                }
                var node = new LinkedListNode<(string, byte[])>((key, value));
                lruList.AddFirst(node);
                map[key] = node;
            }
        }
    }
}
