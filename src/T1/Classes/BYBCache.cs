using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Caching;

namespace T1.Classes
{
    public sealed class BYBCache
    {
        private static readonly Lazy<BYBCache> lazy =
            new Lazy<BYBCache>(() => new BYBCache());

        private static ObjectCache cache = null;
        private CacheItemPolicy policy = null;
       
        
        public static BYBCache Instance
        {
            get
            {
                return lazy.Value;
            }
        }

        public enum objCachePriority
        {
            Default,
            NotRemovable
        }

        private BYBCache()
        {
            cache = MemoryCache.Default;
        }

        public void addToCache(string CacheKeyName, dynamic CacheItem, objCachePriority objCacheItemPriority)
        {
            //Implementar en el futuro algo para que avise porque se bajo del cache el objeto usando callback;
            policy = new CacheItemPolicy();
            policy.Priority = (objCacheItemPriority == objCachePriority.Default) ? CacheItemPriority.Default : CacheItemPriority.NotRemovable;
            cache.Set(CacheKeyName, CacheItem, policy);
           
        }
        
        public dynamic getFromCache(string CacheKeyName)
        {
            return cache[CacheKeyName];
        }

        public void removeFromCache(string CacheKeyName)
        {
            if(cache.Contains(CacheKeyName))
            {
                cache.Remove(CacheKeyName);
            }
        }

        public bool inCache(string CacheKeyName)
        {
            if (cache.Contains(CacheKeyName))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
