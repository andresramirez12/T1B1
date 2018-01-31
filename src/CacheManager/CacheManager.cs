using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Caching;
using log4net;

namespace T1.CacheManager
{
    public class CacheManager
    {
        private static readonly ILog _Logger = T1.Log.Instance.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType, Settings._Main.logLevel);

        private static readonly Lazy<CacheManager> lazy =
          new Lazy<CacheManager>(() => new CacheManager());

        private static ObjectCache cache = null;
        private CacheItemPolicy policy = null;


        public static CacheManager Instance
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

        private CacheManager()
        {
            cache = MemoryCache.Default;
        }

        public void addToCache(string CacheKeyName, dynamic CacheItem, objCachePriority objCacheItemPriority)
        {
            try
            {
                if (!Settings._Main.useAppDomain)
                {

                    //Implementar en el futuro algo para que avise porque se bajo del cache el objeto usando callback;
                    policy = new CacheItemPolicy();
                    policy.Priority = (objCacheItemPriority == objCachePriority.Default) ? CacheItemPriority.Default : CacheItemPriority.NotRemovable;
                    cache.Set(CacheKeyName, CacheItem, policy);
                }
                else
                {
                    AppDomain.CurrentDomain.SetData(CacheKeyName, CacheItem);
                }
            }
            catch(Exception er)
            {
                _Logger.Error("", er);
            }

        }

        public dynamic getFromCache(string CacheKeyName)
        {
            try { 
            if (!Settings._Main.useAppDomain)
            {
                return cache[CacheKeyName];
            }
            else
            {
                return AppDomain.CurrentDomain.GetData(CacheKeyName);
            }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
                return null;
            }
        }

        public void removeFromCache(string CacheKeyName)
        {

            try { 
            if (!Settings._Main.useAppDomain)
            {
                if (cache.Contains(CacheKeyName))
                {
                    cache.Remove(CacheKeyName);
                }
            }
            else
            {
                AppDomain.CurrentDomain.SetData(CacheKeyName, null);
            }
            }
            catch (Exception er)
            {
                _Logger.Error("", er);
            }
        }

    }
}
