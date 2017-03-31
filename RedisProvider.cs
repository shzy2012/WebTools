using StackExchange.Redis;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Project.Utility {
    /// <summary>
    /// Redis内存数据库
    /// </summary>
    public class RedisProvider {
        private static log4net.ILog logger = log4net.LogManager.GetLogger(typeof(RedisProvider));

        /// <summary>
        /// Redis连接字符串
        /// </summary>
        public const string RedisConnectString = "redis_name";
        /// <summary>
        /// Redis 连接配置
        /// </summary>
        private static Lazy<ConfigurationOptions> configOptions = new Lazy<ConfigurationOptions>(() => {
            var configOptions = new ConfigurationOptions();
            configOptions.EndPoints.Add(ConfigurationManager.ConnectionStrings[RedisConnectString].ToString());
            configOptions.ClientName = "ConnectionName";
            configOptions.ConnectTimeout = 100000; //100S
            configOptions.SyncTimeout = 100000;
            return configOptions;
        });

        public static Lazy<IConnectionMultiplexer> clientInstance =
           new Lazy<IConnectionMultiplexer>(CreateInstance, LazyThreadSafetyMode.ExecutionAndPublication);


        /// <summary>
        /// Create Instance
        /// </summary>
        /// <returns></returns>
        private static IConnectionMultiplexer CreateInstance() {
            IConnectionMultiplexer client = null;
            try {
                client = ConnectionMultiplexer.Connect(configOptions.Value);
            } catch(Exception ex) {
                LogProvider.LogException(logger, "RedisProvider.CreateInstance", ex);
                throw;
            }
            return client;
        }

        /// <summary>
        /// ContactKey
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        private static string ContactKey(string key) {
            return key;
        }

        /// <summary>
        /// Hash Add
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <param name="expiration"></param>
        /// <returns></returns>
        public bool HashAdd(int dbIndex, string key, string hashField, string hashValue) {
            if(string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(hashField)) {
                return false;
            }
            try {
                return clientInstance.Value.GetDatabase(dbIndex).HashSet(key, hashField, hashValue);
            } catch(Exception ex) {
                LogProvider.LogException(logger, "RedisProvider.HashAdd", ex, dbIndex, key, hashField, hashValue);
                return false;
            }
        }
    }
}
