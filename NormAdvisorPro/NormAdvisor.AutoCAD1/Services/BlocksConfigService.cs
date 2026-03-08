using System;
using System.IO;
using Newtonsoft.Json;
using NormAdvisor.AutoCAD1.Models;

namespace NormAdvisor.AutoCAD1.Services
{
    /// <summary>
    /// blocks_config.json файлыг уншиж, кэш хийнэ (Singleton)
    /// </summary>
    public class BlocksConfigService
    {
        private static BlocksConfigService _instance;
        private static readonly object _lock = new object();

        private BlocksConfig _config;
        private string _configDirectory;

        public static BlocksConfigService Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                            _instance = new BlocksConfigService();
                    }
                }
                return _instance;
            }
        }

        private BlocksConfigService() { }

        /// <summary>
        /// Тохиргоо ачаалах (DLL-ийн хажуу директороос blocks_config.json олно)
        /// </summary>
        public BlocksConfig LoadConfig()
        {
            if (_config != null) return _config;

            string dllPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            _configDirectory = Path.GetDirectoryName(dllPath);
            string configPath = Path.Combine(_configDirectory, "blocks_config.json");

            if (!File.Exists(configPath))
                throw new FileNotFoundException($"blocks_config.json олдсонгүй: {configPath}");

            string json = File.ReadAllText(configPath);
            _config = JsonConvert.DeserializeObject<BlocksConfig>(json);

            return _config;
        }

        /// <summary>
        /// DWG файлын бүтэн замыг олох
        /// </summary>
        public string GetDwgFullPath(string dwgFileName)
        {
            if (string.IsNullOrEmpty(_configDirectory))
                LoadConfig();

            return Path.Combine(_configDirectory, dwgFileName);
        }

        /// <summary>
        /// Тохиргоо дахин ачаалах (шинэчлэлт хийсэн бол)
        /// </summary>
        public void Reload()
        {
            _config = null;
            LoadConfig();
        }
    }
}
