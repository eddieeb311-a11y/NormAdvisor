using System.Collections.Generic;
using Newtonsoft.Json;

namespace NormAdvisor.AutoCAD1.Models
{
    /// <summary>
    /// blocks_config.json дахь төхөөрөмжийн категори
    /// </summary>
    public class DeviceCategory
    {
        [JsonProperty("name")]
        public string Name { get; set; } = string.Empty;

        [JsonProperty("icon")]
        public string Icon { get; set; } = string.Empty;

        [JsonProperty("dwgFile")]
        public string DwgFile { get; set; } = string.Empty;

        [JsonProperty("blockName")]
        public string BlockName { get; set; } = string.Empty;

        [JsonProperty("devices")]
        public List<DeviceInfo> Devices { get; set; } = new List<DeviceInfo>();

        public override string ToString() => $"{Icon} {Name}";
    }

    /// <summary>
    /// blocks_config.json-ийн үндсэн бүтэц
    /// </summary>
    public class BlocksConfig
    {
        [JsonProperty("categories")]
        public List<DeviceCategory> Categories { get; set; } = new List<DeviceCategory>();
    }
}
