using Newtonsoft.Json;

namespace NormAdvisor.AutoCAD1.Models
{
    /// <summary>
    /// blocks_config.json дахь нэг төхөөрөмжийн мэдээлэл
    /// </summary>
    public class DeviceInfo
    {
        [JsonProperty("name")]
        public string Name { get; set; } = string.Empty;

        [JsonProperty("visibilityState")]
        public string VisibilityState { get; set; } = string.Empty;

        [JsonProperty("tagPrefix")]
        public string TagPrefix { get; set; } = string.Empty;

        public override string ToString() => Name;
    }
}
