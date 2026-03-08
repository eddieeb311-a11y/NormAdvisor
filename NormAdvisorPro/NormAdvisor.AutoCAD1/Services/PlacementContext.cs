using NormAdvisor.AutoCAD1.Models;

namespace NormAdvisor.AutoCAD1.Services
{
    /// <summary>
    /// WPF palette → AutoCAD command өгөгдөл дамжуулах static context
    /// SendStringToExecute-д | тэмдэг, зай ашиглахаас зайлсхийхийн тулд
    /// </summary>
    public static class PlacementContext
    {
        public static DeviceCategory PendingCategory { get; set; }
        public static DeviceInfo PendingDevice { get; set; }

        public static bool HasPending => PendingCategory != null && PendingDevice != null;

        public static void Clear()
        {
            PendingCategory = null;
            PendingDevice = null;
        }
    }
}
