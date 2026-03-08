using NormAdvisor.AutoCAD1.Models;

namespace NormAdvisor.AutoCAD1.Services
{
    /// <summary>
    /// WPF palette → NORMDRAWROOM command өгөгдөл дамжуулах static context
    /// </summary>
    public static class RoomDrawingContext
    {
        public static RoomInfo PendingRoom { get; set; }
        public static bool HasPending => PendingRoom != null;

        public static void Clear()
        {
            PendingRoom = null;
        }
    }
}
