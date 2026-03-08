using System;
using System.Windows.Forms.Integration;
using Autodesk.AutoCAD.Windows;
using NormAdvisor.AutoCAD1.Views;

namespace NormAdvisor.AutoCAD1
{
    /// <summary>
    /// AutoCAD PaletteSet — NormAdvisor Pro палетт цонх
    /// Dock хийх боломжтой (зүүн/баруун тал)
    /// </summary>
    public class NormPaletteSet
    {
        private static NormPaletteSet _instance;
        private PaletteSet _paletteSet;

        // Тогтмол GUID (AutoCAD палетт таних)
        private static readonly Guid PaletteGuid = new Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890");

        public static NormPaletteSet Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new NormPaletteSet();
                return _instance;
            }
        }

        private NormPaletteSet() { }

        /// <summary>
        /// Палетт нээх/хаах toggle
        /// </summary>
        public void Toggle()
        {
            if (_paletteSet == null)
                Create();

            _paletteSet.Visible = !_paletteSet.Visible;
        }

        /// <summary>
        /// Палетт харуулах
        /// </summary>
        public void Show()
        {
            if (_paletteSet == null)
                Create();

            _paletteSet.Visible = true;
        }

        /// <summary>
        /// Палетт нуух
        /// </summary>
        public void Hide()
        {
            if (_paletteSet != null)
                _paletteSet.Visible = false;
        }

        public bool IsVisible => _paletteSet?.Visible ?? false;

        private void Create()
        {
            _paletteSet = new PaletteSet("NormAdvisor Pro", PaletteGuid);
            _paletteSet.Style = PaletteSetStyles.ShowAutoHideButton |
                                PaletteSetStyles.ShowCloseButton |
                                PaletteSetStyles.Snappable;

            _paletteSet.MinimumSize = new System.Drawing.Size(300, 400);
            _paletteSet.DockEnabled = DockSides.Left | DockSides.Right;

            // WPF UserControl-ыг ElementHost-оор оруулах
            var wpfControl = new MainPaletteControl();
            var elementHost = new ElementHost
            {
                Dock = System.Windows.Forms.DockStyle.Fill,
                Child = wpfControl
            };

            _paletteSet.Add("NormAdvisor", elementHost);
            _paletteSet.KeepFocus = false;
        }
    }
}
