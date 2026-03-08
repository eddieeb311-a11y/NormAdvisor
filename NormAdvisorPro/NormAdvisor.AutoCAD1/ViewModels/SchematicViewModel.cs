using System;
using System.Windows.Input;
using Autodesk.AutoCAD.ApplicationServices;

namespace NormAdvisor.AutoCAD1.ViewModels
{
    public class SchematicViewModel : BaseViewModel
    {
        private string _statusText = "Бэлэн";

        public SchematicViewModel()
        {
            DrawLanPrototypeCommand = new RelayCommand(DrawLanPrototype);
        }

        public ICommand DrawLanPrototypeCommand { get; }

        public string StatusText
        {
            get => _statusText;
            set => SetProperty(ref _statusText, value);
        }

        private void DrawLanPrototype()
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    StatusText = "Active зураг олдсонгүй.";
                    return;
                }

                doc.SendStringToExecute("NORM5X10LAN\n", true, false, false);
                StatusText = "NORM5X10LAN ажиллууллаа.";
            }
            catch (Exception ex)
            {
                StatusText = $"Алдаа: {ex.Message}";
            }
        }
    }
}
