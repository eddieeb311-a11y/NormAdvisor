using System.Windows.Controls;
using NormAdvisor.AutoCAD1.ViewModels;

namespace NormAdvisor.AutoCAD1.Views
{
    public partial class MainPaletteControl : UserControl
    {
        public MainPaletteControl()
        {
            InitializeComponent();
            DataContext = new MainViewModel();
        }
    }
}
