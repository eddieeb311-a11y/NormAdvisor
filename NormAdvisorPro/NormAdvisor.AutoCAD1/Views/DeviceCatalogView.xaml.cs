using System.Windows;
using System.Windows.Controls;
using NormAdvisor.AutoCAD1.Models;
using NormAdvisor.AutoCAD1.ViewModels;

namespace NormAdvisor.AutoCAD1.Views
{
    public partial class DeviceCatalogView : UserControl
    {
        public DeviceCatalogView()
        {
            InitializeComponent();
        }

        private void CategorySelected(object sender, RoutedEventArgs e)
        {
            if (sender is RadioButton rb && rb.Tag is DeviceCategory category)
            {
                if (DataContext is DeviceCatalogViewModel vm)
                {
                    vm.SelectedCategory = category;
                }
            }
        }
    }
}
