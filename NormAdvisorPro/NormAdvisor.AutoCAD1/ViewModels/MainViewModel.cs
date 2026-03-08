using System.Windows.Input;

namespace NormAdvisor.AutoCAD1.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        private int _selectedTabIndex;

        public MainViewModel()
        {
            RoomList = new RoomListViewModel();
            DeviceCatalog = new DeviceCatalogViewModel();
            Schematic = new SchematicViewModel();
        }

        public RoomListViewModel RoomList { get; }
        public DeviceCatalogViewModel DeviceCatalog { get; }
        public SchematicViewModel Schematic { get; }

        public int SelectedTabIndex
        {
            get => _selectedTabIndex;
            set => SetProperty(ref _selectedTabIndex, value);
        }
    }
}
