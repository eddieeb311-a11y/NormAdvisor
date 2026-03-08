using System;
using System.Collections.ObjectModel;
using System.Windows.Input;
using Autodesk.AutoCAD.ApplicationServices;
using NormAdvisor.AutoCAD1.Models;
using NormAdvisor.AutoCAD1.Services;

namespace NormAdvisor.AutoCAD1.ViewModels
{
    public class DeviceCatalogViewModel : BaseViewModel
    {
        private ObservableCollection<DeviceCategory> _categories = new ObservableCollection<DeviceCategory>();
        private DeviceCategory _selectedCategory;
        private DeviceInfo _selectedDevice;
        private string _statusText = string.Empty;
        private readonly DevicePlacementService _placementService;

        public DeviceCatalogViewModel()
        {
            _placementService = new DevicePlacementService();
            PlaceDeviceCommand = new RelayCommand<DeviceInfo>(PlaceDevice);
            LoadCategories();
        }

        public ObservableCollection<DeviceCategory> Categories
        {
            get => _categories;
            set => SetProperty(ref _categories, value);
        }

        public DeviceCategory SelectedCategory
        {
            get => _selectedCategory;
            set
            {
                if (SetProperty(ref _selectedCategory, value))
                {
                    OnPropertyChanged(nameof(Devices));
                    SelectedDevice = null;
                }
            }
        }

        public ObservableCollection<DeviceInfo> Devices
        {
            get
            {
                if (_selectedCategory == null) return new ObservableCollection<DeviceInfo>();
                return new ObservableCollection<DeviceInfo>(_selectedCategory.Devices);
            }
        }

        public DeviceInfo SelectedDevice
        {
            get => _selectedDevice;
            set => SetProperty(ref _selectedDevice, value);
        }

        public string StatusText
        {
            get => _statusText;
            set => SetProperty(ref _statusText, value);
        }

        public ICommand PlaceDeviceCommand { get; }

        private void LoadCategories()
        {
            try
            {
                var config = BlocksConfigService.Instance.LoadConfig();
                _categories.Clear();
                foreach (var cat in config.Categories)
                    _categories.Add(cat);

                StatusText = $"{_categories.Count} категори ачааллаа";
            }
            catch (Exception ex)
            {
                StatusText = $"Тохиргоо ачаалахад алдаа: {ex.Message}";
            }
        }

        /// <summary>
        /// CommandParameter-ээс DeviceInfo шууд авна (SelectedDevice-д тулгуурлахгүй)
        /// </summary>
        private void PlaceDevice(DeviceInfo device)
        {
            if (device == null || _selectedCategory == null) return;

            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            try
            {
                PlacementContext.PendingCategory = _selectedCategory;
                PlacementContext.PendingDevice = device;

                doc.SendStringToExecute("NORMPLACE\n", true, false, false);
            }
            catch (Exception ex)
            {
                StatusText = $"Алдаа: {ex.Message}";
            }
        }
    }
}
