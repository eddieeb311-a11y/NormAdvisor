using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Data;
using System.Windows.Input;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using NormAdvisor.AutoCAD1.Models;
using NormAdvisor.AutoCAD1.Services;

namespace NormAdvisor.AutoCAD1.ViewModels
{
    public class RoomListViewModel : BaseViewModel
    {
        private ObservableCollection<RoomInfo> _allRooms = new ObservableCollection<RoomInfo>();
        private ICollectionView _roomsView;
        private string _searchText = string.Empty;
        private RoomInfo _selectedRoom;
        private string _statusText = "Ó¨Ñ€Ó©Ó© ÑƒÐ½ÑˆÐ°Ð°Ð³Ò¯Ð¹ Ð±Ð°Ð¹Ð½Ð°";
        private bool _isMasterLocked;
        private int _copyCount;
        private string _masterStatusText = "";

        /// <summary>
        /// NORMDRAWROOM command-Ð°Ð°Ñ callback Ð°Ð²Ð°Ñ…Ð°Ð´ Ð°ÑˆÐ¸Ð³Ð»Ð°Ð½Ð°
        /// </summary>
        public static RoomListViewModel Current { get; private set; }

        public RoomListViewModel()
        {
            Current = this;

            _roomsView = CollectionViewSource.GetDefaultView(_allRooms);
            _roomsView.Filter = FilterRoom;

            ReadFromTableCommand = new RelayCommand(ReadFromTable);
            ReadFromRegionCommand = new RelayCommand(ReadFromRegion);
            DrawBoundaryCommand = new RelayCommand(DrawBoundary);
            SelectBoundaryCommand = new RelayCommand(SelectBoundary);
            ZoomToRoomCommand = new RelayCommand(ZoomToRoom);
            ToggleMasterLockCommand = new RelayCommand(ToggleMasterLock);
            RefreshCopiesCommand = new RelayCommand(RefreshCopies);
            AutoMatchCommand = new RelayCommand(AutoMatch);
        }

        public ICollectionView RoomsView => _roomsView;

        public string SearchText
        {
            get => _searchText;
            set
            {
                if (SetProperty(ref _searchText, value))
                    _roomsView.Refresh();
            }
        }

        public RoomInfo SelectedRoom
        {
            get => _selectedRoom;
            set
            {
                if (SetProperty(ref _selectedRoom, value))
                {
                    HighlightSelectedRoom();
                }
            }
        }

        public string StatusText
        {
            get => _statusText;
            set => SetProperty(ref _statusText, value);
        }

        public int RoomCount => _allRooms.Count;
        public double TotalArea => _allRooms.Sum(r => r.Area);
        public int LinkedCount => _allRooms.Count(r => r.HasBoundary);

        public bool IsMasterLocked
        {
            get => _isMasterLocked;
            set => SetProperty(ref _isMasterLocked, value);
        }

        public int CopyCount
        {
            get => _copyCount;
            set => SetProperty(ref _copyCount, value);
        }

        public string MasterStatusText
        {
            get => _masterStatusText;
            set => SetProperty(ref _masterStatusText, value);
        }

        /// <summary>
        /// ÐœÐ°ÑÑ‚ÐµÑ€ lock Ñ…Ð¸Ð¹Ñ… Ð±Ð¾Ð»Ð¾Ð¼Ð¶Ñ‚Ð¾Ð¹ ÑÑÑÑ… (Ð±Ò¯Ñ… Ó©Ñ€Ó©Ó© boundary-Ñ‚Ð°Ð¹ Ð±Ð¾Ð»)
        /// </summary>
        public bool CanLockMaster => LinkedCount > 0 && LinkedCount == RoomCount && !IsMasterLocked;

        public ICommand ReadFromTableCommand { get; }
        public ICommand ReadFromRegionCommand { get; }
        public ICommand DrawBoundaryCommand { get; }
        public ICommand SelectBoundaryCommand { get; }
        public ICommand ZoomToRoomCommand { get; }
        public ICommand ToggleMasterLockCommand { get; }
        public ICommand RefreshCopiesCommand { get; }
        public ICommand AutoMatchCommand { get; }

        private void ReadFromTable()
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;

                var reader = new RoomTableReader();
                var rooms = reader.ReadFromTableDirect();

                UpdateRooms(rooms);
            }
            catch (Exception ex)
            {
                StatusText = $"ÐÐ»Ð´Ð°Ð°: {ex.Message}";
            }
        }

        private void ReadFromRegion()
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;

                var reader = new RoomTableReader();
                var rooms = reader.ReadFromRegionDirect();

                UpdateRooms(rooms);
            }
            catch (Exception ex)
            {
                StatusText = $"ÐÐ»Ð´Ð°Ð°: {ex.Message}";
            }
        }

        /// <summary>
        /// "Ð‘Ò¯Ñ Ð·ÑƒÑ€Ð°Ñ…" Ñ‚Ð¾Ð²Ñ‡ â€” ÑÐ¾Ð½Ð³Ð¾ÑÐ¾Ð½ Ó©Ñ€Ó©Ó©Ð½Ð´ polyline Ñ…Ð¾Ð»Ð±Ð¾Ñ…
        /// </summary>
        private void DrawBoundary()
        {
            if (SelectedRoom == null) return;

            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            try
            {
                RoomDrawingContext.PendingRoom = SelectedRoom;
                doc.SendStringToExecute("NORMDRAWROOM\n", true, false, false);
            }
            catch (Exception ex)
            {
                StatusText = $"ÐÐ»Ð´Ð°Ð°: {ex.Message}";
            }
        }

        /// <summary>
        /// âŠ• Ñ‚Ð¾Ð²Ñ‡ â€” Ð±Ð°Ð¹Ð³Ð°Ð° polyline ÑÐ¾Ð½Ð³Ð¾Ð½ Ñ…Ð¾Ð»Ð±Ð¾Ñ…
        /// </summary>
        private void SelectBoundary()
        {
            if (SelectedRoom == null) return;
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            try
            {
                RoomDrawingContext.PendingRoom = SelectedRoom;
                doc.SendStringToExecute("NORMSELECTROOM\n", true, false, false);
            }
            catch (Exception ex)
            {
                StatusText = $"ÐÐ»Ð´Ð°Ð°: {ex.Message}";
            }
        }

        /// <summary>
        /// Zoom Ñ‚Ð¾Ð²Ñ‡ â€” NORMZOOMROOM command-Ð°Ð°Ñ€ zoom Ñ…Ð¸Ð¹Ð½Ñ
        /// </summary>
        private void ZoomToRoom()
        {
            if (SelectedRoom == null || !SelectedRoom.HasBoundary) return;

            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            try
            {
                RoomDrawingContext.PendingRoom = SelectedRoom;
                doc.SendStringToExecute("NORMZOOMROOM\n", true, false, false);
            }
            catch (Exception ex)
            {
                StatusText = $"ÐÐ»Ð´Ð°Ð°: {ex.Message}";
            }
        }

        /// <summary>
        /// SelectedRoom ÑÑÐ»Ð³ÑÑ…ÑÐ´ highlight ÑˆÐ¸Ð½ÑÑ‡Ð»ÑÑ…
        /// WPF thread-ÑÑÑ SetImpliedSelection Ð´ÑƒÑƒÐ´Ð½Ð° (zoom Ñ…Ð¸Ð¹Ñ…Ð³Ò¯Ð¹)
        /// </summary>
        private void HighlightSelectedRoom()
        {
            try
            {
                var service = new RoomBoundaryService();

                if (_selectedRoom != null && _selectedRoom.HasBoundary)
                {
                    service.ClearHighlight();
                    // SetImpliedSelection WPF thread-ÑÑÑ Ð°Ð¶Ð¸Ð»Ð»Ð°Ð½Ð°
                    var doc = Application.DocumentManager.MdiActiveDocument;
                    doc?.Editor.SetImpliedSelection(new ObjectId[] { _selectedRoom.BoundaryId });
                }
                else
                {
                    service.ClearHighlight();
                }
            }
            catch { /* Highlight Ð°Ð»Ð´Ð°Ð°Ð³ Ò¯Ð» Ñ‚Ð¾Ð¾Ð¼ÑÐ¾Ñ€Ð»Ð¾Ñ… */ }
        }

        /// <summary>
        /// NORMDRAWROOM command-Ð°Ð°Ñ Ð´ÑƒÑƒÐ´Ð°Ð³Ð´Ð°Ð½Ð°
        /// </summary>
        public void OnBoundaryDrawn(RoomInfo room, ObjectId polylineId, double drawnArea)
        {
            try
            {
                // WPF Dispatcher Ð´ÑÑÑ€ Ð°Ð¶Ð¸Ð»Ð»ÑƒÑƒÐ»Ð°Ñ… (AutoCAD thread â†’ WPF thread)
                if (System.Windows.Application.Current?.Dispatcher != null)
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        room.BoundaryId = polylineId;
                        room.DrawnArea = drawnArea;
                        OnPropertyChanged(nameof(LinkedCount));
                        OnPropertyChanged(nameof(CanLockMaster));
                        UpdateStatusText();
                    });
                }
                else
                {
                    room.BoundaryId = polylineId;
                    room.DrawnArea = drawnArea;
                    OnPropertyChanged(nameof(LinkedCount));
                    OnPropertyChanged(nameof(CanLockMaster));
                    UpdateStatusText();
                }
            }
            catch { /* UI update Ð°Ð»Ð´Ð°Ð°Ð³ Ò¯Ð» Ñ‚Ð¾Ð¾Ð¼ÑÐ¾Ñ€Ð»Ð¾Ñ… */ }
        }

        /// <summary>
        /// ÐÐ²Ñ‚Ð¾ Ñ‚Ð°Ð½Ð¸Ñ… â€” SendStringToExecute-Ð°Ð°Ñ€ NORMAUTOMATCH command Ð´ÑƒÑƒÐ´Ð½Ð°
        /// </summary>
        private void AutoMatch()
        {
            if (_allRooms.Count == 0)
            {
                StatusText = "Ð­Ñ…Ð»ÑÑÐ´ Ó©Ñ€Ó©Ó©Ð½Ð¸Ð¹ Ñ…Ò¯ÑÐ½ÑÐ³Ñ‚ ÑƒÐ½ÑˆÐ¸Ð½Ð° ÑƒÑƒ";
                return;
            }

            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            // Always do a fresh rematch.
            foreach (var room in _allRooms)
            {
                room.BoundaryId = ObjectId.Null;
                room.DrawnArea = 0;
            }
            OnPropertyChanged(nameof(LinkedCount));
            OnPropertyChanged(nameof(CanLockMaster));
            UpdateStatusText();

            doc.SendStringToExecute("NORMAUTOMATCH\n", true, false, false);
        }

        /// <summary>
        /// NORMAUTOMATCH command-Ð°Ð°Ñ Ð´ÑƒÑƒÐ´Ð°Ð³Ð´Ð°Ð½Ð°
        /// </summary>
        public void OnAutoMatchCompleted(System.Collections.Generic.List<RoomBoundaryService.LinkResult> results)
        {
            try
            {
                var action = new Action(() =>
                {
                    var service = new RoomBoundaryService();

                    foreach (var item in results)
                    {
                        RoomInfo room = null;
                        if (!string.IsNullOrWhiteSpace(item.RoomId))
                        {
                            room = _allRooms.FirstOrDefault(r =>
                                string.Equals(BuildRoomMatchKey(r), item.RoomId, StringComparison.OrdinalIgnoreCase));
                        }

                        if (room == null)
                        {
                            room = _allRooms.FirstOrDefault(r =>
                                !r.HasBoundary && r.Number == item.RoomNumber);
                        }

                        if (room == null) continue;

                        room.BoundaryId = item.PolylineId;
                        room.DrawnArea = item.DrawnArea > 0
                            ? item.DrawnArea
                            : service.GetPolylineArea(item.PolylineId, room.Area);
                    }

                    OnPropertyChanged(nameof(LinkedCount));
                    OnPropertyChanged(nameof(CanLockMaster));
                    UpdateStatusText();
                });

                if (System.Windows.Application.Current?.Dispatcher != null)
                    System.Windows.Application.Current.Dispatcher.Invoke(action);
                else
                    action();
            }
            catch { }
        }

        /// <summary>
        /// Ð‘Ò¯Ñ… Ó©Ñ€Ó©Ó©Ð½Ð¸Ð¹ Ð¶Ð°Ð³ÑÐ°Ð°Ð»Ñ‚ (AutoMatch-Ð´ Ð°ÑˆÐ¸Ð³Ð»Ð°Ð½Ð°)
        /// </summary>
        public System.Collections.Generic.List<RoomInfo> GetAllRooms()
        {
            return _allRooms.ToList();
        }

        /// <summary>
        /// ÐœÐ°ÑÑ‚ÐµÑ€ Lock/Unlock toggle
        /// </summary>
        private void ToggleMasterLock()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            try
            {
                var service = new MasterCopyService();
                using (var lockDoc = doc.LockDocument())
                {
                    if (_isMasterLocked)
                    {
                        service.UnlockMaster(doc.Database);
                        IsMasterLocked = false;
                        MasterStatusText = "";
                    }
                    else
                    {
                        if (LinkedCount == 0)
                        {
                            StatusText = "Lock Ñ…Ð¸Ð¹Ñ…Ð¸Ð¹Ð½ Ó©Ð¼Ð½Ó© Ð±Ò¯Ñ… Ó©Ñ€Ó©Ó©Ð½Ð´ Ñ…Ò¯Ñ€ÑÑ Ð·ÑƒÑ€Ð½Ð° ÑƒÑƒ";
                            return;
                        }
                        service.LockMaster(doc.Database);
                        IsMasterLocked = true;
                        MasterStatusText = "ÐœÐ°ÑÑ‚ÐµÑ€ Lock Ñ…Ð¸Ð¹Ð³Ð´ÑÑÐ½";
                        RefreshCopies();
                    }
                }
                UpdateStatusText();
            }
            catch (Exception ex)
            {
                StatusText = $"ÐÐ»Ð´Ð°Ð°: {ex.Message}";
            }
        }

        /// <summary>
        /// Ð¥ÑƒÑƒÐ»Ð±Ð°Ñ€Ñ‹Ð½ Ñ‚Ð¾Ð¾ ÑˆÐ¸Ð½ÑÑ‡Ð»ÑÑ…
        /// </summary>
        private void RefreshCopies()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            try
            {
                var service = new MasterCopyService();
                var result = service.DetectCopies(doc.Database);
                CopyCount = result.CopyCount;

                if (result.CopyCount > 0)
                {
                    MasterStatusText = $"ÐœÐ°ÑÑ‚ÐµÑ€ Lock  |  {result.CopyCount} Ñ…ÑƒÑƒÐ»Ð±Ð°Ñ€";
                }
                else if (_isMasterLocked)
                {
                    MasterStatusText = "ÐœÐ°ÑÑ‚ÐµÑ€ Lock  |  Ð¥ÑƒÑƒÐ»Ð±Ð°Ñ€ Ð¾Ð»Ð´ÑÐ¾Ð½Ð³Ò¯Ð¹";
                }
            }
            catch (Exception ex)
            {
                MasterStatusText = $"Ð¡ÐºÐ°Ð½ Ð°Ð»Ð´Ð°Ð°: {ex.Message}";
            }
        }

        /// <summary>
        /// ÐœÐ°ÑÑ‚ÐµÑ€ Ñ‚Ó©Ð»Ó©Ð² ÑˆÐ¸Ð½ÑÑ‡Ð»ÑÑ… (Ð·ÑƒÑ€Ð°Ð³Ð½Ð°Ð°Ñ ÑƒÐ½ÑˆÐ¸Ñ…)
        /// </summary>
        public void RefreshMasterState()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            try
            {
                var service = new MasterCopyService();
                IsMasterLocked = service.IsMasterLocked(doc.Database);
                if (_isMasterLocked)
                {
                    RefreshCopies();
                }
                OnPropertyChanged(nameof(CanLockMaster));
            }
            catch { }
        }

        private void UpdateRooms(System.Collections.Generic.List<RoomInfo> rooms)
        {
            _allRooms.Clear();
            foreach (var room in rooms)
                _allRooms.Add(room);

            // ÐžÐ´Ð¾Ð¾ Ð±Ð°Ð¹Ð³Ð°Ð° XData-Ñ‚Ð°Ð¹ polyline-ÑƒÑƒÐ´Ñ‹Ð³ Ð¾Ð»Ð¶ Ñ…Ð¾Ð»Ð±Ð¾Ñ…
            TryReconnectBoundaries();

            OnPropertyChanged(nameof(RoomCount));
            OnPropertyChanged(nameof(TotalArea));
            OnPropertyChanged(nameof(LinkedCount));
            OnPropertyChanged(nameof(CanLockMaster));
            UpdateStatusText();
            RefreshMasterState();
        }

        /// <summary>
        /// XData-тай polyline-уудыг скан хийж, өрөөтэй дахин холбох
        /// </summary>
        private void TryReconnectBoundaries()
        {
            try
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;

                var service = new RoomBoundaryService();
                var existingById = service.FindExistingRoomBoundariesByRoomId(doc.Database);
                var existingByNumber = service.FindExistingRoomBoundaries(doc.Database);

                foreach (var room in _allRooms)
                {
                    ObjectId id = ObjectId.Null;
                    if (!string.IsNullOrWhiteSpace(room.RoomId) && existingById.TryGetValue(room.RoomId, out id))
                    {
                        room.BoundaryId = id;
                        room.DrawnArea = service.GetPolylineArea(id, room.Area);
                        continue;
                    }

                    if (existingByNumber.TryGetValue(room.Number, out id))
                    {
                        room.BoundaryId = id;
                        room.DrawnArea = service.GetPolylineArea(id, room.Area);
                    }
                }
            }
            catch { /* Reconnect алдааг үл тоомсорлох */ }
        }
        private string BuildRoomMatchKey(RoomInfo room)
        {
            if (!string.IsNullOrWhiteSpace(room.RoomId) && System.Text.RegularExpressions.Regex.IsMatch(room.RoomId, "[A-Za-zА-Яа-я]"))
                return room.RoomId.Trim();
            return $"#N{room.Number}:{(room.Name ?? string.Empty).Trim()}";
        }

        private void UpdateStatusText()
        {
            if (RoomCount == 0)
            {
                StatusText = "Ó¨Ñ€Ó©Ó© ÑƒÐ½ÑˆÐ°Ð°Ð³Ò¯Ð¹ Ð±Ð°Ð¹Ð½Ð°";
            }
            else
            {
                StatusText = $"ÐÐ¸Ð¹Ñ‚: {RoomCount} Ó©Ñ€Ó©Ó©  |  Ð¢Ð°Ð»Ð±Ð°Ð¹: {TotalArea:F2} Ð¼Â²  |  Ð—ÑƒÑ€ÑÐ°Ð½: {LinkedCount}/{RoomCount}";
            }
        }

        private bool FilterRoom(object obj)
        {
            if (string.IsNullOrWhiteSpace(_searchText)) return true;
            if (obj is RoomInfo room)
            {
                return room.Name.IndexOf(_searchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                       room.Number.ToString().Contains(_searchText) ||
                       (!string.IsNullOrEmpty(room.SectionName) &&
                        room.SectionName.IndexOf(_searchText, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            return false;
        }
    }
}








