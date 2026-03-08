using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;

namespace NormSchematicStudio;

public partial class MainWindow : Window
{
    private enum ToolMode { Select, AddDevice, AddText, Connect }

    private sealed class NodeModel
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string Type { get; set; } = "DEVICE";
        public string Label { get; set; } = "";
        public double X { get; set; }
        public double Y { get; set; }
        public double W { get; set; } = 90;
        public double H { get; set; } = 34;
    }

    private sealed class EdgeModel
    {
        public string FromId { get; set; } = "";
        public string ToId { get; set; } = "";
        public string Label { get; set; } = "";
    }

    private sealed class FrameSpec
    {
        public double InnerWmm { get; init; } = 26500;
        public double InnerHmm { get; init; } = 39000;
        public double LeftOff { get; init; } = 500;
        public double RightOff { get; init; } = 500;
        public double TopOff { get; init; } = 500;
        public double BottomOff { get; init; } = 500;
    }


    private sealed class SPort
    {
        public string Type { get; set; } = "";
        public int Count { get; set; }
        public string Cable { get; set; } = "";
    }

    private sealed class SUnit
    {
        public string Name { get; set; } = "";
        public List<SPort> Ports { get; set; } = new();
        public bool HasONT { get; set; } = true;
        public int CableACount => Ports.Where(p => p.Cable == "A").Sum(p => p.Count);
        public int CableBCount => Ports.Where(p => p.Cable == "B").Sum(p => p.Count);
    }

    private sealed class SFloor
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
        public List<SUnit> Units { get; set; } = new();
        public string FdbRatio { get; set; } = "1:8";
        public int FdbPortCount { get; set; } = 8;
    }

    private sealed class SBuilding
    {
        public int Floors { get; set; }
        public int LanPerFloor { get; set; }
        public List<SFloor> FloorItems { get; set; } = new();
        public List<SRiserSegment> RiserSegments { get; set; } = new();
        public SCentralBlock CentralBlock { get; set; } = new();
    }

    private sealed class SRiserSegment
    {
        public int FromFloor { get; set; }
        public int ToFloor { get; set; }
        public string Label { get; set; } = "";
    }

    private sealed class SCentralBlock
    {
        public string RackLabel { get; set; } = "Rack";
        public string OdfLabel { get; set; } = "ODF";
        public string ExternalInputLabel { get; set; } = "Оптик оролт";
    }

    private sealed class SchematicLayout
    {
        public double WorkAreaWmm { get; set; }
        public double WorkAreaHmm { get; set; }
        public List<FloorLayout> Floors { get; set; } = new();
        public double RiserXmm { get; set; }
        public double RiserTopYmm { get; set; }
        public double RiserBottomYmm { get; set; }
        public CentralBlockLayout Central { get; set; } = new();
        public double LegendXmm { get; set; }
        public double LegendYmm { get; set; }
        public double LegendWmm { get; set; }
        public double LegendHmm { get; set; }
    }

    private sealed class FloorLayout
    {
        public int FloorIndex { get; set; }
        public string FloorName { get; set; } = "";
        public double Ymm { get; set; }
        public double Hmm { get; set; }
        public double LabelXmm { get; set; }
        public double LabelYmm { get; set; }
        public double FdbXmm { get; set; }
        public double FdbYmm { get; set; }
        public double FdbWmm { get; set; }
        public double FdbHmm { get; set; }
        public string FdbRatio { get; set; } = "1:8";
        public List<UnitLayout> Units { get; set; } = new();
        public string RiserLabel { get; set; } = "";
    }

    private sealed class UnitLayout
    {
        public string Name { get; set; } = "";
        public double Xmm { get; set; }
        public double Ymm { get; set; }
        public double Wmm { get; set; }
        public double Hmm { get; set; }
        public bool HasONT { get; set; }
        public double OntXmm { get; set; }
        public double OntYmm { get; set; }
        public double OntWmm { get; set; }
        public double OntHmm { get; set; }
        public List<PortRowLayout> PortRows { get; set; } = new();
        public string CableSummaryText { get; set; } = "";
        public double SummaryXmm { get; set; }
        public double SummaryYmm { get; set; }
        public double ConnLineStartXmm { get; set; }
        public double ConnLineStartYmm { get; set; }
        public double ConnLineEndXmm { get; set; }
        public double ConnLineEndYmm { get; set; }
        public string ConnLabel { get; set; } = "D";
    }

    private sealed class PortRowLayout
    {
        public string Label { get; set; } = "";
        public string CableType { get; set; } = "";
        public double Ymm { get; set; }
    }

    private sealed class CentralBlockLayout
    {
        public double Xmm { get; set; }
        public double Ymm { get; set; }
        public double Wmm { get; set; }
        public double Hmm { get; set; }
        public double RackXmm { get; set; }
        public double RackYmm { get; set; }
        public double RackWmm { get; set; }
        public double RackHmm { get; set; }
        public double OdfXmm { get; set; }
        public double OdfYmm { get; set; }
        public double OdfWmm { get; set; }
        public double OdfHmm { get; set; }
        public double ArrowStartXmm { get; set; }
        public double ArrowStartYmm { get; set; }
        public double ArrowEndXmm { get; set; }
        public double ArrowEndYmm { get; set; }
    }
    // Schematic layout constants (mm)
    private const double SchMarginMm = 1200;
    private const double FloorLabelWidthMm = 2000;
    private const double RiserWidthMm = 2500;
    private const double CentralBlockHmm = 5000;
    private const double LegendHmm = 3000;
    private const double PortRowHmm = 600;
    private const double OntBoxHmm = 800;
    private const double OntBoxWmm = 1400;
    private const double FdbBoxWmm = 1800;
    private const double FdbBoxHmm = 1000;
    private const double UnitPaddingMm = 400;
    private const double UnitMinWmm = 3000;

    private const double HandleSize = 8.0;
    private const double OuterFrameOriginX = 4;
    private const double OuterFrameOriginY = 4;
    private const double CanvasTailMarginMm = 8;
    private const double CadWheelBase = 1.2;
    private const double MinPanZoomFactor = 1.08;
    private const double PageGapMm = 1200;

    private ToolMode _tool = ToolMode.Select;
    private readonly List<NodeModel> _nodes = new();
    private readonly List<EdgeModel> _edges = new();
    private readonly Stack<string> _undo = new();
    private readonly Stack<string> _redo = new();

    private string? _selectedId;
    private string? _connectStartId;
    private bool _isDragging;
    private bool _isResizing;
    private bool _isPanning;
    private Point _dragOffset;
    private Point _resizeStart;
    private Size _resizeStartSize;
    private Point _panStartMouse;
    private Point _panLastMouse;
    private Point _panAnchorCanvas;
    private double _panStartH;
    private double _panStartV;

    private double _zoom = 1.0;
    private bool _suppressZoomHandler;
    private readonly FrameSpec _frame = new();
    private bool _isLandscape;
    private int _pageCount = 1;
    private SBuilding? _schematic;

    public MainWindow()
    {
        InitializeComponent();
        ZoomText.Text = "100%";
        ToolText.Text = "Select";
        UpdateToolButtonStyles();
        GenerateTemplate();
    }

    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        var wa = SystemParameters.WorkArea;

        if (double.IsNaN(Left) || double.IsInfinity(Left) || Left < wa.Left || Left > wa.Right - 100)
            Left = wa.Left + Math.Max(0, (wa.Width - Width) * 0.5);

        if (double.IsNaN(Top) || double.IsInfinity(Top) || Top < wa.Top || Top > wa.Bottom - 100)
            Top = wa.Top + Math.Max(0, (wa.Height - Height) * 0.5);

        WindowState = WindowState.Normal;
        ShowInTaskbar = true;
        Activate();
        Topmost = true;
        Topmost = false;
        Focus();
    }
    private void SelectTool_Click(object sender, RoutedEventArgs e) => SetTool(ToolMode.Select);
    private void AddDeviceTool_Click(object sender, RoutedEventArgs e) => SetTool(ToolMode.AddDevice);
    private void AddTextTool_Click(object sender, RoutedEventArgs e) => SetTool(ToolMode.AddText);
    private void ConnectTool_Click(object sender, RoutedEventArgs e) => SetTool(ToolMode.Connect);


    private void OrientationBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!IsLoaded) return;
        ApplyLayoutSettings();
    }

    private void ApplyLayout_Click(object sender, RoutedEventArgs e)
    {
        ApplyLayoutSettings();
    }

    private void ApplyLayoutSettings()
    {
        _isLandscape = (OrientationBox?.SelectedIndex ?? 0) == 1;
        _pageCount = Math.Clamp(ParseOrDefault(PagesBox?.Text, 1, 1, 2), 1, 2);
        Redraw();
        FitToFrame();
        StatusText.Text = $"Layout: {(_isLandscape ? "Landscape" : "Portrait")}, pages={_pageCount}";
    }
    private void SetTool(ToolMode mode)
    {
        _tool = mode;
        ToolText.Text = mode.ToString();
        UpdateToolButtonStyles();
        StatusText.Text = $"Tool: {mode}";
    }

    private void UpdateToolButtonStyles()
    {
        if (SelectToolButton == null || DeviceToolButton == null || TextToolButton == null || ConnectToolButton == null) return;

        var normalBg = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1D2734"));
        var normalBorder = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#334256"));
        var activeBg = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1264A3"));
        var activeBorder = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2EA8FF"));

        void SetState(Button b, bool active)
        {
            b.Background = active ? activeBg : normalBg;
            b.BorderBrush = active ? activeBorder : normalBorder;
        }

        SetState(SelectToolButton, _tool == ToolMode.Select);
        SetState(DeviceToolButton, _tool == ToolMode.AddDevice);
        SetState(TextToolButton, _tool == ToolMode.AddText);
        SetState(ConnectToolButton, _tool == ToolMode.Connect);
    }

    private void Generate_Click(object sender, RoutedEventArgs e) => GenerateTemplate();

    private void GenerateTemplate()
    {
        SaveStateForUndo();
        _nodes.Clear();
        _edges.Clear();
        _schematic = BuildSchematicModel();
        StatusText.Text = $"Schematic generated: {_schematic.Floors} floors";
        Redraw();
        FitToFrame();
    }

    private SBuilding BuildSchematicModel()
    {
        int floors = ParseOrDefault(FloorsBox?.Text, 5, 1, 20);
        int lanPerFloor = ParseOrDefault(LanBox?.Text, 10, 1, 256);

        var model = new SBuilding
        {
            Floors = floors,
            LanPerFloor = lanPerFloor
        };

        for (int i = 1; i <= floors; i++)
        {
            int aLan = Math.Max(1, (int)Math.Round(lanPerFloor * 0.35));
            int bLan = Math.Max(1, (int)Math.Round(lanPerFloor * 0.30));
            int cLan = Math.Max(1, lanPerFloor - aLan - bLan);

            var floor = new SFloor
            {
                Index = i,
                Name = $"{i}-р давхар",
                FdbRatio = "1:8",
                FdbPortCount = 8,
                Units = new List<SUnit>
                {
                    new SUnit { Name = "A", HasONT = true, Ports = new List<SPort>{ new(){ Type="LAN", Count=aLan, Cable="A"}, new(){ Type="IPTV", Count=1, Cable="A"}, new(){ Type="TEL", Count=1, Cable="B"} } },
                    new SUnit { Name = "B", HasONT = true, Ports = new List<SPort>{ new(){ Type="LAN", Count=bLan, Cable="A"}, new(){ Type="IPTV", Count=1, Cable="A"}, new(){ Type="TEL", Count=1, Cable="B"} } },
                    new SUnit { Name = "C", HasONT = false, Ports = new List<SPort>{ new(){ Type="LAN", Count=cLan, Cable="A"}, new(){ Type="IPTV", Count=1, Cable="A"} } }
                }
            };
            model.FloorItems.Add(floor);
        }

        for (int i = 0; i < floors; i++)
        {
            model.RiserSegments.Add(new SRiserSegment
            {
                FromFloor = i,
                ToFloor = i + 1,
                Label = $"K-{i + 1}"
            });
        }

        return model;
    }

    private void Validate_Click(object sender, RoutedEventArgs e)
    {
        RuleList.Items.Clear();
        var issues = RunRuleCheck();
        if (issues.Count == 0)
        {
            RuleList.Items.Add("OK: No rule issues.");
            StatusText.Text = "Validation OK";
            return;
        }

        foreach (var i in issues) RuleList.Items.Add(i);
        StatusText.Text = $"Validation: {issues.Count} issue(s)";
    }

    private List<string> RunRuleCheck()
    {
        var result = new List<string>();

        var duplicateLabels = _nodes.Where(n => !string.IsNullOrWhiteSpace(n.Label))
            .GroupBy(n => n.Label.Trim(), StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Count() > 1)
            .Select(g => g.Key)
            .ToList();

        foreach (var d in duplicateLabels) result.Add($"[P1] Duplicate label: {d}");

        var edgeNodeIds = _edges.SelectMany(e => new[] { e.FromId, e.ToId }).ToHashSet();
        foreach (var n in _nodes.Where(n => n.Type == "ONT"))
        {
            if (!edgeNodeIds.Contains(n.Id)) result.Add($"[P2] Unconnected ONT: {n.Label}");
        }

        var (ix, iy, iw, ih) = GetInnerRectPx();
        foreach (var n in _nodes)
        {
            bool outside = n.X < ix || n.Y < iy || (n.X + n.W) > (ix + iw) || (n.Y + n.H) > (iy + ih);
            if (outside) result.Add($"[P2] Node outside frame: {n.Label}");
        }

        foreach (var e in _edges)
        {
            if (_nodes.All(n => n.Id != e.FromId) || _nodes.All(n => n.Id != e.ToId)) result.Add("[P1] Edge references missing node");
        }

        return result;
    }

    private void Undo_Click(object sender, RoutedEventArgs e)
    {
        if (_undo.Count == 0) return;
        _redo.Push(SerializeState());
        DeserializeState(_undo.Pop());
        Redraw();
    }

    private void Redo_Click(object sender, RoutedEventArgs e)
    {
        if (_redo.Count == 0) return;
        _undo.Push(SerializeState());
        DeserializeState(_redo.Pop());
        Redraw();
    }

    private void ZoomSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (ZoomSlider == null) return;
        if (_suppressZoomHandler) return;

        double minFitZoom = GetMinimumAllowedZoom();
        double requested = ZoomSlider.Value;
        double clamped = Math.Clamp(requested, minFitZoom, ZoomSlider.Maximum);
        if (Math.Abs(clamped - requested) > 0.0001)
        {
            _suppressZoomHandler = true;
            ZoomSlider.Value = clamped;
            _suppressZoomHandler = false;
        }

        if (!IsLoaded || PreviewCanvas == null || EditorScroll == null)
        {
            _zoom = clamped;
            if (ZoomText != null)
                ZoomText.Text = $"{_zoom * 100:0.0}%";
            return;
        }

        double oldZoom = _zoom;
        var oldInner = GetInnerRectPxForZoom(oldZoom);

        double viewportW = EditorScroll.ViewportWidth;
        double viewportH = EditorScroll.ViewportHeight;
        double centerX = EditorScroll.HorizontalOffset + viewportW * 0.5;
        double centerY = EditorScroll.VerticalOffset + viewportH * 0.5;

        double rx = oldInner.w > 1 ? (centerX - oldInner.x) / oldInner.w : 0.5;
        double ry = oldInner.h > 1 ? (centerY - oldInner.y) / oldInner.h : 0.5;
        rx = Math.Clamp(rx, -0.25, 1.25);
        ry = Math.Clamp(ry, -0.25, 1.25);

        _zoom = clamped;
        if (ZoomText != null)
            ZoomText.Text = $"{_zoom * 100:0.0}%";

        Redraw();

        var newInner = GetInnerRectPx();
        double targetX = newInner.x + rx * newInner.w - viewportW * 0.5;
        double targetY = newInner.y + ry * newInner.h - viewportH * 0.5;
        ScrollToFrameClamped(targetX, targetY);
    }



    private void EditorScroll_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (ZoomSlider == null || EditorScroll == null || PreviewCanvas == null)
        {
            e.Handled = true;
            return;
        }

        var mouseInViewport = e.GetPosition(EditorScroll);
        var mouseInCanvas = e.GetPosition(PreviewCanvas);

        double oldZoom = _zoom;
        if (oldZoom <= 0.0001) oldZoom = 1.0;

        // AutoCAD-style anchor: keep mouse-point fixed while zooming.
        double anchorMmX = (mouseInCanvas.X - OuterFrameOriginX) * 25.0 / oldZoom;
        double anchorMmY = (mouseInCanvas.Y - OuterFrameOriginY) * 25.0 / oldZoom;

        // High-precision wheel zoom: multiplicative scale like CAD (better anchor stability).
        double notches = e.Delta / 120.0;
        double zoomFactor = Math.Pow(CadWheelBase, notches);
        double minFitZoom = GetMinimumAllowedZoom();
        double nextZoom = Math.Clamp(_zoom * zoomFactor, minFitZoom, ZoomSlider.Maximum);
        if (Math.Abs(nextZoom - _zoom) < 0.0001)
        {
            e.Handled = true;
            return;
        }

        _zoom = nextZoom;
        _suppressZoomHandler = true;
        ZoomSlider.Value = nextZoom;
        _suppressZoomHandler = false;

        if (ZoomText != null)
            ZoomText.Text = $"{_zoom * 100:0.0}%";

        Redraw();

        double anchoredWorldX = OuterFrameOriginX + anchorMmX * _zoom / 25.0;
        double anchoredWorldY = OuterFrameOriginY + anchorMmY * _zoom / 25.0;
        double targetX = anchoredWorldX - mouseInViewport.X;
        double targetY = anchoredWorldY - mouseInViewport.Y;
        ScrollToFrameClamped(targetX, targetY);

        e.Handled = true;
    }

    private void EditorScroll_PreviewMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Middle) return;
        _isPanning = true;
        _panStartMouse = e.GetPosition(EditorScroll);
        _panLastMouse = _panStartMouse;
        _panAnchorCanvas = new Point(
            EditorScroll.HorizontalOffset + _panStartMouse.X,
            EditorScroll.VerticalOffset + _panStartMouse.Y);
        _panStartH = EditorScroll.HorizontalOffset;
        _panStartV = EditorScroll.VerticalOffset;
        Mouse.Capture(EditorScroll);
        Cursor = Cursors.Hand;
        e.Handled = true;
    }

    private void EditorScroll_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (!_isPanning) return;
        var cur = e.GetPosition(EditorScroll);
        double targetX = _panAnchorCanvas.X - cur.X;
        double targetY = _panAnchorCanvas.Y - cur.Y;
        ScrollToFrameClamped(targetX, targetY);
        _panLastMouse = cur;
        e.Handled = true;
    }

    private void EditorScroll_PreviewMouseUp(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Middle) return;
        _isPanning = false;
        if (!_isDragging && !_isResizing)
            Mouse.Capture(null);
        Cursor = Cursors.Arrow;
        e.Handled = true;
    }
    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control && e.Key == Key.D0)
        {
            FitToFrame();
            e.Handled = true;
        }
    }

    private void PreviewCanvas_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton == MouseButton.Middle)
        {
            // Pan is handled by ScrollViewer handlers to avoid double-pan lag.
            e.Handled = true;
        }
    }

    private void PreviewCanvas_MouseUp(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton == MouseButton.Middle)
        {
            e.Handled = true;
        }
    }

    private void PreviewCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        var p = e.GetPosition(PreviewCanvas);

        if (_tool == ToolMode.AddDevice)
        {
            SaveStateForUndo();
            var n = new NodeModel { Type = "DEVICE", Label = $"D{_nodes.Count + 1}", X = p.X, Y = p.Y, W = 90, H = 34 };
            if (!CanPlaceInsideInner(n.X, n.Y, n.W, n.H))
            {
                StatusText.Text = "Хүрээнээс гадуур объект нэмэхгүй.";
                return;
            }
            _nodes.Add(n);
            Redraw();
            return;
        }

        if (_tool == ToolMode.AddText)
        {
            SaveStateForUndo();
            var n = new NodeModel { Type = "TEXT", Label = "Text", X = p.X, Y = p.Y, W = 120, H = 24 };
            if (!CanPlaceInsideInner(n.X, n.Y, n.W, n.H))
            {
                StatusText.Text = "Хүрээнээс гадуур текст нэмэхгүй.";
                return;
            }
            _nodes.Add(n);
            Redraw();
            return;
        }

        var hit = HitTestNode(p);
        _selectedId = hit?.Id;
        SyncSelectionUi();

        if (_tool == ToolMode.Connect)
        {
            if (hit == null) return;
            if (_connectStartId == null)
            {
                _connectStartId = hit.Id;
                StatusText.Text = $"Connect start: {hit.Label}";
            }
            else if (_connectStartId != hit.Id)
            {
                SaveStateForUndo();
                _edges.Add(new EdgeModel { FromId = _connectStartId, ToId = hit.Id, Label = "A-XX-1" });
                _connectStartId = null;
                Redraw();
            }
            return;
        }

        if (_tool == ToolMode.Select && hit != null)
        {
            if (IsOnResizeHandle(hit, p))
            {
                SaveStateForUndo();
                _isResizing = true;
                _resizeStart = p;
                _resizeStartSize = new Size(hit.W, hit.H);
                PreviewCanvas.CaptureMouse();
                Cursor = Cursors.SizeNWSE;
            }
            else
            {
                SaveStateForUndo();
                _isDragging = true;
                _dragOffset = new Point(p.X - hit.X, p.Y - hit.Y);
                PreviewCanvas.CaptureMouse();
                Cursor = Cursors.SizeAll;
            }
        }
    }

    private void PreviewCanvas_MouseMove(object sender, MouseEventArgs e)
    {
        if (_isPanning) return;
        if (_selectedId == null) return;
        var node = _nodes.FirstOrDefault(n => n.Id == _selectedId);
        if (node == null) return;

        var p = e.GetPosition(PreviewCanvas);
        if (_isDragging)
        {
            node.X = p.X - _dragOffset.X;
            node.Y = p.Y - _dragOffset.Y;
            ConstrainNodeToInnerFrame(node);
            Redraw();
        }
        else if (_isResizing)
        {
            var dw = p.X - _resizeStart.X;
            var dh = p.Y - _resizeStart.Y;
            node.W = Math.Max(36, _resizeStartSize.Width + dw);
            node.H = Math.Max(18, _resizeStartSize.Height + dh);
            ConstrainNodeToInnerFrame(node);
            Redraw();
        }
    }

    private void PreviewCanvas_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        _isDragging = false;
        _isResizing = false;
        if (!_isPanning) PreviewCanvas.ReleaseMouseCapture();
        Cursor = Cursors.Arrow;
    }

    private void SelectedLabelBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_selectedId == null) return;
        var n = _nodes.FirstOrDefault(x => x.Id == _selectedId);
        if (n == null) return;
        n.Label = SelectedLabelBox.Text;
        Redraw();
    }

    private NodeModel? HitTestNode(Point p)
    {
        for (int i = _nodes.Count - 1; i >= 0; i--)
        {
            var n = _nodes[i];
            if (p.X >= n.X && p.X <= n.X + n.W && p.Y >= n.Y && p.Y <= n.Y + n.H) return n;
        }
        return null;
    }

    private static bool IsOnResizeHandle(NodeModel n, Point p)
    {
        double hx = n.X + n.W - HandleSize;
        double hy = n.Y + n.H - HandleSize;
        return p.X >= hx && p.X <= hx + HandleSize && p.Y >= hy && p.Y <= hy + HandleSize;
    }

    private void SyncSelectionUi()
    {
        var n = _nodes.FirstOrDefault(x => x.Id == _selectedId);
        if (n == null)
        {
            SelectedLabelBox.Text = string.Empty;
            SelectedTypeText.Text = "-";
            return;
        }

        SelectedLabelBox.Text = n.Label;
        SelectedTypeText.Text = n.Type;
    }

    private void FitToFrame()
    {
        if (EditorScroll.ViewportWidth <= 0 || EditorScroll.ViewportHeight <= 0)
        {
            Dispatcher.BeginInvoke(new Action(FitToFrame), System.Windows.Threading.DispatcherPriority.Background);
            return;
        }

        double z = GetMinimumAllowedZoom();
        _suppressZoomHandler = true;
        ZoomSlider.Value = z;
        _suppressZoomHandler = false;
        _zoom = z;
        Redraw();
        ScrollToFrameClamped(0, 0);
        StatusText.Text = "Fit frame (Ctrl+0)";
    }

    private double GetFrameFitZoom()
    {
        if (EditorScroll == null || ZoomSlider == null) return 1.0;
        if (EditorScroll.ViewportWidth <= 0 || EditorScroll.ViewportHeight <= 0) return _zoom > 0 ? _zoom : 1.0;

        double mmToPx = 1.0 / 25.0;
        var (innerWmm, innerHmm) = GetActiveInnerSizeMm();
        double pageW = (innerWmm + _frame.LeftOff + _frame.RightOff) * mmToPx;
        double pageH = (innerHmm + _frame.TopOff + _frame.BottomOff) * mmToPx;
        double gap = PageGapMm * mmToPx;
        double frameW = pageW * _pageCount + gap * Math.Max(0, _pageCount - 1);
        double frameH = pageH;

        double vw = Math.Max(40, EditorScroll.ViewportWidth - 16);
        double vh = Math.Max(40, EditorScroll.ViewportHeight - 16);
        double z = Math.Min(vw / frameW, vh / frameH);
        return Math.Clamp(z, ZoomSlider.Minimum, ZoomSlider.Maximum);
    }

    private double GetMinimumAllowedZoom()
    {
        double fit = GetFrameFitZoom();
        if (ZoomSlider == null) return fit;
        return Math.Clamp(fit * MinPanZoomFactor, ZoomSlider.Minimum, ZoomSlider.Maximum);
    }

    private void ScrollToFrameClamped(double targetX, double targetY)
    {
        var (ix, iy, iw, ih) = GetOuterRectPx();
        double vw = EditorScroll.ViewportWidth;
        double vh = EditorScroll.ViewportHeight;

        double minX = ix;
        double maxX = ix + Math.Max(0, iw - vw);
        double minY = iy;
        double maxY = iy + Math.Max(0, ih - vh);

        double x = Math.Clamp(targetX, minX, maxX);
        double y = Math.Clamp(targetY, minY, maxY);

        x = Math.Max(0, x);
        y = Math.Max(0, y);
        EditorScroll.ScrollToHorizontalOffset(x);
        EditorScroll.ScrollToVerticalOffset(y);
    }
    private void Redraw()
    {
        PreviewCanvas.Children.Clear();

        var pen = Brushes.Black;
        double U(double mm) => mm * _zoom / 25.0;

        var (innerWmm, innerHmm) = GetActiveInnerSizeMm();
        double outerW = U(innerWmm + _frame.LeftOff + _frame.RightOff);
        double outerH = U(innerHmm + _frame.TopOff + _frame.BottomOff);
        double gap = U(PageGapMm);

        double totalW = outerW * _pageCount + gap * Math.Max(0, _pageCount - 1);
        double outerX0 = OuterFrameOriginX;
        double outerY = OuterFrameOriginY;

        PreviewCanvas.Width = outerX0 + totalW + U(CanvasTailMarginMm);
        PreviewCanvas.Height = outerY + outerH + U(CanvasTailMarginMm);

        for (int page = 0; page < _pageCount; page++)
        {
            double outerX = outerX0 + page * (outerW + gap);
            double innerX = outerX + U(_frame.LeftOff);
            double innerY = outerY + U(_frame.TopOff);
            double innerW = U(innerWmm);
            double innerH = U(innerHmm);

            DrawRect(outerX, outerY, outerW, outerH, pen, 1.2);
            DrawRect(innerX, innerY, innerW, innerH, pen, 1.2);
            DrawFrameBands(innerX, innerY, innerW, innerH, outerX, outerY, outerW, outerH, pen, _zoom);
            if (page == 0)
            {
                if (_schematic != null && _schematic.FloorItems.Count > 0)
                    DrawGeneratedSchematic(innerX, innerY, innerW, innerH, pen);
                else
                    DrawSingleFloorGrid(innerX, innerY, innerW, innerH, pen);
            }
        }

        foreach (var e in _edges)
        {
            var a = _nodes.FirstOrDefault(n => n.Id == e.FromId);
            var b = _nodes.FirstOrDefault(n => n.Id == e.ToId);
            if (a == null || b == null) continue;
            var x1 = a.X + a.W;
            var y1 = a.Y + a.H * 0.5;
            var x2 = b.X;
            var y2 = b.Y + b.H * 0.5;
            DrawLine(x1, y1, x2, y2, Brushes.Black, 1);
            DrawText(e.Label, (x1 + x2) * 0.5, (y1 + y2) * 0.5 - 4, 9, Brushes.Black);
        }

        foreach (var n in _nodes)
        {
            ConstrainNodeToInnerFrame(n);
            if (n.Type == "TEXT")
            {
                DrawText(n.Label, n.X, n.Y + 12, 12, Brushes.Black);
                continue;
            }

            var stroke = n.Id == _selectedId ? Brushes.DodgerBlue : Brushes.Black;
            DrawRect(n.X, n.Y, n.W, n.H, stroke, n.Id == _selectedId ? 2 : 1.1);
            DrawTextCentered(n.Label, n.X, n.Y, n.W, n.H, 10, Brushes.Black);

            if (n.Id == _selectedId)
            {
                DrawRect(n.X + n.W - HandleSize, n.Y + n.H - HandleSize, HandleSize, HandleSize, Brushes.DodgerBlue, 1.5);
            }
        }
    }
    private void DrawFrameBands(double innerX, double innerY, double innerW, double innerH, double outerX, double outerY, double outerW, double outerH, Brush pen, double zoom)
    {
        var letters = new[] { "A", "Б", "В", "Г", "Д", "Е" };
        var numbers = Enumerable.Range(1, 8).ToArray();

        DrawRect(innerX, outerY, innerW, innerY - outerY, pen, 1);
        DrawRect(innerX, innerY + innerH, innerW, (outerY + outerH) - (innerY + innerH), pen, 1);
        DrawRect(outerX, innerY, innerX - outerX, innerH, pen, 1);
        DrawRect(innerX + innerW, innerY, (outerX + outerW) - (innerX + innerW), innerH, pen, 1);

        double colW = innerW / letters.Length;
        double topH = innerY - outerY;
        double bottomH = (outerY + outerH) - (innerY + innerH);

        for (int i = 0; i < letters.Length; i++)
        {
            double x = innerX + i * colW;
            DrawLine(x, outerY, x, outerY + topH, pen, 1);
            DrawLine(x, innerY + innerH, x, innerY + innerH + bottomH, pen, 1);
            DrawTextCentered(letters[i], x, outerY, colW, topH, 9 * zoom, Brushes.Black);
            DrawTextCentered(letters[i], x, innerY + innerH, colW, bottomH, 9 * zoom, Brushes.Black);
        }
        DrawLine(innerX + innerW, outerY, innerX + innerW, outerY + topH, pen, 1);
        DrawLine(innerX + innerW, innerY + innerH, innerX + innerW, innerY + innerH + bottomH, pen, 1);

        double rowH = innerH / numbers.Length;
        double leftW = innerX - outerX;
        double rightW = (outerX + outerW) - (innerX + innerW);

        for (int i = 0; i < numbers.Length; i++)
        {
            double y = innerY + i * rowH;
            DrawLine(outerX, y, outerX + leftW, y, pen, 1);
            DrawLine(innerX + innerW, y, innerX + innerW + rightW, y, pen, 1);
            DrawTextCentered(numbers[i].ToString(), outerX, y, leftW, rowH, 9 * zoom, Brushes.Black);
            DrawTextCentered(numbers[i].ToString(), innerX + innerW, y, rightW, rowH, 9 * zoom, Brushes.Black);
        }

                double stampW = innerW * 0.34;
        double stampH = 120 * zoom;
        double stampX = innerX + innerW - stampW - 8;
        double stampY = innerY + innerH - stampH - 8;
        DrawStamp(stampX, stampY, stampW, stampH, pen);
    }

    private void DrawStamp(double x, double y, double w, double h, Brush pen)
    {
        DrawRect(x, y, w, h, pen, 1);
        for (int i = 1; i <= 4; i++) DrawLine(x, y + (h / 5.0) * i, x + w, y + (h / 5.0) * i, pen, 1);
        double c1 = x + w * 0.3;
        double c2 = x + w * 0.7;
        DrawLine(c1, y, c1, y + h, pen, 1);
        DrawLine(c2, y, c2, y + h, pen, 1);
    }



    private void DrawSingleFloorGrid(double innerX, double innerY, double innerW, double innerH, Brush pen)
    {
        double U(double mm) => mm * _zoom / 25.0;

        // Simple base grid: 4 columns + one fixed floor row (5000 mm)
        double left = innerX + U(1200);
        double top = innerY + U(1200);
        double gridW = innerW - U(2400);
        double headerH = U(1000);
        double rowH = U(5000);
        double gridH = headerH + rowH;

        if (gridW <= U(4000) || (top + gridH) >= (innerY + innerH - U(600)))
            return;

        DrawRect(left, top, gridW, gridH, pen, 1);
        DrawLine(left, top + headerH, left + gridW, top + headerH, pen, 1);

        double c1 = gridW * 0.08; // №
        double c2 = gridW * 0.33; // Өрөө
        double c3 = gridW * 0.27; // Айлын самбар
        double c4 = gridW - (c1 + c2 + c3); // Босоо сувагчлал

        double x1 = left + c1;
        double x2 = x1 + c2;
        double x3 = x2 + c3;

        DrawLine(x1, top, x1, top + gridH, pen, 1);
        DrawLine(x2, top, x2, top + gridH, pen, 1);
        DrawLine(x3, top, x3, top + gridH, pen, 1);

        DrawTextCentered("№", left, top, c1, headerH, 10 * _zoom, Brushes.Black);
        DrawTextCentered("Өрөө", x1, top, c2, headerH, 10 * _zoom, Brushes.Black);
        DrawTextCentered("Айлын самбар", x2, top, c3, headerH, 10 * _zoom, Brushes.Black);
        DrawTextCentered("Босоо сувагчлал", x3, top, c4, headerH, 10 * _zoom, Brushes.Black);

        DrawTextCentered("1-р", left, top + headerH, c1, rowH, 10 * _zoom, Brushes.Black);
    }


    private SchematicLayout CalculateSchematicLayout(SBuilding model, double innerWmm, double innerHmm)
    {
        var layout = new SchematicLayout { WorkAreaWmm = innerWmm, WorkAreaHmm = innerHmm };

        double usableLeft = SchMarginMm;
        double usableTop = SchMarginMm;
        double usableW = innerWmm - 2 * SchMarginMm;
        double usableH = innerHmm - 2 * SchMarginMm;

        double bottomReserve = CentralBlockHmm + LegendHmm + 1000;
        double floorAreaH = usableH - bottomReserve;
        int floorCount = model.Floors;
        double floorH = floorAreaH / floorCount;

        double unitsAreaLeft = usableLeft + FloorLabelWidthMm;
        double riserX = usableLeft + usableW - RiserWidthMm * 0.5;
        double fdbColumnLeft = riserX - FdbBoxWmm - 1500;
        double unitsAreaW = fdbColumnLeft - unitsAreaLeft - 800;

        layout.RiserXmm = riserX;
        layout.RiserTopYmm = usableTop;
        layout.RiserBottomYmm = usableTop + floorAreaH + CentralBlockHmm * 0.5;

        int maxUnits = model.FloorItems.Max(f => f.Units.Count);
        double unitW = maxUnits > 0
            ? Math.Max(UnitMinWmm, (unitsAreaW - UnitPaddingMm * Math.Max(0, maxUnits - 1)) / maxUnits)
            : UnitMinWmm;

        for (int i = 0; i < floorCount; i++)
        {
            var floor = model.FloorItems[i];
            double floorTop = usableTop + (floorCount - floor.Index) * floorH;

            var fl = new FloorLayout
            {
                FloorIndex = floor.Index,
                FloorName = floor.Name,
                Ymm = floorTop,
                Hmm = floorH,
                LabelXmm = usableLeft + 200,
                LabelYmm = floorTop + floorH * 0.5,
                RiserLabel = $"K-{floor.Index}",
                FdbRatio = floor.FdbRatio,
                FdbXmm = fdbColumnLeft,
                FdbYmm = floorTop + (floorH - FdbBoxHmm) * 0.5,
                FdbWmm = FdbBoxWmm,
                FdbHmm = FdbBoxHmm
            };

            for (int u = 0; u < floor.Units.Count; u++)
            {
                var unit = floor.Units[u];
                double unitX = unitsAreaLeft + u * (unitW + UnitPaddingMm);

                var portRows = new List<PortRowLayout>();
                int rowIndex = 0;
                foreach (var port in unit.Ports)
                {
                    for (int p = 0; p < port.Count; p++)
                    {
                        portRows.Add(new PortRowLayout
                        {
                            Label = port.Count > 1 ? $"{port.Type}-{p + 1}" : port.Type,
                            CableType = port.Cable,
                            Ymm = rowIndex * PortRowHmm
                        });
                        rowIndex++;
                    }
                }

                double portRowH = floorH > 0 ? Math.Min(PortRowHmm, (floorH * 0.6) / Math.Max(1, rowIndex)) : PortRowHmm;
                double unitContentH = Math.Max(rowIndex * portRowH + 400, 2000);
                double ontSpace = unit.HasONT ? OntBoxHmm + 200 : 0;
                double totalUnitH = unitContentH + ontSpace;

                if (totalUnitH > floorH * 0.9)
                {
                    unitContentH = floorH * 0.9 - ontSpace;
                    portRowH = rowIndex > 0 ? (unitContentH - 400) / rowIndex : PortRowHmm;
                    for (int r = 0; r < portRows.Count; r++)
                        portRows[r].Ymm = r * portRowH;
                }

                double unitTopY = floorTop + (floorH - unitContentH - ontSpace) * 0.5;
                double ontY = unitTopY;
                double boxY = unit.HasONT ? unitTopY + OntBoxHmm + 200 : unitTopY;

                var ul = new UnitLayout
                {
                    Name = unit.Name,
                    Xmm = unitX,
                    Ymm = boxY,
                    Wmm = unitW,
                    Hmm = unitContentH,
                    HasONT = unit.HasONT,
                    OntXmm = unitX + (unitW - OntBoxWmm) * 0.5,
                    OntYmm = ontY,
                    OntWmm = OntBoxWmm,
                    OntHmm = OntBoxHmm,
                    PortRows = portRows,
                    CableSummaryText = $"A-{unit.CableACount}, B-{unit.CableBCount}",
                    SummaryXmm = unitX,
                    SummaryYmm = boxY + unitContentH + 200,
                    ConnLineStartXmm = unitX + unitW,
                    ConnLineStartYmm = boxY + unitContentH * 0.5,
                    ConnLineEndXmm = fl.FdbXmm,
                    ConnLineEndYmm = fl.FdbYmm + FdbBoxHmm * 0.5,
                    ConnLabel = "D"
                };
                fl.Units.Add(ul);
            }
            layout.Floors.Add(fl);
        }

        double centralTop = usableTop + floorAreaH + 500;
        layout.Central = new CentralBlockLayout
        {
            Xmm = unitsAreaLeft,
            Ymm = centralTop,
            Wmm = fdbColumnLeft + FdbBoxWmm - unitsAreaLeft,
            Hmm = CentralBlockHmm,
            RackXmm = unitsAreaLeft + 400,
            RackYmm = centralTop + 400,
            RackWmm = 5000,
            RackHmm = CentralBlockHmm - 800,
            OdfXmm = unitsAreaLeft + 6000,
            OdfYmm = centralTop + 800,
            OdfWmm = 3000,
            OdfHmm = CentralBlockHmm - 1600,
            ArrowStartXmm = usableLeft - 500,
            ArrowStartYmm = centralTop + CentralBlockHmm * 0.5,
            ArrowEndXmm = unitsAreaLeft + 400,
            ArrowEndYmm = centralTop + CentralBlockHmm * 0.5
        };

        layout.LegendXmm = usableLeft;
        layout.LegendYmm = centralTop + CentralBlockHmm + 500;
        layout.LegendWmm = usableW * 0.5;
        layout.LegendHmm = LegendHmm;

        return layout;
    }

    private void DrawGeneratedSchematic(double innerX, double innerY, double innerW, double innerH, Brush pen)
    {
        if (_schematic == null || _schematic.FloorItems.Count == 0)
        {
            DrawSingleFloorGrid(innerX, innerY, innerW, innerH, pen);
            return;
        }

        var (innerWmm, innerHmm) = GetActiveInnerSizeMm();
        var layout = CalculateSchematicLayout(_schematic, innerWmm, innerHmm);

        var blueBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1264A3"));
        var lightGray = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#CCCCCC"));
        var lightBlueFill = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8F4FD"));
        var ontFill = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFF3E0"));

        DrawSchFloorBackgrounds(innerX, innerY, layout, lightGray);
        DrawSchRiser(innerX, innerY, layout);
        DrawSchFdbBoxes(innerX, innerY, layout);
        DrawSchUnitBlocks(innerX, innerY, layout, blueBrush, lightBlueFill, ontFill);
        DrawSchConnectionLines(innerX, innerY, layout);
        DrawSchCentralBlock(innerX, innerY, layout);
        DrawSchLegendTable(innerX, innerY, layout);
    }

    private void DrawSchFloorBackgrounds(double innerX, double innerY, SchematicLayout layout, Brush dashBrush)
    {
        double U(double mm) => mm * _zoom / 25.0;
        foreach (var fl in layout.Floors)
        {
            double y = innerY + U(fl.Ymm);
            double x1 = innerX + U(SchMarginMm);
            double x2 = innerX + U(layout.WorkAreaWmm - SchMarginMm);
            DrawDashedLine(x1, y, x2, y, dashBrush, 0.8);
            DrawTextRotated(fl.FloorName,
                innerX + U(fl.LabelXmm), innerY + U(fl.LabelYmm),
                9 * _zoom, Brushes.Black, -90);
        }
    }

    private void DrawSchRiser(double innerX, double innerY, SchematicLayout layout)
    {
        double U(double mm) => mm * _zoom / 25.0;
        double x = innerX + U(layout.RiserXmm);
        double y1 = innerY + U(layout.RiserTopYmm);
        double y2 = innerY + U(layout.RiserBottomYmm);
        DrawLine(x, y1, x, y2, Brushes.Black, 2.5);

        foreach (var fl in layout.Floors)
        {
            double labelY = innerY + U(fl.Ymm + fl.Hmm * 0.5);
            DrawText(fl.RiserLabel, x + U(200), labelY, 8 * _zoom, Brushes.Black);
        }
    }

    private void DrawSchFdbBoxes(double innerX, double innerY, SchematicLayout layout)
    {
        double U(double mm) => mm * _zoom / 25.0;
        foreach (var fl in layout.Floors)
        {
            double fx = innerX + U(fl.FdbXmm);
            double fy = innerY + U(fl.FdbYmm);
            double fw = U(fl.FdbWmm);
            double fh = U(fl.FdbHmm);

            DrawRect(fx, fy, fw, fh, Brushes.Black, 1.2);
            DrawTextCentered(fl.FdbRatio, fx, fy, fw, fh * 0.5, 8 * _zoom, Brushes.Black);
            DrawTextCentered("FDB", fx, fy + fh * 0.4, fw, fh * 0.5, 8 * _zoom, Brushes.Black);

            double riserX = innerX + U(layout.RiserXmm);
            DrawLine(fx + fw, fy + fh * 0.5, riserX, fy + fh * 0.5, Brushes.Black, 1.0);
        }
    }

    private void DrawSchUnitBlocks(double innerX, double innerY, SchematicLayout layout,
        Brush blueBrush, Brush lightBlueFill, Brush ontFill)
    {
        double U(double mm) => mm * _zoom / 25.0;
        foreach (var fl in layout.Floors)
        {
            foreach (var ul in fl.Units)
            {
                double ux = innerX + U(ul.Xmm);
                double uy = innerY + U(ul.Ymm);
                double uw = U(ul.Wmm);
                double uh = U(ul.Hmm);

                DrawFilledRect(ux, uy, uw, uh, lightBlueFill, Brushes.Black, 1.0);

                double portAreaTop = uy + U(200);
                foreach (var pr in ul.PortRows)
                {
                    double rowY = portAreaTop + U(pr.Ymm);
                    double rowH = U(PortRowHmm);
                    DrawText(pr.Label, ux + U(200), rowY + rowH * 0.7, 7.5 * _zoom, Brushes.Black);
                    DrawText(pr.CableType, ux + uw - U(600), rowY + rowH * 0.7, 7.5 * _zoom, blueBrush);
                    DrawLine(ux + U(100), rowY + rowH, ux + uw - U(100), rowY + rowH, Brushes.LightGray, 0.5);
                }

                DrawText(ul.CableSummaryText,
                    innerX + U(ul.SummaryXmm), innerY + U(ul.SummaryYmm),
                    7.5 * _zoom, blueBrush);

                DrawTextCentered($"{ul.Name} сууц", ux, uy + uh + U(100), uw, U(600), 9 * _zoom, Brushes.Black);

                if (ul.HasONT)
                {
                    double ox = innerX + U(ul.OntXmm);
                    double oy = innerY + U(ul.OntYmm);
                    double ow = U(ul.OntWmm);
                    double oh = U(ul.OntHmm);
                    DrawFilledRect(ox, oy, ow, oh, ontFill, Brushes.Black, 1.0);
                    DrawTextCentered("ONT", ox, oy, ow, oh, 8 * _zoom, Brushes.Black);
                    DrawLine(ox + ow * 0.5, oy + oh, ox + ow * 0.5, uy, Brushes.Black, 0.8);
                }
            }
        }
    }

    private void DrawSchConnectionLines(double innerX, double innerY, SchematicLayout layout)
    {
        double U(double mm) => mm * _zoom / 25.0;
        foreach (var fl in layout.Floors)
        {
            foreach (var ul in fl.Units)
            {
                double x1 = innerX + U(ul.ConnLineStartXmm);
                double y1 = innerY + U(ul.ConnLineStartYmm);
                double x2 = innerX + U(ul.ConnLineEndXmm);
                double y2 = innerY + U(ul.ConnLineEndYmm);
                DrawLine(x1, y1, x2, y2, Brushes.Black, 0.8);

                double midX = (x1 + x2) * 0.5;
                double midY = (y1 + y2) * 0.5;
                DrawText(ul.ConnLabel, midX + U(100), midY, 7.5 * _zoom, Brushes.Red);
            }
        }
    }

    private void DrawSchCentralBlock(double innerX, double innerY, SchematicLayout layout)
    {
        double U(double mm) => mm * _zoom / 25.0;
        var c = layout.Central;

        DrawDashedRect(innerX + U(c.RackXmm), innerY + U(c.RackYmm), U(c.RackWmm), U(c.RackHmm), Brushes.Black, 1.2);
        DrawTextCentered("№8 Холбооны\nөрөөнд Rack-д",
            innerX + U(c.RackXmm), innerY + U(c.RackYmm), U(c.RackWmm), U(c.RackHmm), 8 * _zoom, Brushes.Black);

        var odfFill = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8F4FD"));
        DrawFilledRect(innerX + U(c.OdfXmm), innerY + U(c.OdfYmm), U(c.OdfWmm), U(c.OdfHmm), odfFill, Brushes.Black, 1.2);
        DrawTextCentered("ODF", innerX + U(c.OdfXmm), innerY + U(c.OdfYmm), U(c.OdfWmm), U(c.OdfHmm), 10 * _zoom, Brushes.Black);

        double riserX = innerX + U(layout.RiserXmm);
        DrawLine(riserX, innerY + U(c.Ymm), riserX, innerY + U(layout.RiserBottomYmm), Brushes.Black, 2.5);

        DrawArrow(innerX + U(c.ArrowStartXmm), innerY + U(c.ArrowStartYmm),
            innerX + U(c.ArrowEndXmm), innerY + U(c.ArrowEndYmm),
            Brushes.Black, 1.5, U(400));
        DrawText("Гадна холбооны оролт",
            innerX + U(c.ArrowStartXmm), innerY + U(c.ArrowStartYmm) - U(600), 8 * _zoom, Brushes.Black);
    }

    private void DrawSchLegendTable(double innerX, double innerY, SchematicLayout layout)
    {
        double U(double mm) => mm * _zoom / 25.0;
        double lx = innerX + U(layout.LegendXmm);
        double ly = innerY + U(layout.LegendYmm);
        double lw = U(layout.LegendWmm);
        double lh = U(layout.LegendHmm);

        DrawRect(lx, ly, lw, lh, Brushes.Black, 1.0);

        double headerH = lh * 0.22;
        DrawLine(lx, ly + headerH, lx + lw, ly + headerH, Brushes.Black, 1.0);
        DrawTextCentered("Тэмдэглэгээ", lx, ly, lw, headerH, 9 * _zoom, Brushes.Black);

        double col1W = lw * 0.15;
        double col2W = lw * 0.45;
        DrawLine(lx + col1W, ly + headerH, lx + col1W, ly + lh, Brushes.Black, 0.5);
        DrawLine(lx + col1W + col2W, ly + headerH, lx + col1W + col2W, ly + lh, Brushes.Black, 0.5);

        var legendItems = new (string code, string mark, string desc)[]
        {
            ("A", "UTP Cat6 4x2x0.5мм", "LAN, IPTV, Камер"),
            ("B", "UTP Cat5 4x2x0.5мм", "Телефон"),
            ("D", "Flat Drop 1C G.657A2", "ДХС-аас айл руу"),
            ("K", "Indoor Riser Fiber", "Босоо сувагчлал")
        };

        double rowH = (lh - headerH) / legendItems.Length;
        for (int i = 0; i < legendItems.Length; i++)
        {
            double ry = ly + headerH + i * rowH;
            if (i > 0) DrawLine(lx, ry, lx + lw, ry, Brushes.Black, 0.5);
            DrawTextCentered(legendItems[i].code, lx, ry, col1W, rowH, 8 * _zoom, Brushes.Black);
            DrawText(legendItems[i].mark, lx + col1W + U(200), ry + rowH * 0.65, 7.5 * _zoom, Brushes.Black);
            DrawText(legendItems[i].desc, lx + col1W + col2W + U(200), ry + rowH * 0.65, 7.5 * _zoom, Brushes.Black);
        }
    }
    private (double wMm, double hMm) GetActiveInnerSizeMm()
    {
        return _isLandscape
            ? (_frame.InnerHmm, _frame.InnerWmm)
            : (_frame.InnerWmm, _frame.InnerHmm);
    }
    private (double x, double y, double w, double h) GetOuterRectPx()
    {
        double U(double mm) => mm * _zoom / 25.0;
        var (innerWmm, innerHmm) = GetActiveInnerSizeMm();
        double outerW = U(innerWmm + _frame.LeftOff + _frame.RightOff);
        double outerH = U(innerHmm + _frame.TopOff + _frame.BottomOff);
        double gap = U(PageGapMm);
        double totalW = outerW * _pageCount + gap * Math.Max(0, _pageCount - 1);
        return (OuterFrameOriginX, OuterFrameOriginY, totalW, outerH);
    }
    private (double x, double y, double w, double h) GetInnerRectPx()
    {
        double U(double mm) => mm * _zoom / 25.0;
        var (innerWmm, innerHmm) = GetActiveInnerSizeMm();
        double outerX = OuterFrameOriginX;
        double outerY = OuterFrameOriginY;
        double innerX = outerX + U(_frame.LeftOff);
        double innerY = outerY + U(_frame.TopOff);
        return (innerX, innerY, U(innerWmm), U(innerHmm));
    }

    private (double x, double y, double w, double h) GetInnerRectPxForZoom(double zoom)
    {
        double U(double mm) => mm * zoom / 25.0;
        var (innerWmm, innerHmm) = GetActiveInnerSizeMm();
        double outerX = OuterFrameOriginX;
        double outerY = OuterFrameOriginY;
        double innerX = outerX + U(_frame.LeftOff);
        double innerY = outerY + U(_frame.TopOff);
        return (innerX, innerY, U(innerWmm), U(innerHmm));
    }

    private bool CanPlaceInsideInner(double x, double y, double w, double h)
    {
        var (ix, iy, iw, ih) = GetInnerRectPx();
        return x >= ix && y >= iy && (x + w) <= (ix + iw) && (y + h) <= (iy + ih);
    }

    private void ConstrainNodeToInnerFrame(NodeModel n)
    {
        var (ix, iy, iw, ih) = GetInnerRectPx();
        if (n.W > iw) n.W = Math.Max(36, iw);
        if (n.H > ih) n.H = Math.Max(18, ih);

        n.X = Math.Clamp(n.X, ix, ix + iw - n.W);
        n.Y = Math.Clamp(n.Y, iy, iy + ih - n.H);
    }

    private static int ParseOrDefault(string? text, int fallback, int min, int max)
    {
        if (!int.TryParse(text, out var v)) return fallback;
        return Math.Clamp(v, min, max);
    }

    private void SaveStateForUndo()
    {
        _undo.Push(SerializeState());
        if (_undo.Count > 80)
        {
            _undo.Clear();
            _undo.Push(SerializeState());
        }
        _redo.Clear();
    }

    private string SerializeState()
    {
        var dto = new StateDto
        {
            Nodes = _nodes.Select(n => new NodeDto { Id = n.Id, Type = n.Type, Label = n.Label, X = n.X, Y = n.Y, W = n.W, H = n.H }).ToList(),
            Edges = _edges.Select(e => new EdgeDto { FromId = e.FromId, ToId = e.ToId, Label = e.Label }).ToList(),
            SelectedId = _selectedId
        };
        return System.Text.Json.JsonSerializer.Serialize(dto);
    }

    private void DeserializeState(string json)
    {
        var dto = System.Text.Json.JsonSerializer.Deserialize<StateDto>(json);
        if (dto == null) return;

        _nodes.Clear();
        _edges.Clear();
        _nodes.AddRange(dto.Nodes.Select(n => new NodeModel { Id = n.Id, Type = n.Type, Label = n.Label, X = n.X, Y = n.Y, W = n.W, H = n.H }));
        _edges.AddRange(dto.Edges.Select(e => new EdgeModel { FromId = e.FromId, ToId = e.ToId, Label = e.Label }));
        _selectedId = dto.SelectedId;
        SyncSelectionUi();
    }

    private sealed class StateDto
    {
        public List<NodeDto> Nodes { get; set; } = new();
        public List<EdgeDto> Edges { get; set; } = new();
        public string? SelectedId { get; set; }
    }

    private sealed class NodeDto
    {
        public string Id { get; set; } = "";
        public string Type { get; set; } = "";
        public string Label { get; set; } = "";
        public double X { get; set; }
        public double Y { get; set; }
        public double W { get; set; }
        public double H { get; set; }
    }

    private sealed class EdgeDto
    {
        public string FromId { get; set; } = "";
        public string ToId { get; set; } = "";
        public string Label { get; set; } = "";
    }

    private void DrawLine(double x1, double y1, double x2, double y2, Brush stroke, double thickness)
    {
        PreviewCanvas.Children.Add(new Line { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Stroke = stroke, StrokeThickness = thickness });
    }

    private void DrawRect(double x, double y, double w, double h, Brush stroke, double thickness)
    {
        var rect = new Rectangle { Width = w, Height = h, Stroke = stroke, StrokeThickness = thickness };
        PreviewCanvas.Children.Add(rect);
        Canvas.SetLeft(rect, x);
        Canvas.SetTop(rect, y);
    }

    private void DrawText(string text, double x, double y, double size, Brush color)
    {
        var tb = new TextBlock { Text = text, Foreground = color, FontSize = Math.Max(8, size), FontFamily = new FontFamily("Segoe UI") };
        PreviewCanvas.Children.Add(tb);
        Canvas.SetLeft(tb, x);
        Canvas.SetTop(tb, y - tb.FontSize);
    }

    private void DrawTextCentered(string text, double x, double y, double w, double h, double size, Brush color)
    {
        var tb = new TextBlock
        {
            Text = text,
            Foreground = color,
            FontSize = Math.Max(8, size),
            FontFamily = new FontFamily("Segoe UI"),
            TextAlignment = TextAlignment.Center,
            Width = w
        };
        PreviewCanvas.Children.Add(tb);
        Canvas.SetLeft(tb, x);
        Canvas.SetTop(tb, y + (h - tb.FontSize) * 0.5 - 1);
    }

    private void DrawDashedLine(double x1, double y1, double x2, double y2, Brush stroke, double thickness, double dash = 6, double gap = 4)
    {
        var line = new Line
        {
            X1 = x1, Y1 = y1, X2 = x2, Y2 = y2,
            Stroke = stroke, StrokeThickness = thickness,
            StrokeDashArray = new DoubleCollection { dash, gap }
        };
        PreviewCanvas.Children.Add(line);
    }

    private void DrawFilledRect(double x, double y, double w, double h, Brush fill, Brush stroke, double thickness)
    {
        var rect = new Rectangle { Width = w, Height = h, Fill = fill, Stroke = stroke, StrokeThickness = thickness };
        PreviewCanvas.Children.Add(rect);
        Canvas.SetLeft(rect, x);
        Canvas.SetTop(rect, y);
    }

    private void DrawDashedRect(double x, double y, double w, double h, Brush stroke, double thickness, double dash = 6, double gap = 4)
    {
        var rect = new Rectangle
        {
            Width = w, Height = h,
            Stroke = stroke, StrokeThickness = thickness,
            StrokeDashArray = new DoubleCollection { dash, gap }
        };
        PreviewCanvas.Children.Add(rect);
        Canvas.SetLeft(rect, x);
        Canvas.SetTop(rect, y);
    }

    private void DrawArrow(double x1, double y1, double x2, double y2, Brush stroke, double thickness, double headSize = 8)
    {
        DrawLine(x1, y1, x2, y2, stroke, thickness);
        double angle = Math.Atan2(y2 - y1, x2 - x1);
        double a1 = angle + Math.PI * 0.82;
        double a2 = angle - Math.PI * 0.82;
        DrawLine(x2, y2, x2 + headSize * Math.Cos(a1), y2 + headSize * Math.Sin(a1), stroke, thickness);
        DrawLine(x2, y2, x2 + headSize * Math.Cos(a2), y2 + headSize * Math.Sin(a2), stroke, thickness);
    }

    private void DrawTextRotated(string text, double x, double y, double size, Brush color, double angleDeg)
    {
        var tb = new TextBlock
        {
            Text = text, Foreground = color,
            FontSize = Math.Max(8, size),
            FontFamily = new FontFamily("Segoe UI"),
            FontWeight = FontWeights.Bold,
            RenderTransformOrigin = new Point(0.5, 0.5),
            RenderTransform = new RotateTransform(angleDeg)
        };
        PreviewCanvas.Children.Add(tb);
        Canvas.SetLeft(tb, x);
        Canvas.SetTop(tb, y);
    }
}




























