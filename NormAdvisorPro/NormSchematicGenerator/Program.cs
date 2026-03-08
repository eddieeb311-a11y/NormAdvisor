using System.Text;
using System.Globalization;

var outputPath = args.Length > 0
    ? args[0]
    : Path.Combine(Environment.CurrentDirectory, "schematic_template.dxf");

int floors = args.Length > 1 && int.TryParse(args[1], out var f) ? Math.Max(1, f) : 5;
int lanPerFloor = args.Length > 2 && int.TryParse(args[2], out var l) ? Math.Max(1, l) : 10;

var dxf = new SimpleDxfWriter();

const double workW = 26500;
const double workH = 39000;
const double margin = 500;

double outerL = 0;
double outerB = 0;
double outerR = outerL + workW + margin * 2;
double outerT = outerB + workH + margin * 2;

double innerL = outerL + margin;
double innerB = outerB + margin;
double innerR = outerR - margin;
double innerT = outerT - margin;

dxf.AddRect(outerL, outerB, outerR, outerT, "FRAME");
dxf.AddRect(innerL, innerB, innerR, innerT, "FRAME");

dxf.AddText(innerL + 900, innerT - 1200, 250, "GURVALSAN UILCHILGEENII UGSRALTYN BUDUUVCH", "TEXT");

const double panelGap = 1200;
const double panelW = 12000;

int leftCount = (int)Math.Ceiling(floors / 2.0);
int rightCount = floors - leftCount;

double panelTop = innerT - 2500;

double leftX = innerL + 800;
DrawPanel(dxf, leftX, panelTop, panelW, leftCount, 1, lanPerFloor, "K-", "PANEL");

if (rightCount > 0)
{
    double rightX = leftX + panelW + panelGap;
    DrawPanel(dxf, rightX, panelTop, panelW, rightCount, leftCount + 1, lanPerFloor, "K-", "PANEL");
}

var basementYTop = innerB + 8500;
var basementYBottom = innerB + 1800;
var basementXLeft = leftX;
var basementXRight = leftX + panelW * 0.62;
dxf.AddRect(basementXLeft, basementYBottom, basementXRight, basementYTop, "PANEL");
dxf.AddText(basementXLeft + 250, basementYTop - 500, 200, "ZOORIIN DAVHAR", "TEXT");

dxf.Save(outputPath);
Console.WriteLine($"Generated: {outputPath}");
Console.WriteLine($"Floors: {floors}, LAN/floor: {lanPerFloor}");

static void DrawPanel(SimpleDxfWriter dxf, double xLeft, double yTop, double width, int floorCount, int floorStart, int lanPerFloor, string riserPrefix, string layer)
{
    if (floorCount <= 0) return;

    const double rowH = 5000;
    const double headerH1 = 900;
    const double headerH2 = 900;

    double totalH = headerH1 + headerH2 + floorCount * rowH;
    double yBottom = yTop - totalH;
    double xRight = xLeft + width;

    double xNo = xLeft + 900;
    double xRoom = xLeft + 5200;
    double xBoard = xLeft + 8000;
    double xRiser = xLeft + 10400;

    dxf.AddRect(xLeft, yBottom, xRight, yTop, layer);
    dxf.AddLine(xLeft, yTop - headerH1, xRight, yTop - headerH1, layer);
    dxf.AddLine(xLeft, yTop - headerH1 - headerH2, xRight, yTop - headerH1 - headerH2, layer);

    dxf.AddLine(xNo, yTop - headerH1, xNo, yBottom, layer);
    dxf.AddLine(xRoom, yTop - headerH1, xRoom, yBottom, layer);
    dxf.AddLine(xBoard, yTop - headerH1, xBoard, yBottom, layer);
    dxf.AddLine(xRiser, yTop - headerH1, xRiser, yBottom, layer);

    dxf.AddText(xLeft + width * 0.38, yTop - 600, 180, "GURVALSAN UILCHILGEE", "TEXT");
    dxf.AddText(xLeft + 250, yTop - 1450, 180, "NO", "TEXT");
    dxf.AddText(xNo + 900, yTop - 1450, 180, "OROO", "TEXT");
    dxf.AddText(xRoom + 500, yTop - 1450, 180, "AILYN SAMBAR", "TEXT");
    dxf.AddText(xBoard + 350, yTop - 1450, 180, "BOSOO SUVAGCHLAL", "TEXT");

    for (int i = 0; i < floorCount; i++)
    {
        int floor = floorStart + i;
        double rowBottom = yTop - headerH1 - headerH2 - (i + 1) * rowH;
        double rowTop = rowBottom + rowH;
        double yMid = (rowTop + rowBottom) * 0.5;

        dxf.AddLine(xLeft, rowBottom, xRight, rowBottom, layer);
        dxf.AddText(xLeft + 120, yMid, 170, $"{floor}-r", "TEXT");

        double roomY1 = rowTop - 800;
        double roomY2 = rowTop - 1700;
        double roomY3 = rowTop - 2600;

        DrawRoomRun(dxf, xNo + 250, xRoom - 350, roomY1, lanPerFloor);
        DrawRoomRun(dxf, xNo + 250, xRoom - 350, roomY2, lanPerFloor);
        DrawRoomRun(dxf, xNo + 250, xRoom - 350, roomY3, lanPerFloor);

        double boardX = xRoom + 300;
        double riserX = xRiser + 300;
        DrawOntRun(dxf, boardX, roomY1, riserX, $"A-{floor},1");
        DrawOntRun(dxf, boardX, roomY2, riserX, $"A-{floor},2");
        DrawOntRun(dxf, boardX, roomY3, riserX, $"A-{floor},3");

        dxf.AddText(xRiser + 100, yMid, 160, $"{riserPrefix}{floor}", "TEXT");
    }
}

static void DrawRoomRun(SimpleDxfWriter dxf, double xL, double xR, double y, int lanPerFloor)
{
    dxf.AddRect(xL, y - 180, xL + 400, y + 180, "DEVICE");
    dxf.AddLine(xL + 450, y, xR - 120, y, "DEVICE");
    dxf.AddCircle(xR - 60, y, 60, "DEVICE");
    dxf.AddText(xL + 500, y + 180, 120, $"LAN-{Math.Max(1, lanPerFloor / 3)}", "TEXT");
}

static void DrawOntRun(SimpleDxfWriter dxf, double x, double y, double riserX, string label)
{
    dxf.AddRect(x, y - 220, x + 700, y + 220, "DEVICE");
    dxf.AddText(x + 220, y + 20, 130, "ONT", "TEXT");
    dxf.AddLine(x + 700, y, riserX, y, "DEVICE");
    dxf.AddText(x - 260, y + 140, 120, label, "TEXT");
}

internal sealed class SimpleDxfWriter
{
    private readonly StringBuilder _entities = new();
    private readonly HashSet<string> _layers = new(StringComparer.OrdinalIgnoreCase) { "0" };

    public void AddLine(double x1, double y1, double x2, double y2, string layer = "0")
    {
        TouchLayer(layer);
        _entities.AppendLine("0");
        _entities.AppendLine("LINE");
        _entities.AppendLine("8");
        _entities.AppendLine(SafeAscii(layer));
        _entities.AppendLine("10");
        _entities.AppendLine(F(x1));
        _entities.AppendLine("20");
        _entities.AppendLine(F(y1));
        _entities.AppendLine("11");
        _entities.AppendLine(F(x2));
        _entities.AppendLine("21");
        _entities.AppendLine(F(y2));
    }

    public void AddRect(double x1, double y1, double x2, double y2, string layer = "0")
    {
        AddLine(x1, y1, x2, y1, layer);
        AddLine(x2, y1, x2, y2, layer);
        AddLine(x2, y2, x1, y2, layer);
        AddLine(x1, y2, x1, y1, layer);
    }

    public void AddCircle(double x, double y, double r, string layer = "0")
    {
        TouchLayer(layer);
        _entities.AppendLine("0");
        _entities.AppendLine("CIRCLE");
        _entities.AppendLine("8");
        _entities.AppendLine(SafeAscii(layer));
        _entities.AppendLine("10");
        _entities.AppendLine(F(x));
        _entities.AppendLine("20");
        _entities.AppendLine(F(y));
        _entities.AppendLine("40");
        _entities.AppendLine(F(r));
    }

    public void AddText(double x, double y, double height, string text, string layer = "0")
    {
        TouchLayer(layer);
        _entities.AppendLine("0");
        _entities.AppendLine("TEXT");
        _entities.AppendLine("8");
        _entities.AppendLine(SafeAscii(layer));
        _entities.AppendLine("10");
        _entities.AppendLine(F(x));
        _entities.AppendLine("20");
        _entities.AppendLine(F(y));
        _entities.AppendLine("40");
        _entities.AppendLine(F(height));
        _entities.AppendLine("1");
        _entities.AppendLine(SafeAscii((text ?? string.Empty).Replace("\r", " ").Replace("\n", " ")));
    }

    public void Save(string path)
    {
        var sb = new StringBuilder();

        sb.AppendLine("0"); sb.AppendLine("SECTION");
        sb.AppendLine("2"); sb.AppendLine("HEADER");
        sb.AppendLine("9"); sb.AppendLine("$ACADVER");
        sb.AppendLine("1"); sb.AppendLine("AC1009");
        sb.AppendLine("0"); sb.AppendLine("ENDSEC");

        sb.AppendLine("0"); sb.AppendLine("SECTION");
        sb.AppendLine("2"); sb.AppendLine("TABLES");
        sb.AppendLine("0"); sb.AppendLine("TABLE");
        sb.AppendLine("2"); sb.AppendLine("LAYER");
        sb.AppendLine("70"); sb.AppendLine(_layers.Count.ToString(CultureInfo.InvariantCulture));

        foreach (var layer in _layers)
        {
            sb.AppendLine("0"); sb.AppendLine("LAYER");
            sb.AppendLine("2"); sb.AppendLine(SafeAscii(layer));
            sb.AppendLine("70"); sb.AppendLine("0");
            sb.AppendLine("62"); sb.AppendLine("7");
            sb.AppendLine("6"); sb.AppendLine("CONTINUOUS");
        }

        sb.AppendLine("0"); sb.AppendLine("ENDTAB");
        sb.AppendLine("0"); sb.AppendLine("ENDSEC");

        sb.AppendLine("0"); sb.AppendLine("SECTION");
        sb.AppendLine("2"); sb.AppendLine("ENTITIES");
        sb.Append(_entities.ToString());
        sb.AppendLine("0"); sb.AppendLine("ENDSEC");
        sb.AppendLine("0"); sb.AppendLine("EOF");

        Directory.CreateDirectory(Path.GetDirectoryName(path) ?? Environment.CurrentDirectory);
        File.WriteAllText(path, sb.ToString(), Encoding.ASCII);
    }

    private void TouchLayer(string layer)
    {
        if (string.IsNullOrWhiteSpace(layer)) layer = "0";
        _layers.Add(layer);
    }

    private static string SafeAscii(string s)
    {
        if (string.IsNullOrEmpty(s)) return string.Empty;
        var chars = s.Select(ch => ch <= 127 ? ch : '_').ToArray();
        return new string(chars);
    }

    private static string F(double v) => v.ToString("0.###", CultureInfo.InvariantCulture);
}

