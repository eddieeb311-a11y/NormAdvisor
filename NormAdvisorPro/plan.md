# NormAdvisorPro — Нийт Төлөвлөгөө

## Хийгдсэн (Done)
- [x] Өрөө унших (Table / Region)
- [x] Boundary зурах / сонгох / AutoMatch
- [x] Zoom to room
- [x] Төхөөрөмж каталог + байршуулах
- [x] MasterCopy (энгийн lock/unlock)
- [x] WPF Palette UI (2 tab)

---

## Фаз 1: Bug fix
- [ ] Area conversion (мм² → м²) AutoMatch-д
- [ ] Status bar encoding засах

## Фаз 2: Smart MasterCopy
- [ ] Ижил BTR блокуудыг автомат илрүүлэх
- [ ] Device блокоор системийн төрөл таних
- [ ] Блокын нэрнээс давхрын нэр таних
- [ ] Boundary автомат хуулах (мастер → бусад instance)
- [ ] NORMMASTERLOCK command
- [ ] UI: систем нэр, тоо, edit боломж

## Фаз 3: Төхөөрөмжийн тооцоо
- [ ] Өрөө бүрд ямар device хэд байгааг автомат тоолох
- [ ] Давхрын нэгтгэл хүснэгт (давхар бүрийн нийт тоо)
- [ ] Бүх барилгын нэгдсэн тооцоо
- [ ] AutoCAD Table оруулах (specification хүснэгт)

## Фаз 4: Авто дугаарлалт (Auto-Tag)
- [ ] Өрөө дотор device блокуудыг олох
- [ ] TagPrefix-ээр дугаарлах (TD-1, TD-2, SZ-1, K-1...)
- [ ] Давхар бүрд дахин эхлэх / үргэлжлүүлэх сонголт
- [ ] Дугаар AttributeReference-д бичих

## Фаз 5: Норм шалгалт
- [ ] БНбД стандарт дүрмүүд тодорхойлох (JSON/config)
- [ ] Өрөө бүрд норм шалгах (талбай → шаардлагатай тоо vs одоогийн тоо)
- [ ] Дутуу/хангасан/илүүдэл үзүүлэлт (улаан/ногоон)
- [ ] Норм tab UI (3-р tab идэвхжүүлэх)
- [ ] Нэгдсэн тайлан

## Фаз 6: Угсралтын бүдүүвч (Riser Diagram)
Dynamic Block template + plugin автомат бөглөх арга:

### Template блокууд (DWG файлууд):
- `RiserFloorBox.dwg` — давхрын хайрцаг
  - Attributes: FLOOR_NAME, TD_COUNT, SZ_COUNT, SI_COUNT, MCP_COUNT, TOTAL
  - Dynamic: stretch action (төхөөрөмж олон бол өргөсөх)
- `RiserMainPanel.dwg` — төв тоноглол хайрцаг
  - Attributes: PANEL_NAME, PANEL_TYPE (FACP/NVR/SW гм)
  - Visibility states: system төрлөөр
- `RiserLine.dwg` — холбох шугам (эсвэл шууд Line зурах)

### Plugin хийх зүйл:
- [ ] Давхар × систем × төхөөрөмж тоо цуглуулах
- [ ] Template блокуудыг зөв байрлалд оруулах (доороос дээш давхар бүр)
- [ ] Attribute-уудыг автомат бөглөх
- [ ] Төв тоноглол ↔ давхар хооронд шугам зурах
- [ ] NORMRISER command

### Бүдүүвчийн жишээ бүтэц:
```
            ┌──────────┐
            │   FACP   │
            └────┬─────┘
                 │
   ┌─────────────┼─────────────┐
   │             │             │
┌──┴───┐    ┌───┴───┐    ┌───┴───┐
│Зоорь │    │1-р дав│    │2-р дав│
│TD × 6│    │TD × 10│    │TD × 10│
│SZ × 3│    │SZ × 5 │    │SZ × 5 │
│SI × 1│    │SI × 2 │    │SI × 2 │
│MCP× 1│    │MCP× 2 │    │MCP× 2 │
│Нийт:11│   │Нийт:19│    │Нийт:19│
└──────┘    └───────┘    └───────┘
```

### 5 системийн бүдүүвч template:

| # | Систем | Төв тоноглол | Topology | Template DWG |
|---|--------|-------------|----------|-------------|
| 1 | Галын дохиолол | FACP, NAC, Repeater | Loop/Zone | RiserFire.dwg |
| 2 | CCTV камер | NVR, PoE Switch | Star | RiserCCTV.dwg |
| 3 | Tel/LAN/IPTV | Switch, PABX, OLT | Star/Tree | RiserLAN.dwg |
| 4 | PA зарлан мэдээлэх | Amplifier | Bus | RiserPA.dwg |
| 5 | УБ Мэдээлэл Холбоо | Rack, Router, FDF | Tree | RiserUB.dwg |

Template бүрд:
- Төв тоноглолын Dynamic Block (visibility state-ээр төрөл солих)
- Давхрын хайрцаг Dynamic Block (stretch + attributes)
- Холбох шугам (Line + арга)

## Фаз 7: Тайлан / Экспорт
- [ ] Excel экспорт (төхөөрөмж, норм, тооцоо)
- [ ] Legend хүснэгт автомат үүсгэх
- [ ] PDF тайлан
- [ ] Кабелийн урт тооцоо
