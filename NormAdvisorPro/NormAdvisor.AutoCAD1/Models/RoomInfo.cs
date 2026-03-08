using Autodesk.AutoCAD.DatabaseServices;
using NormAdvisor.AutoCAD1.ViewModels;

namespace NormAdvisor.AutoCAD1.Models
{
    /// <summary>
    /// AutoCAD зургаас таньсан өрөөний мэдээлэл
    /// Polyline хүрээтэй холбогдож болно (BoundaryId)
    /// </summary>
    public class RoomInfo : BaseViewModel
    {
        private ObjectId _boundaryId = ObjectId.Null;
        private double _drawnArea;

        public int Number { get; set; }
        public string Name { get; set; } = string.Empty;
        public double Area { get; set; }
        public string RawText { get; set; } = string.Empty;

        /// <summary>
        /// Хүснэгтээс уншсан анхны дугаар текст (жишээ: "a1", "b2", "c3")
        /// AutoMatch-д ашиглана
        /// </summary>
        public string RoomId { get; set; } = string.Empty;

        /// <summary>
        /// Секц/бүлгийн нэр (жишээ: "А-Кофешоп", "Б-Хүнсний дэлгүүр")
        /// Хүснэгтийн merge хийсэн гарчиг мөрөөс авна
        /// </summary>
        public string SectionName { get; set; } = string.Empty;

        /// <summary>
        /// Холбогдсон polyline-ийн ObjectId
        /// </summary>
        public ObjectId BoundaryId
        {
            get => _boundaryId;
            set
            {
                if (_boundaryId != value)
                {
                    _boundaryId = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(HasBoundary));
                    OnPropertyChanged(nameof(AreaDifference));
                }
            }
        }

        /// <summary>
        /// Polyline-ийн бодит талбай (м²)
        /// </summary>
        public double DrawnArea
        {
            get => _drawnArea;
            set
            {
                if (SetProperty(ref _drawnArea, value))
                {
                    OnPropertyChanged(nameof(AreaDifference));
                }
            }
        }

        public bool HasBoundary => _boundaryId != ObjectId.Null;

        /// <summary>
        /// Талбайн зөрүү (зурсан - хүснэгт)
        /// </summary>
        public double AreaDifference => HasBoundary && Area > 0
            ? DrawnArea - Area
            : 0;

        public override string ToString()
        {
            if (Area > 0)
                return $"{Number}. {Name} — {Area:F2} м²";
            return $"{Number}. {Name}";
        }
    }
}
