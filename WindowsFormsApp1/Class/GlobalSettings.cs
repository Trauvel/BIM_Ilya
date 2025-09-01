using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;

namespace WindowsFormsApp1.Class
{
    public class GlobalSettings
    {
        public static ExternalCommandData CommandData { get; set; }
        public static ElementId CurrentElementId { get; set; }
        public static ICollection<ElementId> CurrentSelection { get; set; }
        public static bool IsDoubleClick { get; set; }
    }
}
