using System.Drawing;

namespace Prospecta.ConnektHub.Models
{
    public class FieldMetaDataModel
    {
        public string FieldName { get; set; }
        public string Description { get; set; }
        public string TabName { get; set; }
        public string HelpText { get; set; }
        public string Mandatory { get; set; }
        public bool IsDropDown { get; set; }
        public string DropDownType { get; set; }
        public Color Colour { get; set; }
    }
}
