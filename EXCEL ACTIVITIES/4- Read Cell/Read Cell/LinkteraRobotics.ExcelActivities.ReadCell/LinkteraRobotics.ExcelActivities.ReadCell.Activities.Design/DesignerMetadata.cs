using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using LinkteraRobotics.ExcelActivities.ReadCell.Activities.Design.Designers;
using LinkteraRobotics.ExcelActivities.ReadCell.Activities.Design.Properties;

namespace LinkteraRobotics.ExcelActivities.ReadCell.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(ReadCell), categoryAttribute);
            builder.AddCustomAttributes(typeof(ReadCell), new DesignerAttribute(typeof(ReadCellDesigner)));
            builder.AddCustomAttributes(typeof(ReadCell), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
