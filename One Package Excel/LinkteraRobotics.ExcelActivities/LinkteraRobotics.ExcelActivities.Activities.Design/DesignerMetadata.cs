using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using LinkteraRobotics.ExcelActivities.Activities.Design.Designers;
using LinkteraRobotics.ExcelActivities.Activities.Design.Properties;

namespace LinkteraRobotics.ExcelActivities.Activities.Design
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

            builder.AddCustomAttributes(typeof(WriteCell), categoryAttribute);
            builder.AddCustomAttributes(typeof(WriteCell), new DesignerAttribute(typeof(WriteCellDesigner)));
            builder.AddCustomAttributes(typeof(WriteCell), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
