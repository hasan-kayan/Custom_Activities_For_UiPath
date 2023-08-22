using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using LinkteraRobotics.ExcelActivities.OpenWorkbook.Activities.Design.Designers;
using LinkteraRobotics.ExcelActivities.OpenWorkbook.Activities.Design.Properties;

namespace LinkteraRobotics.ExcelActivities.OpenWorkbook.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(OpenWorkbook), categoryAttribute);
            builder.AddCustomAttributes(typeof(OpenWorkbook), new DesignerAttribute(typeof(OpenWorkbookDesigner)));
            builder.AddCustomAttributes(typeof(OpenWorkbook), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
