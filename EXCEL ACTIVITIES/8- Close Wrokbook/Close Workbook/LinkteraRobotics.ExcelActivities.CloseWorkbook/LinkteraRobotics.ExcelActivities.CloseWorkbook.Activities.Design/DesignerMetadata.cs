using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using LinkteraRobotics.ExcelActivities.CloseWorkbook.Activities.Design.Designers;
using LinkteraRobotics.ExcelActivities.CloseWorkbook.Activities.Design.Properties;

namespace LinkteraRobotics.ExcelActivities.CloseWorkbook.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(CloseWorkbook), categoryAttribute);
            builder.AddCustomAttributes(typeof(CloseWorkbook), new DesignerAttribute(typeof(CloseWorkbookDesigner)));
            builder.AddCustomAttributes(typeof(CloseWorkbook), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
