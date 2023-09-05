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

            builder.AddCustomAttributes(typeof(OpenWorkbook), categoryAttribute);
            builder.AddCustomAttributes(typeof(OpenWorkbook), new DesignerAttribute(typeof(OpenWorkbookDesigner)));
            builder.AddCustomAttributes(typeof(OpenWorkbook), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(CopySheet), categoryAttribute);
            builder.AddCustomAttributes(typeof(CopySheet), new DesignerAttribute(typeof(CopySheetDesigner)));
            builder.AddCustomAttributes(typeof(CopySheet), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(ReadBigData), categoryAttribute);
            builder.AddCustomAttributes(typeof(ReadBigData), new DesignerAttribute(typeof(ReadBigDataDesigner)));
            builder.AddCustomAttributes(typeof(ReadBigData), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
