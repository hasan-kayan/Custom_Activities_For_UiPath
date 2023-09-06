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

            builder.AddCustomAttributes(typeof(CloseWorkbook), categoryAttribute);
            builder.AddCustomAttributes(typeof(CloseWorkbook), new DesignerAttribute(typeof(CloseWorkbookDesigner)));
            builder.AddCustomAttributes(typeof(CloseWorkbook), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(ReadBigData), categoryAttribute);
            builder.AddCustomAttributes(typeof(ReadBigData), new DesignerAttribute(typeof(ReadBigDataDesigner)));
            builder.AddCustomAttributes(typeof(ReadBigData), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(CopySheet), categoryAttribute);
            builder.AddCustomAttributes(typeof(CopySheet), new DesignerAttribute(typeof(CopySheetDesigner)));
            builder.AddCustomAttributes(typeof(CopySheet), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
