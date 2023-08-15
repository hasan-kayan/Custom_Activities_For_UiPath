using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using LinkteraRobotics.Copy.Sheet.Activities.Design.Designers;
using LinkteraRobotics.Copy.Sheet.Activities.Design.Properties;

namespace LinkteraRobotics.Copy.Sheet.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(CopyExcelSheetValue), categoryAttribute);
            builder.AddCustomAttributes(typeof(CopyExcelSheetValue), new DesignerAttribute(typeof(CopyExcelSheetValueDesigner)));
            builder.AddCustomAttributes(typeof(CopyExcelSheetValue), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
