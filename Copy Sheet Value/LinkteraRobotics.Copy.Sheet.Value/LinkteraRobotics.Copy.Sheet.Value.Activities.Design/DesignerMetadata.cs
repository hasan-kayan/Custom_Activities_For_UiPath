using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using LinkteraRobotics.Copy.Sheet.Value.Activities.Design.Designers;
using LinkteraRobotics.Copy.Sheet.Value.Activities.Design.Properties;

namespace LinkteraRobotics.Copy.Sheet.Value.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(CopySheetValue), categoryAttribute);
            builder.AddCustomAttributes(typeof(CopySheetValue), new DesignerAttribute(typeof(CopySheetValueDesigner)));
            builder.AddCustomAttributes(typeof(CopySheetValue), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
