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

            builder.AddCustomAttributes(typeof(ReadSpecialRange), categoryAttribute);
            builder.AddCustomAttributes(typeof(ReadSpecialRange), new DesignerAttribute(typeof(ReadSpecialRangeDesigner)));
            builder.AddCustomAttributes(typeof(ReadSpecialRange), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
