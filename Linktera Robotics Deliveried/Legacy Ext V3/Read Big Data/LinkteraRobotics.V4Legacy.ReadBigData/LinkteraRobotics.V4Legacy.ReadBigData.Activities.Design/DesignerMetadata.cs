using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using LinkteraRobotics.V4Legacy.ReadBigData.Activities.Design.Designers;
using LinkteraRobotics.V4Legacy.ReadBigData.Activities.Design.Properties;

namespace LinkteraRobotics.V4Legacy.ReadBigData.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(ReadBigData), categoryAttribute);
            builder.AddCustomAttributes(typeof(ReadBigData), new DesignerAttribute(typeof(ReadBigDataDesigner)));
            builder.AddCustomAttributes(typeof(ReadBigData), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
