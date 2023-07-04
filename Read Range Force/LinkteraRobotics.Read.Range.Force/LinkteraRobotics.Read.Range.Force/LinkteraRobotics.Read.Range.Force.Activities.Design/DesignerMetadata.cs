using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using LinkteraRobotics.Read.Range.Force.Activities.Design.Designers;
using LinkteraRobotics.Read.Range.Force.Activities.Design.Properties;

namespace LinkteraRobotics.Read.Range.Force.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(ReadRangeForce), categoryAttribute);
            builder.AddCustomAttributes(typeof(ReadRangeForce), new DesignerAttribute(typeof(ReadRangeForceDesigner)));
            builder.AddCustomAttributes(typeof(ReadRangeForce), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
