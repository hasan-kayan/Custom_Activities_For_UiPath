using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using LinkteraRobotics.Read.Excel.Range.Activities.Design.Designers;
using LinkteraRobotics.Read.Excel.Range.Activities.Design.Properties;

namespace LinkteraRobotics.Read.Excel.Range.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(Read), categoryAttribute);
            builder.AddCustomAttributes(typeof(Read), new DesignerAttribute(typeof(ReadDesigner)));
            builder.AddCustomAttributes(typeof(Read), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
