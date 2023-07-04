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

            builder.AddCustomAttributes(typeof(ReadRange), categoryAttribute);
            builder.AddCustomAttributes(typeof(ReadRange), new DesignerAttribute(typeof(ReadRangeDesigner)));
            builder.AddCustomAttributes(typeof(ReadRange), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
