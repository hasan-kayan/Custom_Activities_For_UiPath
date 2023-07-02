using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using LinkteraRobotics.Read.Range.Activities.Design.Designers;
using LinkteraRobotics.Read.Range.Activities.Design.Properties;

namespace LinkteraRobotics.Read.Range.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(ReadExcelRange), categoryAttribute);
            builder.AddCustomAttributes(typeof(ReadExcelRange), new DesignerAttribute(typeof(ReadExcelRangeDesigner)));
            builder.AddCustomAttributes(typeof(ReadExcelRange), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
