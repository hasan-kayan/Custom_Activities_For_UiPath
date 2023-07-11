using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using LinkteraRobotics.Read.Big.Data.Activities.Design.Designers;
using LinkteraRobotics.Read.Big.Data.Activities.Design.Properties;

namespace LinkteraRobotics.Read.Big.Data.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(ReadLargeExcelData), categoryAttribute);
            builder.AddCustomAttributes(typeof(ReadLargeExcelData), new DesignerAttribute(typeof(ReadLargeExcelDataDesigner)));
            builder.AddCustomAttributes(typeof(ReadLargeExcelData), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
