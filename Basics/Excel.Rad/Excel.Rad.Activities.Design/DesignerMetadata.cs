using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using Excel.Rad.Activities.Design.Designers;
using Excel.Rad.Activities.Design.Properties;

namespace Excel.Rad.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(ExcelRangeRead), categoryAttribute);
            builder.AddCustomAttributes(typeof(ExcelRangeRead), new DesignerAttribute(typeof(ExcelRangeReadDesigner)));
            builder.AddCustomAttributes(typeof(ExcelRangeRead), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
