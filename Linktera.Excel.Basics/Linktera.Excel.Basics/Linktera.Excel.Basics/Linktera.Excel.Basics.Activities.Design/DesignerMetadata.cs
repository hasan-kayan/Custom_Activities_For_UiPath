using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using Linktera.Excel.Basics.Activities.Design.Designers;
using Linktera.Excel.Basics.Activities.Design.Properties;

namespace Linktera.Excel.Basics.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(OpenExcelFile), categoryAttribute);
            builder.AddCustomAttributes(typeof(OpenExcelFile), new DesignerAttribute(typeof(OpenExcelFileDesigner)));
            builder.AddCustomAttributes(typeof(OpenExcelFile), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
