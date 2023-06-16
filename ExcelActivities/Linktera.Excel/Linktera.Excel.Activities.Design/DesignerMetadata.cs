using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using Linktera.Excel.Activities.Design.Designers;
using Linktera.Excel.Activities.Design.Properties;

namespace Linktera.Excel.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(ExcelActivate), categoryAttribute);
            builder.AddCustomAttributes(typeof(ExcelActivate), new DesignerAttribute(typeof(ExcelActivateDesigner)));
            builder.AddCustomAttributes(typeof(ExcelActivate), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
