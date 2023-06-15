using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using Linktera.Robotics.Excel.Basics.Activities.Design.Designers;
using Linktera.Robotics.Excel.Basics.Activities.Design.Properties;

namespace Linktera.Robotics.Excel.Basics.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(OpenExcelScope), categoryAttribute);
            builder.AddCustomAttributes(typeof(OpenExcelScope), new DesignerAttribute(typeof(OpenExcelScopeDesigner)));
            builder.AddCustomAttributes(typeof(OpenExcelScope), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
