using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using LinkteraRobotics.OpenExcel.Activities.Design.Designers;
using LinkteraRobotics.OpenExcel.Activities.Design.Properties;

namespace LinkteraRobotics.OpenExcel.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(OpenExcel), categoryAttribute);
            builder.AddCustomAttributes(typeof(OpenExcel), new DesignerAttribute(typeof(OpenExcelDesigner)));
            builder.AddCustomAttributes(typeof(OpenExcel), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
