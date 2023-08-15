using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using LinkteraRobotics.CloseExcel.Activities.Design.Designers;
using LinkteraRobotics.CloseExcel.Activities.Design.Properties;

namespace LinkteraRobotics.CloseExcel.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(CloseExcel), categoryAttribute);
            builder.AddCustomAttributes(typeof(CloseExcel), new DesignerAttribute(typeof(CloseExcelDesigner)));
            builder.AddCustomAttributes(typeof(CloseExcel), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
