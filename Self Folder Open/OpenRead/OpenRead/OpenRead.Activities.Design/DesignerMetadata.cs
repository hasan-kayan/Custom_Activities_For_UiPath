using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using OpenRead.Activities.Design.Designers;
using OpenRead.Activities.Design.Properties;

namespace OpenRead.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(ReadData), categoryAttribute);
            builder.AddCustomAttributes(typeof(ReadData), new DesignerAttribute(typeof(ReadDataDesigner)));
            builder.AddCustomAttributes(typeof(ReadData), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
