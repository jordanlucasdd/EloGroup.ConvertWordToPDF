using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using EloGroup.ConvertWordToPDF.Activities.Design.Designers;
using EloGroup.ConvertWordToPDF.Activities.Design.Properties;

namespace EloGroup.ConvertWordToPDF.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(ConvertDocToPDF), categoryAttribute);
            builder.AddCustomAttributes(typeof(ConvertDocToPDF), new DesignerAttribute(typeof(ConvertDocToPDFDesigner)));
            builder.AddCustomAttributes(typeof(ConvertDocToPDF), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
