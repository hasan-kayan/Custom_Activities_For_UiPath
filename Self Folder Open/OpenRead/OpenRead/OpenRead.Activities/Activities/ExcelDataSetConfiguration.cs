namespace OpenRead.Activities
{
    internal class ExcelDataSetConfiguration
    {
        public ExcelDataSetConfiguration()
        {
        }

        public System.Func<object, ExcelDataTableConfiguration> ConfigureDataTable { get; set; }
    }
}