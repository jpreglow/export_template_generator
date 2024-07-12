
using ExportTemplateGenerator.Core2;

namespace ExportTemplateGenerator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var service = new ExcelImportService(args[0]);

            service.GetData();
        }
    }
}
