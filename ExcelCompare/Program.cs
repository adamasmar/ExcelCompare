using ConsoleAppFramework;
using Microsoft.Extensions.Hosting;
using System.Threading.Tasks;
using ExcelCompare.Utils;
using System.Text;

namespace ExcelCompare
{
    public class Program : ConsoleAppBase
    {
        static async Task Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            await Host.CreateDefaultBuilder().RunConsoleAppFrameworkAsync<Commands>(args);
        }
    }
}
