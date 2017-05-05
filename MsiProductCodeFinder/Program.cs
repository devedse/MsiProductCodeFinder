using System;

namespace MsiProductCodeFinder
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            //Check if here are any parameters
            if (args == null || args.Length == 0)
            {
                Console.WriteLine("The tool requires an argument.");
                Environment.Exit(1);
            }

            string strFileName = args[0];

            var result = ProductCodeFinder.ObtainProductCode(strFileName);

            if (result.Success)
            {
                Console.WriteLine(result.Result);
                Environment.Exit(0);
            }
            else
            {
                Console.WriteLine($"Error: {result.Result}");
                Environment.Exit(1);
            }
        }
    }
}
