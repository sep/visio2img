using System;
using System.IO;
using System.Linq;

namespace Visio2Img
{
    class Program
    {
        static int Main(string[] args)
        {
            var stderr = Console.Error;
            var stdout = Console.Out;

            var inputFilename = args[0];

            var options = new NDesk.Options.OptionSet
                {
                    {"<>", x => inputFilename = x}
                };

            // Arguments
            // 0: filename

            if (args.Length < 1)
            {
                PrintUsage();
                return 1;
            }

            if (args.Length > 1)
            {
                PrintUsage();
                return 1;
            }


            if (!inputFilename.EndsWith("vsd"))
            {
                stderr.WriteLine("Must specify a visio file (*.vsd)");
                return 2;
            }

            if (!File.Exists(inputFilename))
            {
                stderr.WriteLine("File does not exist. {0}", inputFilename);
                return 3;
            }

            var inputDirectory = Environment.CurrentDirectory;
            var outputDirectory = Environment.CurrentDirectory;
            var fileName = Path.IsPathRooted(inputFilename) ? inputFilename : Path.Combine(inputDirectory, inputFilename);

            stdout.WriteLine("Attempting export from {0}", fileName);

            var visio = new Microsoft.Office.Interop.Visio.Application();
            var document = visio.Documents.OpenEx(fileName, VisioConstants.VisOpenRo + VisioConstants.VisOpenNoWorkspace + VisioConstants.VisOpenNoWorkspace + VisioConstants.VisOpenMinimized + VisioConstants.VisOpenMacrosDisabled + VisioConstants.VisOpenHidden);

            using (Disposer.Create(visio.Quit))
            using (Disposer.Create(document.Close))
            {
                stdout.WriteLine("{0} pages to export", document.Pages.Count);
                Enumerable.Range(1, document.Pages.Count).ToList().ForEach(i =>
                {
                    stdout.WriteLine("({0}/{1}) Exporting {2}", i, document.Pages.Count, document.Pages[i].Name);
                    document.Pages[i].Export(string.Format(Path.Combine(outputDirectory, "{0}.png"), FixName(document.Pages[i].Name)));
                });
                stdout.WriteLine("Complete");

                return 0;
            }
        }

        private static string FixName(string name)
        {
            return Path.GetInvalidFileNameChars().Aggregate(name, (memo, invalid) => memo.Replace(invalid, '-'));
        }

        private static void PrintUsage()
        {
        }
    }
}
