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


            // Arguments
            // 0: filename

            if (args.Length < 1)
            {
                return 1;
            }

            if (args.Length > 1)
            {
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

            var pageCount = 0;
            new Converter().Convert(
                fileName,
                cnt => { stdout.WriteLine("{0} pages to export", cnt); pageCount = cnt; },
                pageName => string.Format(Path.Combine(outputDirectory, "{0}.png"), FixName(pageName)),
                (filename, i) => stdout.WriteLine("({0}/{1}) Exported {2}", i, pageCount, filename)
                );
            return 0;
        }

        private static string FixName(string name)
        {
            return Path.GetInvalidFileNameChars().Aggregate(name, (memo, invalid) => memo.Replace(invalid, '-'));
        }
    }
}
