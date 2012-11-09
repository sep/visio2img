using System;
using System.Linq;

namespace Visio2Img
{
    public class Converter
    {
        public void Convert(string visioFile,
                            Action<int> pageCount,
                            Func<string, string> getFilename,
                            Action<string, int> fileCreated)
        {
            var visio = new Microsoft.Office.Interop.Visio.Application();
            var document = visio.Documents.OpenEx(visioFile, VisioConstants.VisOpenRo + VisioConstants.VisOpenNoWorkspace + VisioConstants.VisOpenNoWorkspace + VisioConstants.VisOpenMinimized + VisioConstants.VisOpenMacrosDisabled + VisioConstants.VisOpenHidden);

            using (Disposer.Create(visio.Quit))
            using (Disposer.Create(document.Close))
            {
                pageCount(document.Pages.Count);
                Enumerable.Range(1, document.Pages.Count).ToList().ForEach(i =>
                {
                    var page = document.Pages[i];
                    var filename = getFilename(page.Name);

                    page.Export(filename);
                    fileCreated(filename, i);
                });
            }
        }
    }
}