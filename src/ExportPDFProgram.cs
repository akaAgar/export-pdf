using ICSharpCode.SharpZipLib.Zip;
using SautinSoft;
using System;
using System.IO;

namespace ExportPDF
{
    public sealed class ExportPDFProgram
    {
        private static void Main(params string[] args)
        {
            if (args.Length == 0)
                Console.WriteLine($"Aucun fichier fourni.");

            foreach (string filePath in args)
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File \"{Path.GetFileName(filePath)}\" n'existe pas.");
                    continue;
                }

                using (ZipFile zipFile = new ZipFile(filePath))
                {
                    bool pdfFound = false;

                    foreach (ZipEntry zipEntry in zipFile)
                    {
                        if (!zipEntry.IsFile) continue;
                        if (Path.GetFileName(zipEntry.Name).ToLowerInvariant().EndsWith(".pdf"))
                        {
                            PdfFocus pdfFocus = new PdfFocus();
                            pdfFocus.OpenPdf(zipFile.GetInputStream(zipEntry));
                            pdfFocus.WordOptions.Format = PdfFocus.CWordOptions.eWordDocument.Docx;
                            int result = pdfFocus.ToWord($"{Path.GetFileNameWithoutExtension(filePath)}.docx");
                            Console.WriteLine($"Exporté vers \"{Path.GetFileNameWithoutExtension(filePath)}.docx\".");

                            pdfFound = true;
                            break;
                        }
                    }

                    if (!pdfFound)
                        Console.WriteLine($"File \"{Path.GetFileName(filePath)}\" ne contient pas de PDF.");
                }
            }

#if DEBUG
            Console.WriteLine();
            Console.WriteLine("Pressez une touche pour fermer cette fenêtre...");
            Console.ReadKey();
#endif
        }
    }
}
