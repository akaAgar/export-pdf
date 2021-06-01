using ICSharpCode.SharpZipLib.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExportPDF
{
    public sealed class ExportPDFProgram
    {
        private const string HTML_INTRO_END = "<h4 class=\"block-title\"><em>Contenu</em></h4>";
        private const string HTML_OUTRO_START1 = "<h4 class=\"block-title\"><em>Legendes des images</em></h4>";
        private const string HTML_OUTRO_START2 = "<style type=\"text/css\">";
        private const string BLOCK_INTRO = "<div class=\"layout\">";

        private static readonly string[] BLOCKS_TO_EXTRACT = new string[] { "Framed", "Citation", "Photo", "Images" };

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

                switch (Path.GetExtension(filePath).ToLowerInvariant())
                {
                    default:
                        Console.WriteLine(Path.GetExtension(filePath).ToLowerInvariant());
                        break;

                    case ".htm":
                    case ".html":
                        DoExport(File.ReadAllText(filePath), $"{Path.GetFileNameWithoutExtension(filePath)} (trié).html");
                        break;
                    case ".zip":
                        using (ZipFile zipFile = new ZipFile(filePath))
                        {
                            bool fileFound = false;

                            foreach (ZipEntry zipEntry in zipFile)
                            {
                                if (Path.GetFileName(zipEntry.Name).ToLowerInvariant().EndsWith(".html"))
                                {
                                    string html;
                                    using (Stream entryStream = zipFile.GetInputStream(zipEntry))
                                    {
                                        byte[] entryBytes = new byte[zipEntry.Size];
                                        entryStream.Read(entryBytes, 0, entryBytes.Length);
                                        html = Encoding.UTF8.GetString(entryBytes).Replace("\t", "").Replace("\r", "").Replace("\n", "");
                                    }

                                    DoExport(html, $"{Path.GetFileNameWithoutExtension(filePath)}.html");

                                    fileFound = true;
                                    break;
                                }
                            }

                            if (!fileFound)
                                Console.WriteLine($"File \"{Path.GetFileName(filePath)}\" ne contient pas de fichier HTML.");
                        }
                        break;
                }
            }

#if DEBUG
            Console.WriteLine();
            Console.WriteLine("Pressez une touche pour fermer cette fenêtre...");
            Console.ReadKey();
#endif
        }

        private static void DoExport(string html, string fileName)
        {
            string htmlIntro = html.Substring(0, html.IndexOf(HTML_INTRO_END) + HTML_INTRO_END.Length);
            html = html.Substring(htmlIntro.Length);

            string htmlOutro = "";
            if (html.Contains(HTML_OUTRO_START1))
                htmlOutro = html.Substring(html.IndexOf(HTML_OUTRO_START1));
            else if (html.Contains(HTML_OUTRO_START2))
                htmlOutro = html.Substring(html.IndexOf(HTML_OUTRO_START2));
            html = html.Substring(0, html.Length - htmlOutro.Length);

            string[] htmlBlocks = html.Split(new string[] { BLOCK_INTRO }, StringSplitOptions.None);
            List<string> mainBlocks = new List<string>();
            List<string>[] extraBlocks = new List<string>[BLOCKS_TO_EXTRACT.Length];
            for (int i = 0; i < BLOCKS_TO_EXTRACT.Length; i++) extraBlocks[i] = new List<string>();
            foreach (string htmlBlock in htmlBlocks)
            {
                bool extraBlock = false;
                for (int i = 0; i < BLOCKS_TO_EXTRACT.Length; i++)
                {
                    if (htmlBlock.Contains($"<h5 class=\"layout-title\">#{BLOCKS_TO_EXTRACT[i]}#</h5>"))
                    {
                        extraBlock = true;
                        extraBlocks[i].Add(BLOCK_INTRO + htmlBlock);
                        break;
                    }
                }

                if (extraBlock) continue;
                mainBlocks.Add(BLOCK_INTRO + htmlBlock);
            }

            html = "";
            foreach (string block in mainBlocks)
                html += block;

            html += "<h4 class=\"block-title\"><em>Éléments en plus</em></h4>";

            for (int i = 0; i < BLOCKS_TO_EXTRACT.Length; i++)
                foreach (string block in extraBlocks[i])
                    html += block;

            html = htmlIntro + "\r\n\r\n" + html + "\r\n\r\n" + htmlOutro;

            File.WriteAllText(fileName, html);
            Console.WriteLine($"Exporté vers \"{fileName}\".");
        }
    }
}

/*
 * OLD DOCX CONVERSION STUFF
 * 
 *                         if (Path.GetFileName(zipEntry.Name).ToLowerInvariant().EndsWith(".pdf"))
                        {
                            PdfFocus pdfFocus = new PdfFocus();
                            pdfFocus.OpenPdf(zipFile.GetInputStream(zipEntry));
                            pdfFocus.WordOptions.Format = PdfFocus.CWordOptions.eWordDocument.Docx;
                            int result = pdfFocus.ToWord($"{Path.GetFileNameWithoutExtension(filePath)}.docx");
                            Console.WriteLine($"Exporté vers \"{Path.GetFileNameWithoutExtension(filePath)}.docx\".");

                            pdfFound = true;
                            break;
                        }
*/