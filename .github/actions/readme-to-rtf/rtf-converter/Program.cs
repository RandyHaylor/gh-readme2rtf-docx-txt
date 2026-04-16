using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using RtfPipe;

using WpDocument = DocumentFormat.OpenXml.Wordprocessing.Document;
using WpBody = DocumentFormat.OpenXml.Wordprocessing.Body;
using AltChunk = DocumentFormat.OpenXml.Wordprocessing.AltChunk;

class Program
{
    static void Main(string[] args)
    {
        // Register Windows-1252 encoding required by RtfPipe on .NET Core
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: rtf-converter <input.rtf> <output.docx|output.pdf>");
            Console.WriteLine("  Specify multiple outputs: rtf-converter input.rtf output.docx output.pdf");
            return;
        }

        var inputRtf = args[0];
        var outputs = args.Skip(1).ToArray();

        if (!File.Exists(inputRtf))
        {
            Console.Error.WriteLine($"Input file not found: {inputRtf}");
            Environment.Exit(1);
        }

        // Parse RTF to HTML using RtfPipe
        var rtfContent = File.ReadAllText(inputRtf);
        var html = Rtf.ToHtml(rtfContent);

        foreach (var output in outputs)
        {
            var ext = Path.GetExtension(output).ToLowerInvariant();
            switch (ext)
            {
                case ".docx":
                    ConvertToDocx(html, rtfContent, output);
                    Console.WriteLine($"Created: {output}");
                    break;
                case ".pdf":
                    ConvertToPdf(html, output);
                    Console.WriteLine($"Created: {output}");
                    break;
                default:
                    Console.Error.WriteLine($"Unsupported format: {ext}");
                    break;
            }
        }
    }

    static void ConvertToDocx(string html, string rtfContent, string outputPath)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new WpDocument();
            var body = new WpBody();

            // Use the alternative XHTML import approach
            var altChunkId = "htmlChunk";
            var chunk = mainPart.AddAlternativeFormatImportPart(
                AlternativeFormatImportPartType.Html, altChunkId);

            // Wrap HTML in a full document with UTF-8 encoding
            var fullHtml = $"<html><head><meta charset='utf-8'></head><body>{html}</body></html>";
            using (var chunkStream = chunk.GetStream())
            {
                var bytes = Encoding.UTF8.GetBytes(fullHtml);
                chunkStream.Write(bytes, 0, bytes.Length);
            }

            var altChunk = new AltChunk() { Id = altChunkId };
            body.Append(altChunk);
            mainPart.Document.Append(body);
        }

        File.WriteAllBytes(outputPath, ms.ToArray());
    }

    static void ConvertToPdf(string html, string outputPath)
    {
        // Strip HTML tags for text-based PDF (PdfSharpCore doesn't render HTML)
        var plainText = System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", "");
        plainText = System.Net.WebUtility.HtmlDecode(plainText);

        var pdf = new PdfDocument();
        var font = new XFont("Courier", 10);
        var boldFont = new XFont("Courier", 10, XFontStyle.Bold);

        var lines = plainText.Split('\n');
        var pageHeight = 792.0; // Letter
        var pageWidth = 612.0;
        var margin = 72.0;
        var lineHeight = 14.0;
        var usableHeight = pageHeight - 2 * margin;
        var usableWidth = pageWidth - 2 * margin;

        XGraphics gfx = null;
        PdfPage page = null;
        var y = margin;

        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            if (string.IsNullOrEmpty(trimmed))
            {
                y += lineHeight;
            }
            else
            {
                if (gfx == null || y + lineHeight > pageHeight - margin)
                {
                    page = pdf.AddPage();
                    page.Width = pageWidth;
                    page.Height = pageHeight;
                    gfx = XGraphics.FromPdfPage(page);
                    y = margin;
                }

                gfx.DrawString(trimmed, font, XBrushes.Black,
                    new XRect(margin, y, usableWidth, lineHeight),
                    XStringFormats.TopLeft);
                y += lineHeight;
            }
        }

        if (pdf.PageCount == 0)
        {
            pdf.AddPage();
        }

        pdf.Save(outputPath);
    }
}
