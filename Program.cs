using System.Diagnostics;
using System.IO.Compression;
using System.Web;
using System.Xml.Linq;

namespace OpenXmlBrowser;


class Program
{
    static void Main(string[] args)
    {
        if (args.Length != 1 || !File.Exists(args[0]))
        {
            Usage();
            return;
        }

        string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".html");
        using (Stream html = File.OpenWrite(filePath))
        using (StreamWriter htmlFile = new StreamWriter(html))
        using (Stream doc = File.OpenRead(args[0]))
        using (ZipArchive archive = new ZipArchive(doc))
        {
            WriteHtmlFileHeader(args[0], htmlFile, filePath);

            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                WriteTOCEntry(htmlFile, entry);
            }

            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                WriteHtmlEntryHeader(htmlFile, entry);
                if (!entry.FullName.EndsWith("/"))
                {
                    bool isBinary = false;
                    using(var entryStream = entry.Open())
                    using(var sr = new StreamReader(entryStream))
                    {
                        string content = sr.ReadToEnd();
                        if (!IsBinaryFile(content))
                        {
                            htmlFile.Write(HttpUtility.HtmlEncode(FormatXml(content)));
                        }
                        else
                        {
                            isBinary = true;
                        }
                    }

                    if (isBinary)
                    {
                        using(var entryStream = entry.Open())
                        {
                            var binData = ReadFully(entryStream);
                            if (entry.Name.ToLower().EndsWith(".png"))
                            {
                                var base64Data = Convert.ToBase64String(binData);
                                htmlFile.Write($"<img src=\"data:image/png;base64,{base64Data}\"/>");
                            }
                            else if (entry.Name.ToLower().EndsWith(".jpg") || entry.Name.ToLower().EndsWith(".jpeg"))
                            {
                                var base64Data = Convert.ToBase64String(binData);
                                htmlFile.Write($"<img src=\"data:image/jpeg;base64,{base64Data}\"/>");
                            }
                            else
                            {
                                htmlFile.Write("**** BINARY DATA ****");
                            }
                        }
                    }
                }
                WriteHtmlEntryFooter(htmlFile);
            }  
        }
        Process.Start(new ProcessStartInfo(filePath){
            UseShellExecute = true
        });
    }

    // If 85% of the text is text-ish, let's call it text. ()
    private static bool IsBinaryFile(string content)
    {
        int totalCount = Math.Min(2000, content.Length);
        int textChars = 0;
        for(int index = 0; index < totalCount; index++)
        {
            
            if ("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()<>?/.,".IndexOf(content[index].ToString().ToUpper()) >= 0)
            {
                textChars++;
            }
        }
        if (textChars < (totalCount * 0.85))
        {
            return true;
        }
        return false;
    }

    // html & css stuff stolen from these pages: https://renenyffenegger.ch/notes//Microsoft/Office/Open-XML/SpreadsheetML/
    private static void WriteHtmlFileHeader(string filename, StreamWriter htmlFile, string filePath)
    {
        string html = "<!DOCTYPE html>";
        html +="<html>";
        html +="<head>";
        html +="<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">";
        html +="<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">";
        html += $"<title>{filePath}</title>";
        html +="<style type=\"text/css\">";
        html +="@font-face {";
        html +="font-family: Liberation Sans;";
        html +="src:         local(\"Liberation Sans\"),";
        html +="            local(\"Liberation Sans Regular\"),";
        html +="            local(\"LiberationSans-Regular\"),";
        html +="            url(/font/LiberationSans-Regular.ttf);";
        html +="            font-weight: normal; font-style: normal;";
        html +="}";
        html +="@font-face {";
        html +="font-family: Liberation Sans;";
        html +="src:         local(\"Liberation Sans Italic\"),";
        html +="            local(\"LiberationSans-Italic\"),";
        html +="            url(/font/LiberationSans-Italic.ttf);";
        html +="            font-weight: normal; font-style: italic;";
        html +="}";
        html +="@font-face {";
        html +="font-family: Liberation Sans;";
        html +="src:         local(\"Liberation Sans Bold\"),";
        html +="            local(\"LiberationSans-Bold\"),";
        html +="            url(/font/LiberationSans-Bold.ttf);";
        html +="            font-weight: bold; font-style: normal;";
        html +="}";
        html +="@font-face {";
        html +="font-family: Liberation Sans;";
        html +="src:         local(\"Liberation Sans Bold Italic\"),";
        html +="            local(\"Liberation Sans BoldItalic\"),";
        html +="            local(\"LiberationSans-BoldItalic\"),";
        html +="            url(/font/LiberationSans-BoldItalic.ttf);";
        html +="            font-weight: bold; font-style: italic;";
        html +="}";
        html +="div, table, h1  {font-family: Liberation Sans}";
        html +="h2 {";
        html +="margin-bottom: 4px;";
        html +="}";
        html +="div.h {";
        html +="margin-left: 20px;";
        html +="}";
        html +="code {";
        html +="font-family: monospace;";
        html +="white-space: pre";
        html +="}";
        html +="pre.code {";
        html +="    font-family: monospace;";
        html +="    background-color:#f6f9e3;";
        html +="    color:#10205f;";
        html +="    padding-left:10px; ";
        html +="    border:1px solid #004080;";
        html +="    overflow: hidden;";
        html +="}";
        html +="@media screen {";
        html +="body {margin-left: 2em;}";
        html +="pre.code, blockquote {";
        html +="    padding-top:3px;";
        html +="    padding-bottom:5px;";
        html +="    margin-top:2px;";
        html +="    margin-bottom: 2px;";
        html +="}";
        html +="blockquote {";
        html +="    ";
        html +="    background-color: #f9f9f9;";
        html +="    margin-left: 20px;";
        html +="    width: 43em; ";
        html +="    font-family: Garamond;";
        html +="    color: #061;";
        html +="}";
        html +="}";
        html +="@media screen and (max-width: 641px) { ";
        html +="a, h2 { word-break: break-all; } ";
        html +="body {margin-left: 0.5em;";
        html +="}";
        html +="}";
        html +="}";
        html +="@media print {";
        html +="h2 {font-size: 18px}";
        html +="pre.code, blockquote {";
        html +="    padding-top:3px;";
        html +="    padding-bottom:5px;";
        html +="    margin-top:2px;";
        html +="    margin-bottom: 2px;";
        html +="}";
        html +="div.screen-only {display: none}";
        html +="a {text-decoration: none; color: black}";
        html +="}";

        html +="</style>";
        html +="</head>";
        html +="<body>";
        html +=$"<h1>OpenXmlBrowser Dump: {filename}</h1>";
        html +="<h2>Files:</h2>";
        htmlFile.Write(html);
    }

    private static void WriteTOCEntry(StreamWriter htmlFile, ZipArchiveEntry entry)
    {
        string html = $"<a href=\"#{GetEntryId(entry)}\">{entry.FullName}</a><br />";
        htmlFile.Write(html);
    }

    private static string GetEntryId(ZipArchiveEntry entry)
    {
        return entry.Name;
    }


    private static void WriteHtmlEntryHeader(StreamWriter htmlFile, ZipArchiveEntry entry)
    {
        string html =$"<div id=\"{GetEntryId(entry)}\" class='h'>";
        html += $"<h2>{entry.FullName}</h2>";
        html += "<pre class='code'>";
        htmlFile.Write(html);
    }

    private static void WriteHtmlEntryFooter(StreamWriter htmlFile)
    {
        string html ="</pre>";
        html += "</div>";
        htmlFile.Write(html);
    }

    // https://stackoverflow.com/questions/1123718/format-xml-string-to-print-friendly-xml-string
    private static string FormatXml(string xml)
    {
        try
        {
            XDocument doc = XDocument.Parse(xml);
            return doc.ToString();
        }
        catch (Exception)
        {
            return xml;
        }
    }

    private static void Usage()
    {
        Console.WriteLine("OpenXmlBrowser filename.[xlsx,docx,pptx, etc]" + Environment.NewLine);
        Console.WriteLine("  filename.* - OpenDoc file to browse");
    }

    // Stolen from the great Jon Skeet: https://jonskeet.uk/csharp/readbinary.html
    public static byte[] ReadFully (Stream stream)
    {
        byte[] buffer = new byte[32768];
        int read=0;
        int chunk;

        while ( (chunk = stream.Read(buffer, read, buffer.Length-read)) > 0)
        {
            read += chunk;
            
            // If we've reached the end of our buffer, check to see if there's
            // any more information
            if (read == buffer.Length)
            {
                int nextByte = stream.ReadByte();
                
                // End of stream? If so, we're done
                if (nextByte==-1)
                {
                    return buffer;
                }
                
                // Nope. Resize the buffer, put in the byte we've just
                // read, and continue
                byte[] newBuffer = new byte[buffer.Length*2];
                Array.Copy(buffer, newBuffer, buffer.Length);
                newBuffer[read]=(byte)nextByte;
                buffer = newBuffer;
                read++;
            }
        }
        // Buffer is now too big. Shrink it.
        byte[] ret = new byte[read];
        Array.Copy(buffer, ret, read);
        return ret;
    }    
}