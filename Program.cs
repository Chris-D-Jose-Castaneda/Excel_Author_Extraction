// Program.cs to check file authors and extract info with codes to pinpoint author origin

using System;
using System.IO;
using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using System.Diagnostics;

namespace SubmissionAnalyzer
{
    class Program
    {
        static void Main(string[] args)
        {
            // Prompt user for the ZIP path
            Console.Write("Enter full path to the ZIP file: ");
            string zipPath = Console.ReadLine()?.Trim() ?? "";
            
            // Prompt user for the extraction directory
            Console.Write("Enter directory where files should be extracted: ");
            string extractTo = Console.ReadLine()?.Trim() ?? "";

            Console.WriteLine();
            Console.WriteLine($"ZIP file:   {zipPath}");
            Console.WriteLine($"Extract to: {extractTo}");
            Console.WriteLine();

            if (string.IsNullOrWhiteSpace(zipPath) || !File.Exists(zipPath))
            {
                Console.Error.WriteLine("ERROR: ZIP file not found or path is empty. Please check the path and try again.");
                return;
            }

            if (string.IsNullOrWhiteSpace(extractTo))
            {
                Console.Error.WriteLine("ERROR: Extraction directory path is empty. Please provide a valid directory.");
                return;
            }

            try
            {
                var stopwatch = Stopwatch.StartNew();
                ExtractAndAnalyzeFiles(zipPath, extractTo);
                stopwatch.Stop();

                Console.WriteLine($"Total execution time: {stopwatch.ElapsedMilliseconds} ms");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        static void ExtractAndAnalyzeFiles(string zipPath, string extractTo)
        {
            // Ensure the extraction directory exists
            Directory.CreateDirectory(extractTo);

            // Open the ZIP for reading and extract each entry
            using (var archive = ZipFile.OpenRead(zipPath))
            {
                foreach (var entry in archive.Entries)
                {
                    string destinationPath = Path.GetFullPath(Path.Combine(extractTo, entry.FullName));
                    if (!destinationPath.StartsWith(extractTo, StringComparison.Ordinal))
                        throw new IOException("Entry is trying to extract outside of the target directory.");

                    Directory.CreateDirectory(Path.GetDirectoryName(destinationPath)!);
                    entry.ExtractToFile(destinationPath, overwrite: true);
                }
            }

            // Find all .xlsx files (recursively) and analyze each one
            var excelFiles = Directory.GetFiles(extractTo, "*.xlsx", SearchOption.AllDirectories);
            foreach (var file in excelFiles)
            {
                AnalyzeFileMetadata(file);
            }
        }

        static void AnalyzeFileMetadata(string filePath)
        {
            using (var document = SpreadsheetDocument.Open(filePath, false))
            {
                var props = document.PackageProperties;

                Console.WriteLine($"File Name:         {Path.GetFileName(filePath)}");
                Console.WriteLine($"Author:            {props.Creator}");
                Console.WriteLine($"Last Modified By:  {props.LastModifiedBy}");
                Console.WriteLine($"Last Modified On:  {props.Modified:yyyy-MM-dd HH:mm:ss}");
                Console.WriteLine(new string('-', 50));
            }
        }
    }
}

// Ensure that the file names are extracted. Separate NETID & ASSIGN to the name of a candidate.
// A search algorithm that reads each file & groups students that work together
// Search for 4 digit unique Random Code