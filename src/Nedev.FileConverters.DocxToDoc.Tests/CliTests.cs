using System;
using System.IO;
using System.Diagnostics;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests
{
    public class CliTests
    {
        private string? GetCliPath()
        {
            // Try to find the CLI DLL in common build locations
            var possiblePaths = new[]
            {
                Path.Combine("..", "Nedev.FileConverters.DocxToDoc.Cli", "bin", "Debug", "net8.0", "Nedev.FileConverters.DocxToDoc.Cli.dll"),
                Path.Combine("..", "Nedev.FileConverters.DocxToDoc.Cli", "bin", "Release", "net8.0", "Nedev.FileConverters.DocxToDoc.Cli.dll"),
                Path.Combine("..", "..", "Nedev.FileConverters.DocxToDoc.Cli", "bin", "Debug", "net8.0", "Nedev.FileConverters.DocxToDoc.Cli.dll"),
            };

            foreach (var path in possiblePaths)
            {
                var fullPath = Path.GetFullPath(path);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }

            return null;
        }

        private ProcessStartInfo? CreateStartInfo(string arguments)
        {
            var cliPath = GetCliPath();
            if (cliPath == null) return null;

            return new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = $"\"{cliPath}\" {arguments}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
        }

        [Fact]
        public void Cli_Help_ReturnsZero()
        {
            // Arrange
            var psi = CreateStartInfo("--help");
            if (psi == null)
            {
                // Skip test if CLI not built
                return;
            }

            // Act
            using var process = Process.Start(psi);
            process?.WaitForExit();
            string output = process?.StandardOutput.ReadToEnd() ?? "";

            // Assert
            Assert.Equal(0, process?.ExitCode);
            Assert.Contains("Usage:", output);
            Assert.Contains("Options:", output);
        }

        [Fact]
        public void Cli_NoArguments_ReturnsError()
        {
            // Arrange
            var psi = CreateStartInfo("");
            if (psi == null)
            {
                // Skip test if CLI not built
                return;
            }

            // Act
            using var process = Process.Start(psi);
            process?.WaitForExit();

            // Assert
            Assert.Equal(1, process?.ExitCode);
        }

        [Fact]
        public void Cli_NonExistentFile_ReturnsError()
        {
            // Arrange
            var psi = CreateStartInfo("nonexistent.docx output.doc");
            if (psi == null)
            {
                // Skip test if CLI not built
                return;
            }

            // Act
            using var process = Process.Start(psi);
            process?.WaitForExit();
            string error = process?.StandardError.ReadToEnd() ?? "";

            // Assert
            Assert.Equal(2, process?.ExitCode);
            Assert.Contains("does not exist", error);
        }

        [Fact]
        public void Cli_VerboseFlag_ShowsDetails()
        {
            // Arrange - Create a test docx
            string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);
            string inputFile = Path.Combine(tempDir, "test.docx");
            string outputFile = Path.Combine(tempDir, "test.doc");

            try
            {
                // Create minimal docx
                CreateMinimalDocx(inputFile);

                var psi = CreateStartInfo($"-v \"{inputFile}\" \"{outputFile}\"");
                if (psi == null)
                {
                    // Skip test if CLI not built
                    return;
                }

                // Act
                using var process = Process.Start(psi);
                process?.WaitForExit();
                string output = process?.StandardOutput.ReadToEnd() ?? "";

                // Assert
                Assert.Equal(0, process?.ExitCode);
                Assert.Contains("Converting:", output);
                Assert.Contains("bytes", output);
            }
            finally
            {
                // Cleanup
                try { Directory.Delete(tempDir, true); } catch { }
            }
        }

        private void CreateMinimalDocx(string path)
        {
            using var fs = File.Create(path);
            using var archive = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Create);

            var entry = archive.CreateEntry("word/document.xml");
            using var stream = entry.Open();
            using var writer = new StreamWriter(stream);
            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                "<w:body><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:body></w:document>");
        }
    }
}
