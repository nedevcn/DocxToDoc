using System;
using System.Diagnostics;
using System.IO;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests
{
    public class CliTests
    {
        [Fact]
        public void HelpArgument_ReturnsNonZeroAndPrintsUsage()
        {
            var exe = Path.GetFullPath(Path.Combine("..", "Nedev.FileConverters.DocxToDoc.Cli", "bin", "Debug", "net10.0", "Nedev.FileConverters.DocxToDoc.Cli.dll"));
            var process = new Process();
            process.StartInfo.FileName = "dotnet";
            process.StartInfo.Arguments = $