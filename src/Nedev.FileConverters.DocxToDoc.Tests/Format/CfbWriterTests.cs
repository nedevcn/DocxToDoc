using System;
using System.IO;
using OpenMcdf;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class CfbWriterTests
    {
        [Fact]
        public void WriteTo_BasicStreams_ShouldProduceValidCompoundFile()
        {
            // Arrange
            byte[] docData = new byte[] { 0x01, 0x02, 0x03, 0x04 };
            byte[] tableData = new byte[1024]; // Larger than minifat size (4096 is threshold, so this goes to MiniFAT)
            new Random(123).NextBytes(tableData);

            using var cfbWriter = new Nedev.FileConverters.DocxToDoc.Format.CfbWriter();
            cfbWriter.AddStream("WordDocument", docData);
            cfbWriter.AddStream("1Table", tableData);

            using var outputStream = new MemoryStream();

            // Act
            cfbWriter.WriteTo(outputStream);
            byte[] resultBytes = outputStream.ToArray();

            // Assert
            // 1. Minimum size of a CFB file is 3 sectors (Header + FAT + Directory) = 1536 bytes
            Assert.True(resultBytes.Length >= 1536, $"CFB file is too small: {resultBytes.Length} bytes");

            // 2. Validate with an independent CFB parser (OpenMcdf)
            using var inputStream = new MemoryStream(resultBytes);
            using var compoundFile = new CompoundFile(inputStream, CFSUpdateMode.ReadOnly, CFSConfiguration.Default);
            
            // 3. Verify streams exist and have correct data
            var rootStorage = compoundFile.RootStorage;
            
            var wordDocStream = rootStorage.GetStream("WordDocument");
            Assert.Equal(docData.Length, wordDocStream.Size);
            Assert.Equal(docData, wordDocStream.GetData());

            var tableStream = rootStorage.GetStream("1Table");
            Assert.Equal(tableData.Length, tableStream.Size);
            Assert.Equal(tableData, tableStream.GetData());
        }
    }
}
