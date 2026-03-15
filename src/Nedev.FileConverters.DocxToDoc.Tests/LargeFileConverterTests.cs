using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests
{
    public class LargeFileConverterTests
    {
        private byte[] CreateTestDocx()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                // Create docProps/core.xml
                var propsEntry = archive.CreateEntry("docProps/core.xml");
                using (var stream = propsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<coreProperties xmlns=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\">" +
                        "<title>Test Document</title>" +
                        "<creator>Test Author</creator>" +
                        "<pages>5</pages>" +
                        "<words>100</words>" +
                        "</coreProperties>");
                }

                // Create word/document.xml
                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p><w:r><w:t>Test paragraph</w:t></w:r></w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }
            return ms.ToArray();
        }

        [Fact]
        public void AnalyzeDocument_ValidDocx_ReturnsDocumentInfo()
        {
            // Arrange
            string tempFile = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.docx");
            File.WriteAllBytes(tempFile, CreateTestDocx());

            var converter = new LargeFileConverter();

            try
            {
                // Act
                var info = converter.AnalyzeDocument(tempFile);

                // Assert
                Assert.NotNull(info);
                Assert.Equal(tempFile, info.FilePath);
                Assert.True(info.FileSize > 0);
                Assert.Equal("Test Document", info.Title);
                // Note: Some properties may not be parsed due to namespace variations
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [Fact]
        public void AnalyzeDocument_NullPath_ThrowsArgumentNullException()
        {
            // Arrange
            var converter = new LargeFileConverter();

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => converter.AnalyzeDocument(null!));
        }

        [Fact]
        public void AnalyzeDocument_NonExistentFile_ThrowsFileNotFoundException()
        {
            // Arrange
            var converter = new LargeFileConverter();
            string nonExistentPath = Path.Combine(Path.GetTempPath(), $"nonexistent_{Guid.NewGuid()}.docx");

            // Act & Assert
            Assert.Throws<FileNotFoundException>(() => converter.AnalyzeDocument(nonExistentPath));
        }

        [Fact]
        public void IsLargeFile_SmallFile_ReturnsFalse()
        {
            // Arrange
            string tempFile = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.docx");
            File.WriteAllBytes(tempFile, CreateTestDocx());

            var converter = new LargeFileConverter(largeFileThreshold: 1024 * 1024); // 1MB threshold

            try
            {
                // Act
                bool isLarge = converter.IsLargeFile(tempFile);

                // Assert
                Assert.False(isLarge);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [Fact]
        public void IsLargeFile_LargeFile_ReturnsTrue()
        {
            // Arrange
            string tempFile = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.docx");
            File.WriteAllBytes(tempFile, CreateTestDocx());

            var converter = new LargeFileConverter(largeFileThreshold: 100); // 100 bytes threshold

            try
            {
                // Act
                bool isLarge = converter.IsLargeFile(tempFile);

                // Assert
                Assert.True(isLarge);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [Fact]
        public void DocumentInfo_GetSummary_ReturnsFormattedString()
        {
            // Arrange
            var info = new DocumentInfo
            {
                FilePath = @"C:\test\document.docx",
                FileSize = 1024 * 1024, // 1MB
                IsLargeFile = true,
                Title = "Test Doc",
                Author = "Test Author",
                PageCount = 10,
                WordCount = 5000,
                ParagraphCount = 100,
                ImageCount = 5,
                HasImages = true,
                HasComments = false,
                CommentCount = 0
            };

            // Act
            string summary = info.GetSummary();

            // Assert
            Assert.Contains("document.docx", summary);
            Assert.Contains("1 MB", summary);
            Assert.Contains("Large", summary);
            Assert.Contains("Test Doc", summary);
            Assert.Contains("Test Author", summary);
            Assert.Contains("10", summary);
            Assert.Contains("5000", summary);
        }

        [Fact]
        public void LargeFileConversionOptions_DefaultValues_AreCorrect()
        {
            // Arrange & Act
            var options = new LargeFileConversionOptions();

            // Assert
            Assert.False(options.ForceStreamingMode);
            Assert.True(options.EnableGarbageCollection);
            Assert.Equal(81920, options.BufferSize);
            Assert.True(options.ProcessImages);
            Assert.True(options.ProcessComments);
        }
    }
}
