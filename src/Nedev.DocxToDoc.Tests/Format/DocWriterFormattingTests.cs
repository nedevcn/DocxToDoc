using System;
using System.IO;
using System.Text;
using System.IO.Compression;
using Nedev.DocxToDoc.Format;
using Nedev.DocxToDoc.Model;
using Xunit;

namespace Nedev.DocxToDoc.Tests.Format
{
    public class DocWriterFormattingTests
    {
        [Fact]
        public void WriteDocBlocks_FormattingIncluded_CreatesChpx()
        {
            // Register encoding provider for Windows-1252 
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Arrange
            var model = new DocumentModel();
            var para = new ParagraphModel();
            var run = new RunModel { Text = "FormatTest" };
            run.Properties.IsBold = true;
            run.Properties.IsItalic = true;
            run.Properties.FontSize = 24; // 12pt
            
            para.Runs.Add(run);
            model.Content.Add(para);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            // Act
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            // Assert
            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));
            
            byte[] tableData = tableStream.GetData();
            byte[] wordDocData = wordDocStream.GetData();

            // Just rough assertions that the streams are non-empty and larger than minimums 
            Assert.True(tableData.Length > 20); // Should contain Clx and PlcfbteChpx
            
            // Expected WordDoc min size: 1536 (Text Start) + 11 (Text length) 
            // Plus at least one 512 byte FKP since we have formatting
            Assert.True(wordDocData.Length >= 1536 + 11 + 512); 
        }

        [Fact]
        public void WriteDocBlocks_ParagraphFormatting_CreatesPapx()
        {
            // Register encoding provider for Windows-1252 
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Arrange
            var model = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel { Text = "ParaTest" });
            para.Properties.Alignment = ParagraphModel.Justification.Center;
            model.Content.Add(para);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            // Act
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            // Assert
            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));
            
            byte[] tableData = tableStream.GetData();
            byte[] wordDocData = wordDocStream.GetData();

            // Check that the FIB was updated with non-zero Papx PLCF offsets
            // The PLCF for Papx is index 2 in RgFcLcb (offset 154 + 2*8 = 170)
            // Fib size is >= 154 + 744 = 898
            Assert.True(wordDocData.Length >= 898);
            int fcPlcfbtePapx = BitConverter.ToInt32(wordDocData, 170);
            int lcbPlcfbtePapx = BitConverter.ToInt32(wordDocData, 174);
            
            Assert.NotEqual(0, fcPlcfbtePapx);
            Assert.True(lcbPlcfbtePapx > 0);
            
            // Verify that a PAPX FKP was written to WordDocument
            Assert.True(wordDocData.Length >= 1536 + 9 + 512); 
        }
    }
}
