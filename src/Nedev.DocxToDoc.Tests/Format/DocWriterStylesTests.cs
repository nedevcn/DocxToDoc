using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Nedev.DocxToDoc.Format;
using Nedev.DocxToDoc.Model;
using Xunit;

namespace Nedev.DocxToDoc.Tests.Format
{
    public class DocWriterStylesTests
    {
        [Fact]
        public void WriteDocBlocks_WithStylesAndFonts_WritesSttbAndStsh()
        {
            // Register encoding provider for Windows-1252 
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Arrange
            var model = new DocumentModel
            {
                TextBuffer = "StyleTest\r"
            };
            
            model.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "StyleTest" } } });
            
            model.Styles.Add(new StyleModel { Id = "Normal", Name = "Normal", IsParagraphStyle = true });
            model.Fonts.Add(new FontModel { Name = "Arial" });
            model.Fonts.Add(new FontModel { Name = "Times New Roman" });

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

            // Check FIB offsets for STSH and Font Table
            // fcStshf is index 0 in RgFcLcb (offset 154)
            // fcSttbfffn is index 14 in RgFcLcb (offset 154 + 14 * 8 = 266)
            
            int fcStshf = BitConverter.ToInt32(wordDocData, 154);
            int lcbStshf = BitConverter.ToInt32(wordDocData, 158);
            
            int fcSttbfffn = BitConverter.ToInt32(wordDocData, 266);
            int lcbSttbfffn = BitConverter.ToInt32(wordDocData, 270);

            Assert.NotEqual(0, fcStshf);
            Assert.True(lcbStshf > 0);
            Assert.NotEqual(0, fcSttbfffn);
            Assert.True(lcbSttbfffn > 0);

            // Basic check on Font Table (should start with fExtend = 0xFFFF)
            Assert.Equal(0xFF, tableData[fcSttbfffn]);
            Assert.Equal(0xFF, tableData[fcSttbfffn + 1]);
            
            // Count of fonts (2)
            Assert.Equal(2, BitConverter.ToUInt16(tableData, fcSttbfffn + 2));
        }
    }
}
