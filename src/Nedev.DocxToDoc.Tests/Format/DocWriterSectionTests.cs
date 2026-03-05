using System;
using System.IO;
using System.Text;
using Nedev.DocxToDoc.Format;
using Nedev.DocxToDoc.Model;
using Xunit;

namespace Nedev.DocxToDoc.Tests.Format
{
    public class DocWriterSectionTests
    {
        [Fact]
        public void WriteDocBlocks_WithSections_WritesPlcfsed()
        {
            // Register encoding provider for Windows-1252 
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Arrange
            var model = new DocumentModel();
            
            model.Content.Add(new ParagraphModel { Runs = { new RunModel { Text = "SectionTest" } } });
            
            var section = new SectionModel
            {
                PageWidth = 12240,
                PageHeight = 15840,
                MarginLeft = 1440,
                MarginRight = 1440
            };
            model.Sections.Add(section);

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

            // Check FIB offset for Plcfsed (index 6 in RgFcLcb, offset 154 + 6*8 = 202)
            int fcPlcfsed = BitConverter.ToInt32(wordDocData, 202);
            int lcbPlcfsed = BitConverter.ToInt32(wordDocData, 206);

            Assert.NotEqual(0, fcPlcfsed);
            Assert.True(lcbPlcfsed >= 4 + 4 + 12); // CP0, CP_End, 1 SED (12 bytes)

            // Verify CP boundaries in Plcfsed (0 and 12)
            Assert.Equal(0, BitConverter.ToInt32(tableData, fcPlcfsed));
            Assert.Equal(12, BitConverter.ToInt32(tableData, fcPlcfsed + 4));

            // Verify SED (at fcPlcfsed + 8)
            // fn = 0
            Assert.Equal(0, BitConverter.ToInt16(tableData, fcPlcfsed + 8));
            // fcSep (offset in WordDocument)
            int fcSep = BitConverter.ToInt32(tableData, fcPlcfsed + 10);
            Assert.True(fcSep > 1536);

            // Verify SEP sprms in WordDocument at fcSep
            // cb (short)
            short cbSep = BitConverter.ToInt16(wordDocData, fcSep);
            Assert.True(cbSep > 0);
            
            // Check first sprm (sprmSXaPage = 0xB603)
            Assert.Equal(0x03, wordDocData[fcSep + 2]);
            Assert.Equal(0xB6, wordDocData[fcSep + 3]);
            Assert.Equal(12240, BitConverter.ToInt16(wordDocData, fcSep + 4));
        }
    }
}
