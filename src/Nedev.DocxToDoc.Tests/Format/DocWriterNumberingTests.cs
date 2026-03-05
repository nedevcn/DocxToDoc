using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Nedev.DocxToDoc.Format;
using Nedev.DocxToDoc.Model;
using Xunit;

namespace Nedev.DocxToDoc.Tests.Format
{
    public class DocWriterNumberingTests
    {
        [Fact]
        public void WriteDocBlocks_WithLists_WritesNumberingStructures()
        {
            // Register encoding provider for Windows-1252 
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Arrange
            var model = new DocumentModel();
            
            var absNum = new AbstractNumberingModel { Id = 100 };
            absNum.Levels.Add(new NumberingLevelModel { Level = 0, NumberFormat = "decimal", LevelText = "%1." });
            model.AbstractNumbering.Add(absNum);
            
            model.NumberingInstances.Add(new NumberingInstanceModel { Id = 1, AbstractNumberId = 100 });
 
            model.Content.Add(new ParagraphModel 
            { 
                Runs = { new RunModel { Text = "List Item 1" } },
                Properties = { NumberingId = 1, NumberingLevel = 0 }
            });
            model.Content.Add(new ParagraphModel 
            { 
                Runs = { new RunModel { Text = "List Item 2" } },
                Properties = { NumberingId = 1, NumberingLevel = 0 }
            });

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

            // FIB offsets for Numbering
            // fcSttbLst: index 10 (offset 154 + 10*8 = 234)
            // fcPlfLfo: index 53 (offset 154 + 53*8 = 578)
            
            int fcSttbLst = BitConverter.ToInt32(wordDocData, 234);
            int lcbSttbLst = BitConverter.ToInt32(wordDocData, 238);
            
            int fcPlfLfo = BitConverter.ToInt32(wordDocData, 578);
            int lcbPlfLfo = BitConverter.ToInt32(wordDocData, 582);

            Assert.NotEqual(0, fcSttbLst);
            Assert.True(lcbSttbLst > 0);
            Assert.NotEqual(0, fcPlfLfo);
            Assert.True(lcbPlfLfo > 0);

            // Verify SttbLst
            Assert.Equal(0xFFFF, BitConverter.ToUInt16(tableData, fcSttbLst)); // fExtend
            Assert.Equal(1, BitConverter.ToUInt16(tableData, fcSttbLst + 2)); // count

            // Verify PlfLfo
            Assert.Equal(1, BitConverter.ToInt32(tableData, fcPlfLfo)); // lLfo
            Assert.Equal(100, BitConverter.ToInt32(tableData, fcPlfLfo + 4)); // lsid of first instance
        }
    }
}
