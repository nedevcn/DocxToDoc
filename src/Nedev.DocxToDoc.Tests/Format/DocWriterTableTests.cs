using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Nedev.DocxToDoc.Format;
using Nedev.DocxToDoc.Model;
using Xunit;

namespace Nedev.DocxToDoc.Tests.Format
{
    public class DocWriterTableTests
    {
        [Fact]
        public void WriteDocBlocks_WithTable_WritesTableMarkersAndTapx()
        {
            // Register encoding provider for Windows-1252 
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Arrange
            var model = new DocumentModel();
            var table = new TableModel();
            var row = new TableRowModel();
            
            var cell1 = new TableCellModel { Width = 5000 };
            cell1.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Cell 1" } } });
            
            var cell2 = new TableCellModel { Width = 5000 };
            cell2.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Cell 2" } } });
            
            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            table.Rows.Add(row);
            
            model.Content.Add(table);

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            // Act
            try
            {
                writer.WriteDocBlocks(model, ms);
                ms.Position = 0;

                // Assert
                using var compoundFile = new OpenMcdf.CompoundFile(ms);
                Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
                Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));
                
                byte[] tableData = tableStream.GetData();
                byte[] wordDocData = wordDocStream.GetData();

                // Check if cell markers (ASCII 7) are present in the text
                string text = Encoding.GetEncoding(1252).GetString(wordDocData, 1536, (int)wordDocData.Length - 1536);
                Assert.Contains("Cell 1\r\x0007Cell 2\r\x0007\r", text);

                // FIB offsets for Table (TAPX)
                int fcPlcfbteTapx = BitConverter.ToInt32(wordDocData, 186);
                int lcbPlcfbteTapx = BitConverter.ToInt32(wordDocData, 190);

                Assert.NotEqual(0, fcPlcfbteTapx);
                Assert.True(lcbPlcfbteTapx > 0);
            }
            catch (Exception ex)
            {
                throw new Exception($"Test failed with error: {ex.Message}\nStack: {ex.StackTrace}");
            }
            
            // Verify PlcfTapx exists in 1Table at fcPlcfbteTapx
            // (Assuming it was written to 1Table correctly)
        }

        private bool IsWord97Format(byte[] data)
        {
            return BitConverter.ToUInt16(data, 0) == 0xA5EC;
        }
    }
}
