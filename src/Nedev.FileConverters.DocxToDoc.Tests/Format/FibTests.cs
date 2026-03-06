using System.IO;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class FibTests
    {
        [Fact]
        public void Fib_WriteTo_ProducesCorrectBaseStructure()
        {
            // Arrange
            var fib = new Nedev.FileConverters.DocxToDoc.Format.Fib();
            using var ms = new MemoryStream();
            using var writer = new BinaryWriter(ms);

            // Act
            fib.WriteTo(writer);
            byte[] result = ms.ToArray();

            // Assert
            // Base FIB size is 32 + FibRgW97 (28) + FibRgLw97 (88) = 148 bytes minimum
            Assert.True(result.Length >= 148);

            // FibBase
            using var reader = new BinaryReader(new MemoryStream(result));
            Assert.Equal(0xA5EC, reader.ReadUInt16()); // wIdent
            Assert.Equal(0x00C1, reader.ReadUInt16()); // nFib
            Assert.Equal(0x0000, reader.ReadUInt16()); // unused
            Assert.Equal(0x0409, reader.ReadUInt16()); // lid
            Assert.Equal(0, reader.ReadInt16());      // pnNext
            Assert.Equal(0x0060, reader.ReadUInt16()); // fFlags
            
            // Skip rest of base
            reader.BaseStream.Position = 32;

            // FibRgW97
            Assert.Equal(14, reader.ReadUInt16()); // csw
            
            // Skip FibRgW97
            reader.BaseStream.Position = 32 + 2 + 28;

            // FibRgLw97
            Assert.Equal(22, reader.ReadUInt16()); // cslw
        }
    }
}
