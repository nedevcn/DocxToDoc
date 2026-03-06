using System;
using System.Collections.Generic;
using System.IO;
using OpenMcdf;

namespace Nedev.FileConverters.DocxToDoc.Format
{
    /// <summary>
    /// Writes Compound File Binary (OLE2) format streams into a single structure.
    /// Essential for MS-DOC binary files.
    /// </summary>
    public class CfbWriter : IDisposable
    {
        private readonly CompoundFile _compoundFile;

        public CfbWriter()
        {
            _compoundFile = new CompoundFile();
        }

        public void AddStream(string name, byte[] data)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentNullException(nameof(name));
            if (data == null)
                throw new ArgumentNullException(nameof(data));

            var stream = _compoundFile.RootStorage.AddStream(name);
            stream.SetData(data);
        }

        public void WriteTo(Stream outputStream)
        {
            if (outputStream == null)
                throw new ArgumentNullException(nameof(outputStream));

            _compoundFile.Save(outputStream);
        }

        public void Dispose()
        {
            _compoundFile?.Close();
        }
    }
}
