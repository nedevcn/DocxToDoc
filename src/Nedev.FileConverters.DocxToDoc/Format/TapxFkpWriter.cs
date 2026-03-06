using System;
using System.Collections.Generic;
using System.IO;

namespace Nedev.FileConverters.DocxToDoc.Format
{
    /// <summary>
    /// Helps pack Table Properties (TAPX) into a Formatted Disk Page (FKP).
    /// </summary>
    public class TapxFkpWriter
    {
        private const int PageSize = 512;
        
        private readonly List<int> _rgfc = new();
        private readonly List<byte[]> _rgbx = new();

        public void AddRow(int startCp, int endCp, byte[] sprms)
        {
            if (_rgfc.Count == 0 || _rgfc[^1] != startCp)
            {
                _rgfc.Add(startCp);
            }
            _rgfc.Add(endCp);
            _rgbx.Add(sprms);
        }

        public byte[] GeneratePage()
        {
            byte[] page = new byte[PageSize];
            
            if (_rgfc.Count == 0) return page;

            byte cRun = (byte)(_rgfc.Count - 1);
            
            // Last byte of page is cRun
            page[PageSize - 1] = cRun;

            // Write rgfc (cRun + 1 ints)
            int offset = 0;
            foreach (var fc in _rgfc)
            {
                BitConverter.GetBytes(fc).CopyTo(page, offset);
                offset += 4;
            }

            // In TAPX FKP, following rgfc is an array of bx (offset)
            // bx is 1 byte (offset to TAPX in the page / 2? No, for TAPX it might be different.)
            // Actually, for TAPX, the offset is also 1 byte (divided by 2).
            
            int currentTail = PageSize - 1; 
            
            for (int i = 0; i < cRun; i++)
            {
                byte[] sprm = _rgbx[i];
                
                // TAPX structure in FKP: [cb (2 bytes)] [grpprl]
                int grpprlLen = sprm?.Length ?? 0;
                int tapxSize = 2 + grpprlLen; 
                
                // Align to 2-byte boundary
                if (tapxSize % 2 != 0) tapxSize++;

                currentTail -= tapxSize;
                
                // cb (2 bytes)
                BitConverter.GetBytes((ushort)grpprlLen).CopyTo(page, currentTail);

                if (sprm != null)
                {
                    Array.Copy(sprm, 0, page, currentTail + 2, sprm.Length);
                }
                
                // Offset to TAPX is at (rgfc.Count * 4) + i
                page[offset + i] = (byte)(currentTail / 2);
            }

            return page;
        }
    }
}
