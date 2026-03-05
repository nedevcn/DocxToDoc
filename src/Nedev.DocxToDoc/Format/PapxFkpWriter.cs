using System;
using System.Collections.Generic;
using System.IO;

namespace Nedev.DocxToDoc.Format
{
    /// <summary>
    /// Helps pack Paragraph Properties (PAPX) into a Formatted Disk Page (FKP).
    /// </summary>
    public class PapxFkpWriter
    {
        private const int PageSize = 512;
        
        private readonly List<int> _rgfc = new();
        private readonly List<byte[]> _rgbx = new();

        public void AddParagraph(int startCp, int endCp, byte[] sprms)
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

            // In PAPX FKP, following rgfc is an array of bx (offset + ists)
            // For simplicity, we assume ists = 0 (Normal style)
            // bx is 1 byte (offset to PAPX in the page / 2)
            
            int currentTail = PageSize - 1; // End of page
            
            for (int i = 0; i < cRun; i++)
            {
                byte[] sprm = _rgbx[i];
                
                // PAPX = cw (1 byte) + stsh index (2 bytes) + grpprl
                // Actually, PAPX is stored as [cb] [grpprl] where cb is the total length including itself.
                // But in a more specific sense for FKPs, it's [cb] [ISTS] [grpprl]
                // ISTS is 2 bytes (index to stylesheet).
                
                int grpprlLen = sprm?.Length ?? 0;
                int papxSize = 1 + 2 + grpprlLen; // cb(1) + ists(2) + grpprl
                
                // Align to 2-byte boundary
                if (papxSize % 2 != 0) papxSize++;

                currentTail -= papxSize;
                
                page[currentTail] = (byte)(papxSize / 2); // Internal cb is in words? No, it's usually bytes.
                // Wait, per spec: "cb (1 byte): count of words of the property exceptions"
                // Let's use words to be safe.
                page[currentTail] = (byte)((papxSize - 1) / 2); // Exclude cb itself, in words.
                
                // ISTS = 0 (Normal)
                page[currentTail + 1] = 0;
                page[currentTail + 2] = 0;

                if (sprm != null)
                {
                    Array.Copy(sprm, 0, page, currentTail + 3, sprm.Length);
                }
                
                // Offset to PAPX is at (rgfc.Count * 4) + i
                page[offset + i] = (byte)(currentTail / 2);
            }

            return page;
        }
    }
}
