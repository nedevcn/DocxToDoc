using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Text;
using System.Collections.Generic;

namespace Nedev.DocxToDoc.Format
{
    /// <summary>
    /// Reads and coordinates the extraction of features from OpenXML (.docx) files.
    /// Optimized for low-memory, forward-only reading.
    /// </summary>
    public class DocxReader : IDisposable
    {
        private readonly ZipArchive _archive;
        private bool _disposedValue;

        public DocxReader(Stream docxStream)
        {
            // We assume the caller manages the base stream's lifecycle.
            _archive = new ZipArchive(docxStream, ZipArchiveMode.Read, leaveOpen: true);
        }

        public Nedev.DocxToDoc.Model.DocumentModel ReadDocument()
        {
            var documentEntry = _archive.GetEntry("word/document.xml");
            if (documentEntry == null)
            {
                throw new FileNotFoundException("word/document.xml not found in the docx file.");
            }

            using var stream = documentEntry.Open();
            using var xmlReader = XmlReader.Create(stream, new XmlReaderSettings 
            { 
                IgnoreComments = true, 
                IgnoreWhitespace = true 
            });

            var docModel = new Nedev.DocxToDoc.Model.DocumentModel();
            
            // Parse Styles
            var stylesEntry = _archive.GetEntry("word/styles.xml");
            if (stylesEntry != null)
            {
                using var stylesStream = stylesEntry.Open();
                ParseStyles(stylesStream, docModel);
            }

            // Parse Font Table
            var fontTableEntry = _archive.GetEntry("word/fontTable.xml");
            if (fontTableEntry != null)
            {
                using var fontsStream = fontTableEntry.Open();
                ParseFonts(fontsStream, docModel);
            }

            // Parse Numbering
            var numberingEntry = _archive.GetEntry("word/numbering.xml");
            if (numberingEntry != null)
            {
                using var numberingStream = numberingEntry.Open();
                ParseNumbering(numberingStream, docModel);
            }

            StringBuilder textBuffer = new StringBuilder();

            Nedev.DocxToDoc.Model.ParagraphModel currentParagraph = null;
            Nedev.DocxToDoc.Model.RunModel currentRun = null;
            Nedev.DocxToDoc.Model.SectionModel currentSection = null;
            Nedev.DocxToDoc.Model.TableModel currentTable = null;
            Nedev.DocxToDoc.Model.TableRowModel currentRow = null;
            Nedev.DocxToDoc.Model.TableCellModel currentCell = null;

            while (xmlReader.Read())
            {
                if (xmlReader.NodeType == XmlNodeType.Element)
                {
                    string localName = xmlReader.LocalName;
                    if (localName == "tbl")
                    {
                        currentTable = new Nedev.DocxToDoc.Model.TableModel();
                        docModel.Content.Add(currentTable);
                    }
                    else if (localName == "tr" && currentTable != null)
                    {
                        currentRow = new Nedev.DocxToDoc.Model.TableRowModel();
                        currentTable.Rows.Add(currentRow);
                    }
                    else if (localName == "tc" && currentRow != null)
                    {
                        currentCell = new Nedev.DocxToDoc.Model.TableCellModel();
                        currentRow.Cells.Add(currentCell);
                    }
                    else if (localName == "p")
                    {
                        currentParagraph = new Nedev.DocxToDoc.Model.ParagraphModel();
                        if (currentCell != null) currentCell.Paragraphs.Add(currentParagraph);
                        else 
                        {
                            docModel.Content.Add(currentParagraph);
                            docModel.Paragraphs.Add(currentParagraph); // Backwards compatibility? No, needed for textBuffer logic maybe.
                        }
                    }
                    else if (localName == "numId" && currentParagraph != null)
                    {
                        if (int.TryParse(xmlReader.GetAttribute("w:val"), out int id))
                            currentParagraph.Properties.NumberingId = id;
                    }
                    else if (localName == "ilvl" && currentParagraph != null)
                    {
                        if (int.TryParse(xmlReader.GetAttribute("w:val"), out int lvl))
                            currentParagraph.Properties.NumberingLevel = lvl;
                    }
                    else if (localName == "sectPr")
                    {
                        currentSection = new Nedev.DocxToDoc.Model.SectionModel();
                        docModel.Sections.Add(currentSection);
                    }
                    else if (localName == "pgSz" && currentSection != null)
                    {
                        if (int.TryParse(xmlReader.GetAttribute("w:w"), out int w)) currentSection.PageWidth = w;
                        if (int.TryParse(xmlReader.GetAttribute("w:h"), out int h)) currentSection.PageHeight = h;
                    }
                    else if (localName == "pgMar" && currentSection != null)
                    {
                        if (int.TryParse(xmlReader.GetAttribute("w:left"), out int l)) currentSection.MarginLeft = l;
                        if (int.TryParse(xmlReader.GetAttribute("w:right"), out int r)) currentSection.MarginRight = r;
                        if (int.TryParse(xmlReader.GetAttribute("w:top"), out int t)) currentSection.MarginTop = t;
                        if (int.TryParse(xmlReader.GetAttribute("w:bottom"), out int b)) currentSection.MarginBottom = b;
                    }
                    else if (localName == "jc" && currentParagraph != null)
                    {
                        string val = xmlReader.GetAttribute("w:val");
                        currentParagraph.Properties.Alignment = val switch
                        {
                            "center" => Nedev.DocxToDoc.Model.ParagraphModel.Justification.Center,
                            "right" => Nedev.DocxToDoc.Model.ParagraphModel.Justification.Right,
                            "both" => Nedev.DocxToDoc.Model.ParagraphModel.Justification.Both,
                            _ => Nedev.DocxToDoc.Model.ParagraphModel.Justification.Left
                        };
                    }
                    else if (localName == "r" && currentParagraph != null)
                    {
                        currentRun = new Nedev.DocxToDoc.Model.RunModel();
                        currentParagraph.Runs.Add(currentRun);
                    }
                    else if (localName == "b" && currentRun != null)
                    {
                        string val = xmlReader.GetAttribute("w:val");
                        currentRun.Properties.IsBold = (val != "0" && val != "false");
                    }
                    else if (localName == "i" && currentRun != null)
                    {
                        string val = xmlReader.GetAttribute("w:val");
                        currentRun.Properties.IsItalic = (val != "0" && val != "false");
                    }
                    else if (localName == "strike" && currentRun != null)
                    {
                        string val = xmlReader.GetAttribute("w:val");
                        currentRun.Properties.IsStrike = (val != "0" && val != "false");
                    }
                    else if (localName == "sz" && currentRun != null)
                    {
                        string val = xmlReader.GetAttribute("w:val");
                        if (int.TryParse(val, out int size))
                        {
                            currentRun.Properties.FontSize = size;
                        }
                    }
                    else if (localName == "t" && currentRun != null)
                    {
                        string text = xmlReader.ReadElementContentAsString();
                        currentRun.Text = text;
                        textBuffer.Append(text);
                    }
                }
                else if (xmlReader.NodeType == XmlNodeType.EndElement)
                {
                    string localName = xmlReader.LocalName;
                    if (localName == "tbl") currentTable = null;
                    else if (localName == "tr") currentRow = null;
                    else if (localName == "tc") currentCell = null;
                    else if (localName == "p")
                    {
                        textBuffer.Append('\r');
                        currentParagraph = null;
                    }
                    else if (localName == "r")
                    {
                        currentRun = null;
                    }
                }
            }

            docModel.TextBuffer = textBuffer.ToString();
            return docModel;
        }

        private void ParseNumbering(Stream numberingStream, Nedev.DocxToDoc.Model.DocumentModel docModel)
        {
            using var reader = XmlReader.Create(numberingStream, new XmlReaderSettings { IgnoreWhitespace = true });
            Nedev.DocxToDoc.Model.AbstractNumberingModel currentAbstract = null;
            Nedev.DocxToDoc.Model.NumberingLevelModel currentLevel = null;
            Nedev.DocxToDoc.Model.NumberingInstanceModel currentInstance = null;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    string localName = reader.LocalName;
                    if (localName == "abstractNum")
                    {
                        currentAbstract = new Nedev.DocxToDoc.Model.AbstractNumberingModel
                        {
                            Id = int.Parse(reader.GetAttribute("w:abstractNumId") ?? "0")
                        };
                        docModel.AbstractNumbering.Add(currentAbstract);
                    }
                    else if (localName == "lvl" && currentAbstract != null)
                    {
                        currentLevel = new Nedev.DocxToDoc.Model.NumberingLevelModel
                        {
                            Level = int.Parse(reader.GetAttribute("w:ilvl") ?? "0")
                        };
                        currentAbstract.Levels.Add(currentLevel);
                    }
                    else if (localName == "start" && currentLevel != null)
                    {
                        currentLevel.Start = int.Parse(reader.GetAttribute("w:val") ?? "1");
                    }
                    else if (localName == "numFmt" && currentLevel != null)
                    {
                        currentLevel.NumberFormat = reader.GetAttribute("w:val") ?? "decimal";
                    }
                    else if (localName == "lvlText" && currentLevel != null)
                    {
                        currentLevel.LevelText = reader.GetAttribute("w:val") ?? string.Empty;
                    }
                    else if (localName == "num")
                    {
                        currentInstance = new Nedev.DocxToDoc.Model.NumberingInstanceModel
                        {
                            Id = int.Parse(reader.GetAttribute("w:numId") ?? "0")
                        };
                        docModel.NumberingInstances.Add(currentInstance);
                    }
                    else if (localName == "abstractNumId" && currentInstance != null)
                    {
                        currentInstance.AbstractNumberId = int.Parse(reader.GetAttribute("w:val") ?? "0");
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement)
                {
                    if (reader.LocalName == "abstractNum") currentAbstract = null;
                    else if (reader.LocalName == "lvl") currentLevel = null;
                    else if (reader.LocalName == "num") currentInstance = null;
                }
            }
        }

        private void ParseStyles(Stream stylesStream, Nedev.DocxToDoc.Model.DocumentModel docModel)
        {
            using var reader = XmlReader.Create(stylesStream, new XmlReaderSettings { IgnoreWhitespace = true });
            Nedev.DocxToDoc.Model.StyleModel currentStyle = null;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.LocalName == "style")
                    {
                        currentStyle = new Nedev.DocxToDoc.Model.StyleModel
                        {
                            Id = reader.GetAttribute("w:styleId") ?? string.Empty,
                            IsParagraphStyle = reader.GetAttribute("w:type") == "paragraph"
                        };
                        docModel.Styles.Add(currentStyle);
                    }
                    else if (reader.LocalName == "name" && currentStyle != null)
                    {
                        currentStyle.Name = reader.GetAttribute("w:val") ?? string.Empty;
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "style")
                {
                    currentStyle = null;
                }
            }
        }

        private void ParseFonts(Stream fontStream, Nedev.DocxToDoc.Model.DocumentModel docModel)
        {
            using var reader = XmlReader.Create(fontStream, new XmlReaderSettings { IgnoreWhitespace = true });
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "font")
                {
                    string name = reader.GetAttribute("w:name");
                    if (!string.IsNullOrEmpty(name))
                    {
                        docModel.Fonts.Add(new Nedev.DocxToDoc.Model.FontModel { Name = name });
                    }
                }
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposedValue)
            {
                if (disposing)
                {
                    _archive?.Dispose();
                }
                _disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
