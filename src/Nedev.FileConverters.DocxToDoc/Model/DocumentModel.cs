using System.Collections.Generic;

namespace Nedev.FileConverters.DocxToDoc.Model
{
    /// <summary>
    /// Represents the parsed content of a document, ready for serialization to binary MS-DOC format.
    /// </summary>
    public class DocumentModel
    {
        public List<object> Content { get; } = new List<object>(); // Can be ParagraphModel or TableModel
        public List<ParagraphModel> Paragraphs { get; } = new List<ParagraphModel>(); // Keeping for compatibility? No, let's refactor.
        public List<StyleModel> Styles { get; } = new List<StyleModel>();
        public List<FontModel> Fonts { get; } = new List<FontModel>();
        public List<SectionModel> Sections { get; } = new List<SectionModel>();
        public List<AbstractNumberingModel> AbstractNumbering { get; } = new List<AbstractNumberingModel>();
        public List<NumberingInstanceModel> NumberingInstances { get; } = new List<NumberingInstanceModel>();
        
        // As the Document is parsed we will accumulate the plain text here
        // The length of this text will determine the Piece Table and CCP Text
        public string TextBuffer { get; set; } = string.Empty;
    }

    public class SectionModel
    {
        public int PageWidth { get; set; } = 11906; // Default Letter/A4
        public int PageHeight { get; set; } = 16838;
        public int MarginLeft { get; set; } = 1440; 
        public int MarginRight { get; set; } = 1440;
        public int MarginTop { get; set; } = 1440;
        public int MarginBottom { get; set; } = 1440;
    }

    public class AbstractNumberingModel
    {
        public int Id { get; set; }
        public List<NumberingLevelModel> Levels { get; } = new List<NumberingLevelModel>();
    }

    public class NumberingLevelModel
    {
        public int Level { get; set; }
        public string NumberFormat { get; set; } = "decimal";
        public string LevelText { get; set; } = string.Empty;
        public int Start { get; set; } = 1;
    }

    public class NumberingInstanceModel
    {
        public int Id { get; set; }
        public int AbstractNumberId { get; set; }
    }

    public class StyleModel
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public bool IsParagraphStyle { get; set; }
    }

    public class FontModel
    {
        public string Name { get; set; } = string.Empty;
    }

    public class TableModel
    {
        public List<TableRowModel> Rows { get; } = new List<TableRowModel>();
    }

    public class TableRowModel
    {
        public List<TableCellModel> Cells { get; } = new List<TableCellModel>();
    }

    public class TableCellModel
    {
        public List<ParagraphModel> Paragraphs { get; } = new List<ParagraphModel>();
        public int Width { get; set; }
    }

    public class ParagraphModel
    {
        public List<RunModel> Runs { get; } = new List<RunModel>();
        public ParagraphProperties Properties { get; } = new ParagraphProperties();

        public class ParagraphProperties
        {
            public Justification Alignment { get; set; } = Justification.Left;
            public int? NumberingId { get; set; }
            public int? NumberingLevel { get; set; }
        }

        public enum Justification
        {
            Left, Center, Right, Both
        }
    }

    public class RunModel
    {
        public string Text { get; set; } = string.Empty;
        public CharacterProperties Properties { get; } = new CharacterProperties();

        public class CharacterProperties
        {
            public bool IsBold { get; set; }
            public bool IsItalic { get; set; }
            public bool IsStrike { get; set; }
            public int? FontSize { get; set; } // In half-points (e.g., 24 = 12pt)
        }
    }
}
