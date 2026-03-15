using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Text;
using System.Collections.Generic;

namespace Nedev.FileConverters.DocxToDoc
{
    /// <summary>
    /// Provides streaming-based reading of DOCX files for memory-efficient processing of large documents.
    /// </summary>
    public class StreamingDocxReader : IDisposable
    {
        private readonly Stream _stream;
        private ZipArchive? _archive;
        private bool _disposed;

        /// <summary>
        /// Initializes a new instance of the <see cref="StreamingDocxReader"/> class.
        /// </summary>
        /// <param name="stream">The stream containing the DOCX file.</param>
        public StreamingDocxReader(Stream stream)
        {
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        }

        /// <summary>
        /// Opens the DOCX archive for reading.
        /// </summary>
        public void Open()
        {
            if (_archive != null)
                throw new InvalidOperationException("Archive is already open.");

            _archive = new ZipArchive(_stream, ZipArchiveMode.Read, leaveOpen: true);
        }

        /// <summary>
        /// Reads document properties without loading the entire document.
        /// </summary>
        public Model.DocumentProperties? ReadProperties()
        {
            EnsureOpen();

            var propsEntry = _archive!.GetEntry("docProps/core.xml");
            if (propsEntry == null)
                return null;

            var props = new Model.DocumentProperties();
            using var stream = propsEntry.Open();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true });

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    string localName = reader.LocalName;

                    switch (localName)
                    {
                        case "title":
                            props.Title = reader.ReadElementContentAsString();
                            break;
                        case "subject":
                            props.Subject = reader.ReadElementContentAsString();
                            break;
                        case "creator":
                        case "author":
                            props.Author = reader.ReadElementContentAsString();
                            break;
                        case "keywords":
                            props.Keywords = reader.ReadElementContentAsString();
                            break;
                        case "description":
                        case "comments":
                            props.Comments = reader.ReadElementContentAsString();
                            break;
                        case "created":
                            if (DateTime.TryParse(reader.ReadElementContentAsString(), out DateTime created))
                                props.Created = created;
                            break;
                        case "modified":
                            if (DateTime.TryParse(reader.ReadElementContentAsString(), out DateTime modified))
                                props.Modified = modified;
                            break;
                        case "revision":
                            if (int.TryParse(reader.ReadElementContentAsString(), out int revision))
                                props.Revision = revision;
                            break;
                        case "category":
                            props.Category = reader.ReadElementContentAsString();
                            break;
                        case "company":
                            props.Company = reader.ReadElementContentAsString();
                            break;
                    }
                }
            }

            return props;
        }

        /// <summary>
        /// Enumerates paragraphs in a streaming fashion without loading the entire document.
        /// </summary>
        public IEnumerable<Model.ParagraphModel> EnumerateParagraphs()
        {
            EnsureOpen();

            var docEntry = _archive!.GetEntry("word/document.xml");
            if (docEntry == null)
                yield break;

            using var stream = docEntry.Open();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                IgnoreComments = true,
                IgnoreWhitespace = true
            });

            Model.ParagraphModel? currentParagraph = null;
            Model.RunModel? currentRun = null;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    string localName = reader.LocalName;

                    if (localName == "p")
                    {
                        currentParagraph = new Model.ParagraphModel();
                    }
                    else if (localName == "r" && currentParagraph != null)
                    {
                        currentRun = new Model.RunModel();
                        currentParagraph.Runs.Add(currentRun);
                    }
                    else if (localName == "t" && currentRun != null)
                    {
                        string text = reader.ReadElementContentAsString();
                        currentRun.Text = text;
                    }
                    else if (localName == "tab" && currentRun != null)
                    {
                        currentRun.Text += "\t";
                    }
                    else if (localName == "br" && currentRun != null)
                    {
                        currentRun.Text += "\n";
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement)
                {
                    if (reader.LocalName == "p" && currentParagraph != null)
                    {
                        yield return currentParagraph;
                        currentParagraph = null;
                        currentRun = null;
                    }
                }
            }
        }

        /// <summary>
        /// Gets the count of paragraphs without loading them all into memory.
        /// </summary>
        public int GetParagraphCount()
        {
            EnsureOpen();

            var docEntry = _archive!.GetEntry("word/document.xml");
            if (docEntry == null)
                return 0;

            int count = 0;
            using var stream = docEntry.Open();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                IgnoreComments = true,
                IgnoreWhitespace = true
            });

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "p")
                {
                    count++;
                }
            }

            return count;
        }

        /// <summary>
        /// Gets the approximate size of the document content.
        /// </summary>
        public long GetDocumentSize()
        {
            EnsureOpen();

            var docEntry = _archive!.GetEntry("word/document.xml");
            return docEntry?.Length ?? 0;
        }

        /// <summary>
        /// Checks if the document contains images.
        /// </summary>
        public bool HasImages()
        {
            EnsureOpen();

            foreach (var entry in _archive!.Entries)
            {
                if (entry.FullName.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Gets the count of images in the document.
        /// </summary>
        public int GetImageCount()
        {
            EnsureOpen();

            int count = 0;
            foreach (var entry in _archive!.Entries)
            {
                if (entry.FullName.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase))
                    count++;
            }

            return count;
        }

        /// <summary>
        /// Checks if the document has comments.
        /// </summary>
        public bool HasComments()
        {
            EnsureOpen();
            return _archive!.GetEntry("word/comments.xml") != null;
        }

        /// <summary>
        /// Gets the count of comments in the document.
        /// </summary>
        public int GetCommentCount()
        {
            EnsureOpen();

            var commentsEntry = _archive!.GetEntry("word/comments.xml");
            if (commentsEntry == null)
                return 0;

            int count = 0;
            using var stream = commentsEntry.Open();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true });

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "comment")
                {
                    count++;
                }
            }

            return count;
        }

        private void EnsureOpen()
        {
            if (_archive == null)
                throw new InvalidOperationException("Archive is not open. Call Open() first.");
        }

        /// <summary>
        /// Releases all resources used by the streaming reader.
        /// </summary>
        public void Dispose()
        {
            if (!_disposed)
            {
                _archive?.Dispose();
                _disposed = true;
            }
        }
    }
}
