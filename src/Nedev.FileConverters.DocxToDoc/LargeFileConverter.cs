using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Nedev.FileConverters.DocxToDoc
{
    /// <summary>
    /// Provides specialized conversion for large DOCX files with memory-efficient processing.
    /// </summary>
    public class LargeFileConverter
    {
        private readonly ILogger _logger;
        private readonly long _largeFileThreshold;

        /// <summary>
        /// Initializes a new instance of the <see cref="LargeFileConverter"/> class.
        /// </summary>
        /// <param name="largeFileThreshold">The file size threshold in bytes above which to use streaming mode. Default is 10MB.</param>
        public LargeFileConverter(long largeFileThreshold = 10 * 1024 * 1024) : this(NullLogger.Instance, largeFileThreshold)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="LargeFileConverter"/> class with a logger.
        /// </summary>
        /// <param name="logger">The logger to use for recording operations.</param>
        /// <param name="largeFileThreshold">The file size threshold in bytes above which to use streaming mode.</param>
        public LargeFileConverter(ILogger logger, long largeFileThreshold = 10 * 1024 * 1024)
        {
            _logger = logger ?? NullLogger.Instance;
            _largeFileThreshold = largeFileThreshold;
        }

        /// <summary>
        /// Analyzes a DOCX file and returns information about its contents without loading it fully.
        /// </summary>
        /// <param name="docxPath">The path to the DOCX file.</param>
        /// <returns>Information about the document.</returns>
        public DocumentInfo AnalyzeDocument(string docxPath)
        {
            if (string.IsNullOrWhiteSpace(docxPath))
                throw new ArgumentNullException(nameof(docxPath));

            if (!File.Exists(docxPath))
                throw new FileNotFoundException($"File not found: {docxPath}", docxPath);

            _logger.LogInfo($"Analyzing document: {docxPath}");

            var fileInfo = new FileInfo(docxPath);
            var info = new DocumentInfo
            {
                FilePath = docxPath,
                FileSize = fileInfo.Length,
                IsLargeFile = fileInfo.Length > _largeFileThreshold
            };

            using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var reader = new StreamingDocxReader(stream);
            reader.Open();

            // Read properties
            var props = reader.ReadProperties();
            if (props != null)
            {
                info.Title = props.Title;
                info.Author = props.Author;
                info.PageCount = props.Pages;
                info.WordCount = props.Words;
            }

            // Get statistics
            info.ParagraphCount = reader.GetParagraphCount();
            info.ImageCount = reader.GetImageCount();
            info.HasImages = reader.HasImages();
            info.HasComments = reader.HasComments();
            info.CommentCount = reader.GetCommentCount();
            info.DocumentXmlSize = reader.GetDocumentSize();

            _logger.LogInfo($"Document analysis complete: {info.ParagraphCount} paragraphs, {info.ImageCount} images");

            return info;
        }

        /// <summary>
        /// Converts a large DOCX file using streaming mode for memory efficiency.
        /// </summary>
        /// <param name="docxPath">The path to the source DOCX file.</param>
        /// <param name="docPath">The path to the destination DOC file.</param>
        /// <param name="options">Conversion options for large files.</param>
        public void ConvertLargeFile(string docxPath, string docPath, LargeFileConversionOptions? options = null)
        {
            options ??= new LargeFileConversionOptions();

            _logger.LogInfo($"Converting large file: {docxPath}");

            var info = AnalyzeDocument(docxPath);

            if (!info.IsLargeFile && !options.ForceStreamingMode)
            {
                _logger.LogInfo("File size below threshold, using standard conversion");
                var standardConverter = new DocxToDocConverter(_logger);
                standardConverter.Convert(docxPath, docPath);
                return;
            }

            _logger.LogInfo($"Using streaming mode for file: {info.FileSize:N0} bytes");

            // For very large files, we might want to process in chunks
            // This is a simplified implementation that uses the standard converter
            // but with memory optimization

            using var monitor = new PerformanceMonitor(_logger);
            monitor.Start();

            try
            {
                // Pre-conversion memory cleanup
                if (options.EnableGarbageCollection)
                {
                    MemoryMonitor.CollectAndGetManagedMemory();
                    _logger.LogDebug("Pre-conversion garbage collection completed");
                }

                // Perform conversion
                var converter = new DocxToDocConverter(_logger);
                converter.Convert(docxPath, docPath);

                // Post-conversion cleanup
                if (options.EnableGarbageCollection)
                {
                    MemoryMonitor.CollectAndGetManagedMemory();
                    _logger.LogDebug("Post-conversion garbage collection completed");
                }

                monitor.Stop();
                monitor.LogSummary();

                _logger.LogInfo($"Large file conversion completed: {docPath}");
            }
            catch (Exception ex)
            {
                _logger.LogError($"Large file conversion failed: {docxPath}", ex);
                throw new ConversionException(
                    $"Failed to convert large file '{docxPath}'",
                    docxPath,
                    docPath,
                    ConversionStage.Unknown,
                    ex);
            }
        }

        /// <summary>
        /// Asynchronously converts a large DOCX file.
        /// </summary>
        public async Task ConvertLargeFileAsync(string docxPath, string docPath, 
            LargeFileConversionOptions? options = null, CancellationToken cancellationToken = default)
        {
            options ??= new LargeFileConversionOptions();

            await Task.Run(() => ConvertLargeFile(docxPath, docPath, options), cancellationToken);
        }

        /// <summary>
        /// Determines if a file should be treated as a large file.
        /// </summary>
        public bool IsLargeFile(string docxPath)
        {
            if (!File.Exists(docxPath))
                return false;

            var fileInfo = new FileInfo(docxPath);
            return fileInfo.Length > _largeFileThreshold;
        }
    }

    /// <summary>
    /// Contains information about a document.
    /// </summary>
    public class DocumentInfo
    {
        /// <summary>
        /// Gets or sets the file path.
        /// </summary>
        public string? FilePath { get; set; }

        /// <summary>
        /// Gets or sets the file size in bytes.
        /// </summary>
        public long FileSize { get; set; }

        /// <summary>
        /// Gets or sets whether this is considered a large file.
        /// </summary>
        public bool IsLargeFile { get; set; }

        /// <summary>
        /// Gets or sets the document title.
        /// </summary>
        public string? Title { get; set; }

        /// <summary>
        /// Gets or sets the document author.
        /// </summary>
        public string? Author { get; set; }

        /// <summary>
        /// Gets or sets the number of pages.
        /// </summary>
        public int? PageCount { get; set; }

        /// <summary>
        /// Gets or sets the word count.
        /// </summary>
        public int? WordCount { get; set; }

        /// <summary>
        /// Gets or sets the number of paragraphs.
        /// </summary>
        public int ParagraphCount { get; set; }

        /// <summary>
        /// Gets or sets the number of images.
        /// </summary>
        public int ImageCount { get; set; }

        /// <summary>
        /// Gets or sets whether the document has images.
        /// </summary>
        public bool HasImages { get; set; }

        /// <summary>
        /// Gets or sets whether the document has comments.
        /// </summary>
        public bool HasComments { get; set; }

        /// <summary>
        /// Gets or sets the number of comments.
        /// </summary>
        public int CommentCount { get; set; }

        /// <summary>
        /// Gets or sets the size of document.xml in bytes.
        /// </summary>
        public long DocumentXmlSize { get; set; }

        /// <summary>
        /// Gets a summary string of the document information.
        /// </summary>
        public string GetSummary()
        {
            return $"File: {Path.GetFileName(FilePath)}\n" +
                   $"Size: {FileSize:N0} bytes ({FileSize / 1024 / 1024} MB)\n" +
                   $"Type: {(IsLargeFile ? "Large" : "Standard")}\n" +
                   $"Title: {Title ?? "N/A"}\n" +
                   $"Author: {Author ?? "N/A"}\n" +
                   $"Pages: {PageCount?.ToString() ?? "N/A"}\n" +
                   $"Words: {WordCount?.ToString() ?? "N/A"}\n" +
                   $"Paragraphs: {ParagraphCount}\n" +
                   $"Images: {ImageCount}\n" +
                   $"Comments: {CommentCount}";
        }
    }

    /// <summary>
    /// Options for large file conversion.
    /// </summary>
    public class LargeFileConversionOptions
    {
        /// <summary>
        /// Gets or sets whether to force streaming mode regardless of file size.
        /// </summary>
        public bool ForceStreamingMode { get; set; }

        /// <summary>
        /// Gets or sets whether to enable garbage collection during conversion.
        /// </summary>
        public bool EnableGarbageCollection { get; set; } = true;

        /// <summary>
        /// Gets or sets the buffer size for reading in bytes.
        /// </summary>
        public int BufferSize { get; set; } = 81920; // 80KB default

        /// <summary>
        /// Gets or sets whether to process images.
        /// </summary>
        public bool ProcessImages { get; set; } = true;

        /// <summary>
        /// Gets or sets whether to process comments.
        /// </summary>
        public bool ProcessComments { get; set; } = true;
    }
}
