using System;
using System.IO;
using System.Runtime.InteropServices;

public class NativeOutputSuppressor : IDisposable
{
    private TextWriter _oldOut;
    private TextWriter _oldErr;

    // PInvoke to redirect stdout and stderr
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GetStdHandle(int nStdHandle);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool SetStdHandle(int nStdHandle, IntPtr handle);

    /// <summary>
    /// Suppresses native warnings by redirecting stdout and stderr to null.
    /// This ensures no console output is shown for warnings from native libraries (like libpng).
    /// </summary>
    public void SuppressNativeWarnings()
    {
        // Store the current standard output/error streams
        _oldOut = Console.Out;
        _oldErr = Console.Error;

        // Set null streams for stdout and stderr to suppress all output
        Console.SetOut(TextWriter.Null);
        Console.SetError(TextWriter.Null);
    }

    /// <summary>
    /// Restores the original stdout and stderr streams.
    /// </summary>
    public void RestoreNativeWarnings()
    {
        // Restore the original standard output/error streams
        if (_oldOut != null) Console.SetOut(_oldOut);
        if (_oldErr != null) Console.SetError(_oldErr);
    }

    /// <summary>
    /// Ensures stdout and stderr are restored when the object is disposed.
    /// </summary>
    public void Dispose()
    {
        RestoreNativeWarnings(); // Ensure the streams are restored
    }
}

// Part of the DocumentReaderComponent that processes image-based PDFs
// and extracts text from image files using Tesseract OCR.
public class DocumentReaderComponent
{
    private readonly string _tesseractDataPath;

    public DocumentReaderComponent()
    {
        _tesseractDataPath = @"./tessdata/tessdata-main"; // Path to Tesseract OCR data
    }

    /// <summary>
    /// Extracts text from an image using Tesseract OCR. Suppresses native warnings.
    /// </summary>
    /// <param name="imagePath">Path to the image file.</param>
    /// <returns>Extracted text from the image or error message in case of failure.</returns>
    private string ExtractTextFromImage(string imagePath)
    {
        try
        {
            // Suppress native library warnings (e.g., libpng errors)
            using (var suppressor = new NativeOutputSuppressor())
            {
                suppressor.SuppressNativeWarnings(); // Temporarily suppress all native output

                using (var engine = new TesseractEngine(_tesseractDataPath, "eng+heb", EngineMode.Default))
                {
                    // Process the image using Tesseract to extract text
                    using (var img = Pix.LoadFromFile(imagePath)) // Load the image from file
                    {
                        using (var page = engine.Process(img)) // Process image through OCR
                        {
                            return page.GetText(); // Extracted text
                        }
                    }
                }

                suppressor.RestoreNativeWarnings(); // Restore stdout/stderr after extraction
            }
        }
        catch (Exception ex)
        {
            // Handle errors by returning a descriptive message
            return $"Error extracting text from image: {ex.Message}";
        }
    }
}
