using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
Console.WriteLine("Hello, World!");

// Initialize Word application
Word.Application wordApp = new Word.Application();
wordApp.Visible = true; // Make Word application visible
Word.Document doc = null;

try
{
    object missing = System.Reflection.Missing.Value;

    // Open the document

    //string filePath = @"C:\Users\engrb\Desktop\!BUGS.Rules.docx";
    string filePath = @"C:\Users\engrb\Desktop\Hello.docx";
    doc = wordApp.Documents.Open(filePath);

    // Get the total number of pages
    int totalPages = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages, ref missing);

    // Debug print the total pages
    Debug.WriteLine($"Total pages in the document: {totalPages}");

    // Navigate to page 5
    object what = Word.WdGoToItem.wdGoToPage;
    object which = Word.WdGoToDirection.wdGoToAbsolute;
    object count = 4; // Page number
    object name = Type.Missing; // Not used for page navigation

    // Call the GoTo method
    Word.Range range = doc.GoTo(ref what, ref which, ref count, ref name);

    // Scroll into view
    range.Select();

    // Keep the document open (you can add further processing here)
}
catch (Exception ex)
{
    Console.WriteLine("Error: " + ex.Message);
}
finally
{
    // Close the opened document if needed
    // doc.Close();
    // You can also choose to release the doc object if it's no longer needed:
    // if (doc != null)
    // {
    //     System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
    // }

    // Keep the Word application open
}