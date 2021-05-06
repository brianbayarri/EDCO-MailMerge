using Aspose.Words;
using LabelsSystem.Services.IServices;
using System;
using System.Collections.Generic;
using System.Text;

namespace LabelsSystem.Services
{
    public class PdfFileProcess : IPdfFileProcess
    {
        public void GeneratePdfFile(string docFilePath, string pdfFilePath)
        {
            // Load the document from disk.
            Document doc = new Document(docFilePath);
            // Save as PDF
            doc.Save(pdfFilePath);
        }
    }
}
