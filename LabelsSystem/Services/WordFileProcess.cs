using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using GemBox.Document;
using LabelsSystem.Services.IServices;
using LabelsSystem.Utils;
using System;
using System.Data;
using System.Linq;

namespace LabelsSystem.Services
{
    public class WordFileProcess : IWordFileProcess
    {
        public bool ReplaceMergeFields(string templatePath, string outputPath)
        {
            try
            {
                // Set licensing info.
                ComponentInfo.SetLicense("FREE-LIMITED-KEY");
                ComponentInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

                // Create mail merge data source. 
                var dataTable = new DataTable()
                {
                    Columns =
                {
                    new DataColumn("Name"),
                    new DataColumn("Title"),
                    new DataColumn("SKU"),
                    new DataColumn("Invoice")
                },
                    Rows =
                {
                    { "John", "20", "dsad", "dsadsa"},
                    { "Fred", "11", "dsad", "dsadsa"},
                    { "Hans", "12", "dsad", "dsadsa"}
                }
                };

                var document = DocumentModel.Load(templatePath, LoadOptions.DocxDefault);

                // Use this if field names and data column names differ. If they are the same (case-insensitive), then you don't need to define mappings explicitly.
                document.MailMerge.FieldMappings.Add("Name", "Name");
                document.MailMerge.FieldMappings.Add("Title", "Title");
                document.MailMerge.FieldMappings.Add("SKU", "SKU");
                document.MailMerge.FieldMappings.Add("Invoice", "Invoice");

                // Execute mail merge. Each mail merge field will be replaced with the data from the data source's appropriate row and column.
                document.MailMerge.Execute(dataTable);

                // Remove any left mail merge field.
                document.MailMerge.RemoveMergeFields();

                // Save resulting document to a file.
                document.Save(outputPath);

                return true;
            }
            catch(Exception e)
            {
                return false;
            }
            
        }

        public bool RemoveHeader(string filePath)
        {
            try
            {
                // Given a document name, remove all of the headers and footers
                // from the document.
                using (WordprocessingDocument doc =
                    WordprocessingDocument.Open(filePath, true))
                {
                    // Get a reference to the main document part.
                    var docPart = doc.MainDocumentPart;

                    // Count the header and footer parts and continue if there 
                    // are any.
                    if (docPart.HeaderParts.Count() > 0)
                    {
                        // Remove the header part.
                        docPart.DeleteParts(docPart.HeaderParts);

                        // Get a reference to the root element of the main
                        // document part.
                        Document document = docPart.Document;

                        // Remove all references to the headers and footers.

                        // First, create a list of all descendants of type
                        // HeaderReference. Then, navigate the list and call
                        // Remove on each item to delete the reference.
                        var headers = document.Descendants<HeaderReference>().ToList();
                        foreach (var header in headers)
                        {
                            header.Remove();
                        }

                        // Save the changes.
                        document.Save();
                    }
                }

                return true;
            }
            catch(Exception e)
            {
                return false;
            }
            
        }
    }
}
