using System;
using System.Collections.Generic;
using System.Text;

namespace LabelsSystem.Services.IServices
{
    public interface IPdfFileProcess
    {
        public void GeneratePdfFile(string docFilePath, string pdfFilePath);
    }
}
