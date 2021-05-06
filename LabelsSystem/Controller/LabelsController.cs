using LabelsSystem.Services.IServices;
using LabelsSystem.Utils;
using System;
using System.Collections.Generic;
using System.Text;

namespace LabelsSystem.Controller
{
    public class LabelsController
    {
        private readonly IWordFileProcess _wordFileProcess;
        private readonly IPdfFileProcess _pdfFileProcess;
        private string templatePath;
        private string outputPath;
        private string tempPath;

        public LabelsController(IWordFileProcess wordFileProcess, IPdfFileProcess pdfFileProcess, string templatePath, string tempPath, string outputPath)
        {
            _wordFileProcess = wordFileProcess;
            _pdfFileProcess = pdfFileProcess;
            this.templatePath = templatePath;
            this.tempPath = tempPath;
            this.outputPath = outputPath;
        }

        public void GenerateLabelsDocument()
        {
            var result = _wordFileProcess.ReplaceMergeFields(templatePath, tempPath);
            if (result)
            {
                var result2 = _wordFileProcess.RemoveHeader(tempPath);
                if (result2)
                {
                    _pdfFileProcess.GeneratePdfFile(tempPath, outputPath);
                }
            }
        }
    }
}
