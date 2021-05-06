using LabelsSystem.Controller;
using LabelsSystem.Services;
using LabelsSystem.Services.IServices;
using LabelsSystem.Utils;
using System;
using System.Collections.Generic;
using System.Text;

namespace LabelsSystem.Starter
{
    class LabelsSystemMain
    {
        static void Main(string[] args)
        {
            WordFileProcess wfp = new WordFileProcess();
            PdfFileProcess pfp = new PdfFileProcess();
            LabelsController labelsController = new LabelsController(wfp, pfp, Constants.TEMPLATE_PATH, Constants.TEMP_PATH, Constants.OUTPUT_PATH);
            labelsController.GenerateLabelsDocument();
        }
    }
}
