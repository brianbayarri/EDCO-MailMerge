using System;
using System.Collections.Generic;
using System.Text;

namespace LabelsSystem.Services.IServices
{
    public interface IWordFileProcess
    {
        public bool ReplaceMergeFields(string tempaltePath, string outputPath);
        public bool RemoveHeader(string filePath);
    }
}
