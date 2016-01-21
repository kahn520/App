using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace TemplateClern
{
    class Program
    {
        static void Main(string[] args)
        {
            string strFile = @"G:\APP\A000120140530A01KPBG.pptx";
            Application app = new Application();
            _Presentation ppt = app.Presentations.Open(strFile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
            ppt.SlideMaster.
            ppt.Save();
            ppt.Close();
            app.Quit();
        }
    }
}
