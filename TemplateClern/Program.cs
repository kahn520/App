using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace TemplateClern
{
    class Program
    {
        static void Main(string[] args)
        {
            string strFile = @"G:\APP\A000120140530A01KPBG.pptx";
            Application app = new Application();
            _Presentation ppt = app.Presentations.Open(strFile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
            Shapes shapes = ppt.SlideMaster.CustomLayouts[1].Shapes;
            Shape s = shapes["KSO_CT1"];
            MsoAutoShapeType msotype = s.AutoShapeType;
            MsoShapeType typ = s.PlaceholderFormat.ContainedType;
            PpPlaceholderType ppPlaceholderType = s.PlaceholderFormat.Type;
            ppt.Save();
            ppt.Close();
            app.Quit();
        }
    }
}
