using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using System.IO;
using System.Text.RegularExpressions;

namespace TemplateClern
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("continue");
            Console.Read();
            string strPath = @"G:\App\TemplateTest";
            string[] strFiles = Directory.GetFiles(strPath);
            int i = 1, count = strFiles.Length;
            Application app = new Application();
            foreach (string file in strFiles)
            {
                _Presentation ppt = app.Presentations.Open(file, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                foreach (_Slide slide in ppt.Slides)
                {
                    LoopPageShape(slide.Shapes);
                }
                LoopMasterShape(ppt.SlideMaster);
                foreach (CustomLayout layout in ppt.SlideMaster.CustomLayouts)
                {
                    LoopMasterShape(layout);
                }
                ppt.Save();
                ppt.Close();
                Console.WriteLine(i + "/" + count);
                i++;
            }
            app.Quit();
        }

        private static void LoopPageShape(Shapes shps)
        {
            foreach (Shape shp in shps)
            {
                RenameShape(shp);
                DeleteTag(shp);
                DeleteText(shp);
            }
        }

        private static void LoopMasterShape(Master master)
        {
            foreach (Shape shp in master.Shapes)
            {
                RenameShape(shp);
                DeleteText(shp);
            }
        }

        private static void LoopMasterShape(CustomLayout layout)
        {
            foreach (Shape shp in layout.Shapes)
            {
                RenameShape(shp);
                DeleteText(shp);
            }
        }

        private static void LoopNotesShape()
        {
            
        }



        private static void RenameShape(Shape shp)
        {
            shp.Name = shp.Name.Replace("KSO", "APP");
        }

        private static void DeleteTag(Shape shp)
        {
            if (shp.Tags.Count > 0)
            {
                for (int i = shp.Tags.Count - 1; i >= 0; i--)
                {
                    string name = shp.Tags.Name(i);
                    shp.Tags.Delete(name);
                }
            }
        }

        private static void DeleteText(Shape shp)
        {
            if(shp.TextFrame.HasText == MsoTriState.msoFalse)
                return;
            if(shp.TextFrame.TextRange.Text.Trim() == "")
                return;
            //if (shp.Type == MsoShapeType.msoPlaceholder)
            //{
            //    shp.TextFrame.TextRange.Delete();
            //    return;
            //}
            if(HasChinese(shp.TextFrame.TextRange.Text))
                shp.TextFrame.TextRange.Delete();
        }

        public static bool HasChinese(string str)
        {
            return Regex.IsMatch(str, @"[\u4e00-\u9fa5]");
        }


    }
}
