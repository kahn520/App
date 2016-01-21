using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.CSharp;
using Microsoft.Office.Tools;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace ViewAddIn
{
    public partial class ThisAddIn
    {
        private CustomTaskPane pane;
        private PaneShapeInfo paneInfo;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Globals.ThisAddIn.Application.WindowSelectionChange += Application_WindowSelectionChange;
            paneInfo = new PaneShapeInfo();
            pane = Globals.ThisAddIn.CustomTaskPanes.Add(paneInfo, "形状信息");
            pane.Visible = true;
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            if (Sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || Sel.ShapeRange.Count != 1)
                return;
            StringBuilder sb = new StringBuilder("");
            Shape shp = Sel.ShapeRange[1];
            sb.AppendLine("名称:" + shp.Name);
            sb.AppendLine("AutoShapeType:" + Enum.GetName(typeof (Office.MsoAutoShapeType), shp.AutoShapeType));
            try
            {
                sb.AppendLine("PlaceHolderFormat:" + Enum.GetName(typeof(PowerPoint.PpPlaceholderType), shp.PlaceholderFormat.Type));
                sb.AppendLine("ContainedType:" + Enum.GetName(typeof(Office.MsoShapeType), shp.PlaceholderFormat.ContainedType));
            }
            catch
            { }
            paneInfo.SetInfo(sb.ToString());
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            paneInfo.Dispose();
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
