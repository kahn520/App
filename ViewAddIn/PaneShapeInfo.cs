using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ViewAddIn
{
    public partial class PaneShapeInfo : UserControl
    {
        public PaneShapeInfo()
        {
            InitializeComponent();
        }
        public void SetInfo(string strInfo)
        {
            txtInfo.Text = "";
            txtInfo.Text = strInfo;
        }
    }

}
