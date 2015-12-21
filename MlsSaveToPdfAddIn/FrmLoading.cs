using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MlsSaveToPdfAddIn
{
    public partial class FrmLoading : Form
    {
        public FrmLoading()
        {
            InitializeComponent();
        }


        public string Msg
        {
            get { return lblMsg.Text; }
            set { lblMsg.Text = value; }
        }
    }
}
