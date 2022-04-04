using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KodeMagd.Misc
{
    public partial class FrmInstructions : Form
    {
        public enum enumInstructionType 
        { 
            eAddingReference

        }

        public FrmInstructions(enumInstructionType eInstructionType, string sText)
        {
            InitializeComponent();
        }

        private void FrmInstructions_Load(object sender, EventArgs e)
        {
            ClsDefaults.FormatControl(ref ssStatus);
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }
    }
}
