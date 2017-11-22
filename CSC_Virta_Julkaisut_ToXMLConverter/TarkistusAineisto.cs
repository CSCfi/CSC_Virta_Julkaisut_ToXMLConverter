using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSC_Virta_Julkaisut_ToXMLConverter
{
    public partial class TarkistusAineisto : Form
    {
        public TarkistusAineisto()
        {
            InitializeComponent();
        }

        private void TarkistusExcelButton_Click(object sender, EventArgs e)
        {
            CSC_Virta_Julkaisut_ToXMLConverter.CSC_VIRTA_JulkaisutForm.AvaaCSVExcelissa();
        }
    }
}
