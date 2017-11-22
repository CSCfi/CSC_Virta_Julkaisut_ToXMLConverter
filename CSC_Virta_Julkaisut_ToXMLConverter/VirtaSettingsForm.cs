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
    public partial class VirtaSettingsForm : Form
    {
        public VirtaSettingsForm()
        {
            InitializeComponent();

            System.Windows.Forms.ToolTip VuosiButton = new System.Windows.Forms.ToolTip();
            // VuosiButton.SetToolTip(this.VuosiRadioButton2015, "Ilmoitusvuosi 2016 on ajankohtainen vasta 2017");
        }

        private void SettingsTallennaButton_Click(object sender, EventArgs e)
        {

            if ( VuosiRadioButton2015.Checked ) {               
           
                CSC_VIRTA_JulkaisutForm.vuosiIlmo = 2016;              

            }

            if (VuosiRadioButton2016.Checked)
            {

                CSC_VIRTA_JulkaisutForm.vuosiIlmo = 2016;

            }

            if (VuosiRadioButton2017.Checked)
            {

                CSC_VIRTA_JulkaisutForm.vuosiIlmo = 2017;

            }

        }
    }
}
