using NXOpen;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SmartFastener
{
    public partial class SmartFastenerUI : Form
    {
        public string selectedCategory = "" ;
        public SmartFastenerUI()
        {
            InitializeComponent();

            dgvCatalogueResults.AllowUserToAddRows = false;
            dgvCatalogueResults.ReadOnly = true;
            dgvCatalogueResults.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            lblGetCataloguResults.Show();
            HidePanels();
            btnCheckAvailibility.Enabled = false;
            btnGenerateModel.Enabled = false;

            cmbSHCSThreadM.SelectedIndexChanged += cmbSHCSThreadM_SelectedIndexChanged;
            cmbHSCSThreadM.SelectedIndexChanged += cmbHSCSThreadM_SelectedIndexChanged;
            cmbHRNThreadM.SelectedIndexChanged += cmbHRNThreadM_SelectedIndexChanged;

            txtbxSHCSLength.TextChanged += txtbxSHCSLength_TextChanged;
            txtbxHSCSLength.TextChanged += txtbxHSCSLength_TextChanged;

            txtbxPWID.TextChanged += txtbxPWID_TextChanged;
            txtbxPWOD.TextChanged += txtbxPWOD_TextChanged;
            txtbxPWThickness.TextChanged += txtbxPWThickness_TextChanged;
        }

        private void btnHRNCatagory_Click(object sender, EventArgs e)
        {
            HidePanels();
            selectedCategory = "Hexagon Regular Nut";
            lblGetCataloguResults.Hide();
            //dgvCatalogueResults.Columns.Clear();
            //dgvCatalogueResults.Rows.Clear();
            dgvCatalogueResults.DataSource = null;
            DataTable dt = Excel.FetchData("Hexagon Regular Nut");
            dgvCatalogueResults.DataSource = dt;
            pnlHRNInput.Show();
            ValidateAvalibilityButton();
            ValidateGenerateButton();
        }

        private void lblHSCSThreadM_Click(object sender, EventArgs e)
        {

        }

        private void btnHSCSCatagory_Click(object sender, EventArgs e)
        {
            HidePanels();
            selectedCategory = "Hex Socket Countersunk Screw";
            lblGetCataloguResults.Hide();
            //dgvCatalogueResults.Columns.Clear();
            //dgvCatalogueResults.Rows.Clear();
            dgvCatalogueResults.DataSource = null;
            DataTable dt = Excel.FetchData("Hex Socket Countersunk Screw");
            dgvCatalogueResults.DataSource = dt;
            pnlHSCSInput.Show();
            ValidateAvalibilityButton();
            ValidateGenerateButton();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //dgvCatalogueResults.Columns.Add()
        }
    
        private void btnSHCSCatagory_Click(object sender, EventArgs e)
        {
            HidePanels();
            selectedCategory = "Socket Head Cap Screw";
            lblGetCataloguResults.Hide();
            //dgvCatalogueResults.Columns.Clear();
            //dgvCatalogueResults.Rows.Clear();
            dgvCatalogueResults.DataSource = null;
            DataTable dt = Excel.FetchData("Socket Head Cap Screw");
            dgvCatalogueResults.DataSource = dt;
            pnlSHCSInput.Show();
            ValidateAvalibilityButton();
            ValidateGenerateButton();
        }

        private void btnPWCatagory_Click(object sender, EventArgs e)
        {
            HidePanels();
            selectedCategory = "Plain Washer";
            lblGetCataloguResults.Hide();
            //dgvCatalogueResults.Columns.Clear();
            //dgvCatalogueResults.Rows.Clear();
            dgvCatalogueResults.DataSource = null;
            DataTable dt = Excel.FetchData("Plain Washer");
            dgvCatalogueResults.DataSource = dt;
            pnlPWInput.Show();
            ValidateAvalibilityButton();
            ValidateGenerateButton();
        }

        private void lblGetCataloguResults_Click(object sender, EventArgs e)
        {

        }

        private void btnGenerateModel_Click(object sender, EventArgs e)
        {
            
            switch (selectedCategory)
            {
                case "Socket Head Cap Screw":
                    SocketHeadCapScrew.Create(cmbSHCSThreadM.Text, txtbxSHCSLength.Text);
                    Excel.UpdateSHCSData(cmbSHCSThreadM.Text, txtbxSHCSLength.Text);
                    this.Close();
                    break;

                case "Hex Socket Countersunk Screw":
                    HexScoketCountersunkScrew.Create(cmbHSCSThreadM.Text, txtbxHSCSLength.Text);
                    Excel.UpdateHSCSData(cmbHSCSThreadM.Text, txtbxHSCSLength.Text);
                    this.Close();
                    break;

                case "Hexagon Regular Nut":
                    HexagonalRegularNut.Create(cmbHRNThreadM.Text);
                    Excel.UpdateHRNData(cmbHRNThreadM.Text);
                    this.Close();
                    break;

                case "Plain Washer":
                    PlainWasher.Create(txtbxPWID.Text, txtbxPWOD.Text, txtbxPWThickness.Text);
                    Excel.UpdatePWData(txtbxPWID.Text, txtbxPWOD.Text, txtbxPWThickness.Text);
                    this.Close();
                    break;

                default:
                    MessageBox.Show(
                        "Please select a fastener category.",
                        "Generate Model",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    this.Close();
                    break;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
        private void ValidateGenerateButton()
        {
            bool valid = false;

            switch (selectedCategory)
            {
                case "Socket Head Cap Screw":
                    valid =
                        cmbSHCSThreadM.SelectedIndex != -1 &&
                        !string.IsNullOrWhiteSpace(txtbxSHCSLength.Text);

                    break;


                case "Hex Socket Countersunk Screw":

                    valid =
                        cmbHSCSThreadM.SelectedIndex != -1 &&
                        !string.IsNullOrWhiteSpace(txtbxHSCSLength.Text);
                    break;


                case "Hexagon Regular Nut":

                    valid = cmbHRNThreadM.SelectedIndex != -1;
                    break;


                case "Plain Washer":

                    valid =
                        !string.IsNullOrWhiteSpace(txtbxPWID.Text) &&
                        !string.IsNullOrWhiteSpace(txtbxPWOD.Text) &&
                        !string.IsNullOrWhiteSpace(txtbxPWThickness.Text);

                    break;
            }

            btnGenerateModel.Enabled = valid;
        }
        private void ValidateAvalibilityButton()
        {
            bool valid = false;

            switch (selectedCategory)
            {
                case "Socket Head Cap Screw":
                    valid =
                        cmbSHCSThreadM.SelectedIndex != -1 &&
                        !string.IsNullOrWhiteSpace(txtbxSHCSLength.Text);

                    break;


                case "Hex Socket Countersunk Screw":

                    valid =
                        cmbHSCSThreadM.SelectedIndex != -1 &&
                        !string.IsNullOrWhiteSpace(txtbxHSCSLength.Text);
                    break;


                case "Hexagon Regular Nut":

                    valid = cmbHRNThreadM.SelectedIndex != -1;
                    break;


                case "Plain Washer":

                    valid =
                        !string.IsNullOrWhiteSpace(txtbxPWID.Text) &&
                        !string.IsNullOrWhiteSpace(txtbxPWOD.Text) &&
                        !string.IsNullOrWhiteSpace(txtbxPWThickness.Text);

                    break;
            }

            btnCheckAvailibility.Enabled = valid;
        }
        private void cmbSHCSThreadM_SelectedIndexChanged(object sender, EventArgs e)
        {
            ValidateAvalibilityButton();
            ValidateGenerateButton();
        }
        private void cmbHSCSThreadM_SelectedIndexChanged(object sender, EventArgs e)
        {
            ValidateAvalibilityButton();
            ValidateGenerateButton();
        }
        private void cmbHRNThreadM_SelectedIndexChanged(object sender, EventArgs e)
        {
            ValidateAvalibilityButton();
            ValidateGenerateButton();
        }
        private void txtbxSHCSLength_TextChanged(object sender, EventArgs e)
        {
            ValidateAvalibilityButton();
            ValidateGenerateButton();
        }
        private void txtbxHSCSLength_TextChanged(object sender, EventArgs e)
        {
            ValidateAvalibilityButton();
            ValidateGenerateButton();
        }
        private void txtbxPWID_TextChanged(object sender, EventArgs e)
        {
            ValidateAvalibilityButton();
            ValidateGenerateButton();
        }
        private void txtbxPWOD_TextChanged(object sender, EventArgs e)
        {
            ValidateAvalibilityButton();
            ValidateGenerateButton();
        }
        private void txtbxPWThickness_TextChanged(object sender, EventArgs e)
        {
            ValidateAvalibilityButton();
            ValidateGenerateButton();
        }


        private void HidePanels()
        {
            pnlHRNInput.Hide();
            pnlHSCSInput.Hide();
            pnlPWInput.Hide();
            pnlSHCSInput.Hide();
        }

        private void btnCheckAvailibility_Click(object sender, EventArgs e)
        {
            switch (selectedCategory)
            {
                case "Socket Head Cap Screw":
                    Excel.CheckSHCSAvailibilty(cmbSHCSThreadM.Text, txtbxSHCSLength.Text);
                    break;

                case "Hex Socket Countersunk Screw":
                    Excel.CheckHSCSAvailibilty(cmbHSCSThreadM.Text, txtbxHSCSLength.Text);
                    break;

                case "Hexagon Regular Nut":
                    Excel.CheckHRNAvailibilty(cmbHRNThreadM.Text);
                    break;

                case "Plain Washer":
                    Excel.CheckPWAvailibilty(txtbxPWID.Text, txtbxPWOD.Text, txtbxPWThickness.Text);
                    break;

                default:
                    MessageBox.Show(
                        "Please select a fastener category.",
                        "Generate Model",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    break;
            }
        }
    }
}
