using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SyntBotsDesktopAdpter.RPAListener;
using System.IO;

namespace SyntBotsDesktopTestingUtility
{
    public partial class CallAutomationfrm : Form
    {
        public CallAutomationfrm()
        {
            InitializeComponent();
            //txtAssembly.Text = @"C:\Osman\SyntBots\Desktop_RPA\Desktop_RPA\SyntBotsStudio_5.1.0.5216\SyntBotsDLLs\SyntBotsDesktopBase.dll";
            txtAssembly.Text = @"C:\TFS\Syntbots\Dev\MHC\HIM\DNReporting-multibot-Enhancements\Dependencies\SyntBotsDesktopBase.dll";
            txtJson.Text = @"C:\TFS\Syntbots\Dev\MHC\HIM\DNReporting-multibot-Enhancements\DNReporting\SyntBotsDesktopTestingUtility\Json\JsonData.txt";
            
        }

        private void btnRunAutomation_Click(object sender, EventArgs e)
        {
            txtOutPut.Text = string.Empty;
            if (txtAssembly.Text.Trim() != string.Empty && txtJson.Text.Trim() != string.Empty)
            {
                RPAAutomation objAutomation = new RPAAutomation();
                string JsonData = ReadJsonData();
                if (JsonData != string.Empty)
                {
                    txtOutPut.Text = objAutomation.RunAutomation(JsonData, txtAssembly.Text.Trim());
                    txtOutPut.ScrollBars = ScrollBars.Vertical;
                }
                else
                {
                    MessageBox.Show("Json Text file in not proper or missing!");
                }
            }
            else {
                MessageBox.Show("Please enter data!");
            }

        }

        private string ReadJsonData()
        {
             string jsonData = string.Empty;
             if (File.Exists(txtJson.Text.Trim()))
             {
                 jsonData = File.ReadAllText(txtJson.Text.Trim());
             }
            return jsonData;
        }

        private void JsonButton_Click(object sender, EventArgs e)
        {
           

            openFileDialogRPA.Filter = "txt files (*.txt)|*.txt";
            openFileDialogRPA.Title = "Please select text file for json object";
            DialogResult result = openFileDialogRPA.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialogRPA.FileName;
                txtJson.Text = file;
                //txtJson.Text = @"D:\SyntelRPAStarterKit\SyntBotsDesktopComponent\Development\SyntBotsDesktopUtility\JsonData.txt";
            }
        }

        private void BaseAssembly_Click(object sender, EventArgs e)
        {
           
            openFileDialogRPA.Filter = "Assembly files (*.dll)|*.dll";
            openFileDialogRPA.Title = "Please select .net base dll.";
            DialogResult result = openFileDialogRPA.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialogRPA.FileName;
                txtAssembly.Text = file;
               // txtAssembly.Text = @"D:\SyntelRPAStarterKit\SyntelRPA\SyntBotsDesktopBase\bin\Debug\SyntBotsDesktopBase.dll";
            }

        }

        private void Headers_Paint(object sender, PaintEventArgs e)
        {

        }

        private void txtOutPut_TextChanged(object sender, EventArgs e)
        {

        }

        private void CallAutomationfrm_Load(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void lblBaseAssemblyPath_Click(object sender, EventArgs e)
        {

        }

        private void lblOutput_Click(object sender, EventArgs e)
        {

        }
    }
}
