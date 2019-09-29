using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddIn1
{
    public partial class UI
    {
        string FomatFilePath, MdDilePath, OutputFilePath, HeaderDocxPath;
        bool formatDocxSelected = false, mdSelected = false;

        private void UI_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_SelectFormatDocx_Click(object sender, RibbonControlEventArgs e)
        {
            openFileDialog_docx.ShowDialog();
        }

        private void btn_SelectMdFile_Click(object sender, RibbonControlEventArgs e)
        {
            openFileDialog_md.ShowDialog();
        }

        private void saveFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            OutputFilePath = saveFileDialog.FileName;
        }

        private void openFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            HeaderDocxPath = openFileDialog1.FileName;
        }

        private void btn_selecdHeaderDocx_Click(object sender, RibbonControlEventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void btn_BuildDocx_Click(object sender, RibbonControlEventArgs e)
        {
            saveFileDialog.ShowDialog();

            if (formatDocxSelected && mdSelected)
            {
                if (OutputFilePath == null || OutputFilePath == "") OutputFilePath = MdDilePath + ".docx";
                if (Globals.ThisAddIn.processer.process(FomatFilePath, MdDilePath,HeaderDocxPath, OutputFilePath))
                {
                    System.Windows.Forms.MessageBox.Show("Build Succeed!\nWrite to " + OutputFilePath);
                    formatDocxSelected = mdSelected = false;
                    FomatFilePath = MdDilePath = OutputFilePath = HeaderDocxPath = "";
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Build Failed!");
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please Select Files First!");
            }

            
        }

        private void openFileDialog_docx_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            formatDocxSelected = true;
            cb_DocxSelected.Checked = true;
            FomatFilePath = openFileDialog_docx.FileName;

        }

        private void openFileDialog_md_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            mdSelected = true;
            cb_MdSelected.Checked = true;
            MdDilePath = openFileDialog_md.FileName;
        }
    }
}
