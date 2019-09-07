using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddIn1
{
    public partial class UI
    {
        Processer processer;
        string FomatFilePath, MdDilePath, OutputFilePath;
        bool formatDocxSelected = false, mdSelected = false, outputDocxSelected = false;

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
            outputDocxSelected = true;
            OutputFilePath = saveFileDialog.FileName;
        }

        private void btn_BuildDocx_Click(object sender, RibbonControlEventArgs e)
        {
            saveFileDialog.ShowDialog();

            if (formatDocxSelected && mdSelected)
            {
                if (OutputFilePath == null || OutputFilePath == "") OutputFilePath = MdDilePath + ".docx";
                Globals.ThisAddIn.processer.process(FomatFilePath, MdDilePath, OutputFilePath);
                System.Windows.Forms.MessageBox.Show("Build Succeed!\nWrite to " + OutputFilePath);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please Select Files First!");
            }

            formatDocxSelected = mdSelected = outputDocxSelected = false;
            cb_DocxSelected.Name = cb_MdSelected.Name = "";
            FomatFilePath = MdDilePath = OutputFilePath = "";
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
