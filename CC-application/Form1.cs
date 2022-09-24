using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace CC_application
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.ShowDialog();
            string filePath = fileDialog.FileName.ToString();

            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Document doc = app.Documents.Open(filePath);
            try
            {
                if (doc != null)
                {
                    string[] inputArray = {
                        cName.Text, cAddr.Text, cPib.Text, cMb.Text, cCont.Text,
                        clName.Text,clBirth.Text, clNat.Text,clPssn.Text, clPsRelDate.Text, clPsExDate.Text, clPos.Text, clStart.Text, clEnd.Text,
                        agCntP.Text, agCont.Text,
                        wpAddr.Text, empSal.Text, empRep.Text,
                        dTod.Text
                    };
                    string[] parseArray = {
                        "!CNAME!", "!CADDR!","!CPIB!", "!CMB!", "!CCONT!",
                        "!CLNAME!","!CLBIRTH!", "!CLNAT!", "!CLPSSN!", "!CLPSRELDATE!", "!CLPSEXDATE!", "!CLPOS!", "!CLSTART!","!CLEND!",
                        "!AGCNTP!","!AGCONT!",
                        "!WPADDR!", "!EMPSAL!", "!EMPREP!",
                        "!DTOD!"
                    };
                    for (int i=0; i < parseArray.Length;i++)
                    {
                        if (inputArray[i] != "")
                        {
                            doc.Content.Find.Execute(parseArray[i], false, true, false, false, false, true, 1, false, inputArray[i], 2,
                            false, false, false, false);
                        }
                    }


                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "PDF document (*.pdf)|*.pdf";
                    saveFileDialog.ShowDialog();
                    string saveFilePath = saveFileDialog.FileName.ToString();

                    if (saveFilePath != filePath && saveFilePath != "")
                    {
                        doc.ExportAsFixedFormat(saveFilePath.ToString(), WdExportFormat.wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForOnScreen,
                                WdExportRange.wdExportAllDocument, 1, 1, WdExportItem.wdExportDocumentContent, true, true,
                                WdExportCreateBookmarks.wdExportCreateHeadingBookmarks, true, true, false);

                        System.Diagnostics.Process.Start(saveFilePath);
                    }
                }
                doc.Close(false);
                app.Quit();
            }
            catch (Exception)
            {
                app.Quit();
            }
        }


    }
}
