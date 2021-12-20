using System;
using System.Windows.Forms;

namespace WordKiller
{
    public enum TypeDocument
    {
        DefaultDocument,
        LaboratoryWork,
        PracticalWork,
        Coursework,
        Report,
        GraduateWork,
        VKR,
        RGR
    }
    public partial class CustomInterface
    {


        TypeDocument typeDocument;

        void Work_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem toolStripMenuItem = (ToolStripMenuItem)sender;
            if (toolStripMenuItem.Checked)
            {
                return;
            }
            DefaultDocumentMenuItem.Checked = false;
            LabMenuItem.Checked = false;
            PracticeMenuItem.Checked = false;
            CourseworkMenuItem.Checked = false;
            RefMenuItem.Checked = false;
            DiplomMenuItem.Checked = false;
            VKRMenuItem.Checked = false;
            RGRMenuItem.Checked = false;
            toolStripMenuItem.Checked = true;
            TextHeaderUpdate();
        }

        void TextHeaderUpdate()
        {
            if (DefaultDocumentMenuItem.Checked)
            {
                TextHeader("документа");
                typeDocument = TypeDocument.DefaultDocument;
                TitleOffOnMenuItem.Visible = !DefaultDocumentMenuItem.Checked;
                ShowintTitlePanel();
            }
            else
            {
                TitleOffOnMenuItem.Visible = true;
                TitleOffOnMenuItem.Checked = false;
                ShowintTitlePanel();
                if (LabMenuItem.Checked)
                {
                    typeDocument = TypeDocument.LaboratoryWork;
                    TextHeader("лабораторной работы");
                    ShowTitleElems("0.0 1.0 2.1 3.1 0.3 1.3 0.4 1.4 0.6 1.6 0.7 1.7");
                }
                else if (PracticeMenuItem.Checked)
                {
                    typeDocument = TypeDocument.PracticalWork;
                    TextHeader("практической работы");
                    ShowTitleElems("0.0 1.0 2.1 3.1 0.3 1.3 0.4 1.4 0.6 1.6 0.7 1.7");
                }
                else if (CourseworkMenuItem.Checked)
                {
                    typeDocument = TypeDocument.Coursework;
                    TextHeader("курсовой работы");
                    ShowTitleElems("0.0 1.0 0.1 1.1 4.1 5.1 0.3 1.3 0.4 1.4 0.6 1.6 0.7 1.7");
                }
                else if (RefMenuItem.Checked)
                {
                    typeDocument = TypeDocument.Report;
                    TextHeader("реферата");
                }
                else if (DiplomMenuItem.Checked)
                {
                    typeDocument = TypeDocument.GraduateWork;
                    TextHeader("дипломной работы");
                }
                else if (VKRMenuItem.Checked)
                {
                    typeDocument = TypeDocument.VKR;
                    TextHeader("ВКР");
                }
                else if (RGRMenuItem.Checked)
                {
                    typeDocument = TypeDocument.RGR;
                    TextHeader("РГР");
                }
            }

        }

        void TextHeader(string type)
        {
            if (string.IsNullOrEmpty(saveFileName))
            {
                this.Text = "Сотворение " + type + " из небытия";
            }
        }

    }
}
