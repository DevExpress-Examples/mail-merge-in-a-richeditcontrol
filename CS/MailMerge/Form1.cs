using DevExpress.XtraBars;
using DevExpress.XtraEditors;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraTab;
using System.Collections;
using System.Windows.Forms;

namespace MailMerge
{
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public Form1()
        {
            InitializeComponent();
            ribbonControl1.SelectedPage = mailingsRibbonPage1;
            richEditControl1.LoadDocument("invitation.rtf", DocumentFormat.Rtf);

            // Set a data source
            richEditControl1.Options.MailMerge.DataSource = new SampleData();
            richEditControl1.Options.MailMerge.ViewMergedData = true;
        }

        private void mergeToNewDocumentItem_ItemClick(object sender, ItemClickEventArgs e)
        {
            MailMergeOptions myMergeOptions = richEditControl1.Document.CreateMailMergeOptions();
            myMergeOptions.FirstRecordIndex = 1;
            myMergeOptions.LastRecordIndex = 3;
            myMergeOptions.MergeMode = MergeMode.NewSection;

            // Export a mail-merged document in a different Document instance
            richEditControl1.Document.MailMerge(myMergeOptions, richEditControl2.Document);
            xtraTabControl1.SelectedTabPageIndex = 1;
        }
        private void mergeToFileItem_ItemClick(object sender, ItemClickEventArgs e)
        {
            MailMergeOptions myMergeOptions = richEditControl1.Document.CreateMailMergeOptions();
            myMergeOptions.DataSource = new SampleData();
            myMergeOptions.FirstRecordIndex = 1;
            myMergeOptions.LastRecordIndex = 3;
            myMergeOptions.MergeMode = MergeMode.NewSection;

            using (XtraSaveFileDialog fileDialog = new XtraSaveFileDialog())
            {
                fileDialog.Filter = "MS Word 2007 documents (*.docx)|*.docx|All files (*.*)|*.*";
                fileDialog.FilterIndex = 1;
                fileDialog.RestoreDirectory = true;

                DialogResult dialogResult = fileDialog.ShowDialog();
                if (dialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    string fName = fileDialog.FileName;
                    if (fName != "")
                    {
                        // Export a mail-merged document in a new file
                        richEditControl1.Document.MailMerge(myMergeOptions, fileDialog.FileName, DocumentFormat.OpenXml);
                        System.Diagnostics.Process.Start(fileDialog.FileName);
                    }
                }
            }
        }
        private void prevRecordItem_ItemClick(object sender, ItemClickEventArgs e)
        {
            // Navigate to a previous data source record
            if (richEditControl1.Options.MailMerge.ActiveRecord > 0)
                richEditControl1.Options.MailMerge.ActiveRecord--;
        }

        private void nextRecordItem_ItemClick(object sender, ItemClickEventArgs e)
        {
            // Navigate to a next data source record
            int recordCount = (richEditControl1.Options.MailMerge.DataSource as ArrayList).Count;
            if (richEditControl1.Options.MailMerge.ActiveRecord < recordCount - 1)
                richEditControl1.Options.MailMerge.ActiveRecord++;
        }
        #region #TabControlSelection
        private void xtraTabControl1_Selected(object sender, TabPageEventArgs e)
        {
            switch (e.PageIndex)
            {
                case 0:
                    this.richEditBarController1.RichEditControl = this.richEditControl1;
                    break;
                case 1:
                    this.richEditBarController1.RichEditControl = this.richEditControl2;
                    break;
            }
        }
        #endregion
    }
}