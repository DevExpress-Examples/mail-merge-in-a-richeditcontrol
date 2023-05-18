Imports DevExpress.XtraBars
Imports DevExpress.XtraEditors
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraTab
Imports System.Windows.Forms

Namespace MailMerge
    Partial Public Class Form1
        Inherits DevExpress.XtraBars.Ribbon.RibbonForm

        Public Sub New()
            InitializeComponent()
            ribbonControl1.SelectedPage = mailingsRibbonPage1
            richEditControl1.LoadDocument("invitation.rtf", DocumentFormat.Rtf)

            'Set a data source
            richEditControl1.Options.MailMerge.DataSource = New SampleData()
            richEditControl1.Options.MailMerge.ViewMergedData = True
        End Sub

        Private Sub mergeToNewDocumentItem_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs) Handles mergeToNewDocumentItem.ItemClick
            Dim myMergeOptions As MailMergeOptions = richEditControl1.Document.CreateMailMergeOptions()
            myMergeOptions.FirstRecordIndex = 1
            myMergeOptions.LastRecordIndex = 3
            myMergeOptions.MergeMode = MergeMode.NewSection

            'Export the mail-merged document in a different Document instance
            richEditControl1.Document.MailMerge(myMergeOptions, richEditControl2.Document)
            xtraTabControl1.SelectedTabPageIndex = 1
        End Sub

        Private Sub mergeToFileItem_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs) Handles mergeToFileItem.ItemClick
            Dim myMergeOptions As MailMergeOptions = richEditControl1.Document.CreateMailMergeOptions()
            myMergeOptions.DataSource = New SampleData()
            myMergeOptions.FirstRecordIndex = 1
            myMergeOptions.LastRecordIndex = 3
            myMergeOptions.MergeMode = MergeMode.NewSection

            Using fileDialog As XtraSaveFileDialog = New XtraSaveFileDialog()
                fileDialog.Filter = "MS Word 2007 documents (*.docx)|*.docx|All files (*.*)|*.*"
                fileDialog.FilterIndex = 1
                fileDialog.RestoreDirectory = True
                Dim dialogResult As DialogResult = fileDialog.ShowDialog()

                If dialogResult = System.Windows.Forms.DialogResult.OK Then
                    Dim fName As String = fileDialog.FileName

                    If fName <> "" Then
                        'Export the mail-merged document in a new file
                        richEditControl1.Document.MailMerge(myMergeOptions, fileDialog.FileName, DocumentFormat.OpenXml)
                        System.Diagnostics.Process.Start(fileDialog.FileName)
                    End If
                End If
            End Using
        End Sub

        Private Sub prevRecordItem_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs) Handles prevRecordItem.ItemClick
            'Navigate to the previous data source record
            If richEditControl1.Options.MailMerge.ActiveRecord > 0 Then richEditControl1.Options.MailMerge.ActiveRecord -= 1
        End Sub

        Private Sub nextRecordItem_ItemClick(ByVal sender As Object, ByVal e As ItemClickEventArgs) Handles nextRecordItem.ItemClick
            'Navigate to the previous data source record
            Dim recordCount As Integer = (TryCast(richEditControl1.Options.MailMerge.DataSource, ArrayList)).Count
            If richEditControl1.Options.MailMerge.ActiveRecord < recordCount - 1 Then richEditControl1.Options.MailMerge.ActiveRecord += 1
        End Sub

#Region "TabControlSelection"
        Private Sub xtraTabControl1_Selected(ByVal sender As Object, ByVal e As TabPageEventArgs) Handles xtraTabControl1.Selected
            Select Case e.PageIndex
                Case 0
                    Me.richEditBarController1.RichEditControl = Me.richEditControl1
                Case 1
                    Me.richEditBarController1.RichEditControl = Me.richEditControl2
            End Select
        End Sub
#End Region
    End Class

End Namespace