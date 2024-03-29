﻿Namespace MailMerge
    Partial Class Form1
        Private components As System.ComponentModel.IContainer = Nothing

        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (components IsNot Nothing) Then
                components.Dispose()
            End If

            MyBase.Dispose(disposing)
        End Sub

        Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container()
            Dim GalleryItemGroup1 As DevExpress.XtraBars.Ribbon.GalleryItemGroup = New DevExpress.XtraBars.Ribbon.GalleryItemGroup()
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
            Me.galleryChangeStyleItem1 = New DevExpress.XtraRichEdit.UI.GalleryChangeStyleItem()
            Me.xtraTabControl1 = New DevExpress.XtraTab.XtraTabControl()
            Me.xtraTabPage1 = New DevExpress.XtraTab.XtraTabPage()
            Me.richEditControl1 = New DevExpress.XtraRichEdit.RichEditControl()
            Me.ribbonControl1 = New DevExpress.XtraBars.Ribbon.RibbonControl()
            Me.fileNewItem1 = New DevExpress.XtraRichEdit.UI.FileNewItem()
            Me.fileOpenItem1 = New DevExpress.XtraRichEdit.UI.FileOpenItem()
            Me.fileSaveItem1 = New DevExpress.XtraRichEdit.UI.FileSaveItem()
            Me.fileSaveAsItem1 = New DevExpress.XtraRichEdit.UI.FileSaveAsItem()
            Me.quickPrintItem1 = New DevExpress.XtraRichEdit.UI.QuickPrintItem()
            Me.printItem1 = New DevExpress.XtraRichEdit.UI.PrintItem()
            Me.printPreviewItem1 = New DevExpress.XtraRichEdit.UI.PrintPreviewItem()
            Me.undoItem1 = New DevExpress.XtraRichEdit.UI.UndoItem()
            Me.redoItem1 = New DevExpress.XtraRichEdit.UI.RedoItem()
            Me.changeFontNameItem1 = New DevExpress.XtraRichEdit.UI.ChangeFontNameItem()
            Me.repositoryItemFontEdit1 = New DevExpress.XtraEditors.Repository.RepositoryItemFontEdit()
            Me.changeFontSizeItem1 = New DevExpress.XtraRichEdit.UI.ChangeFontSizeItem()
            Me.repositoryItemRichEditFontSizeEdit1 = New DevExpress.XtraRichEdit.Design.RepositoryItemRichEditFontSizeEdit()
            Me.changeFontColorItem1 = New DevExpress.XtraRichEdit.UI.ChangeFontColorItem()
            Me.changeFontBackColorItem1 = New DevExpress.XtraRichEdit.UI.ChangeFontBackColorItem()
            Me.toggleFontBoldItem1 = New DevExpress.XtraRichEdit.UI.ToggleFontBoldItem()
            Me.toggleFontItalicItem1 = New DevExpress.XtraRichEdit.UI.ToggleFontItalicItem()
            Me.fontSizeIncreaseItem1 = New DevExpress.XtraRichEdit.UI.FontSizeIncreaseItem()
            Me.fontSizeDecreaseItem1 = New DevExpress.XtraRichEdit.UI.FontSizeDecreaseItem()
            Me.clearFormattingItem1 = New DevExpress.XtraRichEdit.UI.ClearFormattingItem()
            Me.toggleParagraphAlignmentLeftItem1 = New DevExpress.XtraRichEdit.UI.ToggleParagraphAlignmentLeftItem()
            Me.toggleParagraphAlignmentCenterItem1 = New DevExpress.XtraRichEdit.UI.ToggleParagraphAlignmentCenterItem()
            Me.toggleParagraphAlignmentRightItem1 = New DevExpress.XtraRichEdit.UI.ToggleParagraphAlignmentRightItem()
            Me.toggleParagraphAlignmentJustifyItem1 = New DevExpress.XtraRichEdit.UI.ToggleParagraphAlignmentJustifyItem()
            Me.toggleNumberingListItem1 = New DevExpress.XtraRichEdit.UI.ToggleNumberingListItem()
            Me.toggleBulletedListItem1 = New DevExpress.XtraRichEdit.UI.ToggleBulletedListItem()
            Me.toggleMultiLevelListItem1 = New DevExpress.XtraRichEdit.UI.ToggleMultiLevelListItem()
            Me.decreaseIndentItem1 = New DevExpress.XtraRichEdit.UI.DecreaseIndentItem()
            Me.increaseIndentItem1 = New DevExpress.XtraRichEdit.UI.IncreaseIndentItem()
            Me.toggleShowWhitespaceItem1 = New DevExpress.XtraRichEdit.UI.ToggleShowWhitespaceItem()
            Me.showParagraphFormItem1 = New DevExpress.XtraRichEdit.UI.ShowParagraphFormItem()
            Me.insertMergeFieldItem1 = New DevExpress.XtraRichEdit.UI.InsertMergeFieldItem()
            Me.showAllFieldCodesItem1 = New DevExpress.XtraRichEdit.UI.ShowAllFieldCodesItem()
            Me.showAllFieldResultsItem1 = New DevExpress.XtraRichEdit.UI.ShowAllFieldResultsItem()
            Me.toggleViewMergedDataItem1 = New DevExpress.XtraRichEdit.UI.ToggleViewMergedDataItem()
            Me.mergeToNewDocumentItem = New DevExpress.XtraBars.BarButtonItem()
            Me.mergeToFileItem = New DevExpress.XtraBars.BarButtonItem()
            Me.changeStyleItem1 = New DevExpress.XtraRichEdit.UI.ChangeStyleItem()
            Me.repositoryItemRichEditStyleEdit1 = New DevExpress.XtraRichEdit.Design.RepositoryItemRichEditStyleEdit()
            Me.pasteItem1 = New DevExpress.XtraRichEdit.UI.PasteItem()
            Me.cutItem1 = New DevExpress.XtraRichEdit.UI.CutItem()
            Me.copyItem1 = New DevExpress.XtraRichEdit.UI.CopyItem()
            Me.pasteSpecialItem1 = New DevExpress.XtraRichEdit.UI.PasteSpecialItem()
            Me.barButtonGroup1 = New DevExpress.XtraBars.BarButtonGroup()
            Me.barButtonGroup2 = New DevExpress.XtraBars.BarButtonGroup()
            Me.toggleFontUnderlineItem1 = New DevExpress.XtraRichEdit.UI.ToggleFontUnderlineItem()
            Me.toggleFontDoubleUnderlineItem1 = New DevExpress.XtraRichEdit.UI.ToggleFontDoubleUnderlineItem()
            Me.toggleFontStrikeoutItem1 = New DevExpress.XtraRichEdit.UI.ToggleFontStrikeoutItem()
            Me.toggleFontDoubleStrikeoutItem1 = New DevExpress.XtraRichEdit.UI.ToggleFontDoubleStrikeoutItem()
            Me.toggleFontSuperscriptItem1 = New DevExpress.XtraRichEdit.UI.ToggleFontSuperscriptItem()
            Me.toggleFontSubscriptItem1 = New DevExpress.XtraRichEdit.UI.ToggleFontSubscriptItem()
            Me.barButtonGroup3 = New DevExpress.XtraBars.BarButtonGroup()
            Me.changeTextCaseItem1 = New DevExpress.XtraRichEdit.UI.ChangeTextCaseItem()
            Me.makeTextUpperCaseItem1 = New DevExpress.XtraRichEdit.UI.MakeTextUpperCaseItem()
            Me.makeTextLowerCaseItem1 = New DevExpress.XtraRichEdit.UI.MakeTextLowerCaseItem()
            Me.toggleTextCaseItem1 = New DevExpress.XtraRichEdit.UI.ToggleTextCaseItem()
            Me.barButtonGroup4 = New DevExpress.XtraBars.BarButtonGroup()
            Me.barButtonGroup5 = New DevExpress.XtraBars.BarButtonGroup()
            Me.barButtonGroup6 = New DevExpress.XtraBars.BarButtonGroup()
            Me.barButtonGroup7 = New DevExpress.XtraBars.BarButtonGroup()
            Me.changeParagraphLineSpacingItem1 = New DevExpress.XtraRichEdit.UI.ChangeParagraphLineSpacingItem()
            Me.setSingleParagraphSpacingItem1 = New DevExpress.XtraRichEdit.UI.SetSingleParagraphSpacingItem()
            Me.setSesquialteralParagraphSpacingItem1 = New DevExpress.XtraRichEdit.UI.SetSesquialteralParagraphSpacingItem()
            Me.setDoubleParagraphSpacingItem1 = New DevExpress.XtraRichEdit.UI.SetDoubleParagraphSpacingItem()
            Me.showLineSpacingFormItem1 = New DevExpress.XtraRichEdit.UI.ShowLineSpacingFormItem()
            Me.addSpacingBeforeParagraphItem1 = New DevExpress.XtraRichEdit.UI.AddSpacingBeforeParagraphItem()
            Me.removeSpacingBeforeParagraphItem1 = New DevExpress.XtraRichEdit.UI.RemoveSpacingBeforeParagraphItem()
            Me.addSpacingAfterParagraphItem1 = New DevExpress.XtraRichEdit.UI.AddSpacingAfterParagraphItem()
            Me.removeSpacingAfterParagraphItem1 = New DevExpress.XtraRichEdit.UI.RemoveSpacingAfterParagraphItem()
            Me.changeParagraphBackColorItem1 = New DevExpress.XtraRichEdit.UI.ChangeParagraphBackColorItem()
            Me.findItem1 = New DevExpress.XtraRichEdit.UI.FindItem()
            Me.replaceItem1 = New DevExpress.XtraRichEdit.UI.ReplaceItem()
            Me.barButtonGroup8 = New DevExpress.XtraBars.BarButtonGroup()
            Me.barButtonGroup9 = New DevExpress.XtraBars.BarButtonGroup()
            Me.barButtonGroup10 = New DevExpress.XtraBars.BarButtonGroup()
            Me.barButtonGroup11 = New DevExpress.XtraBars.BarButtonGroup()
            Me.barButtonGroup12 = New DevExpress.XtraBars.BarButtonGroup()
            Me.barButtonGroup13 = New DevExpress.XtraBars.BarButtonGroup()
            Me.barButtonGroup14 = New DevExpress.XtraBars.BarButtonGroup()
            Me.encryptDocumentItem1 = New DevExpress.XtraRichEdit.UI.EncryptDocumentItem()
            Me.showDocumentPropertiesFormItem1 = New DevExpress.XtraRichEdit.UI.ShowDocumentPropertiesFormItem()
            Me.prevRecordItem = New DevExpress.XtraBars.BarButtonItem()
            Me.nextRecordItem = New DevExpress.XtraBars.BarButtonItem()
            Me.mailingsRibbonPage1 = New DevExpress.XtraRichEdit.UI.MailingsRibbonPage()
            Me.mailMergeRibbonPageGroup1 = New DevExpress.XtraRichEdit.UI.MailMergeRibbonPageGroup()
            Me.ribbonPageGroup3 = New DevExpress.XtraBars.Ribbon.RibbonPageGroup()
            Me.ribbonStatusBar1 = New DevExpress.XtraBars.Ribbon.RibbonStatusBar()
            Me.xtraTabPage2 = New DevExpress.XtraTab.XtraTabPage()
            Me.richEditControl2 = New DevExpress.XtraRichEdit.RichEditControl()
            Me.richEditBarController1 = New DevExpress.XtraRichEdit.UI.RichEditBarController(Me.components)
            Me.showFontFormItem1 = New DevExpress.XtraRichEdit.UI.ShowFontFormItem()
            Me.showEditStyleFormItem1 = New DevExpress.XtraRichEdit.UI.ShowEditStyleFormItem()
            Me.ribbonPageGroup2 = New DevExpress.XtraBars.Ribbon.RibbonPageGroup()
            Me.ribbonPageGroup1 = New DevExpress.XtraBars.Ribbon.RibbonPageGroup()
            CType(Me.xtraTabControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.xtraTabControl1.SuspendLayout()
            Me.xtraTabPage1.SuspendLayout()
            CType(Me.ribbonControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.repositoryItemFontEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.repositoryItemRichEditFontSizeEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.repositoryItemRichEditStyleEdit1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.xtraTabPage2.SuspendLayout()
            CType(Me.richEditBarController1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'galleryChangeStyleItem1
            '
            '
            '
            '
            Me.galleryChangeStyleItem1.Gallery.ColumnCount = 10
            Me.galleryChangeStyleItem1.Gallery.Groups.AddRange(New DevExpress.XtraBars.Ribbon.GalleryItemGroup() {GalleryItemGroup1})
            Me.galleryChangeStyleItem1.Gallery.ImageSize = New System.Drawing.Size(76, 57)
            Me.galleryChangeStyleItem1.Id = 39
            Me.galleryChangeStyleItem1.Name = "galleryChangeStyleItem1"
            '
            'xtraTabControl1
            '
            Me.xtraTabControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.xtraTabControl1.Location = New System.Drawing.Point(0, 193)
            Me.xtraTabControl1.Margin = New System.Windows.Forms.Padding(4)
            Me.xtraTabControl1.Name = "xtraTabControl1"
            Me.xtraTabControl1.SelectedTabPage = Me.xtraTabPage1
            Me.xtraTabControl1.Size = New System.Drawing.Size(924, 474)
            Me.xtraTabControl1.TabIndex = 0
            Me.xtraTabControl1.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.xtraTabPage1, Me.xtraTabPage2})
            '
            'xtraTabPage1
            '
            Me.xtraTabPage1.Controls.Add(Me.richEditControl1)
            Me.xtraTabPage1.Margin = New System.Windows.Forms.Padding(4)
            Me.xtraTabPage1.Name = "xtraTabPage1"
            Me.xtraTabPage1.Size = New System.Drawing.Size(922, 444)
            Me.xtraTabPage1.Text = "Template"
            '
            'richEditControl1
            '
            Me.richEditControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.richEditControl1.Location = New System.Drawing.Point(0, 0)
            Me.richEditControl1.Margin = New System.Windows.Forms.Padding(4)
            Me.richEditControl1.MenuManager = Me.ribbonControl1
            Me.richEditControl1.Name = "richEditControl1"
            Me.richEditControl1.Options.DocumentSaveOptions.CurrentFormat = DevExpress.XtraRichEdit.DocumentFormat.PlainText
            Me.richEditControl1.Size = New System.Drawing.Size(922, 444)
            Me.richEditControl1.TabIndex = 0
            '
            'ribbonControl1
            '
            Me.ribbonControl1.AutoSizeItems = True
            Me.ribbonControl1.EmptyAreaImageOptions.ImagePadding = New System.Windows.Forms.Padding(35, 37, 35, 37)
            Me.ribbonControl1.ExpandCollapseItem.Id = 0
            Me.ribbonControl1.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.ribbonControl1.ExpandCollapseItem, Me.ribbonControl1.SearchEditItem, Me.fileNewItem1, Me.fileOpenItem1, Me.fileSaveItem1, Me.fileSaveAsItem1, Me.quickPrintItem1, Me.printItem1, Me.printPreviewItem1, Me.undoItem1, Me.redoItem1, Me.changeFontNameItem1, Me.changeFontSizeItem1, Me.changeFontColorItem1, Me.changeFontBackColorItem1, Me.toggleFontBoldItem1, Me.toggleFontItalicItem1, Me.fontSizeIncreaseItem1, Me.fontSizeDecreaseItem1, Me.clearFormattingItem1, Me.toggleParagraphAlignmentLeftItem1, Me.toggleParagraphAlignmentCenterItem1, Me.toggleParagraphAlignmentRightItem1, Me.toggleParagraphAlignmentJustifyItem1, Me.toggleNumberingListItem1, Me.toggleBulletedListItem1, Me.toggleMultiLevelListItem1, Me.decreaseIndentItem1, Me.increaseIndentItem1, Me.toggleShowWhitespaceItem1, Me.showParagraphFormItem1, Me.insertMergeFieldItem1, Me.showAllFieldCodesItem1, Me.showAllFieldResultsItem1, Me.toggleViewMergedDataItem1, Me.mergeToNewDocumentItem, Me.mergeToFileItem, Me.changeStyleItem1, Me.pasteItem1, Me.cutItem1, Me.copyItem1, Me.pasteSpecialItem1, Me.barButtonGroup1, Me.barButtonGroup2, Me.toggleFontUnderlineItem1, Me.toggleFontDoubleUnderlineItem1, Me.toggleFontStrikeoutItem1, Me.toggleFontDoubleStrikeoutItem1, Me.toggleFontSuperscriptItem1, Me.toggleFontSubscriptItem1, Me.barButtonGroup3, Me.changeTextCaseItem1, Me.makeTextUpperCaseItem1, Me.makeTextLowerCaseItem1, Me.toggleTextCaseItem1, Me.barButtonGroup4, Me.barButtonGroup5, Me.barButtonGroup6, Me.barButtonGroup7, Me.changeParagraphLineSpacingItem1, Me.setSingleParagraphSpacingItem1, Me.setSesquialteralParagraphSpacingItem1, Me.setDoubleParagraphSpacingItem1, Me.showLineSpacingFormItem1, Me.addSpacingBeforeParagraphItem1, Me.removeSpacingBeforeParagraphItem1, Me.addSpacingAfterParagraphItem1, Me.removeSpacingAfterParagraphItem1, Me.changeParagraphBackColorItem1, Me.findItem1, Me.replaceItem1, Me.galleryChangeStyleItem1, Me.barButtonGroup8, Me.barButtonGroup9, Me.barButtonGroup10, Me.barButtonGroup11, Me.barButtonGroup12, Me.barButtonGroup13, Me.barButtonGroup14, Me.encryptDocumentItem1, Me.showDocumentPropertiesFormItem1, Me.prevRecordItem, Me.nextRecordItem})
            Me.ribbonControl1.Location = New System.Drawing.Point(0, 0)
            Me.ribbonControl1.Margin = New System.Windows.Forms.Padding(4)
            Me.ribbonControl1.MaxItemId = 51
            Me.ribbonControl1.Name = "ribbonControl1"
            Me.ribbonControl1.OptionsMenuMinWidth = 385
            Me.ribbonControl1.Pages.AddRange(New DevExpress.XtraBars.Ribbon.RibbonPage() {Me.mailingsRibbonPage1})
            Me.ribbonControl1.RepositoryItems.AddRange(New DevExpress.XtraEditors.Repository.RepositoryItem() {Me.repositoryItemFontEdit1, Me.repositoryItemRichEditFontSizeEdit1, Me.repositoryItemRichEditStyleEdit1})
            Me.ribbonControl1.Size = New System.Drawing.Size(924, 193)
            Me.ribbonControl1.StatusBar = Me.ribbonStatusBar1
            '
            'fileNewItem1
            '
            Me.fileNewItem1.Id = 0
            Me.fileNewItem1.Name = "fileNewItem1"
            '
            'fileOpenItem1
            '
            Me.fileOpenItem1.Id = 1
            Me.fileOpenItem1.Name = "fileOpenItem1"
            '
            'fileSaveItem1
            '
            Me.fileSaveItem1.Id = 2
            Me.fileSaveItem1.Name = "fileSaveItem1"
            '
            'fileSaveAsItem1
            '
            Me.fileSaveAsItem1.Id = 3
            Me.fileSaveAsItem1.Name = "fileSaveAsItem1"
            '
            'quickPrintItem1
            '
            Me.quickPrintItem1.Id = 4
            Me.quickPrintItem1.Name = "quickPrintItem1"
            '
            'printItem1
            '
            Me.printItem1.Id = 5
            Me.printItem1.Name = "printItem1"
            '
            'printPreviewItem1
            '
            Me.printPreviewItem1.Id = 6
            Me.printPreviewItem1.Name = "printPreviewItem1"
            '
            'undoItem1
            '
            Me.undoItem1.Id = 7
            Me.undoItem1.Name = "undoItem1"
            '
            'redoItem1
            '
            Me.redoItem1.Id = 8
            Me.redoItem1.Name = "redoItem1"
            '
            'changeFontNameItem1
            '
            Me.changeFontNameItem1.Edit = Me.repositoryItemFontEdit1
            Me.changeFontNameItem1.EditWidth = 150
            Me.changeFontNameItem1.Id = 9
            Me.changeFontNameItem1.Name = "changeFontNameItem1"
            '
            'repositoryItemFontEdit1
            '
            Me.repositoryItemFontEdit1.AutoHeight = False
            Me.repositoryItemFontEdit1.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
            Me.repositoryItemFontEdit1.Name = "repositoryItemFontEdit1"
            '
            'changeFontSizeItem1
            '
            Me.changeFontSizeItem1.Edit = Me.repositoryItemRichEditFontSizeEdit1
            Me.changeFontSizeItem1.Id = 10
            Me.changeFontSizeItem1.Name = "changeFontSizeItem1"
            '
            'repositoryItemRichEditFontSizeEdit1
            '
            Me.repositoryItemRichEditFontSizeEdit1.AutoHeight = False
            Me.repositoryItemRichEditFontSizeEdit1.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
            Me.repositoryItemRichEditFontSizeEdit1.Control = Me.richEditControl1
            Me.repositoryItemRichEditFontSizeEdit1.Name = "repositoryItemRichEditFontSizeEdit1"
            '
            'changeFontColorItem1
            '
            Me.changeFontColorItem1.Id = 11
            Me.changeFontColorItem1.Name = "changeFontColorItem1"
            '
            'changeFontBackColorItem1
            '
            Me.changeFontBackColorItem1.Id = 12
            Me.changeFontBackColorItem1.Name = "changeFontBackColorItem1"
            '
            'toggleFontBoldItem1
            '
            Me.toggleFontBoldItem1.Id = 13
            Me.toggleFontBoldItem1.Name = "toggleFontBoldItem1"
            '
            'toggleFontItalicItem1
            '
            Me.toggleFontItalicItem1.Id = 14
            Me.toggleFontItalicItem1.Name = "toggleFontItalicItem1"
            '
            'fontSizeIncreaseItem1
            '
            Me.fontSizeIncreaseItem1.Id = 21
            Me.fontSizeIncreaseItem1.Name = "fontSizeIncreaseItem1"
            '
            'fontSizeDecreaseItem1
            '
            Me.fontSizeDecreaseItem1.Id = 22
            Me.fontSizeDecreaseItem1.Name = "fontSizeDecreaseItem1"
            '
            'clearFormattingItem1
            '
            Me.clearFormattingItem1.Id = 23
            Me.clearFormattingItem1.Name = "clearFormattingItem1"
            '
            'toggleParagraphAlignmentLeftItem1
            '
            Me.toggleParagraphAlignmentLeftItem1.Id = 25
            Me.toggleParagraphAlignmentLeftItem1.Name = "toggleParagraphAlignmentLeftItem1"
            '
            'toggleParagraphAlignmentCenterItem1
            '
            Me.toggleParagraphAlignmentCenterItem1.Id = 26
            Me.toggleParagraphAlignmentCenterItem1.Name = "toggleParagraphAlignmentCenterItem1"
            '
            'toggleParagraphAlignmentRightItem1
            '
            Me.toggleParagraphAlignmentRightItem1.Id = 27
            Me.toggleParagraphAlignmentRightItem1.Name = "toggleParagraphAlignmentRightItem1"
            '
            'toggleParagraphAlignmentJustifyItem1
            '
            Me.toggleParagraphAlignmentJustifyItem1.Id = 28
            Me.toggleParagraphAlignmentJustifyItem1.Name = "toggleParagraphAlignmentJustifyItem1"
            '
            'toggleNumberingListItem1
            '
            Me.toggleNumberingListItem1.Id = 29
            Me.toggleNumberingListItem1.Name = "toggleNumberingListItem1"
            '
            'toggleBulletedListItem1
            '
            Me.toggleBulletedListItem1.Id = 30
            Me.toggleBulletedListItem1.Name = "toggleBulletedListItem1"
            '
            'toggleMultiLevelListItem1
            '
            Me.toggleMultiLevelListItem1.Id = 31
            Me.toggleMultiLevelListItem1.Name = "toggleMultiLevelListItem1"
            '
            'decreaseIndentItem1
            '
            Me.decreaseIndentItem1.Id = 32
            Me.decreaseIndentItem1.Name = "decreaseIndentItem1"
            '
            'increaseIndentItem1
            '
            Me.increaseIndentItem1.Id = 33
            Me.increaseIndentItem1.Name = "increaseIndentItem1"
            '
            'toggleShowWhitespaceItem1
            '
            Me.toggleShowWhitespaceItem1.Id = 34
            Me.toggleShowWhitespaceItem1.Name = "toggleShowWhitespaceItem1"
            '
            'showParagraphFormItem1
            '
            Me.showParagraphFormItem1.Id = 35
            Me.showParagraphFormItem1.Name = "showParagraphFormItem1"
            '
            'insertMergeFieldItem1
            '
            Me.insertMergeFieldItem1.Id = 36
            Me.insertMergeFieldItem1.Name = "insertMergeFieldItem1"
            '
            'showAllFieldCodesItem1
            '
            Me.showAllFieldCodesItem1.Id = 37
            Me.showAllFieldCodesItem1.Name = "showAllFieldCodesItem1"
            '
            'showAllFieldResultsItem1
            '
            Me.showAllFieldResultsItem1.Id = 38
            Me.showAllFieldResultsItem1.Name = "showAllFieldResultsItem1"
            '
            'toggleViewMergedDataItem1
            '
            Me.toggleViewMergedDataItem1.Id = 39
            Me.toggleViewMergedDataItem1.Name = "toggleViewMergedDataItem1"
            '
            'mergeToNewDocumentItem
            '
            Me.mergeToNewDocumentItem.Caption = "Merge to New Document"
            Me.mergeToNewDocumentItem.Id = 40
            Me.mergeToNewDocumentItem.ImageOptions.SvgImage = CType(resources.GetObject("mergeToNewDocumentItem.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
            Me.mergeToNewDocumentItem.Name = "mergeToNewDocumentItem"
            '
            'mergeToFileItem
            '
            Me.mergeToFileItem.Caption = "Merge to File"
            Me.mergeToFileItem.Id = 41
            Me.mergeToFileItem.ImageOptions.SvgImage = CType(resources.GetObject("mergeToFileItem.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
            Me.mergeToFileItem.Name = "mergeToFileItem"
            '
            'changeStyleItem1
            '
            Me.changeStyleItem1.Edit = Me.repositoryItemRichEditStyleEdit1
            Me.changeStyleItem1.Id = 2
            Me.changeStyleItem1.Name = "changeStyleItem1"
            '
            'repositoryItemRichEditStyleEdit1
            '
            Me.repositoryItemRichEditStyleEdit1.AutoHeight = False
            Me.repositoryItemRichEditStyleEdit1.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
            Me.repositoryItemRichEditStyleEdit1.Control = Me.richEditControl1
            Me.repositoryItemRichEditStyleEdit1.Name = "repositoryItemRichEditStyleEdit1"
            '
            'pasteItem1
            '
            Me.pasteItem1.Id = 13
            Me.pasteItem1.Name = "pasteItem1"
            '
            'cutItem1
            '
            Me.cutItem1.Id = 14
            Me.cutItem1.Name = "cutItem1"
            '
            'copyItem1
            '
            Me.copyItem1.Id = 15
            Me.copyItem1.Name = "copyItem1"
            '
            'pasteSpecialItem1
            '
            Me.pasteSpecialItem1.Id = 16
            Me.pasteSpecialItem1.Name = "pasteSpecialItem1"
            '
            'barButtonGroup1
            '
            Me.barButtonGroup1.Id = 6
            Me.barButtonGroup1.ItemLinks.Add(Me.changeFontNameItem1, "FF")
            Me.barButtonGroup1.ItemLinks.Add(Me.changeFontSizeItem1)
            Me.barButtonGroup1.ItemLinks.Add(Me.fontSizeIncreaseItem1, "FG")
            Me.barButtonGroup1.ItemLinks.Add(Me.fontSizeDecreaseItem1, "FK")
            Me.barButtonGroup1.Name = "barButtonGroup1"
            Me.barButtonGroup1.Tag = "{97BBE334-159B-44d9-A168-0411957565E8}"
            '
            'barButtonGroup2
            '
            Me.barButtonGroup2.Id = 7
            Me.barButtonGroup2.ItemLinks.Add(Me.toggleFontBoldItem1)
            Me.barButtonGroup2.ItemLinks.Add(Me.toggleFontItalicItem1)
            Me.barButtonGroup2.ItemLinks.Add(Me.toggleFontUnderlineItem1)
            Me.barButtonGroup2.ItemLinks.Add(Me.toggleFontDoubleUnderlineItem1)
            Me.barButtonGroup2.ItemLinks.Add(Me.toggleFontStrikeoutItem1)
            Me.barButtonGroup2.ItemLinks.Add(Me.toggleFontDoubleStrikeoutItem1)
            Me.barButtonGroup2.ItemLinks.Add(Me.toggleFontSuperscriptItem1)
            Me.barButtonGroup2.ItemLinks.Add(Me.toggleFontSubscriptItem1)
            Me.barButtonGroup2.Name = "barButtonGroup2"
            Me.barButtonGroup2.Tag = "{433DA7F0-03E2-4650-9DB5-66DD92D16E39}"
            '
            'toggleFontUnderlineItem1
            '
            Me.toggleFontUnderlineItem1.Id = 17
            Me.toggleFontUnderlineItem1.Name = "toggleFontUnderlineItem1"
            '
            'toggleFontDoubleUnderlineItem1
            '
            Me.toggleFontDoubleUnderlineItem1.Id = 18
            Me.toggleFontDoubleUnderlineItem1.Name = "toggleFontDoubleUnderlineItem1"
            '
            'toggleFontStrikeoutItem1
            '
            Me.toggleFontStrikeoutItem1.Id = 19
            Me.toggleFontStrikeoutItem1.Name = "toggleFontStrikeoutItem1"
            '
            'toggleFontDoubleStrikeoutItem1
            '
            Me.toggleFontDoubleStrikeoutItem1.Id = 20
            Me.toggleFontDoubleStrikeoutItem1.Name = "toggleFontDoubleStrikeoutItem1"
            '
            'toggleFontSuperscriptItem1
            '
            Me.toggleFontSuperscriptItem1.Id = 21
            Me.toggleFontSuperscriptItem1.Name = "toggleFontSuperscriptItem1"
            '
            'toggleFontSubscriptItem1
            '
            Me.toggleFontSubscriptItem1.Id = 22
            Me.toggleFontSubscriptItem1.Name = "toggleFontSubscriptItem1"
            '
            'barButtonGroup3
            '
            Me.barButtonGroup3.Id = 8
            Me.barButtonGroup3.ItemLinks.Add(Me.changeFontColorItem1, "FC")
            Me.barButtonGroup3.ItemLinks.Add(Me.changeFontBackColorItem1, "I")
            Me.barButtonGroup3.Name = "barButtonGroup3"
            Me.barButtonGroup3.Tag = "{DF8C5334-EDE3-47c9-A42C-FE9A9247E180}"
            '
            'changeTextCaseItem1
            '
            Me.changeTextCaseItem1.Id = 23
            Me.changeTextCaseItem1.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.makeTextUpperCaseItem1), New DevExpress.XtraBars.LinkPersistInfo(Me.makeTextLowerCaseItem1), New DevExpress.XtraBars.LinkPersistInfo(Me.toggleTextCaseItem1)})
            Me.changeTextCaseItem1.Name = "changeTextCaseItem1"
            '
            'makeTextUpperCaseItem1
            '
            Me.makeTextUpperCaseItem1.Id = 24
            Me.makeTextUpperCaseItem1.Name = "makeTextUpperCaseItem1"
            '
            'makeTextLowerCaseItem1
            '
            Me.makeTextLowerCaseItem1.Id = 25
            Me.makeTextLowerCaseItem1.Name = "makeTextLowerCaseItem1"
            '
            'toggleTextCaseItem1
            '
            Me.toggleTextCaseItem1.Id = 26
            Me.toggleTextCaseItem1.Name = "toggleTextCaseItem1"
            '
            'barButtonGroup4
            '
            Me.barButtonGroup4.Id = 9
            Me.barButtonGroup4.ItemLinks.Add(Me.toggleBulletedListItem1, "U")
            Me.barButtonGroup4.ItemLinks.Add(Me.toggleNumberingListItem1, "N")
            Me.barButtonGroup4.ItemLinks.Add(Me.toggleMultiLevelListItem1, "M")
            Me.barButtonGroup4.Name = "barButtonGroup4"
            Me.barButtonGroup4.Tag = "{0B3A7A43-3079-4ce0-83A8-3789F5F6DC9F}"
            '
            'barButtonGroup5
            '
            Me.barButtonGroup5.Id = 10
            Me.barButtonGroup5.ItemLinks.Add(Me.decreaseIndentItem1, "AO")
            Me.barButtonGroup5.ItemLinks.Add(Me.increaseIndentItem1, "AI")
            Me.barButtonGroup5.ItemLinks.Add(Me.toggleShowWhitespaceItem1)
            Me.barButtonGroup5.Name = "barButtonGroup5"
            Me.barButtonGroup5.Tag = "{4747D5AB-2BEB-4ea6-9A1D-8E4FB36F1B40}"
            '
            'barButtonGroup6
            '
            Me.barButtonGroup6.Id = 11
            Me.barButtonGroup6.ItemLinks.Add(Me.toggleParagraphAlignmentLeftItem1, "AL")
            Me.barButtonGroup6.ItemLinks.Add(Me.toggleParagraphAlignmentCenterItem1, "AC")
            Me.barButtonGroup6.ItemLinks.Add(Me.toggleParagraphAlignmentRightItem1, "AR")
            Me.barButtonGroup6.ItemLinks.Add(Me.toggleParagraphAlignmentJustifyItem1, "AJ")
            Me.barButtonGroup6.Name = "barButtonGroup6"
            Me.barButtonGroup6.Tag = "{8E89E775-996E-49a0-AADA-DE338E34732E}"
            '
            'barButtonGroup7
            '
            Me.barButtonGroup7.Id = 12
            Me.barButtonGroup7.ItemLinks.Add(Me.changeParagraphLineSpacingItem1, "K")
            Me.barButtonGroup7.ItemLinks.Add(Me.changeParagraphBackColorItem1, "H")
            Me.barButtonGroup7.Name = "barButtonGroup7"
            Me.barButtonGroup7.Tag = "{9A8DEAD8-3890-4857-A395-EC625FD02217}"
            '
            'changeParagraphLineSpacingItem1
            '
            Me.changeParagraphLineSpacingItem1.Id = 27
            Me.changeParagraphLineSpacingItem1.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.setSingleParagraphSpacingItem1), New DevExpress.XtraBars.LinkPersistInfo(Me.setSesquialteralParagraphSpacingItem1), New DevExpress.XtraBars.LinkPersistInfo(Me.setDoubleParagraphSpacingItem1), New DevExpress.XtraBars.LinkPersistInfo(Me.showLineSpacingFormItem1), New DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.KeyTip, Me.addSpacingBeforeParagraphItem1, "", False, True, True, 0, Nothing, DevExpress.XtraBars.BarItemPaintStyle.Standard, "B", ""), New DevExpress.XtraBars.LinkPersistInfo(Me.removeSpacingBeforeParagraphItem1), New DevExpress.XtraBars.LinkPersistInfo(Me.addSpacingAfterParagraphItem1), New DevExpress.XtraBars.LinkPersistInfo(Me.removeSpacingAfterParagraphItem1)})
            Me.changeParagraphLineSpacingItem1.Name = "changeParagraphLineSpacingItem1"
            '
            'setSingleParagraphSpacingItem1
            '
            Me.setSingleParagraphSpacingItem1.Id = 28
            Me.setSingleParagraphSpacingItem1.Name = "setSingleParagraphSpacingItem1"
            '
            'setSesquialteralParagraphSpacingItem1
            '
            Me.setSesquialteralParagraphSpacingItem1.Id = 29
            Me.setSesquialteralParagraphSpacingItem1.Name = "setSesquialteralParagraphSpacingItem1"
            '
            'setDoubleParagraphSpacingItem1
            '
            Me.setDoubleParagraphSpacingItem1.Id = 30
            Me.setDoubleParagraphSpacingItem1.Name = "setDoubleParagraphSpacingItem1"
            '
            'showLineSpacingFormItem1
            '
            Me.showLineSpacingFormItem1.Id = 31
            Me.showLineSpacingFormItem1.Name = "showLineSpacingFormItem1"
            '
            'addSpacingBeforeParagraphItem1
            '
            Me.addSpacingBeforeParagraphItem1.Id = 32
            Me.addSpacingBeforeParagraphItem1.Name = "addSpacingBeforeParagraphItem1"
            '
            'removeSpacingBeforeParagraphItem1
            '
            Me.removeSpacingBeforeParagraphItem1.Id = 33
            Me.removeSpacingBeforeParagraphItem1.Name = "removeSpacingBeforeParagraphItem1"
            '
            'addSpacingAfterParagraphItem1
            '
            Me.addSpacingAfterParagraphItem1.Id = 34
            Me.addSpacingAfterParagraphItem1.Name = "addSpacingAfterParagraphItem1"
            '
            'removeSpacingAfterParagraphItem1
            '
            Me.removeSpacingAfterParagraphItem1.Id = 35
            Me.removeSpacingAfterParagraphItem1.Name = "removeSpacingAfterParagraphItem1"
            '
            'changeParagraphBackColorItem1
            '
            Me.changeParagraphBackColorItem1.Id = 36
            Me.changeParagraphBackColorItem1.Name = "changeParagraphBackColorItem1"
            '
            'findItem1
            '
            Me.findItem1.Id = 37
            Me.findItem1.Name = "findItem1"
            '
            'replaceItem1
            '
            Me.replaceItem1.Id = 38
            Me.replaceItem1.Name = "replaceItem1"
            '
            'barButtonGroup8
            '
            Me.barButtonGroup8.Id = 40
            Me.barButtonGroup8.ItemLinks.Add(Me.changeFontNameItem1, "FF")
            Me.barButtonGroup8.ItemLinks.Add(Me.changeFontSizeItem1)
            Me.barButtonGroup8.ItemLinks.Add(Me.fontSizeIncreaseItem1, "FG")
            Me.barButtonGroup8.ItemLinks.Add(Me.fontSizeDecreaseItem1, "FK")
            Me.barButtonGroup8.Name = "barButtonGroup8"
            Me.barButtonGroup8.Tag = "{97BBE334-159B-44d9-A168-0411957565E8}"
            '
            'barButtonGroup9
            '
            Me.barButtonGroup9.Id = 41
            Me.barButtonGroup9.ItemLinks.Add(Me.toggleFontBoldItem1)
            Me.barButtonGroup9.ItemLinks.Add(Me.toggleFontItalicItem1)
            Me.barButtonGroup9.ItemLinks.Add(Me.toggleFontUnderlineItem1)
            Me.barButtonGroup9.ItemLinks.Add(Me.toggleFontDoubleUnderlineItem1)
            Me.barButtonGroup9.ItemLinks.Add(Me.toggleFontStrikeoutItem1)
            Me.barButtonGroup9.ItemLinks.Add(Me.toggleFontDoubleStrikeoutItem1)
            Me.barButtonGroup9.ItemLinks.Add(Me.toggleFontSuperscriptItem1)
            Me.barButtonGroup9.ItemLinks.Add(Me.toggleFontSubscriptItem1)
            Me.barButtonGroup9.Name = "barButtonGroup9"
            Me.barButtonGroup9.Tag = "{433DA7F0-03E2-4650-9DB5-66DD92D16E39}"
            '
            'barButtonGroup10
            '
            Me.barButtonGroup10.Id = 42
            Me.barButtonGroup10.ItemLinks.Add(Me.changeFontColorItem1, "FC")
            Me.barButtonGroup10.ItemLinks.Add(Me.changeFontBackColorItem1, "I")
            Me.barButtonGroup10.Name = "barButtonGroup10"
            Me.barButtonGroup10.Tag = "{DF8C5334-EDE3-47c9-A42C-FE9A9247E180}"
            '
            'barButtonGroup11
            '
            Me.barButtonGroup11.Id = 43
            Me.barButtonGroup11.ItemLinks.Add(Me.toggleBulletedListItem1, "U")
            Me.barButtonGroup11.ItemLinks.Add(Me.toggleNumberingListItem1, "N")
            Me.barButtonGroup11.ItemLinks.Add(Me.toggleMultiLevelListItem1, "M")
            Me.barButtonGroup11.Name = "barButtonGroup11"
            Me.barButtonGroup11.Tag = "{0B3A7A43-3079-4ce0-83A8-3789F5F6DC9F}"
            '
            'barButtonGroup12
            '
            Me.barButtonGroup12.Id = 44
            Me.barButtonGroup12.ItemLinks.Add(Me.decreaseIndentItem1, "AO")
            Me.barButtonGroup12.ItemLinks.Add(Me.increaseIndentItem1, "AI")
            Me.barButtonGroup12.ItemLinks.Add(Me.toggleShowWhitespaceItem1)
            Me.barButtonGroup12.Name = "barButtonGroup12"
            Me.barButtonGroup12.Tag = "{4747D5AB-2BEB-4ea6-9A1D-8E4FB36F1B40}"
            '
            'barButtonGroup13
            '
            Me.barButtonGroup13.Id = 45
            Me.barButtonGroup13.ItemLinks.Add(Me.toggleParagraphAlignmentLeftItem1, "AL")
            Me.barButtonGroup13.ItemLinks.Add(Me.toggleParagraphAlignmentCenterItem1, "AC")
            Me.barButtonGroup13.ItemLinks.Add(Me.toggleParagraphAlignmentRightItem1, "AR")
            Me.barButtonGroup13.ItemLinks.Add(Me.toggleParagraphAlignmentJustifyItem1, "AJ")
            Me.barButtonGroup13.Name = "barButtonGroup13"
            Me.barButtonGroup13.Tag = "{8E89E775-996E-49a0-AADA-DE338E34732E}"
            '
            'barButtonGroup14
            '
            Me.barButtonGroup14.Id = 46
            Me.barButtonGroup14.ItemLinks.Add(Me.changeParagraphLineSpacingItem1, "K")
            Me.barButtonGroup14.ItemLinks.Add(Me.changeParagraphBackColorItem1, "H")
            Me.barButtonGroup14.Name = "barButtonGroup14"
            Me.barButtonGroup14.Tag = "{9A8DEAD8-3890-4857-A395-EC625FD02217}"
            '
            'encryptDocumentItem1
            '
            Me.encryptDocumentItem1.Id = 47
            Me.encryptDocumentItem1.Name = "encryptDocumentItem1"
            '
            'showDocumentPropertiesFormItem1
            '
            Me.showDocumentPropertiesFormItem1.Id = 48
            Me.showDocumentPropertiesFormItem1.Name = "showDocumentPropertiesFormItem1"
            '
            'prevRecordItem
            '
            Me.prevRecordItem.Caption = "Previous Record"
            Me.prevRecordItem.Id = 49
            Me.prevRecordItem.ImageOptions.SvgImage = CType(resources.GetObject("prevRecordItem.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
            Me.prevRecordItem.Name = "prevRecordItem"
            '
            'nextRecordItem
            '
            Me.nextRecordItem.Caption = "Next Record"
            Me.nextRecordItem.Id = 50
            Me.nextRecordItem.ImageOptions.SvgImage = CType(resources.GetObject("nextRecordItem.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
            Me.nextRecordItem.Name = "nextRecordItem"
            '
            'mailingsRibbonPage1
            '
            Me.mailingsRibbonPage1.Groups.AddRange(New DevExpress.XtraBars.Ribbon.RibbonPageGroup() {Me.mailMergeRibbonPageGroup1, Me.ribbonPageGroup3})
            Me.mailingsRibbonPage1.Name = "mailingsRibbonPage1"
            '
            'mailMergeRibbonPageGroup1
            '
            Me.mailMergeRibbonPageGroup1.CaptionButtonVisible = DevExpress.Utils.DefaultBoolean.[False]
            Me.mailMergeRibbonPageGroup1.ItemLinks.Add(Me.insertMergeFieldItem1)
            Me.mailMergeRibbonPageGroup1.ItemLinks.Add(Me.showAllFieldCodesItem1)
            Me.mailMergeRibbonPageGroup1.ItemLinks.Add(Me.showAllFieldResultsItem1)
            Me.mailMergeRibbonPageGroup1.ItemLinks.Add(Me.toggleViewMergedDataItem1, "P")
            Me.mailMergeRibbonPageGroup1.ItemLinks.Add(Me.mergeToNewDocumentItem)
            Me.mailMergeRibbonPageGroup1.ItemLinks.Add(Me.mergeToFileItem)
            Me.mailMergeRibbonPageGroup1.Name = "mailMergeRibbonPageGroup1"
            '
            'ribbonPageGroup3
            '
            Me.ribbonPageGroup3.ItemLinks.Add(Me.prevRecordItem)
            Me.ribbonPageGroup3.ItemLinks.Add(Me.nextRecordItem)
            Me.ribbonPageGroup3.Name = "ribbonPageGroup3"
            Me.ribbonPageGroup3.Text = "Navigation"
            '
            'ribbonStatusBar1
            '
            Me.ribbonStatusBar1.Location = New System.Drawing.Point(0, 667)
            Me.ribbonStatusBar1.Margin = New System.Windows.Forms.Padding(4)
            Me.ribbonStatusBar1.Name = "ribbonStatusBar1"
            Me.ribbonStatusBar1.Ribbon = Me.ribbonControl1
            Me.ribbonStatusBar1.Size = New System.Drawing.Size(924, 30)
            '
            'xtraTabPage2
            '
            Me.xtraTabPage2.Controls.Add(Me.richEditControl2)
            Me.xtraTabPage2.Margin = New System.Windows.Forms.Padding(4)
            Me.xtraTabPage2.Name = "xtraTabPage2"
            Me.xtraTabPage2.Size = New System.Drawing.Size(922, 444)
            Me.xtraTabPage2.Text = "Merged Document"
            '
            'richEditControl2
            '
            Me.richEditControl2.Dock = System.Windows.Forms.DockStyle.Fill
            Me.richEditControl2.Location = New System.Drawing.Point(0, 0)
            Me.richEditControl2.Margin = New System.Windows.Forms.Padding(4)
            Me.richEditControl2.MenuManager = Me.ribbonControl1
            Me.richEditControl2.Name = "richEditControl2"
            Me.richEditControl2.Options.DocumentSaveOptions.CurrentFormat = DevExpress.XtraRichEdit.DocumentFormat.PlainText
            Me.richEditControl2.Size = New System.Drawing.Size(922, 444)
            Me.richEditControl2.TabIndex = 1
            '
            'richEditBarController1
            '
            Me.richEditBarController1.BarItems.Add(Me.fileNewItem1)
            Me.richEditBarController1.BarItems.Add(Me.fileOpenItem1)
            Me.richEditBarController1.BarItems.Add(Me.fileSaveItem1)
            Me.richEditBarController1.BarItems.Add(Me.fileSaveAsItem1)
            Me.richEditBarController1.BarItems.Add(Me.quickPrintItem1)
            Me.richEditBarController1.BarItems.Add(Me.printItem1)
            Me.richEditBarController1.BarItems.Add(Me.printPreviewItem1)
            Me.richEditBarController1.BarItems.Add(Me.undoItem1)
            Me.richEditBarController1.BarItems.Add(Me.redoItem1)
            Me.richEditBarController1.BarItems.Add(Me.changeFontNameItem1)
            Me.richEditBarController1.BarItems.Add(Me.changeFontSizeItem1)
            Me.richEditBarController1.BarItems.Add(Me.changeFontColorItem1)
            Me.richEditBarController1.BarItems.Add(Me.changeFontBackColorItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleFontBoldItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleFontItalicItem1)
            Me.richEditBarController1.BarItems.Add(Me.fontSizeIncreaseItem1)
            Me.richEditBarController1.BarItems.Add(Me.fontSizeDecreaseItem1)
            Me.richEditBarController1.BarItems.Add(Me.clearFormattingItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleParagraphAlignmentLeftItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleParagraphAlignmentCenterItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleParagraphAlignmentRightItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleParagraphAlignmentJustifyItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleNumberingListItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleBulletedListItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleMultiLevelListItem1)
            Me.richEditBarController1.BarItems.Add(Me.decreaseIndentItem1)
            Me.richEditBarController1.BarItems.Add(Me.increaseIndentItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleShowWhitespaceItem1)
            Me.richEditBarController1.BarItems.Add(Me.showParagraphFormItem1)
            Me.richEditBarController1.BarItems.Add(Me.insertMergeFieldItem1)
            Me.richEditBarController1.BarItems.Add(Me.showAllFieldCodesItem1)
            Me.richEditBarController1.BarItems.Add(Me.showAllFieldResultsItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleViewMergedDataItem1)
            Me.richEditBarController1.BarItems.Add(Me.changeStyleItem1)
            Me.richEditBarController1.BarItems.Add(Me.pasteItem1)
            Me.richEditBarController1.BarItems.Add(Me.cutItem1)
            Me.richEditBarController1.BarItems.Add(Me.copyItem1)
            Me.richEditBarController1.BarItems.Add(Me.pasteSpecialItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleFontUnderlineItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleFontDoubleUnderlineItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleFontStrikeoutItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleFontDoubleStrikeoutItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleFontSuperscriptItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleFontSubscriptItem1)
            Me.richEditBarController1.BarItems.Add(Me.changeTextCaseItem1)
            Me.richEditBarController1.BarItems.Add(Me.makeTextUpperCaseItem1)
            Me.richEditBarController1.BarItems.Add(Me.makeTextLowerCaseItem1)
            Me.richEditBarController1.BarItems.Add(Me.toggleTextCaseItem1)
            Me.richEditBarController1.BarItems.Add(Me.showFontFormItem1)
            Me.richEditBarController1.BarItems.Add(Me.changeParagraphLineSpacingItem1)
            Me.richEditBarController1.BarItems.Add(Me.setSingleParagraphSpacingItem1)
            Me.richEditBarController1.BarItems.Add(Me.setSesquialteralParagraphSpacingItem1)
            Me.richEditBarController1.BarItems.Add(Me.setDoubleParagraphSpacingItem1)
            Me.richEditBarController1.BarItems.Add(Me.showLineSpacingFormItem1)
            Me.richEditBarController1.BarItems.Add(Me.addSpacingBeforeParagraphItem1)
            Me.richEditBarController1.BarItems.Add(Me.removeSpacingBeforeParagraphItem1)
            Me.richEditBarController1.BarItems.Add(Me.addSpacingAfterParagraphItem1)
            Me.richEditBarController1.BarItems.Add(Me.removeSpacingAfterParagraphItem1)
            Me.richEditBarController1.BarItems.Add(Me.changeParagraphBackColorItem1)
            Me.richEditBarController1.BarItems.Add(Me.showEditStyleFormItem1)
            Me.richEditBarController1.BarItems.Add(Me.findItem1)
            Me.richEditBarController1.BarItems.Add(Me.replaceItem1)
            Me.richEditBarController1.BarItems.Add(Me.galleryChangeStyleItem1)
            Me.richEditBarController1.BarItems.Add(Me.encryptDocumentItem1)
            Me.richEditBarController1.BarItems.Add(Me.showDocumentPropertiesFormItem1)
            Me.richEditBarController1.Control = Me.richEditControl1
            '
            'showFontFormItem1
            '
            Me.showFontFormItem1.Name = "showFontFormItem1"
            '
            'showEditStyleFormItem1
            '
            Me.showEditStyleFormItem1.Name = "showEditStyleFormItem1"
            '
            'ribbonPageGroup2
            '
            Me.ribbonPageGroup2.ItemLinks.Add(Me.changeFontNameItem1, "FF")
            Me.ribbonPageGroup2.ItemLinks.Add(Me.changeFontSizeItem1)
            Me.ribbonPageGroup2.ItemLinks.Add(Me.changeFontColorItem1, "FC")
            Me.ribbonPageGroup2.ItemLinks.Add(Me.changeFontBackColorItem1, "I")
            Me.ribbonPageGroup2.ItemLinks.Add(Me.toggleFontBoldItem1)
            Me.ribbonPageGroup2.ItemLinks.Add(Me.toggleFontItalicItem1)
            Me.ribbonPageGroup2.ItemLinks.Add(Me.fontSizeIncreaseItem1, "FG")
            Me.ribbonPageGroup2.ItemLinks.Add(Me.fontSizeDecreaseItem1, "FK")
            Me.ribbonPageGroup2.ItemLinks.Add(Me.clearFormattingItem1, "E")
            Me.ribbonPageGroup2.Name = "ribbonPageGroup2"
            Me.ribbonPageGroup2.Text = "Font"
            '
            'ribbonPageGroup1
            '
            Me.ribbonPageGroup1.ItemLinks.Add(Me.fileNewItem1, "N")
            Me.ribbonPageGroup1.ItemLinks.Add(Me.fileOpenItem1, "O")
            Me.ribbonPageGroup1.ItemLinks.Add(Me.fileSaveItem1, "S")
            Me.ribbonPageGroup1.ItemLinks.Add(Me.fileSaveAsItem1, "A")
            Me.ribbonPageGroup1.ItemLinks.Add(Me.quickPrintItem1)
            Me.ribbonPageGroup1.ItemLinks.Add(Me.printItem1, "P")
            Me.ribbonPageGroup1.ItemLinks.Add(Me.printPreviewItem1)
            Me.ribbonPageGroup1.ItemLinks.Add(Me.undoItem1)
            Me.ribbonPageGroup1.ItemLinks.Add(Me.redoItem1)
            Me.ribbonPageGroup1.Name = "ribbonPageGroup1"
            Me.ribbonPageGroup1.Text = "Common"
            '
            'Form1
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(924, 697)
            Me.Controls.Add(Me.xtraTabControl1)
            Me.Controls.Add(Me.ribbonStatusBar1)
            Me.Controls.Add(Me.ribbonControl1)
            Me.Margin = New System.Windows.Forms.Padding(4)
            Me.Name = "Form1"
            Me.Ribbon = Me.ribbonControl1
            Me.StatusBar = Me.ribbonStatusBar1
            Me.Text = "Form1"
            CType(Me.xtraTabControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.xtraTabControl1.ResumeLayout(False)
            Me.xtraTabPage1.ResumeLayout(False)
            CType(Me.ribbonControl1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.repositoryItemFontEdit1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.repositoryItemRichEditFontSizeEdit1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.repositoryItemRichEditStyleEdit1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.xtraTabPage2.ResumeLayout(False)
            CType(Me.richEditBarController1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub

        Private WithEvents xtraTabControl1 As DevExpress.XtraTab.XtraTabControl
        Private xtraTabPage1 As DevExpress.XtraTab.XtraTabPage
        Private xtraTabPage2 As DevExpress.XtraTab.XtraTabPage
        Private richEditControl1 As DevExpress.XtraRichEdit.RichEditControl
        Private fileNewItem1 As DevExpress.XtraRichEdit.UI.FileNewItem
        Private fileOpenItem1 As DevExpress.XtraRichEdit.UI.FileOpenItem
        Private fileSaveItem1 As DevExpress.XtraRichEdit.UI.FileSaveItem
        Private fileSaveAsItem1 As DevExpress.XtraRichEdit.UI.FileSaveAsItem
        Private quickPrintItem1 As DevExpress.XtraRichEdit.UI.QuickPrintItem
        Private printItem1 As DevExpress.XtraRichEdit.UI.PrintItem
        Private printPreviewItem1 As DevExpress.XtraRichEdit.UI.PrintPreviewItem
        Private undoItem1 As DevExpress.XtraRichEdit.UI.UndoItem
        Private redoItem1 As DevExpress.XtraRichEdit.UI.RedoItem
        Private changeFontNameItem1 As DevExpress.XtraRichEdit.UI.ChangeFontNameItem
        Private changeFontSizeItem1 As DevExpress.XtraRichEdit.UI.ChangeFontSizeItem
        Private changeFontColorItem1 As DevExpress.XtraRichEdit.UI.ChangeFontColorItem
        Private changeFontBackColorItem1 As DevExpress.XtraRichEdit.UI.ChangeFontBackColorItem
        Private toggleFontBoldItem1 As DevExpress.XtraRichEdit.UI.ToggleFontBoldItem
        Private toggleFontItalicItem1 As DevExpress.XtraRichEdit.UI.ToggleFontItalicItem
        Private fontSizeIncreaseItem1 As DevExpress.XtraRichEdit.UI.FontSizeIncreaseItem
        Private fontSizeDecreaseItem1 As DevExpress.XtraRichEdit.UI.FontSizeDecreaseItem
        Private clearFormattingItem1 As DevExpress.XtraRichEdit.UI.ClearFormattingItem
        Private toggleParagraphAlignmentLeftItem1 As DevExpress.XtraRichEdit.UI.ToggleParagraphAlignmentLeftItem
        Private toggleParagraphAlignmentCenterItem1 As DevExpress.XtraRichEdit.UI.ToggleParagraphAlignmentCenterItem
        Private toggleParagraphAlignmentRightItem1 As DevExpress.XtraRichEdit.UI.ToggleParagraphAlignmentRightItem
        Private toggleParagraphAlignmentJustifyItem1 As DevExpress.XtraRichEdit.UI.ToggleParagraphAlignmentJustifyItem
        Private toggleNumberingListItem1 As DevExpress.XtraRichEdit.UI.ToggleNumberingListItem
        Private toggleBulletedListItem1 As DevExpress.XtraRichEdit.UI.ToggleBulletedListItem
        Private toggleMultiLevelListItem1 As DevExpress.XtraRichEdit.UI.ToggleMultiLevelListItem
        Private decreaseIndentItem1 As DevExpress.XtraRichEdit.UI.DecreaseIndentItem
        Private increaseIndentItem1 As DevExpress.XtraRichEdit.UI.IncreaseIndentItem
        Private toggleShowWhitespaceItem1 As DevExpress.XtraRichEdit.UI.ToggleShowWhitespaceItem
        Private showParagraphFormItem1 As DevExpress.XtraRichEdit.UI.ShowParagraphFormItem
        Private insertMergeFieldItem1 As DevExpress.XtraRichEdit.UI.InsertMergeFieldItem
        Private showAllFieldCodesItem1 As DevExpress.XtraRichEdit.UI.ShowAllFieldCodesItem
        Private showAllFieldResultsItem1 As DevExpress.XtraRichEdit.UI.ShowAllFieldResultsItem
        Private toggleViewMergedDataItem1 As DevExpress.XtraRichEdit.UI.ToggleViewMergedDataItem
        Private richEditControl2 As DevExpress.XtraRichEdit.RichEditControl
        Private richEditBarController1 As DevExpress.XtraRichEdit.UI.RichEditBarController
        Private WithEvents mergeToNewDocumentItem As DevExpress.XtraBars.BarButtonItem
        Private WithEvents mergeToFileItem As DevExpress.XtraBars.BarButtonItem
        Private repositoryItemFontEdit1 As DevExpress.XtraEditors.Repository.RepositoryItemFontEdit
        Private repositoryItemRichEditFontSizeEdit1 As DevExpress.XtraRichEdit.Design.RepositoryItemRichEditFontSizeEdit
        Private ribbonControl1 As DevExpress.XtraBars.Ribbon.RibbonControl
        Private ribbonStatusBar1 As DevExpress.XtraBars.Ribbon.RibbonStatusBar
        Private ribbonPageGroup2 As DevExpress.XtraBars.Ribbon.RibbonPageGroup
        Private ribbonPageGroup1 As DevExpress.XtraBars.Ribbon.RibbonPageGroup
        Private mailingsRibbonPage1 As DevExpress.XtraRichEdit.UI.MailingsRibbonPage
        Private mailMergeRibbonPageGroup1 As DevExpress.XtraRichEdit.UI.MailMergeRibbonPageGroup
        Private changeStyleItem1 As DevExpress.XtraRichEdit.UI.ChangeStyleItem
        Private repositoryItemRichEditStyleEdit1 As DevExpress.XtraRichEdit.Design.RepositoryItemRichEditStyleEdit
        Private pasteItem1 As DevExpress.XtraRichEdit.UI.PasteItem
        Private cutItem1 As DevExpress.XtraRichEdit.UI.CutItem
        Private copyItem1 As DevExpress.XtraRichEdit.UI.CopyItem
        Private pasteSpecialItem1 As DevExpress.XtraRichEdit.UI.PasteSpecialItem
        Private barButtonGroup1 As DevExpress.XtraBars.BarButtonGroup
        Private barButtonGroup2 As DevExpress.XtraBars.BarButtonGroup
        Private toggleFontUnderlineItem1 As DevExpress.XtraRichEdit.UI.ToggleFontUnderlineItem
        Private toggleFontDoubleUnderlineItem1 As DevExpress.XtraRichEdit.UI.ToggleFontDoubleUnderlineItem
        Private toggleFontStrikeoutItem1 As DevExpress.XtraRichEdit.UI.ToggleFontStrikeoutItem
        Private toggleFontDoubleStrikeoutItem1 As DevExpress.XtraRichEdit.UI.ToggleFontDoubleStrikeoutItem
        Private toggleFontSuperscriptItem1 As DevExpress.XtraRichEdit.UI.ToggleFontSuperscriptItem
        Private toggleFontSubscriptItem1 As DevExpress.XtraRichEdit.UI.ToggleFontSubscriptItem
        Private barButtonGroup3 As DevExpress.XtraBars.BarButtonGroup
        Private changeTextCaseItem1 As DevExpress.XtraRichEdit.UI.ChangeTextCaseItem
        Private makeTextUpperCaseItem1 As DevExpress.XtraRichEdit.UI.MakeTextUpperCaseItem
        Private makeTextLowerCaseItem1 As DevExpress.XtraRichEdit.UI.MakeTextLowerCaseItem
        Private toggleTextCaseItem1 As DevExpress.XtraRichEdit.UI.ToggleTextCaseItem
        Private barButtonGroup4 As DevExpress.XtraBars.BarButtonGroup
        Private barButtonGroup5 As DevExpress.XtraBars.BarButtonGroup
        Private barButtonGroup6 As DevExpress.XtraBars.BarButtonGroup
        Private barButtonGroup7 As DevExpress.XtraBars.BarButtonGroup
        Private changeParagraphLineSpacingItem1 As DevExpress.XtraRichEdit.UI.ChangeParagraphLineSpacingItem
        Private setSingleParagraphSpacingItem1 As DevExpress.XtraRichEdit.UI.SetSingleParagraphSpacingItem
        Private setSesquialteralParagraphSpacingItem1 As DevExpress.XtraRichEdit.UI.SetSesquialteralParagraphSpacingItem
        Private setDoubleParagraphSpacingItem1 As DevExpress.XtraRichEdit.UI.SetDoubleParagraphSpacingItem
        Private showLineSpacingFormItem1 As DevExpress.XtraRichEdit.UI.ShowLineSpacingFormItem
        Private addSpacingBeforeParagraphItem1 As DevExpress.XtraRichEdit.UI.AddSpacingBeforeParagraphItem
        Private removeSpacingBeforeParagraphItem1 As DevExpress.XtraRichEdit.UI.RemoveSpacingBeforeParagraphItem
        Private addSpacingAfterParagraphItem1 As DevExpress.XtraRichEdit.UI.AddSpacingAfterParagraphItem
        Private removeSpacingAfterParagraphItem1 As DevExpress.XtraRichEdit.UI.RemoveSpacingAfterParagraphItem
        Private changeParagraphBackColorItem1 As DevExpress.XtraRichEdit.UI.ChangeParagraphBackColorItem
        Private findItem1 As DevExpress.XtraRichEdit.UI.FindItem
        Private replaceItem1 As DevExpress.XtraRichEdit.UI.ReplaceItem
        Private showFontFormItem1 As DevExpress.XtraRichEdit.UI.ShowFontFormItem
        Private showEditStyleFormItem1 As DevExpress.XtraRichEdit.UI.ShowEditStyleFormItem
        Private galleryChangeStyleItem1 As DevExpress.XtraRichEdit.UI.GalleryChangeStyleItem
        Private barButtonGroup8 As DevExpress.XtraBars.BarButtonGroup
        Private barButtonGroup9 As DevExpress.XtraBars.BarButtonGroup
        Private barButtonGroup10 As DevExpress.XtraBars.BarButtonGroup
        Private barButtonGroup11 As DevExpress.XtraBars.BarButtonGroup
        Private barButtonGroup12 As DevExpress.XtraBars.BarButtonGroup
        Private barButtonGroup13 As DevExpress.XtraBars.BarButtonGroup
        Private barButtonGroup14 As DevExpress.XtraBars.BarButtonGroup
        Private encryptDocumentItem1 As DevExpress.XtraRichEdit.UI.EncryptDocumentItem
        Private showDocumentPropertiesFormItem1 As DevExpress.XtraRichEdit.UI.ShowDocumentPropertiesFormItem
        Private WithEvents prevRecordItem As DevExpress.XtraBars.BarButtonItem
        Private WithEvents nextRecordItem As DevExpress.XtraBars.BarButtonItem
        Private ribbonPageGroup3 As DevExpress.XtraBars.Ribbon.RibbonPageGroup
    End Class
End Namespace
