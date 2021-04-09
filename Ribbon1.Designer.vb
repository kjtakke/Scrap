Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Email_Scrape = Me.Factory.CreateRibbonTab
        Me.Files = Me.Factory.CreateRibbonGroup
        Me.SingleCSV = Me.Factory.CreateRibbonButton
        Me.MultipleJSON = Me.Factory.CreateRibbonButton
        Me.Attachments = Me.Factory.CreateRibbonGroup
        Me.AllAttachments = Me.Factory.CreateRibbonButton
        Me.JSONAndAttachmens = Me.Factory.CreateRibbonButton
        Me.SaveEmails = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Email_Scrape.SuspendLayout()
        Me.Files.SuspendLayout()
        Me.Attachments.SuspendLayout()
        Me.SuspendLayout()
        '
        'Email_Scrape
        '
        Me.Email_Scrape.Groups.Add(Me.Files)
        Me.Email_Scrape.Groups.Add(Me.Attachments)
        Me.Email_Scrape.Label = "Email Scrape"
        Me.Email_Scrape.Name = "Email_Scrape"
        '
        'Files
        '
        Me.Files.Items.Add(Me.SingleCSV)
        Me.Files.Items.Add(Me.MultipleJSON)
        Me.Files.Items.Add(Me.SaveEmails)
        Me.Files.Label = "Files"
        Me.Files.Name = "Files"
        '
        'SingleCSV
        '
        Me.SingleCSV.Label = "CSV Metadata"
        Me.SingleCSV.Name = "SingleCSV"
        '
        'MultipleJSON
        '
        Me.MultipleJSON.Label = "MultipleJSON"
        Me.MultipleJSON.Name = "MultipleJSON"
        '
        'Attachments
        '
        Me.Attachments.Items.Add(Me.AllAttachments)
        Me.Attachments.Items.Add(Me.JSONAndAttachmens)
        Me.Attachments.Items.Add(Me.Button1)
        Me.Attachments.Label = "Attachments"
        Me.Attachments.Name = "Attachments"
        '
        'AllAttachments
        '
        Me.AllAttachments.Label = "All Attachments"
        Me.AllAttachments.Name = "AllAttachments"
        '
        'JSONAndAttachmens
        '
        Me.JSONAndAttachmens.Label = "JSON And Attachmens"
        Me.JSONAndAttachmens.Name = "JSONAndAttachmens"
        '
        'SaveEmails
        '
        Me.SaveEmails.Label = "Text (.txt)"
        Me.SaveEmails.Name = "SaveEmails"
        '
        'Button1
        '
        Me.Button1.Label = "Text With Attachments"
        Me.Button1.Name = "Button1"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read"
        Me.Tabs.Add(Me.Email_Scrape)
        Me.Email_Scrape.ResumeLayout(False)
        Me.Email_Scrape.PerformLayout()
        Me.Files.ResumeLayout(False)
        Me.Files.PerformLayout()
        Me.Attachments.ResumeLayout(False)
        Me.Attachments.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Email_Scrape As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Files As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Attachments As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents SingleCSV As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents MultipleJSON As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AllAttachments As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents JSONAndAttachmens As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SaveEmails As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
