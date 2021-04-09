Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1


    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub SingleCSV_Click(sender As Object, e As RibbonControlEventArgs) Handles SingleCSV.Click
        Main.CSV()
    End Sub

    Private Sub AllAttachments_Click(sender As Object, e As RibbonControlEventArgs) Handles AllAttachments.Click
        Main.Attachments()
    End Sub

    Private Sub JSONAndAttachmens_Click(sender As Object, e As RibbonControlEventArgs) Handles JSONAndAttachmens.Click
        Main.JSON()
    End Sub

    Private Sub MultipleJSON_Click(sender As Object, e As RibbonControlEventArgs) Handles MultipleJSON.Click
        Main.JSONPlane()
    End Sub

    Private Sub SaveEmails_Click(sender As Object, e As RibbonControlEventArgs) Handles SaveEmails.Click
        Main.save_Emails()
    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Main.save_EmailsWithAttments()
    End Sub
End Class
