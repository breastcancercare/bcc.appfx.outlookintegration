Public Class ThisAddIn

    Protected Friend ribbon As BBECRibbon = Nothing ' Allows background worker to reference current instantiated ribbon

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        bgwWebServiceConnect.RunWorkerAsync()
    End Sub

    Private WithEvents bgwWebServiceConnect As New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = False}

    Private Sub bgwWebServiceConnect_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles bgwWebServiceConnect.DoWork
        ' Put CRM login into background so does not lock Outlook startup
        BBECHelper.PerformCRMLogin()
    End Sub

    Private Sub bgwWebServiceConnect_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bgwWebServiceConnect.RunWorkerCompleted
        If Not ribbon Is Nothing Then
            ribbon.SetTabPermission()   'Update ribbon
        End If
    End Sub

End Class
