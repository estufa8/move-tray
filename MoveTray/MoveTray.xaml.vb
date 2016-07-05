Imports System.Security.Permissions
Imports System.ComponentModel.Design
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports System.Text
Imports System.Windows.Forms
Imports System.Windows.Forms.Design
Imports System.Windows

'''<summary>
'''  Interaction logic for MyControl.xaml
'''</summary>
Partial Public Class UtilitatsDabet
    Inherits System.Windows.Controls.UserControl
    Public dte As EnvDTE80.DTE2
    Public design As IDesignerHost


    Private Sub btnMoveTraytoRight_Click(sender As Object, e As Windows.RoutedEventArgs) Handles btnMoveTraytoRight.Click
        If design IsNot Nothing AndAlso dte IsNot Nothing Then
            Dim tray As ComponentTray = design.GetService(GetType(ComponentTray))
            If tray Is Nothing Then Return
            If Not tray.Dock = DockStyle.Right Then tray.Dock = DockStyle.Right
            If tray.Width = 0 Then tray.Width = 150
            tray.PerformLayout()
            'we sort controls alphabetically
            Dim dict As New SortedDictionary(Of String, Control)
            For Each c As Control In tray.Controls
                dict.Add(c.Text, c)
            Next
            Dim y As Integer = 17
            For Each kv In dict 'we have them sorted
                kv.Value.Location = New System.Drawing.Point(17, y)
                y += 39 '39 looks to be the default separator (when sorting with context menu)
            Next
            tray.PerformLayout()
            Dim frame As Control = tray.Parent 'DesignFrame (hidden type) it contains the form, the splitter and the tray
            Dim overlay = frame.Controls(0) 'is DockStyle.Full
            If overlay.Width = 0 Then overlay.Width = 500 'sometime it ocuppies all the screen
            Dim spl As Splitter = frame.Controls(1) 'the splitter
            If Not spl.Dock = DockStyle.Right Then spl.Dock = DockStyle.Right
        Else
            MsgBox("You should focus a form in design mode to use this.")
        End If

    End Sub


End Class

