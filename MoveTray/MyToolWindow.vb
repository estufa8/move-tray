Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Drawing
Imports System.Data
Imports System.Windows
Imports System.Runtime.InteropServices
Imports Microsoft.VisualStudio.Shell.Interop
Imports Microsoft.VisualStudio.Shell
Imports EnvDTE
Imports System.Text.RegularExpressions
Imports System.ComponentModel.Design
Imports System.IO
Imports System.Windows.Forms.Design
Imports System.Windows.Forms

''' <summary>
''' This class implements the tool window exposed by this package and hosts a user control.
'''
''' In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane, 
''' usually implemented by the package implementer.
'''
''' This class derives from the ToolWindowPane class provided from the MPF in order to use its 
''' implementation of the IVsUIElementPane interface.
''' </summary>// 

<Guid(GuidList.guidToolWindowPersistanceString)> _
Public Class MyToolWindow
    '23d9e749-4fd4-416c-af68-727c7d21dc70
    Inherits ToolWindowPane
    Public dte As EnvDTE80.DTE2

    ''' <summary>
    ''' Standard constructor for the tool window.
    ''' </summary>
    Public Sub New()
        MyBase.New(Nothing)
        ' Set the window title reading it from the resources.
        Me.Caption = Resources.ToolWindowTitle
        ' Set the image that will appear on the tab of the window frame
        ' when docked with an other window
        ' The resource ID correspond to the one defined in the resx file
        ' while the Index is the offset in the bitmap strip. Each image in
        ' the strip being 16x16.
        Me.BitmapResourceID = 301
        Me.BitmapIndex = 1

        'This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
        'we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on 
        'the object returned by the Content property.
        Me.Content = New UtilitatsDabet
    End Sub


    Friend Sub sbVentanaActivada(GotFocus As EnvDTE.Window)
        'si es nuestra ventana no hacemos nada
        If GotFocus.ObjectKind.Equals("{" & GuidList.guidToolWindowPersistanceString & "}", StringComparison.CurrentCultureIgnoreCase) Then
            Return
        End If
        Dim design = TryCast(GotFocus.Object, IDesignerHost)

        Dim ventana As UtilitatsDabet = CType(Me.Content, UtilitatsDabet)
        ventana.design = design
        ventana.dte = dte
        'si es un Form en diseño, buscamos comentarios en el archivo con el mismo nombre (tipo Form1.vb)
        If design IsNot Nothing Then
            ventana.grGrid.RowDefinitions(0).Height = New GridLength(1, GridUnitType.Star)
            ventana.grGrid.RowDefinitions(1).Height = New GridLength(0)
        Else
            'ventana.Comentaris = ""
            'ventana.grGrid.RowDefinitions(0).Height = New GridLength(0)
            'ventana.grGrid.RowDefinitions(1).Height = New GridLength(1, GridUnitType.Star)
        End If
    End Sub

End Class
