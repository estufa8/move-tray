Imports System

Class GuidList
    Private Sub New()
    End Sub

    Public Const guidMoveTrayPkgString As String = "FD29F876-822A-4F51-9146-6AEF98F4F40F"
    Public Const guidMoveTrayCmdSetString As String = "7E1B1567-CE4B-432A-A76D-01D57F0E4155"
    Public Const guidToolWindowPersistanceString As String = "BB0E258C-6065-432C-99D4-DF86EE8186B4"

    Public Shared ReadOnly guidMoveTrayCmdSet As New Guid(guidMoveTrayCmdSetString)
End Class