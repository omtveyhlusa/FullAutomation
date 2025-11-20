Imports EPDM.Interop.epdm
Public Class setupPage
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub
    Public Sub loadData(ByRef poCmd As EdmCmd)
        Dim props As IEdmTaskProperties = poCmd.mpoExtra
        If Not props Is Nothing Then
            fPath.Text = props.GetValEx("fPath").ToString
        End If
    End Sub
    Public Sub storeData(ByRef poCmd As EdmCmd)
        Dim props As IEdmTaskProperties = poCmd.mpoExtra
        props.SetValEx("fPath", fPath.Text)
    End Sub
    Private Sub setupPage_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub
End Class
