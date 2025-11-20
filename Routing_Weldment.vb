Imports SolidWorks.Interop.swconst
Imports SolidWorks.Interop.swdocumentmgr
Imports SolidWorks.Interop.sldworks
Imports EPDM.Interop.epdm
Imports System.Net.WebRequestMethods

Public Class Routing_Weldment
    'Sql
    Public dbECN As New SQLdb.SqlDatabase("omt-sql03", "ENG", "eng_uploadtool", "jj*7^++-PPrqW")
    Public dbPDM As New SQLdb.SqlDatabase("omt-epdm-db22", "OMT-Sandbox", "dbadmin", "ePDMadmin")
    Public dbNAV As New SQLdb.SqlDatabase("omt-sql04", "OMTLIVE2015", "eng_uploadtool", "jj*7^++-PPrqW")
    Public sqlstr As String

    'Datatables
    Public dtColor As DataTable
    Public dtMatThickness As DataTable

    Public vault1 As EdmVault5

    Public fts As Object
    Public ftMgr As FeatureManager


    Sub New(ByVal PartNo As String, ByVal swModel As ModelDoc2, afile As IEdmFile12, ByVal swApp As SldWorks)









        ftMgr = swModel.FeatureManager
        'ftMgr.ViewDependencies = False
        'ftMgr.ShowHierarchyOnly = False
        'ftMgr.ShowFeatureName = True
        'ftMgr.ShowComponentNames = True


        fts = ftMgr.GetFeatures(False)
        For Each ft As Feature In fts

            If ft.GetTypeName = "Reference" Then
                Debug.Print("")
                Debug.Print(ft.GetTypeName)
                Debug.Print(ft.GetID)
                Debug.Print(ft.Name)
                Debug.Print(ft.IGetChildCount)
                Debug.Print(ft.GetChildren)
                'ft.GetTypeName()
            End If
        Next



    End Sub
End Class
