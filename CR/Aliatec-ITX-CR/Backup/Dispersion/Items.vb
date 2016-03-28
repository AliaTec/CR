Imports System.Data

Public Class Items
    Inherits DataTable

    Public Sub New()
        Me.Columns.Add("Label")
        Me.Columns.Add("Accion")
        Me.Columns.Add("Case")

        Dim newRow As DataRow

        newRow = Me.NewRow()
        newRow.Item("Label") = 100014
        'accion
        newRow.Item("Accion") = 100014
        'case del aspx
        newRow.Item("Case") = 100014
        Me.Rows.Add(newRow)

      
    End Sub
End Class