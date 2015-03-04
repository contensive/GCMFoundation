
Imports System.Exception

Namespace Contensive.Addons

    Public Class commonClass
        Public Function isInGroup(ByVal cp As Contensive.BaseClasses.CPBaseClass, ByVal groupName As String, ByVal userId As Long) As Boolean
            '
            Dim groupId As Integer
            Dim cs As Contensive.BaseClasses.CPCSBaseClass
            Dim sql As String = ""
            '
            groupId = cp.Group.GetId(groupName)
            cs = cp.CSNew()
            sql = "select top 1 id from ccmemberrules where memberid=" & userId & " and groupid=" & groupId
            cs.OpenSQL(sql)
            isInGroup = cs.OK
            Call cs.Close()
            '
        End Function
    End Class
End Namespace