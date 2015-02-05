
Imports common = Contensive.Addons.gcmFoundation.commonClass

Namespace Contensive.Addons.gcmFoundation
    '
    Public Class contactListClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = ""
                Dim inS As String = ""
                Dim subS As String = ""
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim groupName As String = CP.Doc.Var("Group Name")
                '
                If cs.OpenGroupUsers(groupName) Then
                    Do While cs.OK()
                        subS = CP.Html.h4(cs.GetText("name"))
                        subS += "<span class=""contactTitle"">" & cs.GetText("title") & "</span>"
                        subS += "<span class=""contactPhone"">" & cs.GetText("phone") & "</span>"
                        subS += "<a onClick=""getStaffContactFormPopUp('" & cs.GetInteger("ID") & "'); return false;"" href=""#"">Contact " & cs.GetText("firstName") & " <span>f</span></a>"
                        '
                        inS += CP.Html.li(subS)
                        cs.GoNext()
                    Loop
                    '
                    s += "<a style=""display:none"" href=""#gcmCFFormContainer"" id=""fakeClicker""></a>"
                    s += CP.Html.div(CP.Html.div(CP.Html.p("The content you requested is currently unavailable.", , "ccError"), , "gcmCFFormContainer", "gcmCFFormContainer"), , "gcmCFWrapper")
                    s += CP.Html.ul(inS, , "gcmCFList")
                End If
                cs.Close()
                '
                s = CP.Html.div(s, , , "gcmCFContainer")
                '
                Return s
            Catch ex As Exception
                CP.Site.ErrorReport(ex, "error in Contensive.Addons.FPI.contactListClass.execute")
            End Try
        End Function
        '
    End Class
    '
    Public Class contactListHandlerClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = ""
                Dim formSumitted As Boolean = CP.Utils.EncodeBoolean(CP.Doc.Var("gcmCFSubmitted"))
                Dim targetMemberID As Integer = CP.Utils.EncodeInteger(CP.Doc.Var("gcmCFTargetMemberID"))
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim emailCopy As String = ""
                Dim hiddenString As String = ""
                '
                If formSumitted Then
                    emailCopy = "<br /><br />From: " & CP.Doc.Var("staffContactEmail") & "<br />"
                    emailCopy += "Message: " & CP.Doc.Var("staffContactMessage") & "<br />"
                    '
                    CP.Email.sendSystem("Contact Notification", emailCopy, targetMemberID)
                    '
                    s = CP.Content.GetCopy("Contact Thank You", "Thank you, your message has been sent.")
                Else
                    If cs.Open("Layouts", "ccGUID='{57228D07-8726-460B-A5C3-084EF4D6D516}'", , , "layout") Then
                        s = cs.GetText("layout")
                    End If
                    cs.Close()
                    '
                    hiddenString = CP.Html.Hidden("gcmCFSubmitted", "1", , "gcmCFSubmitted")
                    hiddenString += CP.Html.Hidden("gcmCFTargetMemberID", targetMemberID, , "gcmCFTargetMemberID")
                    '
                    s = s.Replace("{{Hidden Values}}", hiddenString)
                End If
                '
                s += Common.getHelpWrapper(CP, "")
                '
                Return s
            Catch ex As Exception
                CP.Site.ErrorReport(ex, "error in Contensive.Addons.FPI.contactListHandlerClass.execute")
            End Try
        End Function
        '
    End Class
    '
End Namespace