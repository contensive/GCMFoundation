Namespace Contensive.Addons.gcmFoundation
    '
    Public Class menuFilerClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = CP.Doc.Body
                '
                s = s.Replace("{5010E357-E363-45AD-85DD-FE5C490B355D}", getSectionLink(CP, "{5010E357-E363-45AD-85DD-FE5C490B355D}"))
                s = s.Replace("{F51AE6FA-0BE6-45A8-B968-1884FD430979}", getSectionLink(CP, "{F51AE6FA-0BE6-45A8-B968-1884FD430979}"))
                s = s.Replace("{6BA258A1-647E-4C64-8FD5-6AC5BF6D1913}", getSectionLink(CP, "{6BA258A1-647E-4C64-8FD5-6AC5BF6D1913}"))
                s = s.Replace("{3225D68D-B897-479F-8455-7DF2A8F517C9}", getSectionLink(CP, "{3225D68D-B897-479F-8455-7DF2A8F517C9}"))
                '
                CP.Doc.Body = s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.menuFilerClass.execute")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
        Private Function getSectionLink(ByVal CP As BaseClasses.CPBaseClass, ByVal sectionGUID As String) As String
            Try
                Dim s As String = ""
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim rootPageLink As String = ""
                Dim caption As String = ""
                '
                If cs.Open("Site Sections", "ccGUID=" & CP.Db.EncodeSQLText(sectionGUID), , , "caption,rootPageID") Then
                    rootPageLink = CP.Content.GetPageLink(cs.GetInteger("rootPageID"))
                    caption = cs.GetText("caption")
                    '
                    If caption.Contains("<br />") Then
                        caption = "<span>" & caption.Replace("<br />", "</span>")
                    End If
                    '
                    s = "<a href=""" & rootPageLink & """>" & caption & "</a>"
                End If
                cs.Close()
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.menuFilerClass.getSectionLink")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
    End Class
    '
End Namespace