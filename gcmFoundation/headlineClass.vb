Namespace Contensive.Addons.gcmFoundation
    '
    Public Class headlineClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = ""
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim gcmHeadline As String = ""
                Dim hedlineClass As String = ""
                '
                If cs.Open("Page Content", "id=" & CP.Doc.PageId, , , "gcmHeadline") Then
                    gcmHeadline = cs.GetText("gcmHeadline")
                    '
                    '   some templates use a blue strip vs the default orange
                    '
                    'If CP.Site.GetProperty("Alternate Headline Color In Template ID List").Contains("," & CP.Doc.TemplateId.ToString & ",") Then
                    'hedlineClass = "blue"
                    'End If
                    '
                    '   output headline if set in page record
                    '
                    If gcmHeadline <> "" Then
                        s += CP.Html.h1("<span>" & gcmHeadline & "</span>", , hedlineClass)
                    End If
                End If
                cs.Close()
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.headlineClass.execute")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
    End Class
    '
End Namespace