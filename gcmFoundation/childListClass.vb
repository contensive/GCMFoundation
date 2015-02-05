Namespace Contensive.Addons.gcmFoundation
    '
    Public Class childListClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = ""
                Dim inS As String = ""
                Dim subS As String = ""
                Dim currentPageID As Integer = CP.Doc.PageId
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim allowChildPages As Boolean = False
                Dim pageLink As String = ""
                Dim pageCaption As String = ""
                Dim recordCount As Integer = 0
                Dim brief As String = ""
                '
                If cs.Open("Page Content", "ID=" & currentPageID, , , "gcmAllowChildListDisplay") Then
                    allowChildPages = cs.GetBoolean("gcmAllowChildListDisplay")
                End If
                cs.Close()
                '
                If allowChildPages Then
                    If cs.Open("Page Content", "(parentID=" & currentPageID & ") and ((AllowInMenus<>0) or (AllowInMenus is null))", "sortOrder", , "id,menuHeadline,headline,name,briefFilename") Then
                        Do While cs.OK()
                            pageLink = CP.Content.GetPageLink(cs.GetInteger("id"))
                            pageCaption = cs.GetText("MenuHeadline")
                            brief = cs.GetTextFile("briefFilename")
                            '
                            If pageCaption = "" Then
                                pageCaption = cs.GetText("Headline")
                                If pageCaption = "" Then
                                    pageCaption = cs.GetText("name")
                                End If
                            End If
                            '
                            If pageCaption.Contains("<br />") Then
                                pageCaption = "<span>" & pageCaption.Replace("<br />", "</span>")
                            End If
                            '
                            subS = CP.Html.h4(pageCaption, , "normal")
                            subS += brief
                            subS += "<a class=""cross-link"" href=""" & pageLink & """>LEARN MORE <span>f</span></a>"
                            '
                            inS += CP.Html.li(subS)
                            '
                            cs.GoNext()
                            recordCount += 1
                            If (recordCount > 3) Or (Not cs.OK) Then
                                s += CP.Html.ul(inS, , "call-out")
                                recordCount = 0
                                inS = ""
                            End If
                        Loop
                    End If
                    cs.Close()
                End If
                '
                s += CP.Html.ul(inS, , "call-out")
                s += "<!-- /child list -->"
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.childListClass.execute")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
    End Class
    '
End Namespace