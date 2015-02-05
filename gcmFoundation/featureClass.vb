Namespace Contensive.Addons.gcmFoundation
    '
    Public Class featureClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = ""
                Dim sS As String = ""
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim lO As BaseClasses.CPBlockBaseClass = CP.BlockNew()
                Dim recordCount As Integer = 0
                '
                If cs.Open("Page Feature Stories") Then
                    sS += "var featureImge = [];" & vbCrLf
                    sS += "var featureCopy = [];" & vbCrLf
                    sS += "var slectedNumber = 0;" & vbCrLf
                    '
                    Do While cs.OK
                        sS += "featureImge[" & recordCount & "] = '" & CP.Site.FilePath & cs.GetText("imageFileName") & "'" & vbCrLf
                        sS += "featureCopy[" & recordCount & "] = '" & cs.GetText("copyFileName") & "'" & vbCrLf
                        '
                        cs.GoNext()
                        recordCount += 1
                    Loop
                    sS += "slectedNumber = Math.ceil(Math.random()*featureImge.length);" & vbCrLf
                    sS += "$('.featured-entry').html(featureCopy[slectedNumber]);" & vbCrLf
                    sS += "$('#featureImage').attr('src', featureImge[slectedNumber]);" & vbCrLf
                    '
                    '   add js to head
                    '
                    sS = "$(document).ready(function() {" & vbCrLf & sS & "});"
                    CP.Doc.AddHeadJavascript(sS)
                End If
                cs.Close()
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.featureClass.execute")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
    End Class
    '
End Namespace