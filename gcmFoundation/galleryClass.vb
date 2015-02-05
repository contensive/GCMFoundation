
Imports common = Contensive.Addons.gcmFoundation.commonClass

Namespace Contensive.Addons.gcmFoundation
    '
    Public Class galleryClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = ""
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim galleryID As Integer = CP.Utils.EncodeInteger(CP.Doc.Var("Image Gallery"))
                Dim lO As BaseClasses.CPBlockBaseClass = CP.BlockNew()
                Dim galleryThumbnailList As String = ""
                Dim galleryPhotoListInner As String = ""
                Dim galleryPhotoList As String = ""
                Dim galleryTitle As String = ""
                Dim galleryCopy As String = ""
                '
                If cs.Open("Image Galleries", "id=" & galleryID, , , "name,galleryCopy") Then
                    galleryTitle = CP.Html.p("&lt;" & cs.GetText("name") & "&gt;")
                    galleryCopy = cs.GetText("galleryCopy")
                End If
                cs.Close()
                '
                If cs.Open("Gallery Images", "imageGalleryID=" & galleryID, , , "thumbnailFileName,imageFileName,caption") Then
                    Do While cs.OK()
                        galleryThumbnailList += CP.Html.li("<a href=""#""><img src=""" & CP.Site.FilePath & cs.GetText("thumbnailFileName") & """></a>")
                        galleryPhotoListInner = "<img src=""" & CP.Site.FilePath & cs.GetText("imageFileName") & """>"
                        galleryPhotoListInner += CP.Html.div(CP.Html.p("&lt;" & cs.GetText("caption") & "&gt;"), , "caption")
                        galleryPhotoList += CP.Html.li(galleryPhotoListInner)
                        cs.GoNext()
                    Loop
                End If
                cs.Close()
                '
                lO.OpenLayout("{4E55D38C-2EA6-4F2E-8DE0-A5F2ECCDE3D4}")
                lO.SetInner(".galleryThumbnailList", galleryThumbnailList)
                lO.SetInner(".galleryPhotoList", galleryPhotoList)
                lO.SetInner(".gallery-title", galleryTitle)
                lO.SetInner(".gallery-entry", galleryCopy)
                '
                s += lO.GetHtml()
                '
                s += Common.getHelpWrapper(CP, "")
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.galleryClass.execute")
                Catch errObj As Exception
                    '
                End Try
            End Try
        End Function
        '
    End Class
    '
End Namespace