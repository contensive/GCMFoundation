
Imports common = Contensive.Addons.gcmFoundation.commonClass

Namespace Contensive.Addons.gcmFoundation
    '
    Public Class corporateSponsorClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = CP.Utils.ExecuteAddon("{AB21FB2C-CB96-43A1-8B19-8DFCAC1FAD2B}")
                '
                s = CP.Html.div(s, , , "gcmSFContainer")
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.corporateSponsorClass.execute")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
    End Class
    '
    Public Class corporateSponsorHandlerClass
        Inherits BaseClasses.AddonBaseClass
        '
        Const formInformation As Integer = 100
        Const formPayment As Integer = 200
        Const formThankYou As Integer = 300
        '
        Const rnFormID As String = "gcmSFFormID"
        Const rnSourceFormID As String = "gcmSFSourceFormID"
        '
        Private hiddenString As String = "" '   accumulates all hiddens from throughout process
        Private errMsg As String = ""
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = ""
                Dim sourceFormID As Integer = CP.Utils.EncodeInteger(CP.Doc.Var(rnSourceFormID))
                Dim formID As Integer = processFormGetID(CP, sourceFormID)
                '
                Select Case formID
                    Case formThankYou
                        s += CP.Content.GetCopy("Corporate Sponsorship Form Thank You", "Thank you for your sponsorship. You're request has been received and will be processed shortly.")
                    Case formPayment
                        s += CP.Content.GetCopy("Corporate Sponsorship Form Instructions - Payment", "Editable Instructions")
                        s += getFormPayment(CP)
                    Case Else
                        formID = formInformation
                        s += CP.Content.GetCopy("Corporate Sponsorship Form Instructions - Information", "Editable Instructions")
                        s += getFormInformation(CP)
                End Select
                '
                hiddenString += CP.Html.Hidden("gcmTargetRecordID", common.getCorporateSponsorshipID(CP), , "gcmTargetRecordID")
                hiddenString += CP.Html.Hidden(rnSourceFormID, formID, , rnFormID)
                '
                s = s.Replace("{{Hidden Form Elements}}", hiddenString)
                '
                s += common.getHelpWrapper(CP, "")
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.corporateSponsorHandlerClass.execute")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
        Private Function processFormGetID(ByVal CP As BaseClasses.CPBaseClass, ByVal sourceFormID As Integer) As Integer
            Try
                Dim formID As Integer = 0
                '
                Select Case sourceFormID
                    Case formPayment
                        If processFormPayment(CP) Then
                            formID = formThankYou
                        Else
                            formID = formPayment
                        End If
                    Case formInformation
                        If processFormInformation(CP) Then
                            formID = formPayment
                        Else
                            formID = formInformation
                        End If
                End Select
                '
                Return formID
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.corporateSponsorHandlerClass.processFormGetID")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
        Private Function getFormInformation(ByVal CP As BaseClasses.CPBaseClass) As String
            Try
                Dim s As String = ""
                Dim sS As String = ""
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim imageFileName As String = ""
                Dim imageLink As String = ""
                Dim slashPosition As Integer = 0
                '
                If cs.Open("Layouts", "ccGUID='{646B2B3B-6D79-4E11-9650-818D528CE38A}'", , , "layout") Then
                    s = cs.GetText("layout")
                End If
                '
                s = s.Replace("{{Sponsorship Select}}", CP.Html.SelectContent("gcmSFLevelID", "", "Corporate Sponsorship Levels", , "Select One", , "gcmSFLevelID"))
                s = s.Replace("{{State Select}}", common.getStateSelect(CP, "gcmSFStateID"))
                s = s.Replace("{{Image Upload}}", common.getUploadField(CP, "gcmSFImage"))
                '
                cs.Open("Corporate Sponsorships", "ID=" & common.getCorporateSponsorshipID(CP))
                If cs.OK() Then
                    imageLink = cs.GetText("imageFileName")
                    '
                    sS += "$(document).ready(function(){" & vbCrLf
                    sS += "$('#gcmSFLevelID').val('" & cs.GetInteger("corporateSponsorshipLevelID") & "');" & vbCrLf
                    sS += "$('#gcmSFFirstName').val('" & CP.Utils.EncodeJavascript(cs.GetText("lastName")) & "');" & vbCrLf
                    sS += "$('#gcmSFLastName').val('" & CP.Utils.EncodeJavascript(cs.GetText("firstName")) & "');" & vbCrLf
                    sS += "$('#gcmSFCompany').val('" & CP.Utils.EncodeJavascript(cs.GetText("companyName")) & "');" & vbCrLf
                    sS += "$('#gcmSFAddress').val('" & CP.Utils.EncodeJavascript(cs.GetText("companyAddress")) & "');" & vbCrLf
                    sS += "$('#gcmSFAddress2').val('" & CP.Utils.EncodeJavascript(cs.GetText("companyAddress2")) & "');" & vbCrLf
                    sS += "$('#gcmSFCity').val('" & CP.Utils.EncodeJavascript(cs.GetText("companyCity")) & "');" & vbCrLf
                    sS += "$('#gcmSFStateID').val('" & CP.Content.GetRecordID("States", cs.GetText("CompanyState")) & "');" & vbCrLf
                    sS += "$('#gcmSFZip').val('" & CP.Utils.EncodeJavascript(cs.GetText("companyZip")) & "');" & vbCrLf
                    sS += "$('#gcmSFEmail').val('" & CP.Utils.EncodeJavascript(cs.GetText("email")) & "');" & vbCrLf
                    sS += "$('#gcmSFPhone').val('" & CP.Utils.EncodeJavascript(cs.GetText("phone")) & "');" & vbCrLf
                    If imageLink <> "" Then
                        slashPosition = InStrRev(imageLink, "/", Len(imageLink), vbTextCompare)
                        imageFileName = Right(imageLink, Len(imageLink) - slashPosition)
                        '
                        sS += "$('#gcmSFImageLink').show();" & vbCrLf
                        sS += "$('#gcmSFImageLink').html('Current File: <a target=""_blank"" href=""" & CP.Site.FilePath & imageLink & """>" & imageFileName & "</a>');" & vbCrLf
                    End If
                    sS += "});" & vbCrLf
                    '
                    CP.Doc.AddHeadJavascript(sS)
                End If
                cs.Close()
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.corporateSponsorHandlerClass.getFormInformation")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
        Private Function processFormInformation(ByVal CP As BaseClasses.CPBaseClass) As Boolean
            Try
                Dim processed As Boolean = True
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                '
                cs.Open("Corporate Sponsorships", "ID=" & common.getCorporateSponsorshipID(CP))
                If cs.OK() Then
                    cs.SetField("name", "Incomplete Sponsorship Form")
                    cs.SetFormInput("corporateSponsorshipLevelID", "gcmSFLevelID")
                    cs.SetFormInput("firstName", "gcmSFFirstName")
                    cs.SetFormInput("lastName", "gcmSFLastName")
                    cs.SetFormInput("companyName", "gcmSFCompany")
                    cs.SetFormInput("companyAddress", "gcmSFAddress")
                    cs.SetFormInput("companyAddress2", "gcmSFAddress2")
                    cs.SetFormInput("companyCity", "gcmSFCity")
                    cs.SetField("companyState", CP.Content.GetRecordName("States", CP.Utils.EncodeInteger(CP.Doc.Var("gcmSFStateID"))))
                    cs.SetFormInput("companyZip", "gcmSFZip")
                    cs.SetFormInput("Email", "gcmSFEmail")
                    cs.SetFormInput("Phone", "gcmSFPhone")
                End If
                cs.Close()
                '
                cs.Open("People", "ID=" & CP.User.Id)
                If cs.OK() Then
                    cs.SetField("name", CP.Doc.Var("gcmSFFirstName") & " " & CP.Doc.Var("gcmSFLastName"))
                    cs.SetFormInput("firstName", "gcmSFFirstName")
                    cs.SetFormInput("lastName", "gcmSFLastName")
                    cs.SetFormInput("company", "gcmSFCompany")
                    cs.SetFormInput("email", "gcmSFEmail")
                    cs.SetFormInput("phone", "gcmSFPhone")
                    cs.SetField("allowBulkEmail", 1)
                End If
                cs.Close()
                '
                Return processed
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.corporateSponsorHandlerClass.processFormInformation")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
        Private Function getFormPayment(ByVal CP As BaseClasses.CPBaseClass) As String
            Try
                Dim s As String = ""
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim corporateSponsorshipType As String = ""
                Dim corporateSponsorshipLevelID As Integer = 0
                '
                If cs.Open("Layouts", "ccGUID='{2F4F9843-1252-4774-AFDC-F00FA296DB24}'", , , "layout") Then
                    s = cs.GetText("layout")
                End If
                '
                If cs.Open("Corporate Sponsorships", "id=" & common.getCorporateSponsorshipID(CP), , , "corporateSponsorshipLevelID") Then
                    corporateSponsorshipLevelID = cs.GetInteger("corporateSponsorshipLevelID")
                End If
                cs.Close()
                '
                hiddenString += CP.Html.Hidden("gcmSFLevelID", corporateSponsorshipLevelID.ToString, , "gcmSFLevelID")
                '
                s = s.Replace("{{Sponsorship Level}}", CP.Content.GetRecordName("Corporate Sponsorship Levels", corporateSponsorshipLevelID))
                s = s.Replace("{{Amount Due}}", FormatCurrency(common.getCorporateSponsorshipAmount(CP, corporateSponsorshipLevelID), 2))
                s = s.Replace("{{Expiration Select}}", common.getExpirationSelect(CP, "gcmSFCardExpMonth", "gcmSFCardExpYear", "gcmSFSelectShort"))
                s = s.Replace("{{Check Instruction}}", CP.Content.GetCopy("Corporate Sponsorship Form Check Payment Instructions", "Please send your checks payable to..."))
                '
                If errMsg <> "" Then
                    s = CP.Html.p(errMsg, , "ccError") & s
                End If
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.corporateSponsorHandlerClass.getFormPayment")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
        Private Function processFormPayment(ByVal CP As BaseClasses.CPBaseClass) As Boolean
            Try
                Dim corporateSponsorshipLevelID As Integer = CP.Utils.EncodeInteger(CP.Doc.Var("gcmSFLevelID"))
                Dim corporateSponsorshipType As String = CP.Content.GetRecordName("Corporate Sponsorship Levels", corporateSponsorshipLevelID)
                Dim corporateSponsorshipAmount As String = common.getCorporateSponsorshipAmount(CP, corporateSponsorshipLevelID).ToString
                Dim sponsorshipID As Integer = common.getCorporateSponsorshipID(CP)
                Dim processed As Boolean = False
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim paymenType As String = CP.Doc.Var("gcmSFPayMethod")
                Dim paymentCredit As Boolean = CP.Utils.EncodeBoolean(paymenType = "Credit Card")
                Dim cardNumber As String = CP.Doc.Var("gcmSFCardNumber")
                Dim rS As String = ""   '   response stream
                Dim resultDoc As New System.Xml.XmlDocument
                Dim ppResponse As String = ""
                Dim ppApproved As Boolean = False
                Dim ppReference As String = ""
                Dim ppCSCMatch As Boolean = False
                Dim ppAVSMatch As Boolean = False
                Dim company As String = ""
                Dim copy As String = ""
                Dim copyPlus As String = ""
                Dim br As String = "<br />"
                Dim recordLink As String = ""
                '
                '   verify payment
                '
                If paymentCredit Then
                    CP.Doc.Var("CreditCardNumber") = cardNumber
                    CP.Doc.Var("CreditCardExpiration") = CP.Doc.Var("gcmSFCardExpMonth") & "/" & CP.Doc.Var("gcmSFCardExpYear")
                    CP.Doc.Var("SecurityCode") = CP.Doc.Var("gcmSFCardCVV")
                    CP.Doc.Var("AVSAddress") = CP.Doc.Var("gcmSFCardAddress")
                    CP.Doc.Var("AVSZip") = CP.Doc.Var("gcmSFCardZip")
                    CP.Doc.Var("PaymentAmount") = corporateSponsorshipAmount
                    '
                    rS = CP.Utils.ExecuteAddon("{F71E8C9B-38A4-446E-8CAC-07548EE602BB}")
                    If rS = "" Then
                        errMsg += "There was a problem processing your payment. Please try again.<br />"
                    Else
                        Call resultDoc.LoadXml(rS)
                        If resultDoc.HasChildNodes Then
                            ppApproved = CP.Utils.EncodeBoolean(resultDoc.GetElementsByTagName("status").Item(0).InnerText)
                            ppResponse = resultDoc.GetElementsByTagName("responseMessage").Item(0).InnerText
                            ppReference = resultDoc.GetElementsByTagName("referenceNumber").Item(0).InnerText
                            ppCSCMatch = CP.Utils.EncodeBoolean(resultDoc.GetElementsByTagName("securityMatchPassed").Item(0).InnerText)
                            ppAVSMatch = CP.Utils.EncodeBoolean(resultDoc.GetElementsByTagName("avsMatchPassed").Item(0).InnerText)
                        End If
                        If (Not ppApproved) Then
                            errMsg += "There was a problem processing your payment. -- " & ppResponse & "<br />"
                        Else
                            processed = True
                        End If
                    End If
                Else
                    processed = True
                End If
                '
                If processed Then
                    '
                    '   complete the sponsorship record
                    '
                    cs.Open("Corporate Sponsorships", "ID=" & sponsorshipID)
                    If cs.OK() Then
                        cs.SetField("name", corporateSponsorshipType & " sponsorship for " & cs.GetText("companyName"))
                        cs.SetField("completed", 1)
                        cs.SetField("dateCompleted", Date.Now)
                        cs.SetField("paymentType", paymenType)
                        company = cs.GetText("companyName")
                    End If
                    cs.Close()
                    '
                    '   add to members group
                    '
                    If paymentCredit Then
                        CP.Group.AddUser("GMIC Sponsor Contacts", CP.User.Id, Date.Today.AddYears(1))
                        CP.Group.AddUser(corporateSponsorshipType, CP.User.Id, Date.Today.AddYears(1))
                    End If
                    '
                    '   send notifications
                    '
                    copy += "Sponsorship Type: " & corporateSponsorshipType & br
                    copy += "Company: " & company & br & br
                    copy += "Sponsorship Amount: " & corporateSponsorshipAmount & br
                    If paymentCredit Then
                        copy += "Amount Paid: " & corporateSponsorshipAmount & br
                        copy += "Payment Date: " & Date.Today & br
                        copy += "Account Number: " & Right(cardNumber, 4) & br
                    Else
                        copy += "Amount Due: " & corporateSponsorshipAmount & br & br
                        copy += CP.Content.GetCopy("Corporate Sponsorship Form Check Payment Instructions", "Please send your checks payable to...")
                    End If
                    '
                    '   add link to email
                    '
                    recordLink = CP.Site.DomainPrimary & "/admin/index.php?af=4&cid=" & CP.Content.GetRecordID("Content", "Corporate Sponsorships")
                    recordLink += "&id=" & sponsorshipID
                    '
                    copyPlus += copy
                    copyPlus += CP.Html.p("<a target=""_blank"" href=""http://" & recordLink & """>Click here for details</a>")
                    '
                    CP.Email.sendSystem("Corporate Sponsorship Auto Responder", copy, CP.User.Id)
                    CP.Email.sendSystem("Corporate Sponsorship Notification", copyPlus)
                End If
                '
                Return processed
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.corporateSponsorHandlerClass.processFormPayment")
                Catch errorObj As Exception
                    '
                    '   if something goes wrong - set user error message
                    '
                    errMsg += "There was a problem processing your payment. Please try again.<br />"
                End Try
            End Try
        End Function
        '
    End Class
    '
    Public Class uploadFileClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = ""
                Dim errMsg As String = ""
                Dim msg As String = ""
                Dim file As String = CP.Doc.Var("gcmSFImage")
                Dim recordID As Integer = CP.Utils.EncodeInteger(CP.Doc.Var("targetRecordID"))
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                '
                If file = "" Then
                    errMsg = "No File detected"
                Else
                    msg = "File upload called"
                End If
                '
                cs.Open("Corporate Sponsorships", "ID=" & recordID)
                If cs.OK() Then
                    cs.SetFormInput("imageFileName", "gcmSFImage")
                End If
                cs.Close()
                '
                s += "{"
                s += "error: '" & errMsg & "',"
                s += "msg: '" & msg & "'"
                s += "}"
                '
                Return s
            Catch ex As Exception
                CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.uploadFileClass.execute")
            End Try
        End Function
    End Class
    '
End Namespace