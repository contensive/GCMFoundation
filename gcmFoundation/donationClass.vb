'
Imports common = Contensive.Addons.gcmFoundation.commonClass
'
Namespace Contensive.Addons.gcmFoundation
    '
    Public Class donationClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = CP.Utils.ExecuteAddon("{D82C5786-4BFC-4003-8980-0725B9120572}")
                '
                s = CP.Html.div(s, , , "gcmDFContainer")
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.groupSignupClass.execute")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
    End Class
    '
    Public Class donationHandlerClass
        Inherits BaseClasses.AddonBaseClass
        '
        Private hiddenString As String = "" '   accumulates all hiddens from throughout process
        Private errMsg As String = ""
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = ""
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim processNow As Boolean = CP.Utils.EncodeBoolean(CP.Doc.Var("gcmDFSubmitted"))
                '
                If processNow Then
                    If processForm(CP) Then
                        s += CP.Content.GetCopy("Donation Form Thank You", "Thank you for your donation. Your request has been received and will be processed shortly.")
                    End If
                End If
                '
                If (Not processNow) Or (errMsg <> "") Then
                    If cs.Open("Layouts", "ccGUID='{62DAD87F-A87A-49EA-9C2A-E23E4565C8B6}'", , , "layout") Then
                        s = cs.GetText("layout")
                    End If
                    '
                    If errMsg <> "" Then
                        s = CP.Html.p(errMsg, , "ccError") & s
                    End If
                    '
                    s = s.Replace("{{State Select}}", common.getStateSelect(CP, "gcmDFStateID"))
                    s = s.Replace("{{Expiration Select}}", common.getExpirationSelect(CP, "gcmDFCardExpMonth", "gcmDFCardExpYear", "gcmDFSelectShort"))
                End If
                '
                s += common.getHelpWrapper(CP, "")
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.groupSignupHandlerClass.execute")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
        Private Function processForm(ByVal CP As BaseClasses.CPBaseClass) As Boolean
            Try
                Dim processed As Boolean = False
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim firstName As String = CP.Doc.Var("gcmDFFirstName")
                Dim lastName As String = CP.Doc.Var("gcmDFLastName")
                Dim cardNumber As String = CP.Doc.Var("gcmDFCardNumber")
                Dim cardExpiration As String = CP.Doc.Var("gcmDFCardExpMonth") & "/" & CP.Doc.Var("gcmDFCardExpYear")
                Dim cardCVV As String = CP.Doc.Var("gcmDFCardCVV")
                Dim cardAddress As String = CP.Doc.Var("gcmDFCardAddress")
                Dim cardZip As String = CP.Doc.Var("gcmDFCardZip")
                Dim rS As String = ""   '   response stream
                Dim resultDoc As New System.Xml.XmlDocument
                Dim ppResponse As String = ""
                Dim ppApproved As Boolean = False
                Dim ppReference As String = ""
                Dim ppCSCMatch As Boolean = False
                Dim ppAVSMatch As Boolean = False
                Dim br As String = "<br />"
                Dim copy As String = br & br
                Dim copyPlus As String = ""
                Dim recordLink As String = ""
                Dim amount As Double = CP.Utils.EncodeNumber(CP.Doc.Var("gcmDFAmount"))
                Dim donationID As Integer = 0
                '
                '   verify payment
                '
                CP.Doc.Var("CreditCardNumber") = cardNumber
                CP.Doc.Var("CreditCardExpiration") = cardExpiration
                CP.Doc.Var("SecurityCode") = cardCVV
                CP.Doc.Var("AVSAddress") = cardAddress
                CP.Doc.Var("AVSZip") = cardZip
                CP.Doc.Var("PaymentAmount") = amount
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
                '
                If processed Then
                    '
                    '   insert Donation
                    '
                    If cs.Insert("Donations") Then
                        donationID = cs.GetInteger("id")
                        cs.SetField("name", "Donation made " & Date.Today & " by " & firstName & " " & lastName)
                        cs.SetField("firstName", firstName)
                        cs.SetFormInput("middleName", "gcmDFMiddleName")
                        cs.SetField("lastName", lastName)
                        cs.SetFormInput("address", "gcmDFAddress")
                        cs.SetFormInput("address2", "gcmDFAddress2")
                        cs.SetFormInput("city", "gcmDFCity")
                        cs.SetField("state", CP.Content.GetRecordName("States", CP.Utils.EncodeInteger(CP.Doc.Var("gcmDFStateID"))))
                        cs.SetFormInput("zip", "gcmDFZip")
                        cs.SetFormInput("phone", "gcmDFPhone")
                        cs.SetFormInput("email", "gcmDFEmail")
                        '
                        cs.SetFormInput("amount", "gcmDFAmount")
                        cs.SetFormInput("cardName", "gcmDFCardName")
                        cs.SetField("cardNumber", cardNumber)
                        cs.SetField("cardExpiration", cardExpiration)
                        cs.SetField("cardBillingAddress", cardAddress)
                        cs.SetField("cardBillingZip", cardZip)
                        cs.SetField("cardCode", cardCVV)
                        '
                        cs.SetField("processorReference", ppReference)
                        cs.SetField("processorResponse", ppResponse)
                        '
                        cs.SetField("memberID", CP.User.Id)
                        cs.SetField("visitID", CP.Visit.Id)
                    End If
                    cs.Close()
                    '
                    '   send notifications
                    '
                    copy += "First Name: " & firstName & br
                    copy += "Last Name: " & lastName & br
                    copy += "Total Amount: " & amount & br
                    copy += "Amount Paid: " & amount & br
                    copy += "Payment Date: " & Date.Today & br
                    copy += "Account Number: " & Right(cardNumber, 4) & br
                    '
                    '   add link to email
                    '
                    recordLink = CP.Site.DomainPrimary & "/admin/index.php?af=4&cid=" & CP.Content.GetRecordID("Content", "Donations")
                    recordLink += "&id=" & donationID
                    '
                    copyPlus += copy
                    copyPlus += CP.Html.p("<a target=""_blank"" href=""http://" & recordLink & """>Click here for details</a>")
                    '
                    CP.Email.sendSystem("Dontaion Form Auto Responder", copy, CP.User.Id)
                    CP.Email.sendSystem("Donation Form Notification", copyPlus)
                End If
                '
                Return processed
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.groupSignupHandlerClass.processFormPayment")
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
End Namespace