
Imports common = Contensive.Addons.gcmFoundation.commonClass

Namespace Contensive.Addons.gcmFoundation
    '
    Public Class membershipJoinClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = CP.Utils.ExecuteAddon("{F8A4C76C-264C-4EF0-8E0C-0AAE2A5B4785}")
                '
                s = CP.Html.div(s, , , "gcmMFContainer")
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.membershipJoinClass.execute")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
    End Class
    '
    Public Class membershipJoinHandlerClass
        Inherits BaseClasses.AddonBaseClass
        '
        Const formInformation As Integer = 100
        Const formPayment As Integer = 200
        Const formThankYou As Integer = 300
        '
        Const rnFormID As String = "gcmMFFormID"
        Const rnSourceFormID As String = "gcmMFSourceFormID"
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
                        s += CP.Content.GetCopy("Membership Join Form Thank You", "Thank you for your interest in membership. You're request has been received and will be processed shortly.")
                    Case formPayment
                        s += CP.Content.GetCopy("Membership Join Form Instructions - Payment", "Editable Instructions")
                        s += getFormPayment(CP)
                    Case Else
                        formID = formInformation
                        s += CP.Content.GetCopy("Membership Join Form Instructions - Information", "Editable Instructions")
                        s += getFormInformation(CP)
                End Select
                '
                hiddenString += CP.Html.Hidden("gcmTargetRecordID", CP.User.Id, , "gcmTargetRecordID")
                hiddenString += CP.Html.Hidden(rnSourceFormID, formID, , rnSourceFormID)
                '
                s = s.Replace("{{Hidden Form Elements}}", hiddenString)
                '
                s += common.getHelpWrapper(CP, "")
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.membershipJoinHandlerClass.execute")
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
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.membershipJoinHandlerClass.processFormGetID")
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
                '
                If cs.Open("Layouts", "ccGUID='{F4F056A8-04D5-4ACE-9F3C-F6EFCB150C15}'", , , "layout") Then
                    s = cs.GetText("layout")
                End If
                '
                s = s.Replace("{{Membership Select}}", CP.Html.SelectContent("gcmMFLevelID", "", "Membership Levels", , "Select One", , "gcmMFLevelID"))
                s = s.Replace("{{State Select}}", common.getStateSelect(CP, "gcmMFStateID"))
                '
                cs.Open("Membership Requests", "ID=" & common.getMembershipID(CP))
                If cs.OK() Then
                    '
                    sS += "$(document).ready(function(){" & vbCrLf
                    sS += "$('#gcmMFLevelID').val('" & cs.GetInteger("membershipLevelID") & "');" & vbCrLf
                    sS += "$('#gcmMFTitle').val('" & CP.Utils.EncodeJavascript(cs.GetText("title")) & "');" & vbCrLf
                    sS += "$('#gcmMFFirstName').val('" & CP.Utils.EncodeJavascript(cs.GetText("lastName")) & "');" & vbCrLf
                    sS += "$('#gcmMFMiddleName').val('" & CP.Utils.EncodeJavascript(cs.GetText("middleName")) & "');" & vbCrLf
                    sS += "$('#gcmMFLastName').val('" & CP.Utils.EncodeJavascript(cs.GetText("firstName")) & "');" & vbCrLf
                    sS += "$('#gcmMFAddress').val('" & CP.Utils.EncodeJavascript(cs.GetText("address")) & "');" & vbCrLf
                    sS += "$('#gcmMFAddress2').val('" & CP.Utils.EncodeJavascript(cs.GetText("address2")) & "');" & vbCrLf
                    sS += "$('#gcmMFCity').val('" & CP.Utils.EncodeJavascript(cs.GetText("city")) & "');" & vbCrLf
                    sS += "$('#gcmMFStateID').val('" & CP.Content.GetRecordID("States", cs.GetText("state")) & "');" & vbCrLf
                    sS += "$('#gcmMFZip').val('" & CP.Utils.EncodeJavascript(cs.GetText("zip")) & "');" & vbCrLf
                    sS += "$('#gcmMFEmail').val('" & CP.Utils.EncodeJavascript(cs.GetText("email")) & "');" & vbCrLf
                    sS += "$('#gcmMFPhone').val('" & CP.Utils.EncodeJavascript(cs.GetText("phone")) & "');" & vbCrLf
                    sS += "});" & vbCrLf
                    '
                    CP.Doc.AddHeadJavascript(sS)
                End If
                cs.Close()
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.membershipJoinHandlerClass.getFormInformation")
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
                cs.Open("Membership Requests", "ID=" & common.getMembershipID(CP))
                If cs.OK() Then
                    cs.SetField("name", "Incomplete Membership Join Form")
                    cs.SetFormInput("membershipLevelID", "gcmMFLevelID")
                    cs.SetFormInput("title", "gcmMFTitle")
                    cs.SetFormInput("firstName", "gcmMFFirstName")
                    cs.SetFormInput("middleName", "gcmMFMiddleName")
                    cs.SetFormInput("lastName", "gcmMFLastName")
                    cs.SetFormInput("address", "gcmMFAddress")
                    cs.SetFormInput("address2", "gcmMFAddress2")
                    cs.SetFormInput("city", "gcmMFCity")
                    cs.SetField("state", CP.Content.GetRecordName("States", CP.Utils.EncodeInteger(CP.Doc.Var("gcmMFStateID"))))
                    cs.SetFormInput("zip", "gcmMFZip")
                    cs.SetFormInput("email", "gcmMFEmail")
                    cs.SetFormInput("phone", "gcmMFPhone")
                End If
                cs.Close()
                '
                cs.Open("People", "ID=" & CP.User.Id)
                If cs.OK() Then
                    cs.SetField("name", CP.Doc.Var("gcmMFFirstName") & " " & CP.Doc.Var("gcmMFLastName"))
                    cs.SetFormInput("title", "gcmMFTitle")
                    cs.SetFormInput("firstName", "gcmMFFirstName")
                    cs.SetFormInput("middleName", "gcmMFMiddleName")
                    cs.SetFormInput("lastName", "gcmMFLastName")
                    cs.SetFormInput("company", "gcmMFCompany")
                    cs.SetFormInput("email", "gcmMFEmail")
                    cs.SetFormInput("phone", "gcmMFPhone")
                    cs.SetField("allowBulkEmail", 1)
                End If
                cs.Close()
                '
                Return processed
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.membershipJoinHandlerClass.processFormInformation")
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
                Dim membershipType As String = ""
                Dim membershipLevelID As Integer = 0
                '
                If cs.Open("Layouts", "ccGUID='{40858AF6-103C-4395-B60F-0EFBF702D703}'", , , "layout") Then
                    s = cs.GetText("layout")
                End If
                '
                If cs.Open("Membership Requests", "id=" & common.getMembershipID(CP), , , "membershipLevelID") Then
                    membershipLevelID = cs.GetInteger("membershipLevelID")
                End If
                cs.Close()
                '
                hiddenString += CP.Html.Hidden("gcmMFLevelID", membershipLevelID.ToString, , "gcmMFLevelID")
                '
                s = s.Replace("{{Membership Level}}", CP.Content.GetRecordName("Membership Levels", membershipLevelID))
                s = s.Replace("{{Amount Due}}", FormatCurrency(common.getMembershipAmount(CP, membershipLevelID), 2))
                s = s.Replace("{{Expiration Select}}", common.getExpirationSelect(CP, "gcmMFCardExpMonth", "gcmMFCardExpYear", "gcmMFSelectShort"))
                s = s.Replace("{{Check Instruction}}", CP.Content.GetCopy("Membership Join Form Check Payment Instructions", "Please send your checks payable to..."))
                '
                If errMsg <> "" Then
                    s = CP.Html.p(errMsg, , "ccError") & s
                End If
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.membershipJoinHandlerClass.getFormPayment")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
        Private Function processFormPayment(ByVal CP As BaseClasses.CPBaseClass) As Boolean
            Try
                Dim membershipLevelID As Integer = CP.Utils.EncodeInteger(CP.Doc.Var("gcmMFLevelID"))
                Dim membershipType As String = CP.Content.GetRecordName("Membership Levels", membershipLevelID)
                Dim membershipAmount As String = common.getMembershipAmount(CP, membershipLevelID).ToString
                Dim membershipID As Integer = common.getMembershipID(CP)
                Dim processed As Boolean = False
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim paymenType As String = CP.Doc.Var("gcmMFPayMethod")
                Dim paymentCredit As Boolean = CP.Utils.EncodeBoolean(paymenType = "Credit Card")
                Dim cardNumber As String = CP.Doc.Var("gcmMFCardNumber")
                Dim rS As String = ""   '   response stream
                Dim resultDoc As New System.Xml.XmlDocument
                Dim ppResponse As String = ""
                Dim ppApproved As Boolean = False
                Dim ppReference As String = ""
                Dim ppCSCMatch As Boolean = False
                Dim ppAVSMatch As Boolean = False
                Dim personName As String = ""
                Dim copy As String = ""
                Dim copyPlus As String = ""
                Dim br As String = "<br />"
                Dim recordLink As String = ""
                Dim targetMemberID As Integer = 0
                '
                '   verify payment
                '
                If paymentCredit Then
                    CP.Doc.Var("CreditCardNumber") = cardNumber
                    CP.Doc.Var("CreditCardExpiration") = CP.Doc.Var("gcmMFCardExpMonth") & "/" & CP.Doc.Var("gcmMFCardExpYear")
                    CP.Doc.Var("SecurityCode") = CP.Doc.Var("gcmMFCardCVV")
                    CP.Doc.Var("AVSAddress") = CP.Doc.Var("gcmMFCardAddress")
                    CP.Doc.Var("AVSZip") = CP.Doc.Var("gcmMFCardZip")
                    CP.Doc.Var("PaymentAmount") = membershipAmount
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
                    '   complete the membership record
                    '
                    cs.Open("Membership Requests", "ID=" & membershipID)
                    If cs.OK() Then
                        personName = cs.GetText("firstName") & " " & cs.GetText("lastName")
                        cs.SetField("name", membershipType & " for " & personName)
                        cs.SetField("completed", 1)
                        cs.SetField("dateCompleted", Date.Now)
                        cs.SetField("paymentType", paymenType)
                        cs.SetField("processorResponse", ppResponse)
                        cs.SetField("processorReference", ppReference)
                        targetMemberID = cs.GetInteger("memberID")
                    End If
                    cs.Close()
                    '
                    '   move record to members if paid by credit card
                    '
                    If paymentCredit Then
                        cs.Open("People", "ID=" & targetMemberID)
                        If cs.OK() Then
                            cs.SetField("", CP.Content.GetRecordID("Content", "Members"))
                        End If
                        cs.Close()
                        '
                        '   add to members group
                        '
                        CP.Group.AddUser("GMIC Members", targetMemberID, Date.Today.AddYears(1))
                    End If
                    '
                    '   send notifications
                    '
                    copy += "Membership Type: " & membershipType & br
                    copy += "Name: " & personName & br & br
                    copy += "Membership Amount: " & membershipAmount & br
                    If paymentCredit Then
                        copy += "Amount Paid: " & membershipAmount & br
                        copy += "Payment Date: " & Date.Today & br
                        copy += "Account Number: " & Right(cardNumber, 4) & br
                    Else
                        copy += "Amount Due: " & membershipAmount & br & br
                        copy += CP.Content.GetCopy("Membership Join Form Check Payment Instructions", "Please send your checks payable to...")
                    End If
                    '
                    '   add link to email
                    '
                    recordLink = CP.Site.DomainPrimary & "/admin/index.php?af=4&cid=" & CP.Content.GetRecordID("Content", "Membership Requests")
                    recordLink += "&id=" & membershipID
                    '
                    copyPlus += copy
                    copyPlus += CP.Html.p("<a target=""_blank"" href=""http://" & recordLink & """>Click here for details</a>")
                    '
                    CP.Email.sendSystem("Membership Join Auto Responder", copy, CP.User.Id)
                    CP.Email.sendSystem("Membership Join Notification", copyPlus)
                End If
                '
                Return processed
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.membershipJoinHandlerClass.processFormPayment")
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