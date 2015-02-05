
Imports common = Contensive.Addons.gcmFoundation.commonClass

Namespace Contensive.Addons.gcmFoundation
    '
    Public Class groupSignupClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = CP.Utils.ExecuteAddon("{D4273013-CADB-4813-87D9-84AB16DA8C24}")
                '
                s = CP.Html.div(s, , , "gcmGTContainer")
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
    Public Class groupSignupHandlerClass
        Inherits BaseClasses.AddonBaseClass
        '
        Const formInformation As Integer = 100
        Const formContact As Integer = 200
        Const formPayment As Integer = 300
        Const formThankYou As Integer = 400
        '
        Const rnFormID As String = "gcmGTFormID"
        Const rnSourceFormID As String = "gcmGTSourceFormID"
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
                        s += CP.Content.GetCopy("Group Tour Form Thank You", "Thank you for your interest in our group tours. Your request has been received and will be processed shortly.")
                    Case formPayment
                        s += CP.Content.GetCopy("Group Tour Form Instructions - Payment", "Editable Instructions")
                        s += getFormPayment(CP)
                    Case formContact
                        s += CP.Content.GetCopy("Group Tour Form Instructions - Contact Information", "Editable Instructions")
                        s += getFormContact(CP)
                    Case Else
                        formID = formInformation
                        s += CP.Content.GetCopy("Group Tour Form Instructions - Information", "Editable Instructions")
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
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.groupSignupHandlerClass.execute")
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
                    Case formContact
                        If processFormContact(CP) Then
                            formID = formPayment
                        Else
                            formID = formContact
                        End If
                    Case formInformation
                        If processFormInformation(CP) Then
                            formID = formContact
                        Else
                            formID = formInformation
                        End If
                End Select
                '
                Return formID
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.groupSignupHandlerClass.processFormGetID")
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
                Dim attendeeCount As Integer = 0
                '
                If cs.Open("Layouts", "ccGUID='{B38CD92A-C315-4C69-95E8-680F0440C554}'", , , "layout") Then
                    s = cs.GetText("layout")
                End If
                '
                cs.Open("Group Tour Requests", "ID=" & common.getGroupTourID(CP))
                If cs.OK() Then
                    sS += "$(document).ready(function(){" & vbCrLf
                    sS += "$('#gcmGTGroupName').val('" & CP.Utils.EncodeJavascript(cs.GetText("name")) & "');" & vbCrLf
                    sS += "$('#gcmGTEventName').val('" & CP.Utils.EncodeJavascript(cs.GetText("eventName")) & "');" & vbCrLf
                    sS += "$('#gcmGTDateA').val('" & CP.Utils.EncodeJavascript(cs.GetText("tourDateA")) & "');" & vbCrLf
                    sS += "$('#gcmGTStartA').val('" & CP.Utils.EncodeJavascript(cs.GetText("tourStartA")) & "');" & vbCrLf
                    sS += "$('#gcmGTDateB').val('" & CP.Utils.EncodeJavascript(cs.GetText("tourDateB")) & "');" & vbCrLf
                    sS += "$('#gcmGTStartB').val('" & CP.Utils.EncodeJavascript(cs.GetText("tourStartB")) & "');" & vbCrLf
                    sS += "$('input[name=""gcmGTTransportationType""]').filter('[value=""" & CP.Utils.EncodeJavascript(cs.GetText("groupTransportation")) & """]').attr('checked', true);" & vbCrLf
                    sS += "$('#gcmGTContactFirstName').val('" & CP.Utils.EncodeJavascript(cs.GetText("contactFirstName")) & "');" & vbCrLf
                    sS += "$('#gcmGTContactLastName').val('" & CP.Utils.EncodeJavascript(cs.GetText("contactLastName")) & "');" & vbCrLf
                    sS += "$('#gcmGTContactEmail').val('" & CP.Utils.EncodeJavascript(cs.GetText("contactEmail")) & "');" & vbCrLf
                    sS += "$('#gcmGTContactPhone').val('" & CP.Utils.EncodeJavascript(cs.GetText("contactPhone")) & "');" & vbCrLf
                    sS += "$('#gcmGTRequirements').val('" & CP.Utils.EncodeJavascript(cs.GetText("requirements")) & "');" & vbCrLf
                    sS += "$('input[name=""gcmGTUnable""]').filter('[value=""" & CP.Utils.EncodeJavascript(cs.GetText("unable")) & """]').attr('checked', true);" & vbCrLf
                    sS += "});" & vbCrLf
                    '
                    CP.Doc.AddHeadJavascript(sS)
                End If
                cs.Close()
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.groupSignupHandlerClass.getFormInformation")
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
                Dim primaryContactMemberID As Integer = 0
                '
                cs.Open("Group Tour Requests", "ID=" & common.getGroupTourID(CP))
                If cs.OK() Then
                    cs.SetFormInput("name", "gcmGTGroupName")
                    cs.SetFormInput("eventName", "gcmGTEventName")
                    cs.SetFormInput("tourDateA", "gcmGTDateA")
                    cs.SetFormInput("tourStartA", "gcmGTStartA")
                    cs.SetFormInput("tourDateB", "gcmGTDateB")
                    cs.SetFormInput("tourStartB", "gcmGTStartB")
                    cs.SetField("groupTransportation", CP.Doc.Var("gcmGTTransportationType"))
                    cs.SetFormInput("contactFirstName", "gcmGTContactFirstName")
                    cs.SetFormInput("contactLastName", "gcmGTContactLastName")
                    cs.SetFormInput("contactEmail", "gcmGTContactEmail")
                    cs.SetFormInput("contactPhone", "gcmGTContactPhone")
                    cs.SetFormInput("requirements", "gcmGTRequirements")
                    cs.SetField("unable", CP.Doc.Var("gcmGTUnable"))
                    primaryContactMemberID = cs.GetInteger("primaryContactMemberID")
                End If
                cs.Close()
                '
                '   make a people record for primary contact
                '
                If primaryContactMemberID <> 0 Then
                    cs.Open("People", "id=" & primaryContactMemberID)
                Else
                    cs.Open("People", "email=" & CP.Db.EncodeSQLText(CP.Doc.Var("gcmGTContactEmail")))
                    If Not cs.OK() Then
                        cs.Close()
                        cs.Insert("People")
                    End If
                End If
                If cs.OK() Then
                    cs.SetField("name", CP.Doc.Var("gcmGTContactFirstName") & " " & CP.Doc.Var("gcmGTContactLastName"))
                    cs.SetFormInput("firstName", "gcmGTContactFirstName")
                    cs.SetFormInput("lastName", "gcmGTContactLastName")
                    cs.SetFormInput("email", "gcmGTContactEmail")
                    cs.SetFormInput("phone", "gcmGTContactPhone")
                    primaryContactMemberID = cs.GetInteger("id")
                End If
                cs.Close()
                '
                cs.Open("Group Tour Requests", "ID=" & common.getGroupTourID(CP))
                If cs.OK() Then
                    cs.SetField("primaryContactMemberID", primaryContactMemberID)
                End If
                cs.Close()
                '
                Return processed
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.groupSignupHandlerClass.processFormInformation")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
        Private Function getFormContact(ByVal CP As BaseClasses.CPBaseClass) As String
            Try
                Dim s As String = ""
                Dim sS As String = ""
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim csR As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim groupTourID As Integer = common.getGroupTourID(CP)
                Dim groupTourTypeID As Integer = 0
                Dim htmlName As String = ""
                '
                If cs.Open("Layouts", "ccGUID='{0583C764-68D6-4C7F-AC2C-731A0236192F}'", , , "layout") Then
                    s = cs.GetText("layout")
                End If
                '
                s = s.Replace("{{Attendee Fields}}", common.getGroupAttendeeTypes(CP))
                s = s.Replace("{{State Select}}", common.getStateSelect(CP, "gcmGTStateID"))
                '
                cs.Open("Group Tour Requests", "ID=" & groupTourID)
                If cs.OK() Then
                    sS += "$(document).ready(function(){" & vbCrLf
                    sS += "$('#gcmGTFirstName').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedFirstName")) & "');" & vbCrLf
                    sS += "$('#gcmGTLastName').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedLastName")) & "');" & vbCrLf
                    sS += "$('#gcmGTEmail').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedEmail")) & "');" & vbCrLf
                    sS += "$('#gcmGTPhone1').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedPhone1")) & "');" & vbCrLf
                    sS += "$('#gcmGTPhone2').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedPhone2")) & "');" & vbCrLf
                    sS += "$('#gcmGTAddress').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedAddress")) & "');" & vbCrLf
                    sS += "$('#gcmGTAddress2').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedAddress2")) & "');" & vbCrLf
                    sS += "$('#gcmGTCity').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedCity")) & "');" & vbCrLf
                    sS += "$('#gcmGTStateID').val('" & CP.Content.GetRecordID("States", cs.GetText("requestedState")) & "');" & vbCrLf
                    sS += "$('#gcmGTZip').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedZip")) & "');" & vbCrLf
                    sS += "$('#gcmGTSource').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestSource")) & "');" & vbCrLf
                    sS += "$('#gcmGTInterested').prop('checked', " & cs.GetBoolean("requestMoreInfo").ToString.ToLower & ");" & vbCrLf
                    sS += "$('#gcmGTNewsletter').prop('checked', " & cs.GetBoolean("requestNewsletter").ToString.ToLower & ");" & vbCrLf
                    '
                    '   populate the attendee fields
                    '
                    If csR.Open("Group Tour Attendee Rules", "groupTourID=" & groupTourID, , , "groupTourTypeID,attendeeCount") Then
                        Do While csR.OK()
                            groupTourTypeID = csR.GetInteger("groupTourTypeID")
                            htmlName = "gcmGTAttendees_" & groupTourTypeID
                            '
                            sS += "$('#" & htmlName & "').val('" & csR.GetInteger("attendeeCount") & "');" & vbCrLf
                            csR.GoNext()
                        Loop
                    End If
                    csR.Close()
                    '
                    sS += "});" & vbCrLf
                    '
                    CP.Doc.AddHeadJavascript(sS)
                End If
                cs.Close()
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.groupSignupHandlerClass.getFormContact")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
        Private Function processFormContact(ByVal CP As BaseClasses.CPBaseClass) As Boolean
            Try
                Dim processed As Boolean = True
                Dim groupTourID As Integer = common.getGroupTourID(CP)
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim csR As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim requestInfo As Boolean = CP.Utils.EncodeBoolean(CP.Doc.Var("gcmGTInterested"))
                Dim requestNewsletter As Boolean = CP.Utils.EncodeBoolean(CP.Doc.Var("gcmGTNewsletter"))
                Dim recordName As String = ""
                Dim recordID As Integer = 0
                Dim htmlName As String = ""
                Dim attendeeValue As Integer = 0
                Dim eventName As String = ""
                '
                cs.Open("Group Tour Requests", "ID=" & groupTourID)
                If cs.OK() Then
                    cs.SetFormInput("requestedFirstName", "gcmGTFirstName")
                    cs.SetFormInput("requestedLastName", "gcmGTLastName")
                    cs.SetFormInput("requestedEmail", "gcmGTEmail")
                    cs.SetFormInput("requestedPhone1", "gcmGTPhone1")
                    cs.SetFormInput("requestedPhone2", "gcmGTPhone2")
                    cs.SetFormInput("requestedAddress", "gcmGTAddress")
                    cs.SetFormInput("requestedAddress2", "gcmGTAddress2")
                    cs.SetFormInput("requestedCity", "gcmGTCity")
                    cs.SetField("requestedState", CP.Content.GetRecordName("States", CP.Utils.EncodeInteger(CP.Doc.Var("gcmGTStateID"))))
                    cs.SetFormInput("requestedZip", "gcmGTZip")
                    cs.SetFormInput("requestSource", "gcmGTSource")
                    cs.SetField("requestMoreInfo", requestInfo)
                    cs.SetField("requestNewsletter", requestNewsletter)
                    eventName = cs.GetText("eventName")
                End If
                cs.Close()
                '
                '   update people record with requested by info
                '
                If cs.Open("People", "id=" & CP.User.Id) Then
                    cs.SetField("name", CP.Doc.Var("gcmGTFirstName") & " " & CP.Doc.Var("gcmGTLastName"))
                    cs.SetFormInput("firstName", "gcmGTFirstName")
                    cs.SetFormInput("lastName", "gcmGTLastName")
                    cs.SetFormInput("email", "gcmGTEmail")
                    cs.SetFormInput("phone", "gcmGTPhone1")
                End If
                cs.Close()
                '
                '   add to groups if info was requested
                '
                If requestInfo Then
                    CP.Group.AddUser("Information Requests", CP.User.Id)
                End If
                '
                If requestNewsletter Then
                    CP.Group.AddUser("Newsletter Requests", CP.User.Id)
                End If
                '
                '   create rules for attendee types
                '
                If cs.Open("Group Tour Types") Then
                    Do While cs.OK()
                        recordID = cs.GetInteger("id")
                        recordName = cs.GetText("name")
                        htmlName = "gcmGTAttendees_" & recordID
                        attendeeValue = CP.Utils.EncodeInteger(CP.Doc.Var(htmlName))
                        '
                        If attendeeValue <> 0 Then
                            If Not csR.Open("Group Tour Attendee Rules", "(groupTourTypeID=" & recordID & ") and (groupTourID=" & groupTourID & ")") Then
                                csR.Close()
                                csR.Insert("Group Tour Attendee Rules")
                            End If
                            If csR.OK() Then
                                csR.SetField("name", recordName & " for " & eventName)
                                csR.SetField("groupTourTypeID", recordID)
                                csR.SetField("groupTourID", groupTourID)
                                csR.SetField("attendeeCount", attendeeValue)
                            End If
                            csR.Close()
                        Else
                            CP.Content.Delete("Group Tour Attendee Rules", "(groupTourTypeID=" & recordID & ") and (groupTourID=" & groupTourID & ")")
                        End If
                        cs.GoNext()
                    Loop
                End If
                cs.Close()
                '
                Return processed
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.groupSignupHandlerClass.processFormContact")
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
                Dim studentGroupTourTypeID As Integer = 0
                Dim studentTourCaption As String = ""
                Dim attendeeCount As Integer = common.getStudentTourAttendeeCount(CP)
                '
                If cs.Open("Layouts", "ccGUID='{7A7DBF19-3D00-489D-AA03-D1630BA8822D}'", , , "layout") Then
                    s = cs.GetText("layout")
                End If
                '
                s = s.Replace("{{Group Tour Totals}}", common.getGroupTourTotals(CP))
                s = s.Replace("{{Expiration Select}}", common.getExpirationSelect(CP, "gcmGTCardExpMonth", "gcmGTCardExpYear", "gcmGTSelectShort"))
                s = s.Replace("{{Check Instruction}}", CP.Content.GetCopy("Student Group Tour Request Check Payment Instructions", "Please send your checks payable to..."))
                '
                If errMsg <> "" Then
                    s = CP.Html.p(errMsg, , "ccError") & s
                End If
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.groupSignupHandlerClass.getFormPayment")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
        Private Function processFormPayment(ByVal CP As BaseClasses.CPBaseClass) As Boolean
            Try
                Dim tourAmount As String = common.getGroupTourTotalAmount(CP)
                Dim groupTourID As Integer = common.getGroupTourID(CP)
                Dim processed As Boolean = False
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim paymenType As String = CP.Doc.Var("gcmGTPayMethod")
                Dim paymentCredit As Boolean = CP.Utils.EncodeBoolean(paymenType = "Credit Card")
                Dim cardNumber As String = CP.Doc.Var("gcmGTCardNumber")
                Dim cardExpiration As String = CP.Doc.Var("gcmGTCardExpMonth") & "/" & CP.Doc.Var("gcmGTCardExpYear")
                Dim cardCVV As String = CP.Doc.Var("gcmGTCardCVV")
                Dim cardAddress As String = CP.Doc.Var("gcmGTCardAddress")
                Dim cardZip As String = CP.Doc.Var("gcmGTCardZip")
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
                Dim groupName As String = ""
                Dim eventName As String = ""
                '
                '   verify payment
                '
                If paymentCredit Then
                    CP.Doc.Var("CreditCardNumber") = cardNumber
                    CP.Doc.Var("CreditCardExpiration") = cardExpiration
                    CP.Doc.Var("SecurityCode") = cardCVV
                    CP.Doc.Var("AVSAddress") = cardAddress
                    CP.Doc.Var("AVSZip") = cardZip
                    CP.Doc.Var("PaymentAmount") = tourAmount
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
                    '   complete the tour record
                    '
                    cs.Open("Group Tour Requests", "ID=" & groupTourID)
                    If cs.OK() Then
                        cs.SetField("completed", 1)
                        cs.SetField("dateCompleted", Date.Now)
                        cs.SetField("paymentType", paymenType)
                        cs.SetField("totalAmount", tourAmount)
                        If paymentCredit Then
                            cs.SetField("cardName", paymenType)
                            cs.SetField("cardNumber", cardNumber)
                            cs.SetField("cardExpiration", cardExpiration)
                            cs.SetField("cardBillingAddress", cardAddress)
                            cs.SetField("cardBillingZip", cardZip)
                            cs.SetField("cardCode", cardCVV)
                            cs.SetField("processorResponse", ppResponse)
                            cs.SetField("processorReference", ppReference)
                        End If
                        targetMemberID = cs.GetInteger("memberID")
                        groupName = cs.GetText("name")
                        eventName = cs.GetText("eventName")
                    End If
                    cs.Close()
                    '
                    '   move record to members if paid by credit card
                    '
                    If paymentCredit Then
                        '
                        '   add to user to groups
                        '
                        CP.Group.AddUser("Tour Requests", targetMemberID)
                    End If
                    '
                    '   send notifications
                    '
                    copy += "Group Name: " & groupName & br
                    copy += "Name of Event: " & eventName & br
                    copy += "Attendees: " & common.getGroupTourAttendeeString(CP) & br
                    copy += "Total Amount: " & tourAmount & br
                    If paymentCredit Then
                        copy += "Amount Paid: " & tourAmount & br
                        copy += "Payment Date: " & Date.Today & br
                        copy += "Account Number: " & Right(cardNumber, 4) & br
                    Else
                        copy += "Amount Due: " & tourAmount & br & br
                        copy += CP.Content.GetCopy("Group Tour Request Check Payment Instructions", "Please send your checks payable to...")
                    End If
                    '
                    '   add link to email
                    '
                    recordLink = CP.Site.DomainPrimary & "/admin/index.php?af=4&cid=" & CP.Content.GetRecordID("Content", "Group Tour Requests")
                    recordLink += "&id=" & groupTourID
                    '
                    copyPlus += copy
                    copyPlus += CP.Html.p("<a target=""_blank"" href=""http://" & recordLink & """>Click here for details</a>")
                    '
                    CP.Email.sendSystem("Group Tour Auto Responder", copy, CP.User.Id)
                    CP.Email.sendSystem("Group Tour Notification", copyPlus)
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