
Imports common = Contensive.Addons.gcmFoundation.commonClass

Namespace Contensive.Addons.gcmFoundation
    '
    Public Class studentGroupSignupClass
        Inherits BaseClasses.AddonBaseClass
        '
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            Try
                Dim s As String = CP.Utils.ExecuteAddon("{EA3F2397-F89D-4F0C-91BA-4DEF7E9F5AA9}")
                '
                s = CP.Html.div(s, , , "gcmSGContainer")
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.studentGroupSignupClass.execute")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
    End Class
    '
    Public Class studentGroupSignupHandlerClass
        Inherits BaseClasses.AddonBaseClass
        '
        Const formInformation As Integer = 100
        Const formContact As Integer = 200
        Const formPayment As Integer = 300
        Const formThankYou As Integer = 400
        '
        Const rnFormID As String = "gcmSGormID"
        Const rnSourceFormID As String = "gcmSGSourceFormID"
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
                        s += CP.Content.GetCopy("Student Group Tour Form Thank You", "Thank you for your interest in our group tours. Your request has been received and will be processed shortly.")
                    Case formPayment
                        s += CP.Content.GetCopy("Student Group Tour Form Instructions - Payment", "Editable Instructions")
                        s += getFormPayment(CP)
                    Case formContact
                        s += CP.Content.GetCopy("Student Group Tour Form Instructions - Contact Information", "Editable Instructions")
                        s += getFormContact(CP)
                    Case Else
                        formID = formInformation
                        s += CP.Content.GetCopy("Student Group Tour Form Instructions - Information", "Editable Instructions")
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
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.studentGroupSignupHandlerClass.execute")
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
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.studentGroupSignupHandlerClass.processFormGetID")
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
                If cs.Open("Layouts", "ccGUID='{64F7A942-DB25-4871-B64C-AE9B227D625B}'", , , "layout") Then
                    s = cs.GetText("layout")
                End If
                '
                s = s.Replace("{{Tour Types}}", common.getStudentGroupTourTypes(CP))
                '
                cs.Open("Student Group Tour Requests", "ID=" & common.getStudentGroupTourID(CP))
                If cs.OK() Then
                    attendeeCount = cs.GetInteger("tourAttendeeCountA")
                    '
                    sS += "$(document).ready(function(){" & vbCrLf
                    sS += "$('#gcmSGGroupName').val('" & CP.Utils.EncodeJavascript(cs.GetText("name")) & "');" & vbCrLf
                    sS += "$('input[name=""gcmSGTourTypeID""]').filter('[value=""" & cs.GetInteger("studentGroupTourTypeID") & """]').attr('checked', true);" & vbCrLf
                    sS += "$('#gcmSGDateA').val('" & CP.Utils.EncodeJavascript(cs.GetText("tourDateA")) & "');" & vbCrLf
                    sS += "$('#gcmSGStartA').val('" & CP.Utils.EncodeJavascript(cs.GetText("tourStartA")) & "');" & vbCrLf
                    If attendeeCount <> 0 Then
                        sS += "$('#gcmSGAttendeeCountA').val('" & cs.GetInteger("tourAttendeeCountA") & "');" & vbCrLf
                    End If
                    sS += "$('#gcmSGDateB').val('" & CP.Utils.EncodeJavascript(cs.GetText("tourDateB")) & "');" & vbCrLf
                    sS += "$('#gcmSGStartB').val('" & CP.Utils.EncodeJavascript(cs.GetText("tourStartB")) & "');" & vbCrLf
                    sS += "$('input[name=""gcmSGTransportationType""]').filter('[value=""" & CP.Utils.EncodeJavascript(cs.GetText("groupTransportation")) & """]').attr('checked', true);" & vbCrLf
                    sS += "$('#gcmSGContactFirstName').val('" & CP.Utils.EncodeJavascript(cs.GetText("contactFirstName")) & "');" & vbCrLf
                    sS += "$('#gcmSGContactLastName').val('" & CP.Utils.EncodeJavascript(cs.GetText("contactLastName")) & "');" & vbCrLf
                    sS += "$('#gcmSGContactEmail').val('" & CP.Utils.EncodeJavascript(cs.GetText("contactEmail")) & "');" & vbCrLf
                    sS += "$('#gcmSGContactPhone').val('" & CP.Utils.EncodeJavascript(cs.GetText("contactPhone")) & "');" & vbCrLf
                    sS += "$('#gcmSGRequirements').val('" & CP.Utils.EncodeJavascript(cs.GetText("requirements")) & "');" & vbCrLf
                    sS += "$('input[name=""gcmSGUnable""]').filter('[value=""" & CP.Utils.EncodeJavascript(cs.GetText("unable")) & """]').attr('checked', true);" & vbCrLf
                    sS += "});" & vbCrLf
                    '
                    CP.Doc.AddHeadJavascript(sS)
                End If
                cs.Close()
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.studentGroupSignupHandlerClass.getFormInformation")
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
                cs.Open("Student Group Tour Requests", "ID=" & common.getStudentGroupTourID(CP))
                If cs.OK() Then
                    cs.SetFormInput("name", "gcmSGGroupName")
                    cs.SetField("studentGroupTourTypeID", CP.Utils.EncodeInteger(CP.Doc.Var("gcmSGTourTypeID")))
                    cs.SetFormInput("tourDateA", "gcmSGDateA")
                    cs.SetFormInput("tourStartA", "gcmSGStartA")
                    cs.SetFormInput("tourAttendeeCountA", "gcmSGAttendeeCountA")
                    cs.SetFormInput("tourDateB", "gcmSGDateB")
                    cs.SetFormInput("tourStartB", "gcmSGStartB")
                    cs.SetField("groupTransportation", CP.Doc.Var("gcmSGTransportationType"))
                    cs.SetFormInput("contactFirstName", "gcmSGContactFirstName")
                    cs.SetFormInput("contactLastName", "gcmSGContactLastName")
                    cs.SetFormInput("contactEmail", "gcmSGContactEmail")
                    cs.SetFormInput("contactPhone", "gcmSGContactPhone")
                    cs.SetFormInput("requirements", "gcmSGRequirements")
                    cs.SetField("unable", CP.Doc.Var("gcmSGUnable"))
                    primaryContactMemberID = cs.GetInteger("primaryContactMemberID")
                End If
                cs.Close()
                '
                '   make a people record for primary contact
                '
                If primaryContactMemberID <> 0 Then
                    cs.Open("People", "id=" & primaryContactMemberID)
                Else
                    cs.Open("People", "email=" & CP.Db.EncodeSQLText(CP.Doc.Var("gcmSGContactEmail")))
                    If Not cs.OK() Then
                        cs.Close()
                        cs.Insert("People")
                    End If
                End If
                If cs.OK() Then
                    cs.SetField("name", CP.Doc.Var("gcmSGContactFirstName") & " " & CP.Doc.Var("gcmSGContactLastName"))
                    cs.SetFormInput("firstName", "gcmSGContactFirstName")
                    cs.SetFormInput("lastName", "gcmSGContactLastName")
                    cs.SetFormInput("email", "gcmSGContactEmail")
                    cs.SetFormInput("phone", "gcmSGContactPhone")
                    primaryContactMemberID = cs.GetInteger("id")
                End If
                cs.Close()
                '
                Return processed
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.studentGroupSignupHandlerClass.processFormInformation")
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
                '
                If cs.Open("Layouts", "ccGUID='{7B1A4916-0481-413F-ACB7-18FBD2DB3086}'", , , "layout") Then
                    s = cs.GetText("layout")
                End If
                '
                s = s.Replace("{{State Select}}", common.getStateSelect(CP, "gcmSGStateID"))
                '
                cs.Open("Student Group Tour Requests", "ID=" & common.getStudentGroupTourID(CP))
                If cs.OK() Then
                    sS += "$(document).ready(function(){" & vbCrLf
                    sS += "$('#gcmSGFirstName').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedFirstName")) & "');" & vbCrLf
                    sS += "$('#gcmSGLastName').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedLastName")) & "');" & vbCrLf
                    sS += "$('#gcmSGEmail').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedEmail")) & "');" & vbCrLf
                    sS += "$('#gcmSGPhone1').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedPhone1")) & "');" & vbCrLf
                    sS += "$('#gcmSGPhone2').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedPhone2")) & "');" & vbCrLf
                    sS += "$('#gcmSGAddress').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedAddress")) & "');" & vbCrLf
                    sS += "$('#gcmSGAddress2').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedAddress2")) & "');" & vbCrLf
                    sS += "$('#gcmSGCity').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedCity")) & "');" & vbCrLf
                    sS += "$('#gcmSGStateID').val('" & CP.Content.GetRecordID("States", cs.GetText("requestedState")) & "');" & vbCrLf
                    sS += "$('#gcmSGZip').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestedZip")) & "');" & vbCrLf
                    sS += "$('#gcmSGSource').val('" & CP.Utils.EncodeJavascript(cs.GetText("requestSource")) & "');" & vbCrLf
                    sS += "$('#gcmSGInterested').prop('checked', " & cs.GetBoolean("requestMoreInfo").ToString.ToLower & ");" & vbCrLf
                    sS += "$('#gcmSGNewsletter').prop('checked', " & cs.GetBoolean("requestNewsletter").ToString.ToLower & ");" & vbCrLf
                    sS += "});" & vbCrLf
                    '
                    CP.Doc.AddHeadJavascript(sS)
                End If
                cs.Close()
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.studentGroupSignupHandlerClass.getFormContact")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
        Private Function processFormContact(ByVal CP As BaseClasses.CPBaseClass) As Boolean
            Try
                Dim processed As Boolean = True
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim requestInfo As Boolean = CP.Utils.EncodeBoolean(CP.Doc.Var("gcmSGInterested"))
                Dim requestNewsletter As Boolean = CP.Utils.EncodeBoolean(CP.Doc.Var("gcmSGNewsletter"))
                '
                cs.Open("Student Group Tour Requests", "ID=" & common.getStudentGroupTourID(CP))
                If cs.OK() Then
                    cs.SetFormInput("requestedFirstName", "gcmSGFirstName")
                    cs.SetFormInput("requestedLastName", "gcmSGLastName")
                    cs.SetFormInput("requestedEmail", "gcmSGEmail")
                    cs.SetFormInput("requestedPhone1", "gcmSGPhone1")
                    cs.SetFormInput("requestedPhone2", "gcmSGPhone2")
                    cs.SetFormInput("requestedAddress", "gcmSGAddress")
                    cs.SetFormInput("requestedAddress2", "gcmSGAddress2")
                    cs.SetFormInput("requestedCity", "gcmSGCity")
                    cs.SetField("requestedState", CP.Content.GetRecordName("States", CP.Utils.EncodeInteger(CP.Doc.Var("gcmSGStateID"))))
                    cs.SetFormInput("requestedZip", "gcmSGZip")
                    cs.SetFormInput("requestSource", "gcmSGSource")
                    cs.SetField("requestMoreInfo", requestInfo)
                    cs.SetField("requestNewsletter", requestNewsletter)
                End If
                cs.Close()
                '
                '   update people record with requested by info
                '
                If cs.Open("People", "id=" & CP.User.Id) Then
                    cs.SetField("name", CP.Doc.Var("gcmSGFirstName") & " " & CP.Doc.Var("gcmSGLastName"))
                    cs.SetFormInput("firstName", "gcmSGFirstName")
                    cs.SetFormInput("lastName", "gcmSGLastName")
                    cs.SetFormInput("email", "gcmSGEmail")
                    cs.SetFormInput("phone", "gcmSGPhone1")
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
                Return processed
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.studentGroupSignupHandlerClass.processFormContact")
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
                If cs.Open("Layouts", "ccGUID='{CAC975FB-3382-4026-AB35-271148A30CE0}'", , , "layout") Then
                    s = cs.GetText("layout")
                End If
                '
                If cs.Open("Student Group Tour Requests", "id=" & common.getStudentGroupTourID(CP), , , "studentGroupTourTypeID") Then
                    studentGroupTourTypeID = cs.GetInteger("studentGroupTourTypeID")
                End If
                cs.Close()
                '
                If cs.Open("Student Group Tour Types", "id=" & studentGroupTourTypeID, , , "name,tourLength") Then
                    studentTourCaption = cs.GetText("name") & " (" & cs.GetText("tourLength") & ")"
                End If
                cs.Close()
                '
                hiddenString += CP.Html.Hidden("gcmSGLevelID", studentGroupTourTypeID.ToString, , "gcmSGLevelID")
                '
                s = s.Replace("{{Tour Type}}", CP.Content.GetRecordName("Student Group Tour Types", studentGroupTourTypeID))
                s = s.Replace("{{Number of Attendees}}", attendeeCount)
                s = s.Replace("{{Amount Due}}", FormatCurrency(attendeeCount * common.getStudentTourIndividualAmount(CP, studentGroupTourTypeID), 2))
                s = s.Replace("{{Expiration Select}}", common.getExpirationSelect(CP, "gcmSGCardExpMonth", "gcmSGCardExpYear", "gcmSGSelectShort"))
                s = s.Replace("{{Check Instruction}}", CP.Content.GetCopy("Student Group Tour Request Check Payment Instructions", "Please send your checks payable to..."))
                '
                If errMsg <> "" Then
                    s = CP.Html.p(errMsg, , "ccError") & s
                End If
                '
                Return s
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.studentGroupSignupHandlerClass.getFormPayment")
                Catch errorObj As Exception
                    '
                End Try
            End Try
        End Function
        '
        Private Function processFormPayment(ByVal CP As BaseClasses.CPBaseClass) As Boolean
            Try
                Dim studentGroupTourTypeID As Integer = CP.Utils.EncodeInteger(CP.Doc.Var("gcmSGLevelID"))
                Dim studentGroupTourTypeName As String = CP.Content.GetRecordName("Student Group Tour Types", studentGroupTourTypeID)
                Dim attendeeCount As Integer = common.getStudentTourAttendeeCount(CP)
                Dim tourAmount As String = attendeeCount * common.getStudentTourIndividualAmount(CP, studentGroupTourTypeID)
                Dim studentGroupTourID As Integer = common.getStudentGroupTourID(CP)
                Dim processed As Boolean = False
                Dim cs As BaseClasses.CPCSBaseClass = CP.CSNew()
                Dim paymenType As String = CP.Doc.Var("gcmSGPayMethod")
                Dim paymentCredit As Boolean = CP.Utils.EncodeBoolean(paymenType = "Credit Card")
                Dim cardNumber As String = CP.Doc.Var("gcmSGCardNumber")
                Dim cardExpiration As String = CP.Doc.Var("gcmSGCardExpMonth") & "/" & CP.Doc.Var("gcmSGCardExpYear")
                Dim cardCVV As String = CP.Doc.Var("gcmSGCardCVV")
                Dim cardAddress As String = CP.Doc.Var("gcmSGCardAddress")
                Dim cardZip As String = CP.Doc.Var("gcmSGCardZip")
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
                    '   complete the membership record
                    '
                    cs.Open("Student Group Tour Requests", "ID=" & studentGroupTourID)
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
                    End If
                    cs.Close()
                    '
                    '   move record to members if paid by credit card
                    '
                    If paymentCredit Then
                        '
                        '   add to user to groups
                        '
                        CP.Group.AddUser("Student Tour Requests", targetMemberID)
                    End If
                    '
                    '   send notifications
                    '
                    copy += "Tour Type: " & studentGroupTourTypeName & br
                    copy += "Attendees: " & attendeeCount & br
                    copy += "Total Amount: " & tourAmount & br
                    If paymentCredit Then
                        copy += "Amount Paid: " & tourAmount & br
                        copy += "Payment Date: " & Date.Today & br
                        copy += "Account Number: " & Right(cardNumber, 4) & br
                    Else
                        copy += "Amount Due: " & tourAmount & br & br
                        copy += CP.Content.GetCopy("Student Group Tour Request Check Payment Instructions", "Please send your checks payable to...")
                    End If
                    '
                    '   add link to email
                    '
                    recordLink = CP.Site.DomainPrimary & "/admin/index.php?af=4&cid=" & CP.Content.GetRecordID("Content", "Student Group Tour Requests")
                    recordLink += "&id=" & studentGroupTourID
                    '
                    copyPlus += copy
                    copyPlus += CP.Html.p("<a target=""_blank"" href=""http://" & recordLink & """>Click here for details</a>")
                    '
                    CP.Email.sendSystem("Student Group Tour Auto Responder", copy, CP.User.Id)
                    CP.Email.sendSystem("Student Group Tour Notification", copyPlus)
                End If
                '
                Return processed
            Catch ex As Exception
                Try
                    CP.Site.ErrorReport(ex, "error in Contensive.Addons.gcmFoundation.studentGroupSignupHandlerClass.processFormPayment")
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