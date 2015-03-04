Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses
Imports Contensive.Addons.aoAccountBilling
Imports System.Exception
Imports System.IO

Namespace Contensive.Addons

    Public Class gcmPaymentClass
        Inherits AddonBaseClass



        '
        ' - update references to your installed version of cpBase
        ' - Verify project root name space is empty
        ' - Change the namespace to the collection name
        ' - Change this class name to the addon name
        ' - Create a Contensive Addon record with the namespace apCollectionName.ad
        ' - add reference to CPBase.DLL, typically installed in c:\program files\kma\contensive\
        '
        '=====================================================================================
        ' addon api
        '=====================================================================================
        '

        Dim orderId As Object

       

        Private Property returnHtml As String
        Dim paymentComment2 As Object
        Dim orderIdList As Object
        Dim paymentComment1 As Object


        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String
            Dim cs As CPCSBaseClass = CP.CSNew()
            Dim accountId As Integer = 0
            '
            '
            If cs.Open("people", "id=" & CP.User.Id) Then
                accountId = cs.GetInteger("accountId")
            End If
            Call cs.Close()

            Dim accountStatus As New apiClass.accountStatusStruct
            accountStatus = api.getAccountStatus(CP, returnUserError, accountId)
            ' accountStatus.closed is a valid account that has been closed
            ' accountStatus.currentBalance is the balance due
            ' accountStatus.exists is true if the account exists
            ' accountStatus.pastDueAmount is the past due balance
            ' returnUserError if not blank, this is a user appropriate message
            If (Not accountStatus.exists) Or (accountStatus.closed) Then
                If cs.Open("people", "id=" & CP.User.Id) Then
                    Call cs.SetField("accountId", 0)
                End If
                Call cs.Close()
                accountId = api.createAccount(CP.User.Id, "New Account Name")
            End If
            '
            '
            Dim orderId As Integer
            orderId = api.createOrder()
            Call setOrderAccount(orderId, accountId)
            Call addOrderItem(orderId, itemId)

            Try
                returnHtml = "Visual Studio 2005 Contensive Addon - OK response"
            Catch ex As Exception
                errorReport(CP, ex, "execute")
                returnHtml = "Visual Studio 2005 Contensive Addon - Error response"
            End Try
            Return returnHtml
        End Function
        '
        '


        '
        '=====================================================================================
        ' common report for this class
        '=====================================================================================
        '
        Private Sub errorReport(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)

            Try
                cp.Site.ErrorReport(ex, "Unexpected error in sampleClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                returnHtml = cp.Html.Form(api.getOnlinePaymentFields())
                '

                '
                '
                If Not api.payOrdersByOnlinePaymentFields(cp, returnUserError, orderIdList, paymentComment1, paymentComment2) Then
                    ' there was a problem with the order, a user compatible message is in returnUserError
                Else
                    ' the order processed correctly.
                End If
                '

                If Not api.chargeOrder(cp, returnUserError, orderId, paymentComment1, paymentComment2) Then
                    ' there was a problem with the order, a user compatible message is in returnUserError
                Else
                    ' the order processed correctly.
                End If
            End Try


        End Sub



        Private Function returnUserError() As Object
            Throw New NotImplementedException
        End Function

        Private Function api() As Object
            Throw New NotImplementedException
        End Function

        Private Sub setOrderAccount(orderId As Integer, accountId As Integer)
            Throw New NotImplementedException
        End Sub

        Private Sub addOrderItem(orderId As Integer, p2 As Object)
            Throw New NotImplementedException
        End Sub

        Private Function itemId() As Object
            Throw New NotImplementedException
        End Function

    End Class
End Namespace
