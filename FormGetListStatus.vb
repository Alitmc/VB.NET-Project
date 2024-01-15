Imports System.ComponentModel
Imports PCommonTools.ExceptionHandler
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports System.Threading
Imports DevExpress.XtraBars.Docking2010

Public Class FormGetListStatus

    Public Property CanClose As Boolean = False
    Public Property TrnList As List(Of Trn_Transaction) = Nothing
    Public Property TriList As List(Of Trn_TransactionItem) = Nothing
    Public Property ftrnList As List(Of Trn_Transaction) = Nothing
    Public Property Trn As Trn_Transaction = Nothing
    Public Property PartnerList As BindingList(Of Com_Partner)

    Private WithEvents _BLElectronicInvoice As TaxLIbrary.ElectronicInvoice
    Private WithEvents _BLTransaction As BLTransaction
    Private th As Thread
    Private th1 As Thread
    Public Property TrnListCount As Integer = 0
    Public Property TrnListSentCount As Integer = 0

    Private Delegate Sub SetStateFunc(state As Boolean)

    Delegate Sub SetlblCount(ByVal text As String)

    Delegate Sub SetlblRepair(ByVal text As String)
    Delegate Sub SetlblIgnore(ByVal text As String)

    Delegate Sub SetTxt(ByVal text As String)

    Private Sub Setlbl1(ByVal text As String)
        If lblCount.InvokeRequired Then
            Dim d As SetlblIgnore = New SetlblIgnore(AddressOf Setlbl1)
            Me.Invoke(d, New Object() {text})
        End If
    End Sub

    Private Sub Setlbl(ByVal text As String)
        If lblCount.InvokeRequired Then
            Dim d As SetlblRepair = New SetlblRepair(AddressOf Setlbl)
            Me.Invoke(d, New Object() {text})
        End If
    End Sub
    Private Sub SetText(ByVal text As String)
        If lblCount.InvokeRequired Then
            Dim d As SetlblCount = New SetlblCount(AddressOf SetText)
            Me.Invoke(d, New Object() {text})
        Else
            If text = "" Then
                lblCount.Visible = False

            Else
                lblCount.Visible = True

            End If
            lblCount.Text = text
        End If
    End Sub
    Private Sub SetText1(ByVal text As String)
        If lblCount.InvokeRequired Then
            Dim d As SetTxt = New SetTxt(AddressOf SetText1)
            Me.Invoke(d, New Object() {text})
        End If
    End Sub

    Private Sub SetText2(ByVal text As String)
        If lblCount.InvokeRequired Then
            Dim d As SetTxt = New SetTxt(AddressOf SetText2)
            Me.Invoke(d, New Object() {text})
        End If
    End Sub

    Private Sub timer1_tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            If Trn Is Nothing Then
                SetText("")
                Setlbl("")
                Setlbl1("")
                SetText1("اطلاعاتی یافت نشد")
                SetText2("")
                If Ignore = True Then
                    Setlbl1("a")
                End If
                Return
            End If
            Setlbl("")
            Setlbl1("")
            If Trn._RepairTrn = True Then
                Setlbl("a")
            End If
            If Ignore = True Then
                Setlbl1("a")
            End If
            If Trn.FinancialStatmentState = 5 Then
                SetText(Trn.ElectronicInvoiceErrorMessage)
                SetText1("رد شده توسط سامانه")
                SetText2(Trn.FinancialStatmentTaxID)
            End If
            If Trn.FinancialStatmentState = 4 Then
                SetText("")
                SetText1("تأیید شده توسط سامانه")
                SetText2(Trn.FinancialStatmentTaxID)
            End If
            If Trn.FinancialStatmentState = 2 Then
                SetText("")
                SetText1("در انتظار پاسخ از سامانه")
                SetText2(Trn.FinancialStatmentTaxID)
            End If

        Catch ex As Exception
            CustomException.ShowDialogue(ex)
        End Try
    End Sub
    Public Property Ignore As Boolean = False

    Private Sub DoPaymentCalculationAction()
        Try
            Timer1.Start()

            _BLTransaction = New BLTransaction(BLTransactionForm.GetFormByID(cmbTransactionForm.EditValue), LookUpEdit1.EditValue)
            Dim trnLst As New BindingList(Of Trn_Transaction)
            If PTextBox1.Text <> "" Then
                If CType(PTextBox1.Text, Integer) > 0 Then
                    trnLst = _BLTransaction.GetTransactionListForListFinancialStatment(CType(PTextBox1.Text, Integer))
                    If trnLst.Count > 100 Then
                        Throw New CustomException("سقف تعداد در هر بار استعلام 100 عدد می باشد")
                    End If
                    If trnLst.Count = 0 Then
                        Throw New CustomException("اطلاعاتی جهت استعلام وجود ندارد.")
                    End If
                Else
                    Return
                End If
            Else
                Return
            End If
            If trnLst.Count <> 0 Then
                For Each i In trnLst
                    Trn = Ganjineh.Net.TaxLIbrary.ElectronicInvoice.GetStatus(i, BLAttachment.DownloadAttachedFile(BLAttachment.GetAttachmentList(0, 0).FirstOrDefault.AttachmentID))
                Next
            Else
                Trn = Nothing
            End If
            InformationMessageBox("استعلام اطلاعات با موفقیت انجام شد.")
        Catch ex As Exception
            If TypeOf ex Is CalculationExceptionWithoutMessage Then
                Return
            End If
            If ex.Message = "Thread was being aborted." Or ex.Message = "اشكال در ذخيره اطلاعات." Then
                Return
            End If
            CustomException.ShowDialogue(ex)

        End Try
    End Sub

    Private Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
        Try

            If btnSend.Text = "محاسبه تعداد فاکتور ها" Then

                _BLTransaction = New BLTransaction(BLTransactionForm.GetFormByID(cmbTransactionForm.EditValue), LookUpEdit1.EditValue)
                Dim trnLst As New BindingList(Of Trn_Transaction)
                If PTextBox1.Text <> "" Then
                    If CType(PTextBox1.Text, Integer) > 0 Then
                        trnLst = _BLTransaction.GetTransactionListForListFinancialStatment(CType(PTextBox1.Text, Integer))
                        If trnLst.Count > 100 Then
                            Label3.Text = trnLst.Count
                            Label3.ForeColor = System.Drawing.Color.DarkRed
                            Throw New CustomException("سقف تعداد در هر بار استعلام 100 عدد می باشد")
                        End If
                        btnSend.Text = "دریافت اطلاعات از سامانه"
                        Label3.ForeColor = System.Drawing.Color.ForestGreen
                        Label3.Text = trnLst.Count
                        Return
                    Else
                        PTextBox1.Text = Nothing
                        Return
                    End If
                Else
                    PTextBox1.Text = Nothing
                    Return
                End If
            End If
            DoPaymentCalculationAction()
            Label3.Text = 0
            PTextBox1.Text = Nothing
        Catch ex As Exception
            CustomException.ShowDialogue(ex)
        End Try

    End Sub


    Function IsValidNationalCode(Partner As Com_Partner, ByVal nationalCode As String) As Boolean
        If String.IsNullOrEmpty(nationalCode) Then Throw New Exception("طرف تجاری " & Partner.PartnerCode & " فاقد شناسه یا کد ملی می باشد.")
        If nationalCode.Length <> 10 Then Throw New Exception("طول کد ملی طرف تجاری " & Partner.PartnerCode & " باید ده کاراکتر باشد")
        Dim regex = New Regex("\d{10}")
        If Not regex.IsMatch(nationalCode) Then Throw New Exception("کد ملی طرف تجاری " & Partner.PartnerCode & " باید از ده رقم عددی تشکیل شده باشد؛ لطفا کد ملی را اصلاح کنید")
        Dim allDigitEqual = {"0000000000", "1111111111", "2222222222", "3333333333", "4444444444", "5555555555", "6666666666", "7777777777", "8888888888", "9999999999"}
        If allDigitEqual.Contains(nationalCode) Then Return False
        Dim chArray = nationalCode.ToCharArray()
        Dim num0 = Convert.ToInt32(chArray(0).ToString) * 10
        Dim num2 = Convert.ToInt32(chArray(1).ToString) * 9
        Dim num3 = Convert.ToInt32(chArray(2).ToString) * 8
        Dim num4 = Convert.ToInt32(chArray(3).ToString) * 7
        Dim num5 = Convert.ToInt32(chArray(4).ToString) * 6
        Dim num6 = Convert.ToInt32(chArray(5).ToString) * 5
        Dim num7 = Convert.ToInt32(chArray(6).ToString) * 4
        Dim num8 = Convert.ToInt32(chArray(7).ToString) * 3
        Dim num9 = Convert.ToInt32(chArray(8).ToString) * 2
        Dim a = Convert.ToInt32(chArray(9).ToString)
        Dim b = (((((((num0 + num2) + num3) + num4) + num5) + num6) + num7) + num8) + num9
        Dim c = b Mod 11
        Return (((c < 2) AndAlso (a = c)) OrElse ((c >= 2) AndAlso ((11 - c) = a)))
    End Function

    Private Sub _BLVoucher_VoucherDeleting(sender As Object, e As VoucherDeletingEventArgs) Handles _BLElectronicInvoice.TrnSending
        Try

            Application.DoEvents()
        Catch ex As Exception
            PCommonTools.ExceptionHandler.CustomException.ShowDialogue(ex)
        End Try
    End Sub

    Private Sub FormFinancialStatment_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Label3.ForeColor = System.Drawing.Color.Black

            Invoke(New CloseFunc(AddressOf CloseMe))
            cmbTransactionForm.Focus()

            PartnerList = BLPartner.GetPartnerList
            Dim source As List(Of Trn_TransactionForm) = GlobalParam.TransactionFormsList.Where(Function(a) a.FormCode.ToString.Remove(3) = 103 Or a.FormCode.ToString.Remove(3) = 102).ToList
            cmbTransactionForm.Properties.DataSource = source
            cmbTransactionForm.EditValue = source.FirstOrDefault(Function(a) a.FormCode.ToString.Remove(3) = 102).FormID


            LookUpEdit1.Properties.DataSource = BLFinancialYear.GetAllFinancialYearList
            LookUpEdit1.EditValue = Context.CurrentYear.FinancialYearID

            Dim Dic As New Dictionary(Of Integer, String)
            Dic.Add(30, "خرید و فروش")


        Catch ex As Exception
            CustomException.ShowDialogue(ex)
        End Try
    End Sub

    Private Delegate Sub CloseFunc()
    Private Sub CloseMe()
        If InvokeRequired Then
            Invoke(New CloseFunc(AddressOf CloseMe))
        Else
            If th IsNot Nothing Then
                th.Abort()
            End If
            If th1 IsNot Nothing Then
                th1.Abort()
            End If
            If CanClose = True Then
                Dispose(True)
            End If

        End If
    End Sub

    Private Sub FormFinancialStatment_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Try
            If CanClose = False Then
                e.Cancel = True
            End If

        Catch ex As Exception
            CustomException.ShowDialogue(ex)
        End Try
    End Sub



    Private Sub FormPostProcess_KeyDown(sender As Object, e As KeyEventArgs) Handles btnSend.KeyDown, cmbTransactionForm.KeyDown
        Try

            If e.KeyCode = Keys.Escape Then
                CanClose = True
                If th IsNot Nothing Then
                    th.Abort()
                End If
                If th1 IsNot Nothing Then
                    th1.Abort()
                End If
                CloseMe()
            End If
        Catch ex As Exception
            CustomException.ShowDialogue(ex)
        End Try
    End Sub

    Private Sub GroupControl2_CustomButtonClick(sender As Object, e As BaseButtonEventArgs) Handles GroupControl2.CustomButtonClick

        CanClose = True
        If th IsNot Nothing Then
            th.Abort()
        End If
        If th1 IsNot Nothing Then
            th1.Abort()
        End If
        CloseMe()
        Me.Close()
    End Sub

    Private Sub PTextBox1_TextChanged(sender As Object, e As EventArgs) Handles PTextBox1.TextChanged
        btnSend.Text = "محاسبه تعداد فاکتور ها"
        Label3.Text = 0
        Label3.ForeColor = System.Drawing.Color.Black
    End Sub
End Class