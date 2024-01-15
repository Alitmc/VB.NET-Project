Imports System.ComponentModel
Imports PCommonTools.ExceptionHandler
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports System.Threading
Imports DevExpress.XtraBars.Docking2010

Public Class FormGetStatus

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
    Delegate Sub SetlblReject(ByVal text As String)
    Delegate Sub SetlblRepair(ByVal text As String)
    Delegate Sub SetlblIgnore(ByVal text As String)

    Delegate Sub SetTxt(ByVal text As String)

    Private Sub Setlbl1(ByVal text As String)
        If lblCount.InvokeRequired Then
            Dim d As SetlblIgnore = New SetlblIgnore(AddressOf Setlbl1)
            Me.Invoke(d, New Object() {text})
        Else
            If text = "" Then
                PLabel5.Visible = False

            Else
                PLabel5.Visible = True

            End If
        End If
    End Sub
    Private Sub SetPLabel6(ByVal text As String)
        If lblCount.InvokeRequired Then
            Dim d As SetlblReject = New SetlblReject(AddressOf SetPLabel6)
            Me.Invoke(d, New Object() {text})
        Else
            If text = "" Then
                PLabel6.Visible = False

            Else
                PLabel6.Visible = True

            End If
        End If
    End Sub
    Private Sub Setlbl(ByVal text As String)
        If lblCount.InvokeRequired Then
            Dim d As SetlblRepair = New SetlblRepair(AddressOf Setlbl)
            Me.Invoke(d, New Object() {text})
        Else
            If text = "" Then
                PLabel4.Visible = False

            Else
                PLabel4.Visible = True

            End If
        End If
    End Sub
    Private Sub SetText(ByVal text As String)
        If lblCount.InvokeRequired Then
            Dim d As SetlblCount = New SetlblCount(AddressOf SetText)
            Me.Invoke(d, New Object() {text})
        Else
            If text = "" Then
                lblCount.Visible = False
                Me.Width = 327
                Me.Height = 180
            Else
                lblCount.Visible = True
                Me.Width = 327
                Me.Height = 327
            End If
            lblCount.Text = text
        End If
    End Sub
    Private Sub SetText1(ByVal text As String)
         
    Private Sub SetText2(ByVal text As String)
        If lblCount.InvokeRequired Then
            Dim d As SetTxt = New SetTxt(AddressOf SetText2)
            Me.Invoke(d, New Object() {text})
        Else
            PTextBox3.Text = text
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
         

        Catch ex As Exception
            CustomException.ShowDialogue(ex)
        End Try
    End Sub
    Public Property Ignore As Boolean = False
    Private Sub DoPaymentCalculationAction()
        Try
            Timer1.Start()

            _BLTransaction = New BLTransaction(BLTransactionForm.GetFormByID(cmbTransactionForm.EditValue), LookUpEdit1.EditValue)
            Dim trn1 As Trn_Transaction = Nothing
            If PTextBox1.Text <> "" Then
                trn1 = _BLTransaction.GetTransactionByNumber(CType(PTextBox1.Text, Integer))
                If trn1.IgnoreFinancialStatment = True Then
                    Ignore = True
                Else
                    Ignore = False
                End If
            End If
            If trn1 IsNot Nothing Then
                Trn = Ganjineh.Net.TaxLIbrary.ElectronicInvoice.GetStatus(trn1, BLAttachment.DownloadAttachedFile(BLAttachment.GetAttachmentList(0, 0).FirstOrDefault.AttachmentID))
            Else
                Trn = Nothing
            End If

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
            DoPaymentCalculationAction()
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
            Me.Width = 327
            Me.Height = 180
            AddKeyDownEvent(Me)
            PLabel4.Visible = False
            PLabel5.Visible = False
            PLabel6.Visible = False
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



    Private Sub FormPostProcess_KeyDown(sender As Object, e As KeyEventArgs) Handles btnSend.KeyDown, cmbTransactionForm.KeyDown, SimpleButton2.KeyDown, SimpleButton1.KeyDown
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


    Private Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click
        Try


            Dim lstColumn As New List(Of ColumnsSelector)
            lstColumn.Add(New ColumnsSelector("تاریخ ارسال فاکتور ↓↓", "JalaliDate", 70))
            lstColumn.Add(New ColumnsSelector("وضعیت فاکتور", "StatusTitle", 350))


            Dim lst As IEnumerable(Of Trn_FinancialStatementLog)

            If cmbTransactionForm.EditValue = 0 Then
                Throw New CustomException("نوع فرم را مشخص نمایید.")
            End If

            _BLTransaction = New BLTransaction(BLTransactionForm.GetFormByID(cmbTransactionForm.EditValue), LookUpEdit1.EditValue)
            Dim trn1 As Trn_Transaction = Nothing
            If PTextBox1.Text <> "" Then
                trn1 = _BLTransaction.GetTransactionByNumber(CType(PTextBox1.Text, Integer))

            End If

            If trn1 IsNot Nothing Then
                lst = BLFinancialStatementLog.GetFinancialStatementLogs(trn1.TransactionID)
            Else
                Throw New CustomException("اطلاعاتی جهت نمایش وجود ندارد.")
            End If

            Dim frmItem As New FormItemSelector(lst, lstColumn, "سوابق ارسال", "", "BankAc12222countCode", 0,,,,, False, False, True)
            frmItem.ShowDialog()
        Catch ex As Exception
            CustomException.ShowDialogue(ex)
        End Try
    End Sub


    Public Sub AddKeyDownEvent(ByVal obj As Control)
        Dim ctl As Control = obj.GetNextControl(obj, True)
        Do Until ctl Is Nothing
            AddHandler ctl.KeyDown, AddressOf AllControls_KeyDown
            ctl = obj.GetNextControl(ctl, True)
        Loop
    End Sub

    Public Sub AllControls_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.SendWait("{TAB}")
        End If
    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        Try


          
            Dim lst As IEnumerable(Of Trn_Transaction)

            If cmbTransactionForm.EditValue = 0 Then
                Throw New CustomException("نوع فرم را مشخص نمایید.")
            End If
            Dim ProductList = BLProductDef.GetAllProductList

            _BLTransaction = New BLTransaction(BLTransactionForm.GetFormByID(cmbTransactionForm.EditValue), LookUpEdit1.EditValue)

            If Val(PTextBox1.Text) = 0 Then
                Throw New CustomException("اطلاعاتی جهت نمایش وجود ندارد.")
            End If

            Dim TransactionObject = _BLTransaction.GetTransactionByNumber(Val(PTextBox1.Text))
            If TransactionObject.FinancialStatmentState = 0 Then
                TrnList = _BLTransaction.TransactionListForFinancialStatment(RangeTypeEnum.NumberRange, 30, BLTransactionForm.GetFormByCode(TransactionObject.FormCode).FormID, TransactionObject.FinancialYearID,
                                                                     TransactionObject.Number,
                                                                     TransactionObject.Number,
                                                                     Nothing,
                                                                     Nothing, True)
            ElseIf TransactionObject.FinancialStatmentState = 5 Then
                TrnList = BLTransaction.TransactionListForFinancialStatmentRejected(RangeTypeEnum.NumberRange, 30, BLTransactionForm.GetFormByCode(TransactionObject.FormCode).FormID, TransactionObject.FinancialYearID,
                                                                     TransactionObject.Number,
                                                                     TransactionObject.Number,
                                                                     Nothing,
                                                                     Nothing, True)
            ElseIf TransactionObject.FinancialStatmentState = 2 Then
                TrnList = BLTransaction.TransactionListForFinancialStatmentPendding(RangeTypeEnum.NumberRange, 30, BLTransactionForm.GetFormByCode(TransactionObject.FormCode).FormID, TransactionObject.FinancialYearID,
                                                                     TransactionObject.Number,
                                                                     TransactionObject.Number,
                                                                     Nothing,
                                                                     Nothing, True)
            ElseIf TransactionObject.FinancialStatmentState = 4 Or TransactionObject.FinancialStatmentState = 6 Then
                TrnList = BLTransaction.TransactionListForFinancialStatmentSent(RangeTypeEnum.NumberRange, 30, BLTransactionForm.GetFormByCode(TransactionObject.FormCode).FormID, TransactionObject.FinancialYearID,
                                                                     TransactionObject.Number,
                                                                     TransactionObject.Number,
                                                                     Nothing,
                                                                     Nothing, True, TransactionObject.TransactionID)
            End If

            If TrnList.Count = 0 Then
                Throw New CustomException("در بازه انتخاب شده اطلاعاتی جهت نمایش وجود ندارد.")
            End If
            Dim _blTri As New BLTransactionItem(TrnList.FirstOrDefault)
            TriList = New List(Of Trn_TransactionItem)
            For Each TrnObj In TrnList
                If TrnObj.triList Is Nothing Then
                    TrnObj.triList = BLTransactionItem.GetTransactionItemList(TrnObj.TransactionID).ToList
                    TrnObj.triList.FirstOrDefault.ProductFinancialStatementCode = ProductList.FirstOrDefault(Function(a) a.ProductID = TrnObj.triList.FirstOrDefault.ProductID).FinancialStatementCode
                End If
                TriList.AddRange(TrnObj.triList)
            Next

            Dim ftrnList As New List(Of Trn_Transaction)
            Dim FirstFinancialStatment As Integer = 0
            Dim TransactionCostList As BindingList(Of Trn_TransactionCost)
            TransactionCostList = BLTransaction.GetTrnCostList


            For Each TrnObj In TrnList
                Dim Serial As String = Nothing
                If (TrnObj.FinancialStatmentGenNo.ToString.Length = 5) Then
                    Serial = "00000" + TrnObj.FinancialStatmentGenNo.ToString
                End If

                If (TrnObj.FinancialStatmentGenNo.ToString.Length = 6) Then
                    Serial = "0000" + TrnObj.FinancialStatmentGenNo.ToString
                End If

                If (TrnObj.FinancialStatmentGenNo.ToString.Length = 7) Then
                    Serial = "000" + TrnObj.FinancialStatmentGenNo.ToString
                End If

                If (TrnObj.FinancialStatmentGenNo.ToString.Length = 8) Then
                    Serial = "00" + TrnObj.FinancialStatmentGenNo.ToString
                End If

                If (TrnObj.FinancialStatmentGenNo.ToString.Length = 9) Then
                    Serial = "0" + TrnObj.FinancialStatmentGenNo.ToString
                End If
                TrnObj.FinancialStatmentGenNoSerial = Serial

                If TrnObj.DstPartnerID IsNot Nothing Then
                    TrnObj.tob = (If(PartnerList.FirstOrDefault(Function(a) TrnObj.DstPartnerID = a.PartnerID).PartnerType = "R", 1, 2)).ToString
                End If
                If TrnObj.SrcPartnerID IsNot Nothing Then
                    TrnObj.tob = (If(PartnerList.FirstOrDefault(Function(a) TrnObj.SrcPartnerID = a.PartnerID).PartnerType = "R", 1, 2)).ToString
                End If
                If TrnObj.DstPartnerID IsNot Nothing Then
                    TrnObj.PartnerNationalCode = PartnerList.FirstOrDefault(Function(a) TrnObj.DstPartnerID = a.PartnerID).PartnerNationalCode
                    TrnObj.PartnerPostalCode = PartnerList.FirstOrDefault(Function(a) TrnObj.DstPartnerID = a.PartnerID).PartnerPostalCode
                End If
                If TrnObj.SrcPartnerID IsNot Nothing Then
                    TrnObj.PartnerNationalCode = PartnerList.FirstOrDefault(Function(a) TrnObj.SrcPartnerID = a.PartnerID).PartnerNationalCode
                    TrnObj.PartnerPostalCode = PartnerList.FirstOrDefault(Function(a) TrnObj.SrcPartnerID = a.PartnerID).PartnerPostalCode
                End If
                Dim NumericFormat As String = "#,##0.## (#,##0.##)"


                TrnObj.tprdis = (TrnObj.triList.Where(Function(a) a.TransactionID = TrnObj.TransactionID).Sum(Function(b) b.Amount * b.SaleUnitPrice))
                TrnObj.tdis = (TrnObj.triList.Where(Function(a) a.TransactionID = TrnObj.TransactionID).Sum(Function(b) b.Discount))
                TrnObj.tadis = (TrnObj.triList.Where(Function(a) a.TransactionID = TrnObj.TransactionID).Sum(Function(b) b.TotalPriceWithDiscount))
                TrnObj.tvam = (TrnObj.triList.Where(Function(a) a.TransactionID = TrnObj.TransactionID).Sum(Function(b) b.VATAmountForElectronicInvoice))
                TrnObj.tbill = (TrnObj.triList.Where(Function(a) a.TransactionID = TrnObj.TransactionID).Sum(Function(b) b.TotalAmountWithVATForElectronicInvoice))
                TrnObj.cap = ((TrnObj.triList.Where(Function(a) a.TransactionID = TrnObj.TransactionID).Sum(Function(b) b.TotalAmountWithVATForElectronicInvoice)) - (TrnObj.triList.Where(Function(a) a.TransactionID = TrnObj.TransactionID).Sum(Function(b) b.VATAmountForElectronicInvoice)))
                TrnObj.tvop = (TrnObj.triList.Where(Function(a) a.TransactionID = TrnObj.TransactionID).Sum(Function(b) b.VATAmountForElectronicInvoice))


                If TrnObj.triList.Any = False Then
                    TrnObj.FinancialTotalPrice = "0"
                    TrnObj.FinancialTotalPriceWithVAT = "0"
                Else
                    Dim SharedCostFactorSum As Decimal
                    Dim SetInfo As Func(Of Trn_TransactionCost, Trn_TransactionCost) =
                      Function(transactionCost)
                          transactionCost.Partner = transactionCost.Com_Partner
                          transactionCost.Com_Partner = Nothing
                          Return transactionCost
                      End Function

                    Dim TransactionCostLists = TransactionCostList.Where(Function(a) a.TransactionID = TrnObj.TransactionID).ToList.Select(Function(s) SetInfo(s)).ToList
                    If TransactionCostLists Is Nothing OrElse CType(TransactionCostLists, IEnumerable(Of Trn_TransactionCost)).Count = 0 Then
                        SharedCostFactorSum = 0
                    Else
                        SharedCostFactorSum = CType(TransactionCostLists, IEnumerable(Of Trn_TransactionCost)).Sum(Function(cf) cf.IsShared = True)
                    End If


                    Dim List As List(Of Trn_TransactionItem) = TrnObj.triList

                    If List.Any(Function(ti) ti.SaleUnitPriceCurrency IsNot Nothing) Then
                        Dim TotalPrice As Decimal = List.Sum(Function(Rows) Rows.SaleUnitPriceCurrency * Rows.Amount * Rows.TransactionObject.ExchangeRate)
                        TrnObj.FinancialTotalPrice = Math.Truncate(TotalPrice).ToString(NumericFormat)
                        If TrnObj.HasVoucher Then
                            Dim tmpTotalPriceWithVAT As Decimal = List.Sum(Function(Rows) Rows.PriceWithDiscount + Rows.VATAmount)
                            TrnObj.FinancialTotalPriceWithVAT = Math.Truncate(tmpTotalPriceWithVAT).ToString(NumericFormat)
                        Else
                            TrnObj.FinancialTotalPriceWithVAT = (Math.Floor((List.Sum(Function(Rows) Rows.PriceWithDiscount) + SharedCostFactorSum)).ToString(NumericFormat))
                        End If
                    Else
                        Dim tmpTotalPrice As Decimal = List.Sum(Function(Rows) Rows.UnitPrice_Wrapper * Rows.Amount)
                        TrnObj.FinancialTotalPrice = Math.Truncate(tmpTotalPrice).ToString(NumericFormat)

                        If TrnObj.HasVoucher Then
                            Dim tmpTotalPriceWithVAT As Decimal = List.Sum(Function(Rows) Rows.PriceWithDiscount + Rows.VATAmount)
                            TrnObj.FinancialTotalPriceWithVAT = Math.Truncate(tmpTotalPriceWithVAT).ToString(NumericFormat)
                        Else
                            TrnObj.FinancialTotalPriceWithVAT = (List.Sum(Function(Rows) Rows.PriceWithDiscount) + SharedCostFactorSum).ToString(NumericFormat)
                        End If
                    End If
                End If


                TrnObj.inty = "نوع اول"
                TrnObj.tins = GlobalParam.FinancialStatmentUserID
                If TrnObj.DstPartnerID IsNot Nothing Then
                    TrnObj.tob = (If(PartnerList.FirstOrDefault(Function(a) TrnObj.DstPartnerID = a.PartnerID).PartnerType = "R", "حقیقی", "حقوقی")).ToString
                    TrnObj.PartnerName = PartnerList.FirstOrDefault(Function(s) s.PartnerID = TrnObj.DstPartnerID).PartnerName
                End If
                If TrnObj.SrcPartnerID IsNot Nothing Then
                    TrnObj.tob = (If(PartnerList.FirstOrDefault(Function(a) TrnObj.SrcPartnerID = a.PartnerID).PartnerType = "R", "حقیقی", "حقوقی")).ToString
                    TrnObj.PartnerName = PartnerList.FirstOrDefault(Function(s) s.PartnerID = TrnObj.SrcPartnerID).PartnerName
                End If
                If TrnObj.FinancialStatmentGenNo IsNot Nothing Then
                    TrnObj.inno = TrnObj.FinancialStatmentGenNo
                End If
                Select Case TrnObj.FinancialStatmentState

                End Select

                ftrnList.Add(TrnObj)
            Next

            lst = ftrnList

            Dim frmItem As New FormItemSelector(lst, lstColumn, "برگه های انتخاب شده", "", "BankAc2222countCode", 0,,,,, False, False, True)
            frmItem.ShowDialog()
            CanClose = False
        Catch ex As Exception
            CustomException.ShowDialogue(ex)
        End Try
    End Sub
End Class
