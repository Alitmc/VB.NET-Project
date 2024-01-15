Imports System.ComponentModel
Imports PCommonTools.ExceptionHandler
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports System.Threading

Public Class FormRejectElectronicInvoice

    Public Property CanClose As Boolean = False
    Public Property TrnList As List(Of Trn_Transaction) = Nothing
    Public Property TriList As List(Of Trn_TransactionItem) = Nothing
    Public Property ftrnList As List(Of Trn_Transaction) = Nothing
    Public Property PartnerList As BindingList(Of Com_Partner)
    Public Property RangeTypeNo As Integer?
    Public Property FormID As Integer?
    Public Property FinancialYearID As Integer?
    Public Property Number As Integer?

    Private WithEvents _BLTransaction As BLTransaction
    Private th As Thread

    Private Delegate Sub SetStateFunc(state As Boolean)
    Public Sub New(_RangeType As Integer, _FormID As Integer, _FinancialYearID As Integer, _Number As Integer)
        InitializeComponent()
        RangeTypeNo = _RangeType
        FormID = _FormID
        FinancialYearID = _FinancialYearID
        Number = _Number
    End Sub

    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub SetProgressBarState(state)
        Try


            If InvokeRequired Then
                Invoke(New SetStateFunc(AddressOf SetProgressBarState), state)
            Else
                btnSend.Enabled = Not state
                Cursor = If(state, Cursors.WaitCursor, DefaultCursor)

            End If
        Catch ex As Exception

            Return

        End Try
    End Sub

    Private Sub DoPaymentCalculationAction()
        Try

            Dim ProductList = BLProductDef.GetAllProductList

            PartnerList = BLPartner.GetPartnerList
            If cmbTransactionForm.EditValue = 0 Then
                Throw New CustomException("نوع فرم را مشخص نمایید.")
            End If
            _BLTransaction = New BLTransaction(BLTransactionForm.GetFormByID(cmbTransactionForm.EditValue), slcRangeType.FinancialYearID)
            TrnList = _BLTransaction.GetTransactionListForRejectFinancialStatment(slcRangeType.SelectedRangeType, cmbSystem.EditValue, cmbTransactionForm.EditValue, slcRangeType.FinancialYearID,
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetStartNumber > 0, slcRangeType.GetStartNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetEndNumber > 0, slcRangeType.GetEndNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetStartDate, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetEndDate, Nothing), True)
            If TrnList.Count = 0 Then
                Throw New CustomException("اطلاعاتی جهت ابطال وجود ندارد.")
            End If

            'For Each trn In TrnList
            '    If trn.
            'Next
            If TrnList.Count > 1 Then
                Throw New CustomException("امکان ابطال بیش از 1 برگه به صورت یکجا وجود ندارد.")
            End If
            If ConfirmMessageBox("آیا از ابطال صورت حساب اطمینان دارید؟" & vbCr & "تعداد برگه های انتخابی: " & TrnList.Count, True, "") = System.Windows.Forms.DialogResult.Cancel Then
                Return
            End If


            TriList = New List(Of Trn_TransactionItem)
            For Each trn In TrnList
                TriList.AddRange(trn.triList)
            Next
            If TriList.Any(Function(a) a.ProductFinancialStatementCode Is Nothing OrElse (Not a.ProductFinancialStatementCode.Reverse.ToList.Contains("1") AndAlso Not a.ProductFinancialStatementCode.Reverse.ToList.Contains("2") AndAlso Not a.ProductFinancialStatementCode.Reverse.ToList.Contains("3") AndAlso Not a.ProductFinancialStatementCode.Reverse.ToList.Contains("4") AndAlso Not a.ProductFinancialStatementCode.Reverse.ToList.Contains("5") AndAlso Not a.ProductFinancialStatementCode.Reverse.ToList.Contains("6") AndAlso Not a.ProductFinancialStatementCode.Reverse.ToList.Contains("7") AndAlso Not a.ProductFinancialStatementCode.Reverse.ToList.Contains("8") AndAlso Not a.ProductFinancialStatementCode.Reverse.ToList.Contains("9") AndAlso Not a.ProductFinancialStatementCode.Reverse.ToList.Contains("0"))) Then
                Throw New CustomException("کالا با کد " & ProductList.FirstOrDefault(Function(v) v.ProductID = TriList.FirstOrDefault(Function(a) a.ProductFinancialStatementCode Is Nothing).ProductID).ProductCode & " فاقد کد سامانه میباشد و امکان ابطال وجود ندارد.")
            End If


            If TriList.Any(Function(a) a.ProductFinancialStatementCode.Length <> 13) Then
                Throw New CustomException("طول کد سامانه نامعتبر میباشد", "FinancialStatementCode")
            End If

            SetProgressBarState(True)

            Dim transactionItemlist As New List(Of Trn_TransactionItem)
            ftrnList = New List(Of Trn_Transaction)

            Dim FirstFinancialStatment As Integer = 0
            Dim TPrice As Decimal
            Dim TPriceWithVAT As Decimal
            Dim SharedCostFactorSum As Decimal
            Dim TransactionCostList As BindingList(Of Trn_TransactionCost)
            TransactionCostList = BLTransaction.GetTrnCostList
            For Each trn In TrnList
                If trn.FinancialStatmentState <> 4 Then
                    Throw New CustomException("امکان ابطال فاکتور شماره " & trn.Number & " وجود ندارد.")
                End If
                For Each tri In trn.triList
                    tri.TrnNumber = trn.Number.ToString
                    tri.JalaliDate = trn.JalaliTransactionDate.Remove(0, 2)
                    tri.FormInfo = trn.FormInfo

                    Dim SetInfo As Func(Of Trn_TransactionCost, Trn_TransactionCost) =
              Function(transactionCost)
                  transactionCost.Partner = transactionCost.Com_Partner
                  transactionCost.Com_Partner = Nothing
                  Return transactionCost
              End Function

                    Dim TransactionCostLists = TransactionCostList.Where(Function(a) a.TransactionID = tri.TransactionID).ToList.Select(Function(s) SetInfo(s)).ToList
                    If TransactionCostLists Is Nothing OrElse CType(TransactionCostLists, IEnumerable(Of Trn_TransactionCost)).Count = 0 Then
                        SharedCostFactorSum = 0
                    Else
                        SharedCostFactorSum = CType(TransactionCostLists, IEnumerable(Of Trn_TransactionCost)).Sum(Function(cf) cf.IsShared = True)
                    End If


                    If trn.triList.Any(Function(ti) ti.SaleUnitPriceCurrency IsNot Nothing) Then
                        Dim TotalPrice As Decimal = trn.triList.Sum(Function(Rows) Rows.SaleUnitPriceCurrency * Rows.Amount * Rows.TransactionObject.ExchangeRate)
                        TPrice = Math.Truncate(TotalPrice).ToString("#,##0.##;(#,##0.##)")
                        If trn.IsOfficial IsNot Nothing AndAlso trn.IsOfficial Then
                            Dim tmpTotalPriceWithVAT As Decimal = trn.triList.Sum(Function(Rows) Rows.PriceWithDiscount + Rows.VATAmount)
                            TPriceWithVAT = Math.Truncate(tmpTotalPriceWithVAT).ToString("#,##0.##;(#,##0.##)")
                        Else
                            TPriceWithVAT = (Math.Floor((trn.triList.Sum(Function(Rows) Rows.PriceWithDiscount) + SharedCostFactorSum)).ToString("#,##0.##;(#,##0.##)"))
                        End If
                    Else
                        Dim tmpTotalPrice As Decimal = trn.triList.Sum(Function(Rows) Rows.UnitPrice_Wrapper * Rows.Amount)
                        TPrice = Math.Truncate(tmpTotalPrice).ToString("#,##0.##;(#,##0.##)")

                        If trn.IsOfficial IsNot Nothing AndAlso trn.IsOfficial Then
                            Dim tmpTotalPriceWithVAT As Decimal = trn.triList.Sum(Function(Rows) Rows.PriceWithDiscount + Rows.VATAmount)
                            TPriceWithVAT = Math.Truncate(tmpTotalPriceWithVAT).ToString("#,##0.##;(#,##0.##)")
                        Else
                            TPriceWithVAT = (trn.triList.Sum(Function(Rows) Rows.PriceWithDiscount) + SharedCostFactorSum).ToString("#,##0.##;(#,##0.##)")
                        End If
                    End If

                    tri.adis = tri.TotalPrice_Wrapper - tri.Discount
                    tri.am = String.Format("{0:#,0.##}", tri.Amount_Wrapper)
                    tri.dis = tri.Discount
                    tri.fee = tri.UnitPrice_Wrapper
                    tri.ins = If(trn.FormCode = 102000, 1, 4)
                    tri.prdis = tri.TotalPrice_Wrapper
                    tri.sstid = (ProductList.FirstOrDefault(Function(a) a.ProductID = tri.ProductID).FinancialStatementCode).ToString
                    tri.tadis = TPriceWithVAT
                    tri.sstt = If(tri.ProductCommercialName IsNot Nothing, tri.ProductCommercialName, tri.ProductName)
                    tri.tbill = tri.tadis
                    If trn.DstPartnerID IsNot Nothing Then
                        trn.tob = (If(PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerType = "R", 1, 2)).ToString
                    End If
                    If trn.SrcPartnerID IsNot Nothing Then
                        trn.tob = (If(PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerType = "R", 1, 2)).ToString
                    End If
                    tri.tprdis = TPrice
                    tri.tsstam = tri.TotalAmountWith_VAT_Discount_ImpactFactors
                    tri.tvam = trn.triList.Sum(Function(r) r.VATAmount)
                    tri.vam = tri.VATAmount
                    transactionItemlist.Add(tri)
                Next

                Dim unixTimestamp As Integer = CInt(trn.TransactionDate.Subtract(New DateTime(1970, 1, 1)).TotalSeconds) - 16200
                trn.indatim = unixTimestamp

                If FirstFinancialStatment = 0 Then
                    Dim LastFinancialStatment = BLTransaction.GetLastFinancialStatmentNo()
                    If LastFinancialStatment = 1 Then
                        trn.FinancialStatmentGenNo = 10000
                    Else
                        trn.FinancialStatmentGenNo = LastFinancialStatment + 1
                    End If

                    FirstFinancialStatment = trn.FinancialStatmentGenNo
                Else
                    trn.FinancialStatmentGenNo = Val(FirstFinancialStatment) + 1
                    FirstFinancialStatment = trn.FinancialStatmentGenNo
                End If


                ftrnList.Add(trn)
            Next

            TriList = transactionItemlist


            Dim att = BLAttachment.GetAttachmentList(0, 0)
            If att.Count = 0 Then
                InformationMessageBox("کلید خصوصی یافت نشد.")
                Return
            End If



            For Each trn In ftrnList
                If GlobalParam.OtherPartnerID <> If(trn.DstPartnerID, trn.SrcPartnerID) Then
                    If trn.DstPartnerID IsNot Nothing Then
                        If PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerType = "R" Then
                            trn.PartnerType = "1"
                            If PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerNationalCode Is Nothing Then
                                Throw New CustomException("طرف تجاری به کد " & PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerCode & " فاقد شناسه/کد ملی میباشد.")
                            End If
                            If PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerNationalCode.Length <> 10 Then
                                Throw New CustomException("فرمت مربوط به شناسه/کد ملی طرف تجاری " & PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerCode & " با نوع طرف تجاری همخوانی ندارد.")
                            End If
                            If IsValidNationalCode(PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID), PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerNationalCode) = False Then
                                Throw New CustomException("فرمت مربوط به شناسه/کد ملی طرف تجاری " & PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerCode & " صحیح نیست.")
                            End If
                            If PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerPostalCode Is Nothing Then
                                Throw New CustomException("طرف تجاری به کد " & PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerCode & " فاقد کد پستی میباشد.")
                            End If
                        End If
                        If PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerType = "L" Then
                            trn.PartnerType = "2"
                            If PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerNationalCode Is Nothing Then
                                Throw New CustomException("طرف تجاری به کد " & PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerCode & " فاقد شناسه/کد ملی میباشد.")
                            End If
                            If PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerNationalCode.Length <> 11 Then
                                Throw New CustomException("فرمت مربوط به شناسه/کد ملی طرف تجاری " & PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerCode & " با نوع طرف تجاری همخوانی ندارد.")
                            End If
                        End If
                    Else
                        If trn.SrcPartnerID IsNot Nothing Then
                            If PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerType = "R" Then
                                trn.PartnerType = "1"
                                If PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerNationalCode Is Nothing Then
                                    Throw New CustomException("طرف تجاری به کد " & PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerCode & " فاقد شناسه/کد ملی میباشد.")
                                End If
                                If PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerNationalCode.Length <> 10 Then
                                    Throw New CustomException("فرمت مربوط به شناسه/کد ملی طرف تجاری " & PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerCode & " با نوع طرف تجاری همخوانی ندارد.")
                                End If
                                If IsValidNationalCode(PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID), PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerNationalCode) = False Then
                                    Throw New CustomException("فرمت مربوط به شناسه/کد ملی طرف تجاری " & PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerCode & " صحیح نیست.")
                                End If
                                If PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerPostalCode Is Nothing Then
                                    Throw New CustomException("طرف تجاری به کد " & PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerCode & " فاقد کد پستی میباشد.")
                                End If
                            End If
                            If PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerType = "L" Then
                                trn.PartnerType = "2"
                                If PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerNationalCode Is Nothing Then
                                    Throw New CustomException("طرف تجاری به کد " & PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerCode & " فاقد شناسه/کد ملی میباشد.")
                                End If
                                If PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerNationalCode.Length <> 11 Then
                                    Throw New CustomException("فرمت مربوط به شناسه/کد ملی طرف تجاری " & PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerCode & " با نوع طرف تجاری همخوانی ندارد.")
                                End If
                            End If
                        End If
                    End If
                End If
            Next

            If BLAttachment.GetAttachmentList(0, 0).FirstOrDefault.AttachmentID = Nothing Then
                Throw New CustomException("کلید خصوصی مشخص نشده است.")
            End If
            If GlobalParam.UniqueTaxMemoryID Is Nothing Then
                Throw New CustomException("شناسه یکتا حافظه مالیاتی مشخص نشده است.")
            End If
            If GlobalParam.FinancialStatmentUserID Is Nothing Then
                Throw New CustomException("شناسه ملی شرکت مشخص نشده است.")
            End If
            Dim UnitList = BLUnit.GetUnitList()
            For Each tri In TriList
                If (UnitList.FirstOrDefault(Function(a) a.UnitID = tri.UnitID).FinancialStatementUnitID) Is Nothing Then
                    Throw New CustomException("واحد معادل سامانه مودیان برای واحد " & UnitList.FirstOrDefault(Function(a) a.UnitID = tri.UnitID).UnitName & "  مشخص نشده است")
                End If
            Next

            For Each tri In TriList
                If tri.sstid Is Nothing Then
                    Throw New CustomException("کد کالا معادل سامانه مودیان برای کالا " & ProductList.FirstOrDefault(Function(a) a.ProductID = tri.ProductID).ProductCode & "  مشخص نشده است")
                End If
            Next

            Ganjineh.Net.TaxLIbrary.ElectronicInvoice.Reject(BLAttachment.DownloadAttachedFile(BLAttachment.GetAttachmentList(0, 0).FirstOrDefault.AttachmentID), TriList, ftrnList, PartnerList, cmbTransactionForm.EditValue, slcRangeType.SelectedRangeType, cmbSystem.EditValue, cmbTransactionForm.EditValue, slcRangeType.FinancialYearID,
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetStartNumber > 0, slcRangeType.GetStartNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetEndNumber > 0, slcRangeType.GetEndNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetStartDate, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetEndDate, Nothing), True)


            SetProgressBarState(False)

        Catch ex As Exception
            If TypeOf ex Is CalculationExceptionWithoutMessage Then
                Return
            End If
            If ex.Message = "Thread was being aborted." Or ex.Message = "اشكال در ذخيره اطلاعات." Then
                Return
            End If
            CustomException.ShowDialogue(ex)
        Finally
            SetProgressBarState(False)
        End Try
    End Sub

    Private Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
        Try

            'th = New Thread(New ThreadStart(AddressOf DoPaymentCalculationAction))
            'th.Start()
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
        Dim num0 = Convert.ToInt32(chArray(0).ToString()) * 10
        Dim num2 = Convert.ToInt32(chArray(1).ToString()) * 9
        Dim num3 = Convert.ToInt32(chArray(2).ToString()) * 8
        Dim num4 = Convert.ToInt32(chArray(3).ToString()) * 7
        Dim num5 = Convert.ToInt32(chArray(4).ToString()) * 6
        Dim num6 = Convert.ToInt32(chArray(5).ToString()) * 5
        Dim num7 = Convert.ToInt32(chArray(6).ToString()) * 4
        Dim num8 = Convert.ToInt32(chArray(7).ToString()) * 3
        Dim num9 = Convert.ToInt32(chArray(8).ToString()) * 2
        Dim a = Convert.ToInt32(chArray(9).ToString())
        Dim b = (((((((num0 + num2) + num3) + num4) + num5) + num6) + num7) + num8) + num9
        Dim c = b Mod 11
        Return (((c < 2) AndAlso (a = c)) OrElse ((c >= 2) AndAlso ((11 - c) = a)))
    End Function


    Private Sub FormFinancialStatment_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Invoke(New CloseFunc(AddressOf CloseMe))
            cmbTransactionForm.Focus()
            slcRangeType.Initialize()
            PartnerList = BLPartner.GetPartnerList
            Dim source As List(Of Trn_TransactionForm) = GlobalParam.TransactionFormsList.Where(Function(a) a.FormCode.ToString.Remove(3) = 103 Or a.FormCode.ToString.Remove(3) = 102).ToList
            cmbTransactionForm.Properties.DataSource = source
            cmbTransactionForm.EditValue = source.FirstOrDefault(Function(a) a.FormCode.ToString.Remove(3) = 102).FormID

            Dim Dic As New Dictionary(Of Integer, String)
            Dic.Add(30, "خرید و فروش")
            cmbSystem.Properties.DataSource = Dic
            cmbSystem.EditValue = 30

            If RangeTypeNo IsNot Nothing Then
                slcRangeType.SetRangeInfo(New RangeInfo(BLFinancialYear.GetAllFinancialYearList.FirstOrDefault(Function(a) a.FinancialYearID = FinancialYearID), RangeTypeEnum.NumberRange,
                                              Nothing, Nothing, Nothing, Nothing,
                                              Number,
                                              Number))
                slcRangeType.Enabled = False
                cmbTransactionForm.Enabled = False
            End If

        Catch ex As Exception
            CustomException.ShowDialogue(ex)
        End Try
    End Sub

    Private Delegate Sub CloseFunc()
    Private Sub CloseMe()
        If InvokeRequired Then
            Invoke(New CloseFunc(AddressOf CloseMe))
        Else
            If th Is Nothing Then Return
            th.Abort()
            Dispose(True)
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


    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs)

        CanClose = True
        If th IsNot Nothing Then
            th.Abort()
        End If
        CloseMe()
    End Sub

    Private Sub FormPostProcess_KeyDown(sender As Object, e As KeyEventArgs) Handles cmbTransactionForm.KeyDown, slcRangeType.KeyDown, btnSend.KeyDown
        Try
            If e.KeyCode = Keys.Escape Then
                CanClose = True
                If th IsNot Nothing Then
                    th.Abort()
                End If
                CloseMe()
            End If
        Catch ex As Exception
            CustomException.ShowDialogue(ex)
        End Try
    End Sub


    Private Sub GroupControl2_CustomButtonClick(sender As Object, e As DevExpress.XtraBars.Docking2010.BaseButtonEventArgs) Handles GroupControl2.CustomButtonClick

        CanClose = True
        If th IsNot Nothing Then
            th.Abort()
        End If
        CloseMe()
        Me.Close()
    End Sub

End Class