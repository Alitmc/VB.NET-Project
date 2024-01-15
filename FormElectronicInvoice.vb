Imports System.ComponentModel
Imports PCommonTools.ExceptionHandler
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports System.Threading
Imports DevExpress.XtraBars.Docking2010

Public Class FormElectronicInvoice

    Public Property CanClose As Boolean = False
    Public Property TrnList As List(Of Trn_Transaction) = Nothing
    Public Property TriList As List(Of Trn_TransactionItem) = Nothing
    Public Property ftrnList As List(Of Trn_Transaction) = Nothing
    Public Property PartnerList As BindingList(Of Com_Partner)

    Public Property RangeTypeNo As Integer?
    Public Property FormID As Integer?
    Public Property FinancialYearID As Integer?
    Public Property Number As Integer?


    Private WithEvents _BLElectronicInvoice As TaxLIbrary.ElectronicInvoice
    Private WithEvents _BLTransaction As BLTransaction
    Private th As Thread
    Private th1 As Thread

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

    Public Property TrnListCount As Integer = 0
    Public Property TrnListSentCount As Integer = 0

    Private Delegate Sub SetStateFunc(state As Boolean)
    Private Sub SetProgressBarState(state)
        Try


            If InvokeRequired Then
                Invoke(New SetStateFunc(AddressOf SetProgressBarState), state)

            Else
                btnSend.Enabled = Not state
                Cursor = If(state, Cursors.WaitCursor, DefaultCursor)
                prgPaymentCalculation.Visible = state
                If prgPaymentCalculation.Visible = False Then
                    lblCount.Text = ""
                End If

            End If

        Catch ex As Exception

            Return

        End Try
    End Sub


    Delegate Sub SetlblCount(ByVal text As String)

    Private Sub SetText(ByVal text As String)
        If lblCount.InvokeRequired Then
            Dim d As SetlblCount = New SetlblCount(AddressOf SetText)
            Me.Invoke(d, New Object() {text})
        Else
            prgPaymentCalculation.Text = Math.Floor(TrnListSentCount / TrnListCount * 100).ToString & " % "
            lblCount.Text = text
        End If
    End Sub

    Private Sub timer1_tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            If TrnListCount = 0 Then
                Return
            End If

            TrnListSentCount = Ganjineh.Net.TaxLIbrary.ElectronicInvoice.SentList.Count
            SetText("تعداد ارسالی: " & TrnListSentCount & " / " & TrnListCount)


            If TrnListSentCount = TrnListCount Then
                Ganjineh.Net.TaxLIbrary.ElectronicInvoice.SentList = New List(Of Integer)
                TrnListCount = 0
            End If
        Catch ex As Exception
            CustomException.ShowDialogue(ex)
        End Try
    End Sub

    Private Sub DoPaymentCalculationAction()
        Try
            Timer1.Start()
            CanClose = False

            SetProgressBarState(True)

            Dim ProductList = BLProductDef.GetAllProductList

            PartnerList = BLPartner.GetPartnerList
            If cmbTransactionForm.EditValue = 0 Then
                Throw New CustomException("نوع فرم را مشخص نمایید.")
            End If

            _BLTransaction = New BLTransaction(BLTransactionForm.GetFormByID(cmbTransactionForm.EditValue), slcRangeType.FinancialYearID)


            If RangeTypeNo Is Nothing Then
                'ارسال از مسیر تب سامانه مودیان
                TrnList = _BLTransaction.TransactionListForFinancialStatment(slcRangeType.SelectedRangeType, cmbSystem.EditValue, cmbTransactionForm.EditValue, slcRangeType.FinancialYearID,
                                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetStartNumber > 0, slcRangeType.GetStartNumber, Nothing),
                                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetEndNumber > 0, slcRangeType.GetEndNumber, Nothing),
                                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetStartDate, Nothing),
                                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetEndDate, Nothing), True)

            Else   'Transaction ارسال از مسیر فرم
                Dim TransactionObject = _BLTransaction.GetTransactionByNumber(Number)
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
                End If
            End If

            If TrnList.Count = 0 Then
                Throw New CustomException("در بازه انتخاب شده اطلاعاتی جهت ارسال وجود ندارد.")
            End If

            TrnListCount = TrnList.Count
            If RangeTypeNo Is Nothing Then
                If ConfirmMessageBox("آیا از ارسال صورت حساب اطمینان دارید؟" & vbCr & "تعداد برگه های انتخابی: " & TrnList.Count, True, "") = System.Windows.Forms.DialogResult.Cancel Then
                    Return
                End If
            End If

            If TrnList.Count > 30 Then
                Throw New CustomException("امکان ارسال بیش از 30 برگه به صورت یکجا وجود ندارد.")
            End If

            TriList = New List(Of Trn_TransactionItem)
            For Each trn In TrnList
                If trn.triList Is Nothing Then
                    trn.triList = BLTransactionItem.GetTransactionItemList(trn.TransactionID).ToList
                    trn.triList.FirstOrDefault.ProductFinancialStatementCode = ProductList.FirstOrDefault(Function(a) a.ProductID = trn.triList.FirstOrDefault.ProductID).FinancialStatementCode
                End If
                TriList.AddRange(trn.triList)
            Next
            If TriList.Any(Function(a) a.ProductFinancialStatementCode Is Nothing OrElse (a.ProductFinancialStatementCode.Contains(" "))) Then
                Throw New CustomException("کالا با کد " & ProductList.FirstOrDefault(Function(v) v.ProductID = TriList.FirstOrDefault(Function(a) a.ProductFinancialStatementCode Is Nothing OrElse (a.ProductFinancialStatementCode.Contains(" "))).ProductID).ProductCode & " فاقد کد سامانه میباشد و امکان ارسال وجود ندارد.")
            End If


            If TriList.Any(Function(a) a.ProductFinancialStatementCode.Length <> 13) Then
                Throw New CustomException("طول کد سامانه نامعتبر میباشد", "FinancialStatementCode")
            End If


            Dim transactionItemlist As New List(Of Trn_TransactionItem)
            ftrnList = New List(Of Trn_Transaction)

            Dim FirstFinancialStatment As Integer = 0
            Dim TPrice As Decimal
            Dim TPriceWithVAT As Decimal
            Dim SharedCostFactorSum As Decimal
            Dim TransactionCostList As BindingList(Of Trn_TransactionCost)
            TransactionCostList = BLTransaction.GetTrnCostList
            For Each trn In TrnList
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
                If trn.FinancialStatmentState = 0 Then
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
                            If Not (PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerNationalCode.Length >= 11 AndAlso PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerNationalCode.Length <= 14) Then
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
                                If Not (PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerNationalCode.Length >= 11 AndAlso PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerNationalCode.Length <= 14) Then
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

            If RangeTypeNo Is Nothing Then
                Ganjineh.Net.TaxLIbrary.ElectronicInvoice.Send(BLAttachment.DownloadAttachedFile(BLAttachment.GetAttachmentList(0, 0).FirstOrDefault.AttachmentID), TriList, ftrnList, PartnerList, cmbTransactionForm.EditValue, slcRangeType.SelectedRangeType, cmbSystem.EditValue, cmbTransactionForm.EditValue, slcRangeType.FinancialYearID,
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetStartNumber > 0, slcRangeType.GetStartNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetEndNumber > 0, slcRangeType.GetEndNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetStartDate, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetEndDate, Nothing), True)
            Else
                If ftrnList.FirstOrDefault.FinancialStatmentState = 0 Then
                    Ganjineh.Net.TaxLIbrary.ElectronicInvoice.Send(BLAttachment.DownloadAttachedFile(BLAttachment.GetAttachmentList(0, 0).FirstOrDefault.AttachmentID), TriList, ftrnList, PartnerList, cmbTransactionForm.EditValue, slcRangeType.SelectedRangeType, cmbSystem.EditValue, cmbTransactionForm.EditValue, slcRangeType.FinancialYearID,
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetStartNumber > 0, slcRangeType.GetStartNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetEndNumber > 0, slcRangeType.GetEndNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetStartDate, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetEndDate, Nothing), True)


                ElseIf ftrnList.FirstOrDefault.FinancialStatmentState = 5 Then
                    Ganjineh.Net.TaxLIbrary.ElectronicInvoice.Send(BLAttachment.DownloadAttachedFile(BLAttachment.GetAttachmentList(0, 0).FirstOrDefault.AttachmentID), TriList, ftrnList, PartnerList, cmbTransactionForm.EditValue, slcRangeType.SelectedRangeType, cmbSystem.EditValue, cmbTransactionForm.EditValue, slcRangeType.FinancialYearID,
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetStartNumber > 0, slcRangeType.GetStartNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetEndNumber > 0, slcRangeType.GetEndNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetStartDate, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetEndDate, Nothing), True)


                ElseIf ftrnList.FirstOrDefault.FinancialStatmentState = 2 Then
                    TrnList = BLTransaction.TransactionListForFinancialStatmentPendding(RangeTypeEnum.NumberRange, 30, BLTransactionForm.GetFormByCode(ftrnList.FirstOrDefault.FormCode).FormID, ftrnList.FirstOrDefault.FinancialYearID,
                                                                     ftrnList.FirstOrDefault.Number,
                                                                     ftrnList.FirstOrDefault.Number,
                                                                     Nothing,
                                                                     Nothing, True)
                    Ganjineh.Net.TaxLIbrary.ElectronicInvoice.GetPenddings(TrnList, BLAttachment.DownloadAttachedFile(BLAttachment.GetAttachmentList(0, 0).FirstOrDefault.AttachmentID))

                    Dim trn1 As Trn_Transaction = BLTransaction.GetByID(TrnList.FirstOrDefault.TransactionID)

                    If trn1 IsNot Nothing Then
                        Dim Trnobj = Ganjineh.Net.TaxLIbrary.ElectronicInvoice.GetStatus(trn1, BLAttachment.DownloadAttachedFile(BLAttachment.GetAttachmentList(0, 0).FirstOrDefault.AttachmentID))
                        Select Case Trnobj.FinancialStatmentState
                            Case 2
                                PCommonTools.ErrorMessageBox("در انتظار پاسخ از سامانه")
                            Case 4
                                PCommonTools.InformationMessageBox("ارسال با موفقیت انجام شد." & vbCr & "وضعیت برگه: " + "ثبت موفق در سامانه",)
                            Case 5
                                PCommonTools.ErrorMessageBox("ارسال با موفقیت انجام شد." & vbCr & "وضعیت برگه: " + "رد شده توسط سامانه")
                        End Select
                    Else
                        Dim Trnobj = Nothing
                    End If

                End If

            End If

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
            Timer1.Stop()
            SetProgressBarState(False)
        End Try
    End Sub

    Private Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
        Try
            Ganjineh.Net.TaxLIbrary.ElectronicInvoice.SentList = New List(Of Integer)
            th = New Thread(New ThreadStart(AddressOf DoPaymentCalculationAction))
            th.Start()
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

    Private Sub _BLVoucher_VoucherDeleteEnded(sender As Object, e As EventArgs) Handles _BLElectronicInvoice.TrnSent

    End Sub


    Private Sub _BLVoucher_VoucherDeleting(sender As Object, e As VoucherDeletingEventArgs) Handles _BLElectronicInvoice.TrnSending
        Try

            Application.DoEvents()
        Catch ex As Exception
            PCommonTools.ExceptionHandler.CustomException.ShowDialogue(ex)
        End Try
    End Sub

    Private Sub _BLVoucher_VoucherDeleteStarted(sender As Object, e As EventArgs) Handles _BLElectronicInvoice.TrnSendStarted
        Try


        Catch ex As Exception
            PCommonTools.ExceptionHandler.CustomException.ShowDialogue(ex)
        End Try
    End Sub

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


    Private Sub btnShow_Click(sender As Object, e As EventArgs) Handles btnShow.Click
        Try


            Dim lstColumn As New List(Of ColumnsSelector)
            lstColumn.Add(New ColumnsSelector("شماره فاکتور", "Number", 70))
            lstColumn.Add(New ColumnsSelector("شماره فاکتور چاپی", "TmpNumber", 70))
            lstColumn.Add(New ColumnsSelector("تاریخ فاکتور", "JalaliTransactionDate", 450))
            lstColumn.Add(New ColumnsSelector("مجموع مبلغ قبل از کسر تخفیف", "tprdis", 500, "{0:#,0.##}", FormatTypeEnum.Numeric))
            lstColumn.Add(New ColumnsSelector("مجموع تخفیفات", "tdis", 500, "{0:#,0.##}", FormatTypeEnum.Numeric))
            lstColumn.Add(New ColumnsSelector("مجموع مبلغ پس از کسر تخفیف", "tadis", 500, "{0:#,0.##}", FormatTypeEnum.Numeric))
            lstColumn.Add(New ColumnsSelector("مجموع مالیات بر ارزش افزوده", "tvam", 400, "{0:#,0.##}", FormatTypeEnum.Numeric))
            lstColumn.Add(New ColumnsSelector("مجموع صورتحساب", "tbill", 500, "{0:#,0.##}", FormatTypeEnum.Numeric))
            lstColumn.Add(New ColumnsSelector("مبلغ پرداختی نقدی", "cap", 500, "{0:#,0.##}", FormatTypeEnum.Numeric))
            lstColumn.Add(New ColumnsSelector("نوع صورتحساب", "inty", 300))
            lstColumn.Add(New ColumnsSelector("نوع شخص خریدار", "tob", 350))
            lstColumn.Add(New ColumnsSelector("نام خریدار", "PartnerName", 350))
            lstColumn.Add(New ColumnsSelector("شماره اقتصادی/کد ملی خریدار", "PartnerNationalCode", 350))
            lstColumn.Add(New ColumnsSelector("کد پستی خریدار", "PartnerPostalCode", 400))
            lstColumn.Add(New ColumnsSelector("سریال صورتحساب داخلی", "FinancialStatmentGenNo", 300))
            lstColumn.Add(New ColumnsSelector("شماره مالیاتی", "FinancialStatmentTaxID", 300))
            lstColumn.Add(New ColumnsSelector("وضعیت فاکتور", "FinancialStatmentStateTitle", 350))

            Dim lst As IEnumerable(Of Trn_Transaction)

            If cmbTransactionForm.EditValue = 0 Then
                Throw New CustomException("نوع فرم را مشخص نمایید.")
            End If
            Dim ProductList = BLProductDef.GetAllProductList

            _BLTransaction = New BLTransaction(BLTransactionForm.GetFormByID(cmbTransactionForm.EditValue), slcRangeType.FinancialYearID)
            If RangeTypeNo Is Nothing Then
                TrnList = _BLTransaction.TransactionListForFinancialStatment(slcRangeType.SelectedRangeType, cmbSystem.EditValue, cmbTransactionForm.EditValue, slcRangeType.FinancialYearID,
                                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetStartNumber > 0, slcRangeType.GetStartNumber, Nothing),
                                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetEndNumber > 0, slcRangeType.GetEndNumber, Nothing),
                                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetStartDate, Nothing),
                                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetEndDate, Nothing), True)

            Else
                Dim TransactionObject = _BLTransaction.GetTransactionByNumber(Number)
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
                End If
            End If
            If TrnList.Count = 0 Then
                Throw New CustomException("در بازه انتخاب شده اطلاعاتی جهت نمایش وجود ندارد.")
            End If

            TriList = New List(Of Trn_TransactionItem)
            For Each trn In TrnList
                If trn.triList Is Nothing Then
                    trn.triList = BLTransactionItem.GetTransactionItemList(trn.TransactionID).ToList
                    trn.triList.FirstOrDefault.ProductFinancialStatementCode = ProductList.FirstOrDefault(Function(a) a.ProductID = trn.triList.FirstOrDefault.ProductID).FinancialStatementCode
                End If
                TriList.AddRange(trn.triList)
            Next

            Dim ftrnList As New List(Of Trn_Transaction)

            Dim TransactionCostList As BindingList(Of Trn_TransactionCost)
            TransactionCostList = BLTransaction.GetTrnCostList


            For Each trn In TrnList
                If trn.DstPartnerID IsNot Nothing Then
                    trn.tob = (If(PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerType = "R", 1, 2)).ToString
                End If
                If trn.SrcPartnerID IsNot Nothing Then
                    trn.tob = (If(PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerType = "R", 1, 2)).ToString
                End If
                If trn.DstPartnerID IsNot Nothing Then
                    trn.PartnerNationalCode = PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerNationalCode
                    trn.PartnerPostalCode = PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerPostalCode
                End If
                If trn.SrcPartnerID IsNot Nothing Then
                    trn.PartnerNationalCode = PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerNationalCode
                    trn.PartnerPostalCode = PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerPostalCode
                End If
                Dim NumericFormat As String = "#,##0.##;(#,##0.##)"


                trn.tprdis = (trn.triList.Where(Function(a) a.TransactionID = trn.TransactionID).Sum(Function(b) b.Amount * b.SaleUnitPrice))
                trn.tdis = (trn.triList.Where(Function(a) a.TransactionID = trn.TransactionID).Sum(Function(b) b.Discount))
                trn.tadis = (trn.triList.Where(Function(a) a.TransactionID = trn.TransactionID).Sum(Function(b) b.TotalPriceWithDiscount))
                trn.tvam = (trn.triList.Where(Function(a) a.TransactionID = trn.TransactionID).Sum(Function(b) b.VATAmountForElectronicInvoice))
                trn.tbill = (trn.triList.Where(Function(a) a.TransactionID = trn.TransactionID).Sum(Function(b) b.TotalAmountWithVATForElectronicInvoice))
                trn.cap = ((trn.triList.Where(Function(a) a.TransactionID = trn.TransactionID).Sum(Function(b) b.TotalAmountWithVATForElectronicInvoice)) - (trn.triList.Where(Function(a) a.TransactionID = trn.TransactionID).Sum(Function(b) b.VATAmountForElectronicInvoice)))
                trn.tvop = (trn.triList.Where(Function(a) a.TransactionID = trn.TransactionID).Sum(Function(b) b.VATAmountForElectronicInvoice))


                If trn.triList.Any = False Then
                    trn.FinancialTotalPrice = "0"
                    trn.FinancialTotalPriceWithVAT = "0"
                Else
                    Dim SharedCostFactorSum As Decimal
                    Dim SetInfo As Func(Of Trn_TransactionCost, Trn_TransactionCost) =
                      Function(transactionCost)
                          transactionCost.Partner = transactionCost.Com_Partner
                          transactionCost.Com_Partner = Nothing
                          Return transactionCost
                      End Function

                    Dim TransactionCostLists = TransactionCostList.Where(Function(a) a.TransactionID = trn.TransactionID).ToList.Select(Function(s) SetInfo(s)).ToList
                    If TransactionCostLists Is Nothing OrElse CType(TransactionCostLists, IEnumerable(Of Trn_TransactionCost)).Count = 0 Then
                        SharedCostFactorSum = 0
                    Else
                        SharedCostFactorSum = CType(TransactionCostLists, IEnumerable(Of Trn_TransactionCost)).Sum(Function(cf) cf.IsShared = True)
                    End If


                    Dim List As List(Of Trn_TransactionItem) = trn.triList

                    If List.Any(Function(ti) ti.SaleUnitPriceCurrency IsNot Nothing) Then
                        Dim TotalPrice As Decimal = List.Sum(Function(Rows) Rows.SaleUnitPriceCurrency * Rows.Amount * Rows.TransactionObject.ExchangeRate)
                        trn.FinancialTotalPrice = Math.Truncate(TotalPrice).ToString(NumericFormat)
                        If trn.HasVoucher Then
                            Dim tmpTotalPriceWithVAT As Decimal = List.Sum(Function(Rows) Rows.PriceWithDiscount + Rows.VATAmount)
                            trn.FinancialTotalPriceWithVAT = Math.Truncate(tmpTotalPriceWithVAT).ToString(NumericFormat)
                        Else
                            trn.FinancialTotalPriceWithVAT = (Math.Floor((List.Sum(Function(Rows) Rows.PriceWithDiscount) + SharedCostFactorSum)).ToString(NumericFormat))
                        End If
                    Else
                        Dim tmpTotalPrice As Decimal = List.Sum(Function(Rows) Rows.UnitPrice_Wrapper * Rows.Amount)
                        trn.FinancialTotalPrice = Math.Truncate(tmpTotalPrice).ToString(NumericFormat)

                        If trn.HasVoucher Then
                            Dim tmpTotalPriceWithVAT As Decimal = List.Sum(Function(Rows) Rows.PriceWithDiscount + Rows.VATAmount)
                            trn.FinancialTotalPriceWithVAT = Math.Truncate(tmpTotalPriceWithVAT).ToString(NumericFormat)
                        Else
                            trn.FinancialTotalPriceWithVAT = (List.Sum(Function(Rows) Rows.PriceWithDiscount) + SharedCostFactorSum).ToString(NumericFormat)
                        End If
                    End If
                End If



                Dim BidPartnerEconomicalNo As String = Nothing
                Dim TinbPartnerEconomicalNo As String = Nothing
                Dim BpcPartner As String = Nothing
                Dim InvoiceType As Integer = 1
                Dim PartnerType As Integer? = 0
                Dim InsNumber As Integer = 1
                Dim partner = PartnerList.FirstOrDefault(Function(a) a.PartnerID = trn.DstPartnerID)



                If GlobalParam.OtherPartnerID = 0 Then
                    trn.inty = "نوع اول"

                    If PartnerType = 1 Then
                BidPartnerEconomicalNo = partner.PartnerNationalCode
            Else
                TinbPartnerEconomicalNo = partner.PartnerNationalCode
                BpcPartner = Nothing
            End If
        Else

            If partner.PartnerID = GlobalParam.OtherPartnerID Then
                PartnerType = Nothing
                        trn.inty = "نوع دوم"
                    Else
                        trn.inty = "نوع اول"

                        If PartnerType = 1 Then
                    BidPartnerEconomicalNo = partner.PartnerNationalCode
                Else
                    TinbPartnerEconomicalNo = partner.PartnerNationalCode
                    BpcPartner = Nothing
                End If
            End If
        End If





                trn.tins = GlobalParam.FinancialStatmentUserID
                If trn.DstPartnerID IsNot Nothing Then
                    trn.tob = (If(PartnerList.FirstOrDefault(Function(a) trn.DstPartnerID = a.PartnerID).PartnerType = "R", "حقیقی", "حقوقی")).ToString
                    trn.PartnerName = PartnerList.FirstOrDefault(Function(s) s.PartnerID = trn.DstPartnerID).PartnerName
                End If
                If trn.SrcPartnerID IsNot Nothing Then
                    trn.tob = (If(PartnerList.FirstOrDefault(Function(a) trn.SrcPartnerID = a.PartnerID).PartnerType = "R", "حقیقی", "حقوقی")).ToString
                    trn.PartnerName = PartnerList.FirstOrDefault(Function(s) s.PartnerID = trn.SrcPartnerID).PartnerName
                End If
                If trn.FinancialStatmentGenNo IsNot Nothing Then
                    trn.inno = trn.FinancialStatmentGenNo
                End If
                Select Case trn.FinancialStatmentState
                    Case 0
                        trn.FinancialStatmentStateTitle = "ارسال نشده"
                    Case 1
                        trn.FinancialStatmentStateTitle = "ارسال شده"
                    Case 2
                        trn.FinancialStatmentStateTitle = "در انتظار پاسخ"
                    Case 4
                        trn.FinancialStatmentStateTitle = "تأیید شده"
                    Case 5
                        trn.FinancialStatmentStateTitle = "رد شده"
                    Case 6
                        trn.FinancialStatmentStateTitle = "ابطال شده در سامانه"
                End Select

                ftrnList.Add(trn)
            Next

            lst = ftrnList
            Dim frmItem As New FormItemSelector(lst, lstColumn, "برگه های انتخاب شده", "", "BankAc2222countCode", 0,,,,, False, False, True)
            frmItem.ShowDialog()
            CanClose = False
        Catch ex As Exception
            CustomException.ShowDialogue(ex)
        End Try
    End Sub

    Private Sub FormPostProcess_KeyDown(sender As Object, e As KeyEventArgs) Handles cmbTransactionForm.KeyDown, slcRangeType.KeyDown, btnSend.KeyDown, btnShow.KeyDown
        Try
            If prgPaymentCalculation.Visible = True Then
                If ConfirmMessageBox("آیا از بستن فرم اطمینان دارید؟", True, "") = System.Windows.Forms.DialogResult.Cancel Then
                    Return
                End If
            End If

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
        If prgPaymentCalculation.Visible = True Then
            If ConfirmMessageBox("آیا از بستن فرم اطمینان دارید؟", True, "") = System.Windows.Forms.DialogResult.Cancel Then
                Return
            End If
        End If
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


End Class