Imports System.ComponentModel
Imports PCommonTools.ExceptionHandler
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports DevExpress.XtraEditors.Controls


Public Class FormElectronicInvoiceChecking

    Public Property TrnList As List(Of Trn_Transaction) = Nothing
    Public Property TriList As List(Of Trn_TransactionItem) = Nothing
    Public Property PartnerList As BindingList(Of Com_Partner)
    Public Property ftrnList As List(Of Trn_Transaction) = Nothing

    Private Sub FormFinancialStatmentsChecking_Load(sender As Object, e As EventArgs) Handles Me.Load
        slcRangeType.Initialize()
        PartnerList = BLPartner.GetPartnerList
        Dim source As List(Of Trn_TransactionForm) = GlobalParam.TransactionFormsList.Where(Function(a) a.FormCode.ToString.Remove(3) = 103 Or a.FormCode.ToString.Remove(3) = 102).ToList
        cmbTransactionForm.Properties.DataSource = source
        cmbTransactionForm.EditValue = source.FirstOrDefault(Function(a) a.FormCode.ToString.Remove(3) = 102).FormID

        Dim Dic As New Dictionary(Of Integer, String)
        Dic.Add(30, "خرید و فروش")
        cmbSystem.Properties.DataSource = Dic
        cmbSystem.EditValue = 30
        gvTransaction.OptionsDetail.EnableMasterViewMode = False
    End Sub

    Private Sub btnRefresh_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnRefresh.ItemClick
        Try

            If cmbTransactionForm.EditValue = 0 Then
                Throw New CustomException("نوع فرم را مشخص نمایید.")
            End If
            Dim ProductList = BLProductDef.GetAllProductList
            TrnList = BLTransaction.TransactionListForFinancialStatmentRejected(slcRangeType.SelectedRangeType, cmbSystem.EditValue, cmbTransactionForm.EditValue, slcRangeType.FinancialYearID,
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetStartNumber > 0, slcRangeType.GetStartNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetEndNumber > 0, slcRangeType.GetEndNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetStartDate, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetEndDate, Nothing), True)
            If TrnList.Count = 0 Then
                Throw New CustomException("در بازه انتخاب شده اطلاعاتی جهت نمایش وجود ندارد.")
            End If

            For Each i In TrnList
                i.FinancialStatmentStateTitle = "رد شده"
            Next
            gvTransaction.Focus()
            For Each i In TrnList
                If i.RejectOrRepairStatus Is Nothing Then
                    i.RequestKind = "عادی"
                End If
                If i.RejectOrRepairStatus = 1 Then
                    i.RequestKind = "جهت ابطال"
                End If
                If i.RejectOrRepairStatus = 2 Then
                    i.RequestKind = "جهت اصلاح"
                End If
            Next
            gcTransaction.DataSource = TrnList

            gvTransaction.FocusedRowHandle = 0
            gvTransaction.FocusedColumn = GridColumn1
        Catch ex As Exception
            CustomException.ShowDialogue(ex)
        End Try
    End Sub

    Private Sub btnSend1_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnSend1.ItemClick
        Dim FocusedRow As Trn_Transaction = gvTransaction.GetFocusedRow
        If FocusedRow Is Nothing Then Return
        Try

            Dim ProductList = BLProductDef.GetAllProductList
            If ConfirmMessageBox("آیا از ارسال صورت حساب اطمینان دارید؟", True, "حذف") = System.Windows.Forms.DialogResult.Cancel Then
                Return
            End If
            If FocusedRow.RejectOrRepairStatus = 1 Then
                Throw New CustomException("امکان ارسال مجدد فاکتور ابطالی وجود ندارد،ابتدا لغو ابطال کرده و سپس مجدد فاکتور را ابطال نمایید .")
            End If
            If FocusedRow.RejectOrRepairStatus = 2 Then
                Throw New CustomException("امکان ارسال مجدد فاکتور اصلاحی وجود ندارد،ابتدا لغو اصلاح کرده و سپس مجدد فاکتور را اصلاح نمایید .")
            End If
            Label4.Visible = True
            Label5.Visible = True
            Label6.Visible = True
            Label7.Visible = True
            Label8.Visible = True
            FocusedRow.triList = BLTransactionItem.GetTransactionItemList(FocusedRow.TransactionID).ToList
            For Each i In FocusedRow.triList
                i.ProductFinancialStatementCode = ProductList.FirstOrDefault(Function(a) a.ProductID = i.ProductID).FinancialStatementCode
            Next
            PartnerList = BLPartner.GetPartnerList
            If cmbTransactionForm.EditValue = 0 Then
                Throw New CustomException("نوع فرم را مشخص نمایید.")
            End If
            Dim TrnLst As New List(Of Trn_Transaction)
            TrnLst.Add(FocusedRow)
            If TrnLst.Count = 0 Then
                Throw New CustomException("اطلاعاتی جهت ارسال وجود ندارد.")
            End If

            TriList = New List(Of Trn_TransactionItem)
            For Each trn In TrnLst
                TriList.AddRange(trn.triList)
            Next

            If TriList.Any(Function(a) a.ProductFinancialStatementCode Is Nothing OrElse (a.ProductFinancialStatementCode.Contains(" "))) Then
                Throw New CustomException("کالا با کد " & ProductList.FirstOrDefault(Function(v) v.ProductID = TriList.FirstOrDefault(Function(a) a.ProductFinancialStatementCode Is Nothing OrElse a.ProductFinancialStatementCode.Contains(" ")).ProductID).ProductCode & " فاقد کد سامانه میباشد و امکان ارسال وجود ندارد.")
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
            For Each trn In TrnLst
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
                            trn.FinancialStatmentGenNo = 1
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

           
            For Each i In TrnList
                i.FinancialStatmentStateTitle = "رد شده"
            Next
            For Each i In TrnList
                If i.RejectOrRepairStatus Is Nothing Then
                    i.RequestKind = "عادی"
                End If
                If i.RejectOrRepairStatus = 1 Then
                    i.RequestKind = "جهت ابطال"
                End If
                If i.RejectOrRepairStatus = 2 Then
                    i.RequestKind = "جهت اصلاح"
                End If
            Next
            gcTransaction.DataSource = TrnList
            Label4.Visible = False
            Label5.Visible = False
            Label6.Visible = False
            Label7.Visible = False
            Label8.Visible = False
        Catch ex As Exception
            CustomException.ShowDialogue(ex)

        End Try
    End Sub

    Private Sub cmbTransactionForm_EditValueChanged(sender As Object, e As EventArgs) Handles cmbTransactionForm.EditValueChanged
        gcTransaction.DataSource = Nothing
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

    Private Sub slcRangeType_Validated(sender As Object, e As EventArgs) Handles slcRangeType.Validated
        gcTransaction.DataSource = Nothing
    End Sub

    Private Sub RepositoryItemButtonEdit1_ButtonClick(sender As Object, e As ButtonPressedEventArgs) Handles RepositoryItemButtonEdit1.ButtonClick
        Dim FocusedRow As Trn_Transaction = gvTransaction.GetFocusedRow
        If FocusedRow Is Nothing Then Return
        Dim frm As New FormDescription(FocusedRow.ErrorMessage, Me.DisplayRectangle, False)
        frm.ShowDialog()

    End Sub

    Private Sub BarButtonItem1_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem1.ItemClick
        Dim FocusedRow As Trn_Transaction = gvTransaction.GetFocusedRow
        If FocusedRow Is Nothing Then Return
        Try

            Dim ProductList = BLProductDef.GetAllProductList
            If ConfirmMessageBox("آیا از لغو اصلاح صورت حساب اطمینان دارید؟", True, "حذف") = System.Windows.Forms.DialogResult.Cancel Then
                Return
            End If

            FocusedRow.triList = BLTransactionItem.GetTransactionItemList(FocusedRow.TransactionID).ToList
            For Each i In FocusedRow.triList
                i.ProductFinancialStatementCode = ProductList.FirstOrDefault(Function(a) a.ProductID = i.ProductID).FinancialStatementCode
            Next
            PartnerList = BLPartner.GetPartnerList
            If cmbTransactionForm.EditValue = 0 Then
                Throw New CustomException("نوع فرم را مشخص نمایید.")
            End If
            Dim TrnLst As New List(Of Trn_Transaction)
            TrnLst.Add(FocusedRow)
            Dim HistoryTrn As Trn_TransactionHistory = Nothing
            If TrnLst.Count = 0 Then
                Throw New CustomException("اطلاعاتی جهت لغو اصلاح وجود ندارد.")
            Else
                HistoryTrn = BLTransactionHistory.GetTransactionHistoryList(FocusedRow.TransactionID).LastOrDefault
                If HistoryTrn Is Nothing Then
                    Throw New CustomException("فاکتور انتخابی جهت اصلاح ارسال نشده است و لغو اصلاح امکان پذیر نیست.")
                End If
            End If


        

            Dim HistoryTrnList = BLTransactionHistory.GetTransactionHistoryList(FocusedRow.TransactionID)

            For Each trn In HistoryTrnList
                Dim _BLTransactionHistory As New BLTransactionHistory
                Dim _BLTransactionItemHistory As New BLTransactionItemHistory
                Dim AllTransactionItemHistoryList = BLTransactionItemHistory.GetAllTransactionItemHistoryList(trn.TransactionID)
                For Each tri In AllTransactionItemHistoryList
                    _BLTransactionItemHistory.Delete(tri)
                Next
                _BLTransactionHistory.Delete(trn)
            Next

            TrnList = BLTransaction.TransactionListForFinancialStatmentRejected(slcRangeType.SelectedRangeType, cmbSystem.EditValue, cmbTransactionForm.EditValue, slcRangeType.FinancialYearID,
                                                    If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetStartNumber > 0, slcRangeType.GetStartNumber, Nothing),
                                                    If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetEndNumber > 0, slcRangeType.GetEndNumber, Nothing),
                                                    If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetStartDate, Nothing),
                                                    If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetEndDate, Nothing), True)

            For Each i In TrnList
                i.FinancialStatmentStateTitle = "رد شده"
                If i.RejectOrRepairStatus Is Nothing Then
                    i.RequestKind = "عادی"
                End If
                If i.RejectOrRepairStatus = 1 Then
                    i.RequestKind = "جهت ابطال"
                End If
                If i.RejectOrRepairStatus = 2 Then
                    i.RequestKind = "جهت اصلاح"
                End If
            Next

            gcTransaction.DataSource = TrnList


            MessageBox.Show("لغو اصلاح با موفقیت انجام شد")
        Catch ex As Exception
            CustomException.ShowDialogue(ex)

        End Try
    End Sub

    Private Sub BarButtonItem2_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem2.ItemClick

        Dim TransactionObject As Trn_Transaction = gvTransaction.GetFocusedRow
        If TransactionObject Is Nothing Then Return
        Try

            Dim ProductList = BLProductDef.GetAllProductList
            If ConfirmMessageBox("آیا از لغو ابطال صورت حساب اطمینان دارید؟", True, "حذف") = System.Windows.Forms.DialogResult.Cancel Then
                Return
            End If
            If If(TransactionObject.RejectOrRepairStatus, 0) <> 1 Then
                Throw New CustomException("فاکتور انتخابی جهت ابطال ارسال نشده است و لغو ابطال امکان پذیر نیست.")
            End If

            TransactionObject.triList = BLTransactionItem.GetTransactionItemList(TransactionObject.TransactionID).ToList
            For Each i In TransactionObject.triList
                i.ProductFinancialStatementCode = ProductList.FirstOrDefault(Function(a) a.ProductID = i.ProductID).FinancialStatementCode
            Next
            Dim PartnerList = BLPartner.GetPartnerList

            Dim TrnLst As New List(Of Trn_Transaction)
            TrnLst.Add(TransactionObject)
            Dim HistoryTrn As Trn_TransactionHistory = Nothing
            If TrnLst.Count = 0 Then
                Throw New CustomException("اطلاعاتی جهت لغو ابطال وجود ندارد.")
            Else
                HistoryTrn = BLTransactionHistory.GetTransactionHistoryList(TransactionObject.TransactionID).LastOrDefault
                If HistoryTrn Is Nothing Then
                    Throw New CustomException("فاکتور انتخابی جهت اصلاح ارسال نشده است و لغو اصلاح امکان پذیر نیست.")
                End If
            End If

            Dim _BL As New BLTransactionHistory
            Dim _BLItem As New BLTransactionItemHistory
            Dim TransactionItemHistoryList = BLTransactionItemHistory.GetTransactionItemHistoryList(HistoryTrn.TransactionID)
            Dim TransactionItemList = BLTransactionItem.GetTransactionItemList(TransactionObject.TransactionID)

            Dim FinancialStatementLog As New Trn_FinancialStatementLog
            FinancialStatementLog.TransactionID = TransactionObject.TransactionID
            FinancialStatementLog.UserID = Context.CurrentUser.UserID
            FinancialStatementLog.EventDate = Now
            FinancialStatementLog.EventType = "ReturnToPast1"
            FinancialStatementLog.OriginalValue = TransactionObject.FinancialStatmentState.ToString()
            FinancialStatementLog.NewValue = TransactionObject.FinancialStatmentState.ToString()
            FinancialStatementLog.TryNo = BLFinancialStatementLog.GetTryNo(TransactionObject.TransactionID)
            Dim _BLFinancialStatementLog As New BLFinancialStatementLog
            _BLFinancialStatementLog.Add(FinancialStatementLog)

            For Each Historytri In TransactionItemHistoryList
                For Each tri In TransactionItemList
                    If Historytri.TransactionItemID = tri.TransactionItemID Then

                        Dim triBeforRestore As Trn_TransactionItem = tri
                        triBeforRestore.TransactionItemID = Historytri.TransactionItemID
                        triBeforRestore.TransactionID = Historytri.TransactionID
                        triBeforRestore.RowNo = Historytri.RowNo
                        triBeforRestore.ProductID = Historytri.ProductID
                        triBeforRestore.ProductDescription = Historytri.ProductDescription
                        triBeforRestore.Amount = Historytri.Amount
                        triBeforRestore.UnitID = Historytri.UnitID
                        triBeforRestore.Discount = Historytri.Discount
                        triBeforRestore.VATPercent = Historytri.VATPercent
                        triBeforRestore.AccountID = Historytri.AccountID
                        triBeforRestore.FreeAccountID1 = Historytri.FreeAccountID1
                        triBeforRestore.FreeAccountID2 = Historytri.FreeAccountID2
                        triBeforRestore.Priced = Historytri.Priced
                        triBeforRestore.PriceRefTransactionItemID = Historytri.PriceRefTransactionItemID
                        triBeforRestore.CountRefTransactionItemID = Historytri.CountRefTransactionItemID
                        triBeforRestore.RemainedPricingAmount = Historytri.RemainedPricingAmount
                        triBeforRestore.CreatorID = Historytri.CreatorID
                        triBeforRestore.CreateDate = Historytri.CreateDate
                        triBeforRestore.LastUpdatorID = Historytri.LastUpdatorID
                        triBeforRestore.LastUpdateDate = Historytri.LastUpdateDate
                        triBeforRestore.WarehouseUnitPrice = Historytri.WarehouseUnitPrice
                        triBeforRestore.SaleUnitPrice = Historytri.SaleUnitPrice
                        triBeforRestore.ImpactFactors = Historytri.ImpactFactors
                        triBeforRestore.TransactionItemControlCodeID = Historytri.TransactionItemControlCodeID
                        triBeforRestore.SaleUnitPriceCurrency = Historytri.SaleUnitPriceCurrency
                        triBeforRestore.FreeAccountID = Historytri.FreeAccountID
                        triBeforRestore.TotalWarehousePrice = Historytri.TotalWarehousePrice
                        triBeforRestore.TotalSalePrice = Historytri.TotalSalePrice
                        triBeforRestore.PricingDate = Historytri.PricingDate
                        triBeforRestore.ComplicationsPercent = Historytri.ComplicationsPercent
                        triBeforRestore.TaxesPercent = Historytri.TaxesPercent
                        triBeforRestore.Comment = Historytri.Comment
                        triBeforRestore.SystemicPriced = Historytri.SystemicPriced
                        triBeforRestore.UnitConversionRate = Historytri.UnitConversionRate
                        _BLItem.Update(triBeforRestore)

                    End If

                Next

            Next

            Dim trnBeforRestore As Trn_Transaction = TransactionObject

            trnBeforRestore.FormCode = HistoryTrn.FormCode
            trnBeforRestore.FinancialYearID = HistoryTrn.FinancialYearID
            trnBeforRestore.Number = HistoryTrn.Number
            trnBeforRestore.TransactionDesc = HistoryTrn.TransactionDesc
            trnBeforRestore.TmpNumber = HistoryTrn.TmpNumber
            trnBeforRestore.TransactionDate = HistoryTrn.TransactionDate
            trnBeforRestore.PriceRefTransactionID = HistoryTrn.PriceRefTransactionID
            trnBeforRestore.CountRefTransactionID = HistoryTrn.CountRefTransactionID
            trnBeforRestore.SaleVoucherID = HistoryTrn.SaleVoucherID
            trnBeforRestore.SrcPartnerID = HistoryTrn.SrcPartnerID
            trnBeforRestore.SrcWarehouseID = HistoryTrn.SrcWarehouseID
            trnBeforRestore.SrcCostCenterID = HistoryTrn.SrcCostCenterID
            trnBeforRestore.DstPartnerID = HistoryTrn.DstPartnerID
            trnBeforRestore.DstWarehouseID = HistoryTrn.DstWarehouseID
            trnBeforRestore.DstCostCenterID = HistoryTrn.DstCostCenterID
            trnBeforRestore.PartnerName = HistoryTrn.PartnerName
            trnBeforRestore.PartnerAddress = HistoryTrn.PartnerAddress
            trnBeforRestore.PartnerPhone = HistoryTrn.PartnerPhone
            trnBeforRestore.PartnerMobile = HistoryTrn.PartnerMobile
            trnBeforRestore.MiddleManID = HistoryTrn.MiddleManID
            trnBeforRestore.ValidityDays = HistoryTrn.ValidityDays
            trnBeforRestore.CreatorID = HistoryTrn.CreatorID
            trnBeforRestore.CreateDate = HistoryTrn.CreateDate
            trnBeforRestore.LastUpdatorID = HistoryTrn.LastUpdatorID
            trnBeforRestore.LastUpdateDate = HistoryTrn.LastUpdateDate
            trnBeforRestore.Registered = HistoryTrn.Registered
            trnBeforRestore.RegistererID = HistoryTrn.RegistererID
            trnBeforRestore.RegisterDate = HistoryTrn.RegisterDate
            trnBeforRestore.Confirmed = HistoryTrn.Confirmed
            trnBeforRestore.ConfirmerID = HistoryTrn.ConfirmerID
            trnBeforRestore.ConfirmDate = HistoryTrn.ConfirmDate
            trnBeforRestore.Canceled = HistoryTrn.Canceled
            trnBeforRestore.CancelerID = HistoryTrn.CancelerID
            trnBeforRestore.CancelDate = HistoryTrn.CancelDate
            trnBeforRestore.IsOfficial = HistoryTrn.IsOfficial
            trnBeforRestore.WarehouseVoucherID = HistoryTrn.WarehouseVoucherID
            trnBeforRestore.HasCurrency = HistoryTrn.HasCurrency
            trnBeforRestore.CurrencyID = HistoryTrn.CurrencyID
            trnBeforRestore.ExchangeRate = HistoryTrn.ExchangeRate
            trnBeforRestore.PrintDate = HistoryTrn.PrintDate
            trnBeforRestore.RegisteredPrice = HistoryTrn.RegisteredPrice
            trnBeforRestore.RegistererPriceID = HistoryTrn.RegistererPriceID
            trnBeforRestore.RegisterPriceDate = HistoryTrn.RegisterPriceDate
            trnBeforRestore.FreeAccountID = HistoryTrn.FreeAccountID
            trnBeforRestore.TransactionFixNo = HistoryTrn.TransactionFixNo
            trnBeforRestore.CreatedBySystem = HistoryTrn.CreatedBySystem
            trnBeforRestore.CreatorSystemID = HistoryTrn.CreatorSystemID
            trnBeforRestore.OriginEntityID = HistoryTrn.OriginEntityID
            trnBeforRestore.OriginEntityRecordID = HistoryTrn.OriginEntityRecordID
            trnBeforRestore.ConfirmedLogistic = HistoryTrn.ConfirmedLogistic
            trnBeforRestore.ContractNumber = HistoryTrn.ContractNumber
            trnBeforRestore.TransactionDesc1 = HistoryTrn.TransactionDesc1
            trnBeforRestore.FinancialStatmentState = HistoryTrn.FinancialStatmentState
            trnBeforRestore.FinancialStatmentGenNo = HistoryTrn.FinancialStatmentGenNo
            trnBeforRestore.FinancialStatmentReferenceNo = HistoryTrn.FinancialStatmentReferenceNo
            trnBeforRestore.FinancialStatmentTaxID = HistoryTrn.FinancialStatmentTaxID
            trnBeforRestore.FinancialStatmentUID = HistoryTrn.FinancialStatmentUID
            trnBeforRestore.SalemanID = HistoryTrn.SalemanID
            trnBeforRestore.RejectOrRepairStatus = Nothing
            _BL.Update(trnBeforRestore)

            Dim HistoryTrnList = BLTransactionHistory.GetTransactionHistoryList(TransactionObject.TransactionID)

            For Each trn In HistoryTrnList
                Dim _BLTransactionHistory As New BLTransactionHistory
                Dim _BLTransactionItemHistory As New BLTransactionItemHistory
                Dim AllTransactionItemHistoryList = BLTransactionItemHistory.GetAllTransactionItemHistoryList(trn.TransactionID)
                For Each tri In AllTransactionItemHistoryList
                    _BLTransactionItemHistory.Delete(tri)
                Next
                _BLTransactionHistory.Delete(trn)
            Next

            TrnList = BLTransaction.TransactionListForFinancialStatmentRejected(slcRangeType.SelectedRangeType, cmbSystem.EditValue, cmbTransactionForm.EditValue, slcRangeType.FinancialYearID,
                                                    If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetStartNumber > 0, slcRangeType.GetStartNumber, Nothing),
                                                    If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetEndNumber > 0, slcRangeType.GetEndNumber, Nothing),
                                                    If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetStartDate, Nothing),
                                                    If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetEndDate, Nothing), True)

            For Each i In TrnList
                i.FinancialStatmentStateTitle = "رد شده"
                If i.RejectOrRepairStatus Is Nothing Then
                    i.RequestKind = "عادی"
                End If
                If i.RejectOrRepairStatus = 1 Then
                    i.RequestKind = "جهت ابطال"
                End If
                If i.RejectOrRepairStatus = 2 Then
                    i.RequestKind = "جهت اصلاح"
                End If
            Next

            gcTransaction.DataSource = TrnList

            MessageBox.Show("لغو ابطال با موفقیت انجام شد")
        Catch ex As Exception
            CustomException.ShowDialogue(ex)

        End Try
    End Sub
End Class
