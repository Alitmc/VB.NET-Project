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
                Dim NumericFormat As String = "#,##0.##;(#,##0.##)
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
