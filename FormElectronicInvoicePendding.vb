Imports System.ComponentModel
Imports PCommonTools.ExceptionHandler
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports System.Threading
Public Class FormElectronicInvoicePendding

    Private th As Thread
    Public Property TrnList As List(Of Trn_Transaction) = Nothing
    Public Property TriList As List(Of Trn_TransactionItem) = Nothing
    Public Property PartnerList As BindingList(Of Com_Partner)



    Private Sub btnSaveCashier_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnSaveCashier.ItemClick
        Try

            If cmbTransactionForm.EditValue = 0 Then
                Throw New CustomException("نوع فرم را مشخص نمایید.")
            End If
            Dim ProductList = BLProductDef.GetAllProductList
            TrnList = BLTransaction.TransactionListForFinancialStatmentPendding(slcRangeType.SelectedRangeType, cmbSystem.EditValue, cmbTransactionForm.EditValue, Context.CurrentYear.FinancialYearID,
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetStartNumber > 0, slcRangeType.GetStartNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetEndNumber > 0, slcRangeType.GetEndNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetStartDate, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetEndDate, Nothing), True)
            If TrnList.Count = 0 Then
                Throw New CustomException("اطلاعاتی جهت ارسال وجود ندارد.")
            End If

            th = New Thread(New ThreadStart(AddressOf GetData))
            th.Start()

            'TrnList = Nothing
            TrnList = BLTransaction.TransactionListForFinancialStatmentPendding(slcRangeType.SelectedRangeType, cmbSystem.EditValue, cmbTransactionForm.EditValue, Context.CurrentYear.FinancialYearID,
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetStartNumber > 0, slcRangeType.GetStartNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetEndNumber > 0, slcRangeType.GetEndNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetStartDate, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetEndDate, Nothing), True)
            For Each i In TrnList
                Select Case i.FinancialStatmentState
                    Case 0
                        i.FinancialStatmentStateTitle = "ارسال نشده"
                    Case 1
                        i.FinancialStatmentStateTitle = "ارسال شده"
                    Case 2
                        i.FinancialStatmentStateTitle = "در انتظار پاسخ"
                    Case 4
                        i.FinancialStatmentStateTitle = "تأیید شده"
                    Case 5
                        i.FinancialStatmentStateTitle = "رد شده"
                End Select
            Next
            PCommonTools.InformationMessageBox("ارسال انجام شد.")
            If cmbTransactionForm.EditValue = 0 Then
                Throw New CustomException("نوع فرم را مشخص نمایید.")
            End If

            TrnList = BLTransaction.TransactionListForFinancialStatmentPendding(slcRangeType.SelectedRangeType, cmbSystem.EditValue, cmbTransactionForm.EditValue, Context.CurrentYear.FinancialYearID,
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetStartNumber > 0, slcRangeType.GetStartNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetEndNumber > 0, slcRangeType.GetEndNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetStartDate, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetEndDate, Nothing), True)
            If TrnList.Count = 0 Then
                gcTransaction.DataSource = Nothing
                Throw New CustomException("در بازه انتخاب شده اطلاعاتی جهت نمایش وجود ندارد.")

            End If

            For Each i In TrnList
                Select Case i.FinancialStatmentState
                    Case 0
                        i.FinancialStatmentStateTitle = "ارسال نشده"
                    Case 1
                        i.FinancialStatmentStateTitle = "ارسال شده"
                    Case 2
                        i.FinancialStatmentStateTitle = "در انتظار پاسخ"
                    Case 4
                        i.FinancialStatmentStateTitle = "تأیید شده"
                    Case 5
                        i.FinancialStatmentStateTitle = "رد شده"
                    Case 6
                        i.FinancialStatmentStateTitle = "ابطال شده در سامانه"
                End Select
            Next
            gcTransaction.DataSource = TrnList

        Catch ex As Exception
            CustomException.ShowDialogue(ex)
        End Try
    End Sub
    Private Sub btnRefresh_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles btnRefresh.ItemClick
        Try

            If cmbTransactionForm.EditValue = 0 Then
                Throw New CustomException("نوع فرم را مشخص نمایید.")
            End If
            Dim ProductList = BLProductDef.GetAllProductList
            TrnList = BLTransaction.TransactionListForFinancialStatmentPendding(slcRangeType.SelectedRangeType, cmbSystem.EditValue, cmbTransactionForm.EditValue, Context.CurrentYear.FinancialYearID,
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetStartNumber > 0, slcRangeType.GetStartNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType = RangeTypeEnum.NumberRange AndAlso slcRangeType.GetEndNumber > 0, slcRangeType.GetEndNumber, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetStartDate, Nothing),
                                                     If(slcRangeType.SelectedRangeType <> RangeTypeEnum.NumberRange, slcRangeType.GetEndDate, Nothing), True)
            If TrnList.Count = 0 Then
                gcTransaction.DataSource = Nothing
                Throw New CustomException("اطلاعاتی جهت نمایش وجود ندارد.")

            End If

            For Each i In TrnList
                Select Case i.FinancialStatmentState
                    Case 0
                        i.FinancialStatmentStateTitle = "ارسال نشده"
                    Case 1
                        i.FinancialStatmentStateTitle = "ارسال شده"
                    Case 2
                        i.FinancialStatmentStateTitle = "در انتظار پاسخ"
                    Case 4
                        i.FinancialStatmentStateTitle = "تأیید شده"
                    Case 5
                        i.FinancialStatmentStateTitle = "رد شده"
                    Case 6
                        i.FinancialStatmentStateTitle = "ابطال شده در سامانه"
                End Select
            Next
            gcTransaction.DataSource = TrnList
        Catch ex As Exception
            CustomException.ShowDialogue(ex)
        End Try
    End Sub
    Public Sub GetData()
        Dim CurrentRow = gvTransaction.GetFocusedRow
        If CurrentRow Is Nothing Then Return
        Dim lst As New List(Of Trn_Transaction)
        lst.Add(CurrentRow)
        Ganjineh.Net.TaxLIbrary.ElectronicInvoice.GetPenddings(lst, BLAttachment.DownloadAttachedFile(BLAttachment.GetAttachmentList(0, 0).FirstOrDefault.AttachmentID))
    End Sub

    Private Sub FormFinancialStatmentsChecking_Load(sender As Object, e As EventArgs) Handles Me.Load
        slcRangeType.Initialize()
        PartnerList = BLPartner.GetPartnerList
        Dim source As List(Of Trn_TransactionForm) = GlobalParam.TransactionFormsList.Where(Function(a) a.FormCode.ToString.Remove(3) = 103 Or a.FormCode.ToString.Remove(3) = 102).ToList
        cmbTransactionForm.Properties.DataSource = source
        cmbTransactionForm.EditValue = 0

        Dim Dic As New Dictionary(Of Integer, String)
        Dic.Add(30, "خرید و فروش")
        cmbSystem.Properties.DataSource = Dic
        cmbSystem.EditValue = 30
        gvTransaction.OptionsDetail.EnableMasterViewMode = False
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

    Private Sub BarButtonItem1_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem1.ItemClick
        If th Is Nothing Then Return
        CloseMe()
    End Sub

    Private Sub FormFinancialStatmentsPendding_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If th Is Nothing Then Return
        th.Abort()

    End Sub

    Private Sub cmbTransactionForm_EditValueChanged(sender As Object, e As EventArgs) Handles cmbTransactionForm.EditValueChanged
        gcTransaction.DataSource = Nothing
    End Sub

    Private Sub slcRangeType_ValuesChanged(sender As Object, e As ValueChangedEventArgs) Handles slcRangeType.ValuesChanged
        gcTransaction.DataSource = Nothing
    End Sub
End Class