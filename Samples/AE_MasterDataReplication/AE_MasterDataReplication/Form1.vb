Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        Try
            Dim oCurrencyCodes As SAPbobsCOM.Currencies
            oCurrencyCodes = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCurrencyCodes)
            oCurrencyCodes.Code = "IND"
            oCurrencyCodes.Name = "India"
            oCurrencyCodes.InternationalDescription = "Indian Rupees"
            oCurrencyCodes.DocumentsCode = "IND"
            Dim imjs As Integer = oCurrencyCodes.Add()


            'While Not oCurrencyCodes.Browser.EoF
            '    MsgBox(oCurrencyCodes.Code)
            '    oCurrencyCodes.Browser.MoveNext()
            'End While


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim sErrDesc As String = String.Empty
        If GetSystemIntializeInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If ConnectToTargetCompany(p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim oSBObob As SAPbobsCOM.SBObob
        oSBObob = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oRecordSet = oSBObob.GetLocalCurrency()

        oRecordSet = oSBObob.GetSystemCurrency()
        If result.Currency.ToString() = USD Then

            oSBObob.SetCurrencyRate(USD, DateTime.Now, Convert.ToDouble(result.Rate), True)
        End If

        If result.Currency.ToString() = AUD Then

            oSBObob.SetCurrencyRate(AUD, DateTime.Now, Convert.ToDouble(result.Rate), True)
        End If

        If result.Currency.ToString() = GBP Then

            oSBObob.SetCurrencyRate(GBP, DateTime.Now, Convert.ToDouble(result.Rate), True)
        End If

        If result.Currency.ToString() = [TRY] Then

            oSBObob.SetCurrencyRate([TRY], DateTime.Now, Convert.ToDouble(result.Rate), True)
        End If

    End Sub
End Class
