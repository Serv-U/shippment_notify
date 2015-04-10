Option Strict On

Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Net.Mail
Imports System.Xml

Module Shipment_Notify
    Dim mageService As New MagentoService.Mage_Api_Model_Server_V2_HandlerPortTypeClient
    Dim sessionId As String
    Dim solutionsConnectionString As String = "Data Source=sqlserver;Initial Catalog=SERV-U;User ID=sa;Password=!get2it"
    Dim solutionsConnection As New SqlClient.SqlConnection
    'Dim solutionsSQL As String = "SELECT a.[NO], a.[Date], a.AUX_ORDER_NO, b.Product_#, b.QTY, b.QTY_SHIPPED, b.SIZE, b.COLOR, b.FIRST_SHIPPED as liFirstShipped, c.FIRST_SHIPPED as shiptoFirstShipped, c.PO_NUMBER, d.tracer# " & _
    '"FROM [ORDERS] AS a INNER JOIN [LITEM] AS b ON a.no = b.order_no INNER JOIN [SHIPTO] AS c ON a.NO = c.order_no " & _
    '"INNER JOIN [BOXDEF] AS d ON a.no = d.order_no where a.aux_order_no = "
    Dim solutionsSQL As String = "SELECT a.[NO], a.[Date], a.AUX_ORDER_NO, b.Product_#, b.QTY, b.QTY_SHIPPED, b.SIZE, b.COLOR, b.FIRST_SHIPPED as liFirstShipped, c.FIRST_SHIPPED as shiptoFirstShipped, c.PO_NUMBER " & _
                        "FROM [ORDERS] AS a INNER JOIN [LITEM] AS b ON a.no = b.order_no INNER JOIN [SHIPTO] AS c ON a.NO = c.order_no where b.QTY_SHIPPED > 0 AND c.PO_NUMBER = "
    Dim trackingNumbers As String = "Select a.aux_order_no, d.tracer# FROm [ORDERS] AS a INNER JOIN [SHIPTO] AS c ON a.NO = c.order_no INNER JOIN [BOXDEF] AS d on a.no = d.order_no where c.PO_NUMBER = "
    Dim solutionOrderDetails As New Collections.Generic.Dictionary(Of String, orderResultSet)
    Dim orderSkus As New Collections.Generic.Dictionary(Of String, shipmentSkuQty)
    Dim prepareShipment As New Collections.Generic.Dictionary(Of String, shipmentSkuQty)
    Dim shippingTracking As New Collections.Generic.Dictionary(Of String, String)
    Dim mailTo, SMTPAddress, code, title As String
    Dim trackingNumber As Int16
    Dim logString As String = ""

    Sub Main()
        readConfig()
        sessionId = GetSessionId()

        Dim fltr As MagentoService.filters
        Dim shipmentFilter As MagentoService.filters
        Dim orderFilters As MagentoService.complexFilter()
        Dim ae As New MagentoService.associativeEntity
        Dim shipFilters As MagentoService.complexFilter()
        Dim now As DateTime = DateTime.Now()
        Dim timeSpan As New TimeSpan(30, 0, 0, 0)
        Dim maximumOrderDate As DateTime = now.Subtract(timeSpan)
        Dim dateString As String
        Dim logID As String = "no"
        dateString = CStr(maximumOrderDate.ToString("yyyy-MM-dd HH:mm:ss"))

        Try
            fltr = New MagentoService.filters()
            shipmentFilter = New MagentoService.filters()

            shipFilters = New MagentoService.complexFilter(0) {}

            orderFilters = New MagentoService.complexFilter(1) {}
            ae = New MagentoService.associativeEntity()

            orderFilters(0) = New MagentoService.complexFilter()
            orderFilters(1) = New MagentoService.complexFilter()
            orderFilters(0).key = "status"
            ae.key = "status"
            ae.value = "fulfilling"
            orderFilters(0).value = ae
            'orderFilters(1).key = "created_at"
            'ae.key = "from"
            'ae.value = dateString
            'orderFilters(1).value = ae
            fltr.complex_filter = orderFilters

            Dim orders As MagentoService.salesOrderListEntity()
            Dim shipments As MagentoService.salesOrderShipmentEntity()
            Dim orderInfo As MagentoService.salesOrderEntity

            orders = mageService.salesOrderList(sessionId, fltr)

            For Each soe In orders
                'logString += "PRE: " + soe.increment_id + Environment.NewLine

                If solutionsResult(CType(soe.increment_id, Integer)) Then
                    orderSkus = New Collections.Generic.Dictionary(Of String, shipmentSkuQty)
                    shippingTracking = New Collections.Generic.Dictionary(Of String, String)
                    orderInfo = mageService.salesOrderInfo(sessionId, soe.increment_id)

                    For Each soie In orderInfo.items
                        Dim orderSkuQtyId = New shipmentSkuQty
                        Dim skus As String = ""
                        skus = formatSku(soie.sku)
                        If orderSkus.ContainsKey(skus) And soie.product_type = "simple" Then
                            'nothing
                        Else
                            orderSkuQtyId.skuId = CInt(soie.item_id)
                            orderSkuQtyId.qty = 0
                            orderSkuQtyId.ordered = CInt(soie.qty_ordered)
                            'logString += CStr(orderSkuQtyId.ordered)
                            orderSkus(skus) = orderSkuQtyId
                        End If

                    Next

                    ae = New MagentoService.associativeEntity()

                    shipFilters(0) = New MagentoService.complexFilter()
                    shipFilters(0).key = "order_id"
                    ae.key = "order_id"
                    ae.value = soe.order_id
                    shipFilters(0).value = ae
                    shipmentFilter.complex_filter = shipFilters

                    Dim hasTracking As Boolean = getTrackingNumbers(CType(soe.increment_id, Integer))

                    shipments = mageService.salesOrderShipmentList(sessionId, shipmentFilter)

                    For shipmentCount As Integer = shipments.GetLowerBound(0) To shipments.GetUpperBound(0)
                        Dim shipment As MagentoService.salesOrderShipmentEntity = mageService.salesOrderShipmentInfo(sessionId, shipments(shipmentCount).increment_id)
                        'logString += "Order ID: " & shipment.order_id + Environment.NewLine
                        'logString += "QTY: " & shipment.items(0).qty + Environment.NewLine
                        'logString += "SKU: " & shipment.items(0).sku + Environment.NewLine

                        For itemCount As Integer = shipment.items.GetLowerBound(0) To shipment.items.GetUpperBound(0)

                            Dim skuf As String = ""
                            Dim orderSkuQtyId = New shipmentSkuQty

                            skuf = formatSku(shipment.items(itemCount).sku)
                            'logString += skuf

                            If orderSkus.ContainsKey(skuf) Then
                                orderSkuQtyId.qty = CInt(shipment.items(itemCount).qty) + orderSkus(skuf).qty
                                orderSkuQtyId.skuId = orderSkus(skuf).skuId
                                orderSkuQtyId.ordered = orderSkus(skuf).ordered
                            Else
                                orderSkuQtyId.qty = CInt(shipment.items(itemCount).qty)
                                orderSkuQtyId.skuId = orderSkus(skuf).skuId
                                orderSkuQtyId.ordered = orderSkus(skuf).ordered
                            End If

                            orderSkus(skuf) = orderSkuQtyId

                        Next

                        'logString += soe.increment_id + Environment.NewLine

                        'logString += "HAS TRACKING: " & hasTracking & Environment.NewLine

                        If hasTracking Then
                            For counter As Integer = shipment.tracks.GetLowerBound(0) To shipment.tracks.GetUpperBound(0)
                                shippingTracking.Remove(RTrim(shipment.tracks(counter).number))
                            Next
                        End If

                    Next

                    For Each pair In orderSkus
                        logString += "Pair: " & pair.Key + Environment.NewLine
                        If solutionOrderDetails.ContainsKey(pair.Key) Then
                            logString += solutionOrderDetails(pair.Key).qtyShipped & " " & pair.Value.qty & " " & pair.Value.ordered
                            If CDbl(solutionOrderDetails(pair.Key).qtyShipped) > CInt(pair.Value.qty) And ((CDbl(solutionOrderDetails(pair.Key).qtyShipped) - pair.Value.qty) <= CInt(pair.Value.ordered)) Then
                                Dim prepareShipmentSkuQtyId = New shipmentSkuQty
                                prepareShipmentSkuQtyId.skuId = pair.Value.skuId
                                prepareShipmentSkuQtyId.qty = CInt(CDbl(solutionOrderDetails(pair.Key).qtyShipped) - pair.Value.qty)
                                prepareShipment(pair.Key) = prepareShipmentSkuQtyId
                                logString += CStr(prepareShipmentSkuQtyId.qty)
                            End If
                        End If
                    Next

                    If prepareShipment.Count > 0 Then
                        logString += "Prepared:"
                        Dim preparedArray(100) As MagentoService.orderItemIdQty
                        Dim preparedLength As Integer = 0
                        Dim shipmentId As String
                        For Each pair In prepareShipment
                            logString += pair.Key & " with " & pair.Value.qty & " shipping and an ID of: " & pair.Value.skuId
                            preparedArray(preparedLength) = New MagentoService.orderItemIdQty
                            preparedArray(preparedLength).order_item_id = pair.Value.skuId
                            preparedArray(preparedLength).qty = pair.Value.qty
                            preparedLength += 1
                        Next

                        logID = soe.increment_id
                        shipmentId = mageService.salesOrderShipmentCreate(sessionId, soe.increment_id, preparedArray, "", CInt(False), CInt(False))

                        logString += "Shipping: " & shippingTracking.Count & " " & preparedArray.ToString + Environment.NewLine


                        If shippingTracking.Count > 0 Then
                            For Each trackingNum In shippingTracking
                                'logString += "TRACKING LOOP: " & trackingNum.Value & Environment.NewLine
                                Dim upsMatch As Boolean = UCase(trackingNum.Value) Like "1Z*"
                                If upsMatch Then
                                    mageService.salesOrderShipmentAddTrack(sessionId, shipmentId, "ups", "", trackingNum.Value)
                                    'logString += "UPS" & Environment.NewLine
                                Else
                                    mageService.salesOrderShipmentAddTrack(sessionId, shipmentId, "custom", "Conway Freight Tracking", trackingNum.Value)
                                    'logString += "CONWAY" & Environment.NewLine
                                End If
                            Next

                        End If

                    End If

                End If
                prepareShipment = New Collections.Generic.Dictionary(Of String, shipmentSkuQty)

            Next

        Catch ex As Exception
            LogError("Exception Occurred: " & " " & logID & " " & ex.ToString)
        End Try

        'LogError(logString)
    End Sub

    Private Function formatSku(ByVal sku As String) As String
        Dim formatedSku As String = ""
        Dim splitSku() As String = Regex.Split(sku, "([=^#*])")
        Dim previousPart As String = "#FIRSTPART#"

        For Each skuPart As String In splitSku
            If previousPart = "#FIRSTPART#" Then
                formatedSku = skuPart
            ElseIf previousPart = Chr(35) Then
                formatedSku &= skuPart
            End If
            previousPart = skuPart
        Next
        logString += "FORMATED: " & formatedSku + Environment.NewLine
        Return RTrim(formatedSku)

    End Function

    Private Function GetSessionId() As String

        If sessionId <> "" Then
            Return sessionId
        End If

        Try
            sessionId = mageService.login("dmillerapi", "5ERaL253201S")
        Catch e As Exception
            LogError("Session ID: " & e.ToString())
        End Try

        Return sessionId

    End Function

    Private Function solutionsResult(ByVal OrderNo As Integer) As Boolean
        Dim returnValue As Boolean = False

        Try
            solutionsConnection.ConnectionString = solutionsConnectionString
            solutionsConnection.Open()

            Dim sql As SqlCommand = New SqlCommand(solutionsSQL & "'" & OrderNo & "'", solutionsConnection)
            logString += solutionsSQL & "'" & OrderNo & "'" + Environment.NewLine

            Try
                Dim sqlReader As SqlDataReader = sql.ExecuteReader()

                If sqlReader.HasRows() Then
                    returnValue = True

                    While sqlReader.Read()
                        Dim orderResult = New orderResultSet
                        orderResult.orderDate = IfNull(sqlReader, "Date", "")
                        orderResult.magentoOrderNumber = IfNull(sqlReader, "AUX_ORDER_NO", "")
                        orderResult.sku = RTrim(IfNull(sqlReader, "Product_#", "")) & RTrim(IfNull(sqlReader, "SIZE", "")) & RTrim(IfNull(sqlReader, "COLOR", ""))
                        orderResult.qtyOrdered = IfNull(sqlReader, "QTY", "")
                        orderResult.qtyShipped = IfNull(sqlReader, "QTY_SHIPPED", "")
                        orderResult.lineItemShipDate = IfNull(sqlReader, "liFirstShipped", "")
                        orderResult.shipToShipDate = IfNull(sqlReader, "shiptoFirstShipped", "")
                        orderResult.poNumber = IfNull(sqlReader, "PO_NUMBER", "")
                        orderResult.solutionOrderNumber = IfNull(sqlReader, "no", "")
                        logString += orderResult.sku + Environment.NewLine
                        solutionOrderDetails(orderResult.sku) = orderResult
                    End While
                Else
                    logString += "NO MATCHING ROWS FOR " & OrderNo & " " + Environment.NewLine
                End If


                sqlReader.Close()
            Catch e As Exception
                LogError("Solution Read Error: " & e.ToString())
            End Try

        Catch ex As Exception
            LogError("Solution Connection Error: " & ex.ToString())
        Finally
            solutionsConnection.Close()
        End Try

        Return returnValue

    End Function

    Private Function getTrackingNumbers(ByVal orderNo As Integer) As Boolean
        Dim returnValue As Boolean = False

        logString += trackingNumbers & "'" & orderNo & "'" + Environment.NewLine

        Try
            solutionsConnection.ConnectionString = solutionsConnectionString
            solutionsConnection.Open()

            Dim sql As SqlCommand = New SqlCommand(trackingNumbers & "'" & orderNo & "'", solutionsConnection)

            Try
                Dim sqlReader As SqlDataReader = sql.ExecuteReader()

                If sqlReader.HasRows() Then
                    returnValue = True
                    logString += "Has rows" + Environment.NewLine

                    While sqlReader.Read()
                        logString += (RTrim(IfNull(sqlReader, "tracer#", ""))) + Environment.NewLine
                        shippingTracking(RTrim(IfNull(sqlReader, "tracer#", ""))) = RTrim(IfNull(sqlReader, "tracer#", ""))
                    End While
                End If

                sqlReader.Close()
            Catch e As Exception
                LogError("Tracking Number Read Error: " & e.ToString())
            End Try

        Catch ex As Exception
            LogError("Tracking Number Connection Error: " & ex.ToString())
        Finally
            solutionsConnection.Close()
        End Try

        Return returnValue
    End Function
    Sub readConfig()
        SMTPAddress = "192.168.1.50"
        mailTo = "dustinmiller@servu-online.com"

        Try
            Dim doc As New System.Xml.XmlDocument
            doc.Load(My.Application.Info.DirectoryPath & "\config.xml")
            Dim list = doc.GetElementsByTagName("name")

            If (doc.GetElementsByTagName("SMTP")(0).InnerText <> "") Then
                SMTPAddress = doc.GetElementsByTagName("SMTP")(0).InnerText
            End If

            If (doc.GetElementsByTagName("MailTo")(0).InnerText <> "") Then
                mailTo = doc.GetElementsByTagName("MailTo")(0).InnerText
            End If

            If (doc.GetElementsByTagName("SolutionsConnection")(0).InnerText <> "") Then
                solutionsConnectionString = doc.GetElementsByTagName("SolutionsConnection")(0).InnerText
            End If
        Catch e As Exception
            LogError("Error in xml: " & e.ToString())
        End Try

    End Sub
    Private Sub LogError(ByVal e As String)
        'Utility function to log errors
        Dim fileName As String = My.Application.Info.DirectoryPath & "\logs\log.txt"

        If File.Exists(fileName) Then
            Using fileWriter As StreamWriter = New StreamWriter(fileName, True)
                fileWriter.Write(e & vbTab)
            End Using
        Else
            Using fileWriter As StreamWriter = New StreamWriter(fileName)
                fileWriter.Write(e & vbTab)
            End Using
        End If

        sendException("An Error Occured With The Shipment Notify", "Please check the log files at " & fileName)

    End Sub

    Public Function IfNull(Of T)(ByVal dr As SqlDataReader, ByVal fieldName As String, ByVal _default As T) As T
        logString += dr(fieldName).ToString + Environment.NewLine
        If IsDBNull(dr(fieldName)) Then
            Return _default
        Else
            Return CType(dr(fieldName), T)
        End If

    End Function

    Private Structure orderResultSet
        Public orderDate As String
        Public magentoOrderNumber As String
        Public sku As String
        Public qtyOrdered As String
        Public qtyShipped As String
        Public lineItemShipDate As String
        Public shipToShipDate As String
        Public poNumber As String
        Public trackNumber As String
        Public solutionOrderNumber As String
    End Structure

    Private Structure shipmentSkuQty
        Public qty As Integer
        Public skuId As Integer
        Public ordered As Integer
    End Structure

    Sub sendException(ByVal subject As String, ByVal body As String)
        Dim mail As New MailMessage()
        Dim smtp As New SmtpClient(SMTPAddress)

        mail.From = New MailAddress(mailTo)
        mail.To.Add(mailTo)

        mail.Subject = subject
        mail.Body = body

        smtp.Send(mail)
    End Sub

End Module