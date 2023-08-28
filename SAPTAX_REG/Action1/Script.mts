SystemUtil.CloseProcessByName("EXCEL.EXE")
Set SAPWindowObject = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
Set SAPWindowObject1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
strDataExcelPath = "C:\Users\dchinnasam\OneDrive\Documents\01. Offical\EHP8\TaxValidation\TAX-EQ3-New.xlsx"
strCapImages= ""
'Create Excel object 
Call CreateExcelObject(strDataExcelPath,"Inputs",intExcelRowCount,intExcelColumnCount,objExcelSheet,objExcelWorkbook)
'Read all Values from Login Sheet
Call ReadAllValuesFromInputExcel(objExcelSheet)
Set objColumnNumberdictionary = CreateObject("Scripting.Dictionary")
For intLoop = 1 To intExcelColumnCount
	strColumnNameToAdd  = objExcelSheet.Cells(1,intLoop).Value
	objColumnNumberdictionary.Add strColumnNameToAdd,intLoop
Next
For intTestCaseCount = 1 To UBOUND(arrProductCategory)
	If arrExecutionFlag(intTestCaseCount) = "Y" Then
		SAPWindowObject.SAPGuiOKCode("guicomponenttype:=35","name:=okcd").Set "/nVA41"
		SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
		SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBAK-AUART").set "ZSUB"
		SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBAK-VKORG").set arrSellingEntity(intTestCaseCount)
		SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBAK-SPART").set "00"   
		SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBAK-VTWEG").set "00"
		Call CaptureAUTScreenShot
		SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
		SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=KUAGV-KUNNR").set arrSoldToCustomerID(intTestCaseCount)
		SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=KUWEV-KUNNR").set arrShipToCustomerID(intTestCaseCount)
		SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
		Call CaptureAUTScreenShot
		If SAPWindowObject1.Exist Then
			SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
		End If
		strStatusMessage = SAPWindowObject.SAPGuiStatusBar("guicomponenttype:=103","name:=sbar").GetROProperty("messagetype")
		strMessageText = SAPWindowObject.SAPGuiStatusBar("guicomponenttype:=103","name:=sbar").GetROProperty("text")
		arrBPID = Split(strMessageText," ")
		If Instr (strMessageText,"block") Then
			intBPID = arrBPID(1)
			SAPWindowObject.SAPGuiOKCode("guicomponenttype:=35","name:=okcd").Set "/nBP"
			Call CaptureAUTScreenShot
			SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
			SAPWindowObject.SAPGuiButton("guicomponenttype:=40","tooltip:=Open.*").Click()
			SAPWindowObject1.SAPGuiEdit("guicomponenttype:=32","name:=BUS_JOEL_MAIN-OPEN_NUMBER").set intBPID
			SAPWindowObject1.SAPGuiButton("guicomponenttype:=40","tooltip:=Enter.*").Click()
			If SAPWindowObject1.Exist Then
				SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
			End If
			SAPWindowObject.SAPGuiComboBox("guicomponenttype:=34","name:=BUS_JOEL_MAIN-PARTNER_ROLE").SelectKey "ISM000"
			SAPWindowObject.SAPGuiButton("guicomponenttype:=40","tooltip:=Sales and Distribution   (Ctrl+F3)").Click()
			'Status Tab
			SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=GS_SCREEN_1100_TABSTRIP").Select "Status"
			SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=PUSH_SA").Click()   
			SAPWindowObject1.SAPGuiTable("guicomponenttype:=80","name:=SAPLCVI_FS_UI_CUSTOMER_SALESTCTRL_SALES_AREA").SelectRow(1)
			SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=PUSH_SA_OKAY").Click() 
			Call CaptureAUTScreenShot
			SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=GS_KNVV-FAKSD").Set ""
			SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=GS_KNVV-LIFSD").Set ""
			Call CaptureAUTScreenShot
			SAPWindowObject.SAPGuiButton("guicomponenttype:=40","tooltip:=Save.*").Click()
			wait 3
			Call CaptureAUTScreenShot
		End If
		SAPWindowObject.SAPGuiOKCode("guicomponenttype:=35","name:=okcd").Set "/nVA41"
		SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
		SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBAK-AUART").set "ZSUB"
		SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBAK-VKORG").set arrSellingEntity(intTestCaseCount)
		SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBAK-SPART").set "00"   
		SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBAK-VTWEG").set "00"
		Call CaptureAUTScreenShot
		SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
		SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=KUAGV-KUNNR").set arrSoldToCustomerID(intTestCaseCount)
		SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=KUWEV-KUNNR").set arrShipToCustomerID(intTestCaseCount)
		SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
		Call CaptureAUTScreenShot
		If SAPWindowObject1.Exist Then
			SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
		End If
		strStatusMessage = SAPWindowObject.SAPGuiStatusBar("guicomponenttype:=103","name:=sbar").GetROProperty("messagetype")
		strMessageText = SAPWindowObject.SAPGuiStatusBar("guicomponenttype:=103","name:=sbar").GetROProperty("text")
		If strStatusMessage = "" Then
			SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP.*").Select ("Item overview")
			SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPMV45ATCTRL_U_ERF_KONTRAKT").SetCellData 1,"Material",arrMaterialID(intTestCaseCount)
			SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPMV45ATCTRL_U_ERF_KONTRAKT").SetCellData 1,"Target quantity","1"
			Call CaptureAUTScreenShot
			SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
			If SAPWindowObject1.SAPGUiButton("guicomponenttype:=40").Exist Then
				SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
			End If
			wait 3
			If SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","tooltip:=Continue.*").Exist Then
				SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
			End If
			SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=BT_HEAD").Click()
			Call CaptureAUTScreenShot
			strMaterialMessage = SAPWindowObject.SAPGuiStatusBar("guicomponenttype:=103","name:=sbar").GetROProperty("messagetype")
			If strMaterialMessage <> "E" Then
'				SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=BT_HEAD").Click()
				SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP_HEAD").Select("Sales")
				SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBAK-VKBUR").set "0050"
				Call CaptureAUTScreenShot
				SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP_HEAD").Select("Order Data") 
				SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBKD-BSARK").set "0020"
				Call CaptureAUTScreenShot
				SAPWindowObject.SendKey F3
				strAmount = SAPWindowObject.SAPGuiEdit("guicomponenttype:=31","name:=VBAK-NETWR").GetROProperty("value")
				If Replace(strAmount," ","") = "0.00" Then
					SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP.*").Select ("Item overview")
					SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPMV45ATCTRL.*").SelectCell 1,"Material"
					SAPWindowObject.SendKey F2
					SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP.*").Select ("Conditions") 
					strCurrency = SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=KOMK-WAERK").GetROProperty("value")
			'		strAmount = SAPWindowObject.SAPGuiEdit("guicomponenttype:=31","name:=KOMK-WAERK").GetTOProperty("value")
			'		strTax = SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=KOMK-WAERK").GetTOProperty("value")
					set arrZITRRowCount = SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPLV69ATCTRL_KONDITIONEN").FindAllRowsByCellContent("Curr.",strCurrency)
					intZITRRowCount = arrZITRRowCount.count
					SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=BT_KOAN").Click()
					SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPLV69ATCTRL_KONDITIONEN").SetCellData intZITRRowCount+1,"CnTy","ZMPR"
	'				strAmountToEnter = arrNetAmount(intTestCaseCount)
					SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPLV69ATCTRL_KONDITIONEN").SetCellData intZITRRowCount+1,"Amount","100"
					SAPWindowObject.SendKey F3
					SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
				End If
				SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP.*").Select ("Item overview")
				SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPMV45ATCTRL.*").SelectCell 1,"Material"
				SAPWindowObject.SendKey F2
				SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP.*").Select ("Conditions") 
				Call CaptureAUTScreenShot
				strAmount = SAPWindowObject.SAPGuiEdit("guicomponenttype:=31","name:=KOMP-NETWR").GetROProperty("value")
				strAmount = Replace(strAmount," ","")
				strTax = SAPWindowObject.SAPGuiEdit("guicomponenttype:=31","name:=KOMP-MWSBP").GetROProperty("value")
				strTax = Replace(strTax," ","")
				SAPWindowObject.SAPGuiMenubar("guicomponenttype:=111","name:=mbar").Select  "Edit;Incompletion log"
				Call CaptureAUTScreenShot
				strCreditStatusMessage = SAPWindowObject.SAPGuiStatusBar("guicomponenttype:=103","name:=sbar").GetROProperty("text")
				Else
					strCreditStatusMessage = "Error"
			End  If
			If strCreditStatusMessage = "Document is complete" Then
				SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[11]").Click()
'				SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
				If SAPWindowObject1.SAPGUiButton("guicomponenttype:=40").Exist Then
					SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
				End If
				If SAPWindowObject1.SAPGUiButton("guicomponenttype:=40").Exist Then
					SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
				End If
				If SAPWindowObject1.SAPGUiButton("guicomponenttype:=40").Exist Then
					SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
				End If
'				SAPWindowObject1.SAPGUiButton("guicomponenttype:=40").Highlight
				If SAPWindowObject1.SAPGUiButton("guicomponenttype:=40").Exist Then
					SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
				End If
'				On error resume next 
'				SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
				
				wait 2
'				SAPWindowObject.SendKey()
				If SAPWindowObject1.SAPGuiButton("guicomponenttype:=40","name:=SPOP-VAROPTION2").Exist Then
					SAPWindowObject1.SAPGuiButton("guicomponenttype:=40","name:=SPOP-VAROPTION2").Click()
					Call CaptureAUTScreenShot
					intBillingDoc = "ERROR"
					intStatusColumnNum = objColumnNumberdictionary.Item("Status")
					objExcelSheet.Cells(intTestCaseCount+1,intStatusColumnNum).Value=  "Data Missing in BP"
					Else
						strContractText = SAPWindowObject.SAPGuiStatusBar("guicomponenttype:=103","name:=sbar").GetROProperty("text")
						Call CaptureAUTScreenShot
						intBillingDoc = SAPWindowObject.SAPGuiStatusBar("guicomponenttype:=103","name:=sbar").GetROProperty("item2")
				End If
				If IsNumeric(intBillingDoc) Then
					intContractColumnNum = objColumnNumberdictionary.Item("ContractNo")
					objExcelSheet.Cells(intTestCaseCount+1,intContractColumnNum).Value=  intBillingDoc
					SAPWindowObject.SAPGuiOKCode("guicomponenttype:=35","name:=okcd").Set "/nVA43"
					SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
					If SAPWindowObject1.Exist Then
						SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
					End If
					SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBAK-VBELN").set intBillingDoc
					SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
					Call CaptureAUTScreenShot
					strExpectedPrice = SAPWindowObject.SAPGuiEdit("guicomponenttype:=31","name:=VBAK-NETWR").GetROProperty("value")
					SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP.*").Select ("Item overview")
					SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPMV45ATCTRL.*").SelectCell 1,"Material"
					SAPWindowObject.SendKey F2
					SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP.*").Select ("Conditions") 
					Call CaptureAUTScreenShot
'					intCountryRowCount = SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPLV69ATCTRL_KONDITIONEN").FindRowByCellContent("Name","One Source Tax Total")
'					intCountryRowCount = SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPLV69ATCTRL_KONDITIONEN").FindRowByCellContent("Name",arrBillToParty(intTestCaseCount))
'					strTaxPercentage = SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPLV69ATCTRL_KONDITIONEN").GetCellData(intCountryRowCount,"Amount")
					strAmount = SAPWindowObject.SAPGuiEdit("guicomponenttype:=31","name:=KOMP-NETWR").GetROProperty("value")
					strAmount = Replace(strAmount," ","")
					strTax = SAPWindowObject.SAPGuiEdit("guicomponenttype:=31","name:=KOMP-MWSBP").GetROProperty("value")
					strTax = Replace(strTax," ","")
					intPriceColumnNum =  objColumnNumberdictionary.Item("NetAmount")
'					If Replace(strExpectedPrice," ","") = Replace(strAmount,",","") Then
					If Replace(strExpectedPrice," ","") = strAmount Then
						objExcelSheet.Cells(intTestCaseCount+1,intPriceColumnNum).Interior.ColorIndex = 35
						objExcelSheet.Cells(intTestCaseCount+1,intPriceColumnNum).Value = strExpectedPrice
						Else
							objExcelSheet.Cells(intTestCaseCount+1,intPriceColumnNum).Value= "Expected-"&strExpectedPrice&"--Actual-"&strAmount
'							objExcelSheet.Cells(intTestCaseCount+1,intPriceColumnNum).Interior.ColorIndex = 36
							objExcelSheet.Cells(intTestCaseCount+1,intPriceColumnNum).font.color = vbRed
					End If
					strTaxPercentage = arrExpectedItemTax(intTestCaseCount)
					If Replace(strTaxPercentage," ","") = "0" Then
						strExpectedTax = strTax
						Else
							strTaxPercentage = Replace(strTaxPercentage," ","")
							strExpectedTax = strAmount*strTaxPercentage/100
							If instr(strExpectedTax,".") Then
								arrExpectedTax = split(strExpectedTax,".")
								strExpectedTax = arrExpectedTax(0)&"."&Left(arrExpectedTax(1),2)
								Else
									strExpectedTax = strExpectedTax&".00"
							End If
					End If
					
					intTaxColumnNum =  objColumnNumberdictionary.Item("TaxAmount")
					If strExpectedTax = Replace(strTax,",","") Then
						objExcelSheet.Cells(intTestCaseCount+1,intTaxColumnNum).Interior.ColorIndex = 35
						objExcelSheet.Cells(intTestCaseCount+1,intTaxColumnNum).Value = "'"&strTax
						Else
							objExcelSheet.Cells(intTestCaseCount+1,intTaxColumnNum).Value = "Expected-"&strExpectedTax&"-- Actual -"&strTax
'							objExcelSheet.Cells(intTestCaseCount+1,intTaxColumnNum).Interior.ColorIndex = 36
							objExcelSheet.Cells(intTestCaseCount+1,intTaxColumnNum).font.color =vbRed
					End If
					SAPWindowObject.SAPGuiOKCode("guicomponenttype:=35","name:=okcd").Set "/nse16n"
					SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
					SAPWindowObject.SAPGUiEdit("guicomponenttype:=32","name:=GD-TAB").Set "/IDT/D_TAX_DATA"
					SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
					intTableRowNumber =  GetSE16NTableFieldNameRowCount("Document Number")
					SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPLSE16NSELFIELDS_TC").SetCellData  intTableRowNumber,"Fr.Value",intBillingDoc
					SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[8]").Click
					Call CaptureAUTScreenShot
					strTableError = SAPWindowObject.SAPGuiStatusBar("guicomponenttype:=103","name:=sbar").GetROProperty("messagetype")
					strExpectedSeller = arrSellerNo(intTestCaseCount)
					If strTableError <> "E" Then
						intTableRowCount = SAPWindowObject.SAPGuiGrid("guicomponenttype:=201","name:=shell").RowCount
						Call CaptureAUTScreenShot
						For intTableIterator = 1 To intTableRowCount
							strSeller = SAPWindowObject.SAPGuiGrid("guicomponenttype:=201","name:=shell").GetCellData(intTableIterator,"Seller VAT Registration Number")
							SAPWindowObject.SAPGuiGrid("guicomponenttype:=201","name:=shell").SelectCell intTableIterator,"Seller VAT Registration Number"
							Call CaptureAUTScreenShot
							If strSeller = "" Then
								strCheckFlag = "N"
								Else
									If  strExpectedSeller = strSeller Then
										intSellerColumnNum = objColumnNumberdictionary.Item("ExpectedSellerNo")
										objExcelSheet.Cells(intTestCaseCount+1,intSellerColumnNum).Interior.ColorIndex = 35
										objExcelSheet.Cells(intTestCaseCount+1,intSellerColumnNum).Value = strSeller
										strCheckFlag = "Y"
										Exit for
										Else
											strCheckFlag = "N"
									End If
							End If	
						Next
						If strCheckFlag = "N" Then
							intSellerColumnNum = objColumnNumberdictionary.Item("ExpectedSellerNo")
							objExcelSheet.Cells(intTestCaseCount+1,intSellerColumnNum).Value = "Expected-"&strExpectedSeller&"-- Actual -"&strSeller
							objExcelSheet.Cells(intTestCaseCount+1,intSellerColumnNum).font.color = vbRed
						End If
						strExpectedSeller = arrSellerNo(intTestCaseCount)
						If strExpectedSeller = "" Then
							If strSeller = "" Then
								intSellerColumnNum = objColumnNumberdictionary.Item("ExpectedSellerNo")
								objExcelSheet.Cells(intTestCaseCount+1,intSellerColumnNum).Value = "Expected-Empty-- Actual -Empty"
								objExcelSheet.Cells(intTestCaseCount+1,intSellerColumnNum).font.color = vbGreen
							End If
						End If
						Else
							intStatusColumnNum = objColumnNumberdictionary.Item("Status")
							objExcelSheet.Cells(intTestCaseCount+1,intStatusColumnNum).Value = "No Values in SE16N"
					End  If 
					Else
						intStatusColumnNum = objColumnNumberdictionary.Item("Status")
						objExcelSheet.Cells(intTestCaseCount+1,intStatusColumnNum).Value=  strContractText&"-Error in Contract Creation"
				End  If 						
				Else
					intStatusColumnNum = objColumnNumberdictionary.Item("Status")
					objExcelSheet.Cells(intTestCaseCount+1,intStatusColumnNum).Value=  strCreditStatusMessage&"-Error in Contract Creation"
			End If	
			Else
				intStatusColumnNum = objColumnNumberdictionary.Item("Status")
				objExcelSheet.Cells(intTestCaseCount+1,intStatusColumnNum).Value=  strMessageText&"-Error in Contract Creation"
		End If
		intDateColumnNum = objColumnNumberdictionary.Item("ExecutionDate")
		objExcelSheet.Cells(intTestCaseCount+1,intDateColumnNum).Value=  Now
		strTesCaseName = arrTestCaseNumber(intTestCaseCount)
		strTesCaseName = strTesCaseName&".docx"
		Call CopyImagesToWord("C:\Users\dchinnasam\OneDrive\Documents\01. Offical\EHP8\Report\"&strTesCaseName)
		strCapImages= ""
	End If 
	objExcelWorkbook.Save	
Next
Call CloseExcelObject(objExcelWorkbook,objExcelObject)



'SystemUtil.CloseProcessByName("EXCEL.EXE")
'Set SAPWindowObject = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
'Set SAPWindowObject1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
'strDataExcelPath = "C:\Users\dchinnasam\OneDrive\Documents\01. Offical\EHP8\TaxValidation\TAX-MultiLine.xlsx"
'strCapImages= ""
''Create Excel object 
'Call CreateExcelObject(strDataExcelPath,"Muliline",intExcelRowCount,intExcelColumnCount,objExcelSheet,objExcelWorkbook)
''Read all Values from Login Sheet
'Call ReadAllValuesFromInputExcel(objExcelSheet)
'intRowToWriteInExcel = ""
'Call FindIterationCount(objExcelObject,objExcelSheet,arrSplitRowNo)
'For intTestIterator = 1 To Ubound(arrSplitRowNo)-1
'	intIterationStartRow = arrSplitRowNo(intTestIterator)
'	intIterationEndRow = arrSplitRowNo(intTestIterator+1)
'	intRowToSet = 1
'	intRowToWriteInExcel = arrSplitRowNo(intTestIterator)
''	For intSplitIterator = intIterationStartRow To intIterationEndRow-1
'		If arrExecutionFlag(intTestIterator) = "Y" or arrExecutionFlag(intTestIterator) = "" Then
'			'If intSplitIterator =  intIterationStartRow Then
'				SAPWindowObject.SAPGuiOKCode("guicomponenttype:=35","name:=okcd").Set "/nVA41"
'				SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'				SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBAK-AUART").set "ZSUB"
'				SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBAK-VKORG").set arrSellingEntity(intTestIterator)
'				SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBAK-SPART").set "00"   
'				SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBAK-VTWEG").set "00"
'				Call CaptureAUTScreenShot
'				SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'				SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=KUAGV-KUNNR").set arrSoldToCustomerID(intTestIterator)
'				SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=KUWEV-KUNNR").set arrShipToCustomerID(intTestIterator)
'				SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'				Call CaptureAUTScreenShot
'				If SAPWindowObject1.Exist Then
'					SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'				End If
'				strStatusMessage = SAPWindowObject.SAPGuiStatusBar("guicomponenttype:=103","name:=sbar").GetROProperty("messagetype")
'				strMessageText = SAPWindowObject.SAPGuiStatusBar("guicomponenttype:=103","name:=sbar").GetROProperty("text")
'			'End  If 
'			If strStatusMessage = "" Then
'				For intSplitIterator = intIterationStartRow To intIterationEndRow-1
'					SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP.*").Select ("Item overview")
'					SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPMV45ATCTRL_U_ERF_KONTRAKT").SetCellData intRowToSet,"Material",arrMaterialID(intSplitIterator-1)
'					SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPMV45ATCTRL_U_ERF_KONTRAKT").SetCellData intRowToSet,"Target quantity","1"
'					Call CaptureAUTScreenShot
'					SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'					If SAPWindowObject1.SAPGUiButton("guicomponenttype:=40").Exist Then
'						SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'					End If
'					wait 3
'					If SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","tooltip:=Continue.*").Exist Then
'						SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'					End If
'					SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=BT_HEAD").Click()
'					Call CaptureAUTScreenShot
'					strMaterialMessage = SAPWindowObject.SAPGuiStatusBar("guicomponenttype:=103","name:=sbar").GetROProperty("messagetype")
'					If strMaterialMessage <> "E" Then
'		'				SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=BT_HEAD").Click()
'						SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP_HEAD").Select("Sales")
'						SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBAK-VKBUR").set "0050"
'						Call CaptureAUTScreenShot
'						SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP_HEAD").Select("Order Data") 
'						SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBKD-BSARK").set "0020"
'						Call CaptureAUTScreenShot
'						SAPWindowObject.SendKey F3
'						strAmount = SAPWindowObject.SAPGuiEdit("guicomponenttype:=31","name:=VBAK-NETWR").GetROProperty("value")
'						If Replace(strAmount," ","") = "0.00" Then
'							SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP.*").Select ("Item overview")
'							SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPMV45ATCTRL.*").SelectCell intRowToSet,"Material"
'							SAPWindowObject.SendKey F2
'							SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP.*").Select ("Conditions") 
'							strCurrency = SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=KOMK-WAERK").GetROProperty("value")
'					'		strAmount = SAPWindowObject.SAPGuiEdit("guicomponenttype:=31","name:=KOMK-WAERK").GetTOProperty("value")
'					'		strTax = SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=KOMK-WAERK").GetTOProperty("value")
'							set arrZITRRowCount = SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPLV69ATCTRL_KONDITIONEN").FindAllRowsByCellContent("Curr.",strCurrency)
'							intZITRRowCount = arrZITRRowCount.count
'							SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=BT_KOAN").Click()
'							SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPLV69ATCTRL_KONDITIONEN").SetCellData intZITRRowCount+1,"CnTy","ZMPR"
'			'				strAmountToEnter = arrNetAmount(intTestCaseCount)
'							SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPLV69ATCTRL_KONDITIONEN").SetCellData intZITRRowCount+1,"Amount","100"
'							SAPWindowObject.SendKey F3
'							SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'						End If
'						SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP.*").Select ("Item overview")
'						SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPMV45ATCTRL.*").SelectCell intRowToSet,"Material"
'						SAPWindowObject.SendKey F2
'						SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP.*").Select ("Conditions") 
'						Call CaptureAUTScreenShot
'						intRowToSet = intRowToSet+1
'					End If 
'				Next
'				SAPWindowObject.SAPGuiMenubar("guicomponenttype:=111","name:=mbar").Select  "Edit;Incompletion log"
'				Call CaptureAUTScreenShot
'				strCreditStatusMessage = SAPWindowObject.SAPGuiStatusBar("guicomponenttype:=103","name:=sbar").GetROProperty("text")
''					Else
''						strCreditStatusMessage = "Error"
'				'End  If
'				If strCreditStatusMessage = "Document is complete" Then
'					SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[11]").Click()
'	'				SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'					If SAPWindowObject1.SAPGUiButton("guicomponenttype:=40").Exist Then
'						SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'					End If
'					If SAPWindowObject1.SAPGUiButton("guicomponenttype:=40").Exist Then
'						SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'					End If
'					If SAPWindowObject1.SAPGUiButton("guicomponenttype:=40").Exist Then
'						SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'					End If
'	'				SAPWindowObject1.SAPGUiButton("guicomponenttype:=40").Highlight
'					If SAPWindowObject1.SAPGUiButton("guicomponenttype:=40").Exist Then
'						SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'					End If
'	'				On error resume next 
'	'				SAPWindowObject1.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'					
'					wait 2
'	'				SAPWindowObject.SendKey()
'					If SAPWindowObject1.SAPGuiButton("guicomponenttype:=40","name:=SPOP-VAROPTION2").Exist Then
'						SAPWindowObject1.SAPGuiButton("guicomponenttype:=40","name:=SPOP-VAROPTION2").Click()
'						Call CaptureAUTScreenShot
'						intBillingDoc = "ERROR"
'						intStatusColumnNum = objColumnNumberdictionary.Item("Status")
'						objExcelSheet.Cells(intTestCaseCount+1,intStatusColumnNum).Value=  "Data Missing in BP"
'						Else
'							strContractText = SAPWindowObject.SAPGuiStatusBar("guicomponenttype:=103","name:=sbar").GetROProperty("text")
'							Call CaptureAUTScreenShot
'							intBillingDoc = SAPWindowObject.SAPGuiStatusBar("guicomponenttype:=103","name:=sbar").GetROProperty("item2")
'					End If
'					If IsNumeric(intBillingDoc) Then
'						intContractColumnNum = objColumnNumberdictionary.Item("ContractNo")
'						objExcelSheet.Cells(intRowToWriteInExcel,intContractColumnNum).Value=  intBillingDoc
'						SAPWindowObject.SAPGuiOKCode("guicomponenttype:=35","name:=okcd").Set "/nVA43"
'						SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'						SAPWindowObject.SAPGuiEdit("guicomponenttype:=32","name:=VBAK-VBELN").set intBillingDoc
'						SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'						Call CaptureAUTScreenShot
'						strExpectedPrice = SAPWindowObject.SAPGuiEdit("guicomponenttype:=31","name:=VBAK-NETWR").GetROProperty("value")
'						SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP.*").Select ("Item overview")
'						intRowToSet = 1
'						For intSplitIterator = intIterationStartRow To intIterationEndRow-1
'							SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPMV45ATCTRL.*").SelectCell intRowToSet,"Material"
'							SAPWindowObject.SendKey F2
'							SAPWindowObject.SAPGuiTabStrip("guicomponenttype:=90","name:=TAXI_TABSTRIP.*").Select ("Conditions") 
'							Call CaptureAUTScreenShot
'		'					intCountryRowCount = SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPLV69ATCTRL_KONDITIONEN").FindRowByCellContent("Name","One Source Tax Total")
'		'					intCountryRowCount = SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPLV69ATCTRL_KONDITIONEN").FindRowByCellContent("Name",arrBillToParty(intTestCaseCount))
'		'					strTaxPercentage = SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPLV69ATCTRL_KONDITIONEN").GetCellData(intCountryRowCount,"Amount")
'							strAmount = SAPWindowObject.SAPGuiEdit("guicomponenttype:=31","name:=KOMP-NETWR").GetROProperty("value")
'							strAmount = Replace(strAmount," ","")
'							strTax = SAPWindowObject.SAPGuiEdit("guicomponenttype:=31","name:=KOMP-MWSBP").GetROProperty("value")
'							strTax = Replace(strTax," ","")
'							intPriceColumnNum =  objColumnNumberdictionary.Item("NetAmount")
'		'					If Replace(strExpectedPrice," ","") = Replace(strAmount,",","") Then
'							If Replace(strExpectedPrice," ","") = strAmount Then
'								objExcelSheet.Cells(intSplitIterator,intPriceColumnNum).Interior.ColorIndex = 35
'								objExcelSheet.Cells(intSplitIterator,intPriceColumnNum).Value = strExpectedPrice
'								Else
'									objExcelSheet.Cells(intSplitIterator,intPriceColumnNum).Value= "Expected-"&strExpectedPrice&"--Actual-"&strAmount
'		'							objExcelSheet.Cells(intTestCaseCount+1,intPriceColumnNum).Interior.ColorIndex = 36
'									objExcelSheet.Cells(intSplitIterator,intPriceColumnNum).font.color = vbRed
'							End If
'							strTaxPercentage = arrExpectedItemTax(intTestCaseCount)
'							If Replace(strTaxPercentage," ","") = "0" Then
'								strExpectedTax = strTax
'								Else
'									strTaxPercentage = Replace(strTaxPercentage," ","")
'									strExpectedTax = strAmount*strTaxPercentage/100
'									If instr(strExpectedTax,".") Then
'										arrExpectedTax = split(strExpectedTax,".")
'										strExpectedTax = arrExpectedTax(0)&"."&Left(arrExpectedTax(1),2)
'										Else
'											strExpectedTax = strExpectedTax&".00"
'									End If
'							End If
'							
'							intTaxColumnNum =  objColumnNumberdictionary.Item("TaxAmount")
'							If strExpectedTax = Replace(strTax,",","") Then
'								objExcelSheet.Cells(intSplitIterator,intTaxColumnNum).Interior.ColorIndex = 35
'								objExcelSheet.Cells(intSplitIterator,intTaxColumnNum).Value = "'"&strTax
'								Else
'									objExcelSheet.Cells(intSplitIterator,intTaxColumnNum).Value = "Expected-"&strExpectedTax&"-- Actual -"&strTax
'		'							objExcelSheet.Cells(intTestCaseCount+1,intTaxColumnNum).Interior.ColorIndex = 36
'									objExcelSheet.Cells(intSplitIterator,intTaxColumnNum).font.color =vbRed
'							End If
'							intRowToSet = intRowToSet+1
'						Next
'						SAPWindowObject.SAPGuiOKCode("guicomponenttype:=35","name:=okcd").Set "/nse16n"
'						SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'						SAPWindowObject.SAPGUiEdit("guicomponenttype:=32","name:=GD-TAB").Set "/IDT/D_TAX_DATA"
'						SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'						intTableRowNumber =  GetSE16NTableFieldNameRowCount("Document Number")
'						SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPLSE16NSELFIELDS_TC").SetCellData  intTableRowNumber,"Fr.Value",intBillingDoc
'						SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[8]").Click
'						Call CaptureAUTScreenShot
'						strTableError = SAPWindowObject.SAPGuiStatusBar("guicomponenttype:=103","name:=sbar").GetROProperty("messagetype")
'						strExpectedSeller = arrSellerNo(intTestCaseCount)
'						If strTableError <> "E" Then
'							intTableRowCount = SAPWindowObject.SAPGuiGrid("guicomponenttype:=201","name:=shell").RowCount
'							Call CaptureAUTScreenShot
'							For intSplitIterator = intIterationStartRow To intIterationEndRow-1
'								For intTableIterator = 1 To intTableRowCount
'									strSeller = SAPWindowObject.SAPGuiGrid("guicomponenttype:=201","name:=shell").GetCellData(intTableIterator,"Seller VAT Registration Number")
'									SAPWindowObject.SAPGuiGrid("guicomponenttype:=201","name:=shell").SelectCell intTableIterator,"Seller VAT Registration Number"
'									Call CaptureAUTScreenShot
'									If strSeller = "" Then
'										strCheckFlag = "N"
'										Else
'											If  strExpectedSeller = strSeller Then
'												intSellerColumnNum = objColumnNumberdictionary.Item("ExpectedSellerNo")
'												objExcelSheet.Cells(intSplitIterator,intSellerColumnNum).Interior.ColorIndex = 35
'												objExcelSheet.Cells(intSplitIterator,intSellerColumnNum).Value = strSeller
'												strCheckFlag = "Y"
'												Exit for
'												Else
'													strCheckFlag = "N"
'											End If
'									End If	
'								Next
'								If strCheckFlag = "N" Then
'									intSellerColumnNum = objColumnNumberdictionary.Item("ExpectedSellerNo")
'									objExcelSheet.Cells(intSplitIterator,intSellerColumnNum).Value = "Expected-"&strExpectedSeller&"-- Actual -"&strSeller
'									objExcelSheet.Cells(intSplitIterator,intSellerColumnNum).font.color = vbRed
'								End If
'								strExpectedSeller = arrSellerNo(intTestCaseCount)
'								If strExpectedSeller = "" Then
'									If strSeller = "" Then
'										intSellerColumnNum = objColumnNumberdictionary.Item("ExpectedSellerNo")
'										objExcelSheet.Cells(intSplitIterator,intSellerColumnNum).Value = "Expected-Empty-- Actual -Empty"
'										objExcelSheet.Cells(intSplitIterator,intSellerColumnNum).font.color = vbGreen
'									End If
'								End If
'							Next
'							Else
'								intStatusColumnNum = objColumnNumberdictionary.Item("Status")
'								objExcelSheet.Cells(intRowToWriteInExcel,intStatusColumnNum).Value = "No Values in SE16N"
'						End  If 
'						Else
'							intStatusColumnNum = objColumnNumberdictionary.Item("Status")
'							objExcelSheet.Cells(intRowToWriteInExcel,intStatusColumnNum).Value=  strContractText&"-Error in Contract Creation"
'					End  If 						
'					Else
'						intStatusColumnNum = objColumnNumberdictionary.Item("Status")
'						objExcelSheet.Cells(intRowToWriteInExcel,intStatusColumnNum).Value=  strCreditStatusMessage&"-Error in Contract Creation"
'				End If	
'				Else
'					intStatusColumnNum = objColumnNumberdictionary.Item("Status")
'					objExcelSheet.Cells(intRowToWriteInExcel,intStatusColumnNum).Value=  strMessageText&"-Error in Contract Creation"
'			End If
'			intDateColumnNum = objColumnNumberdictionary.Item("ExecutionDate")
'			objExcelSheet.Cells(intTestCaseCount+1,intDateColumnNum).Value=  Now
'			strTestCaseName = arrSoldToCustomerID(intRowToWriteInExcel-1)
'			'strTesCaseName = arrTestCaseNumber(intRowToWriteInExcel-1)
'			strTesCaseName = strTesCaseName&".docx"
'			Call CopyImagesToWord("C:\Users\dchinnasam\OneDrive\Documents\01. Offical\EHP8\Report\"&strTesCaseName)
'			strCapImages= ""
'		End If 
'	Next
''Next
'
Public Function FindIterationCount(objExcelObject,objExcelSheet,arrSplitRowNo)
	Set objUsedRange=objExcelSheet.usedrange
	strRowCount =objUsedRange.rows.count
	strColumnCount =objUsedRange.columns.count
	Set strFindValue = objExcelSheet.Range("A2:A"&strRowCount).Find("END")
	strAddress = strFindValue.address
	strSplitValue = Split(strAddress, "$")
	strRowTillValue = strSplitValue(2)
	strValuesFound = ""
	For intCount = 2 To strRowTillValue - 1
		strCellValue = objExcelSheet.Cells(intCount,1).Value
		If strCellValue <> "" Then
			strValuesFound = strValuesFound&";"&intCount
		End If
	Next
	strValuesFound = strValuesFound&";"&strRowTillValue
	arrSplitRowNo = Split(strValuesFound, ";")
End Function

'****************************************************************************************************************
'Name of the Function   :GetSE16NTableFieldNameRowCount(strColumnName)
'Author     :DeepanRaj
'Description    :
'Input Parameters    :
'Output Parameters      :
'Creation Date :26-Julu-2022
'****************************************************************************************************************
'GetSE16NTableFieldNameRowCount(strColumnName)
'****************************************************************************************************************
 Public Function GetSE16NTableFieldNameRowCount(strColumnName)
	intTableRowCount = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable("guicomponenttype:=80","name:=SAPLSE16NSELFIELDS_TC").RowCount
	For intRowCount = 1 To intTableRowCount
		strColumnValue = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21").SAPGuiTable("guicomponenttype:=80","name:=SAPLSE16NSELFIELDS_TC").GetCellData(intRowCount,"Fld name")
		If strColumnValue = strColumnName Then
			GetSE16NTableFieldNameRowCount = intRowCount
		Exit For
	End If
Next
End Function



''****************************************************************************************************************
'Name of the Function   :CloseExcelObject(objExcelWorkbook,objExcelObject)
'Author     :DeepanRaj
'Description    :CloseExcelObject(objExcelWorkbook,objExcelObject)
'Input Parameters    :objExcelSheet
'Output Parameters      :
'Creation Date :
'****************************************************************************************************************
'CloseExcelObject(objExcelWorkbook,objExcelObject)
'****************************************************************************************************************
Public Function CloseExcelObject(objExcelWorkbook,objExcelObject)
	On Error Resume Next
	objExcelWorkbook.Save
	objExcelWorkbook.Close	
	Set objExcelWorkbook = Nothing 
	Set objExcelObject = Nothing
End Function

Public Function CreateExcelObject(strExcelFilePath,strSheetName,intExcelRowCount,intExcelColumnCount,objExcelSheet,objExcelWorkbook)
	Set objExcelObject = CreateObject("Excel.Application")
	Set objExcelWorkbook=objExcelObject.Workbooks.Open (strExcelFilePath)
	Set objExcelSheet = objExcelObject.WorkSheets(strSheetName) 
	Set excelUsedRange=objExcelSheet.usedrange
	intExcelRowCount = excelUsedRange.rows.count
	intExcelColumnCount=excelUsedRange.Columns.count
End Function




Public Function ReadAllValuesFromInputExcel(objExcelSheet)
	Set excelUsedRange=objExcelSheet.usedrange
	excelRowCount = excelUsedRange.rows.count
	excelColumnCount=excelUsedRange.Columns.count
	For intColumnLoop = 1 To excelColumnCount
		strFlagForFirstForExit = ""
		If strFlagForFirstForExit = "YES" Then
			Exit For 		
		End If
		For inRowLoop = 2 To excelRowCount - 1
			strFinalRowValue = objExcelSheet.Cells(inRowLoop,1).Value
	    		If strFinalRowValue = "END" Then
	    			strFlagForFirstForExit = "YES"
	    			Exit For
	   	 	End If
	    		strIterationCellValue = objExcelSheet.Cells(inRowLoop,intColumnLoop).Value
	   		strEntireColumnValue = strEntireColumnValue & ";" & strIterationCellValue
		Next
		strExcelColumnName = objExcelSheet.Cells(1,intColumnLoop).Value
		strEntireColumnValue = Replace(strEntireColumnValue,VBCR&VBLF,"")
		strEntireColumnValue = Replace(strEntireColumnValue,VBCR,"")
		strEntireColumnValue = Replace(strEntireColumnValue,VBLF,"")
		strTempExcelColumnName = Replace(strExcelColumnName,"ID_","str")
		Execute strTempExcelColumnName & " =  " & Chr(34) & strEntireColumnValue & Chr(34) 
		Execute "arr" &strExcelColumnName & " = " & "Split(" & strTempExcelColumnName & Chr(44) & Chr (34) & Chr(59) & Chr(34) & ")"
		strEntireColumnValue = ""
	Next
End Function

'Set SAPWindowObject = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
'Set SAPWindowObject1 = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=22")
''SAPWindowObject1.SAPGUiButton("guicomponenttype:=40").Highlight
'''msgbox 
'''msgbox SAPWindowObject.SAPGuiTable("guicomponenttype:=80","name:=SAPLV69ATCTRL_KONDITIONEN").GetCellData(13,"Name")
''
''
'strCapImages = ""
'SAPWindowObject.SAPGuiOKCode("guicomponenttype:=35","name:=okcd").Set "/nVA43"
'SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'Call CaptureAUTScreenShot
'SAPWindowObject.SAPGuiOKCode("guicomponenttype:=35","name:=okcd").Set "/nVA42"
'SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'Call CaptureAUTScreenShot
'SAPWindowObject.SAPGuiOKCode("guicomponenttype:=35","name:=okcd").Set "/nVA41"
'SAPWindowObject.SAPGUiButton("guicomponenttype:=40","name:=btn[0]").Click()
'Call CaptureAUTScreenShot
'Call CopyImagesToWord("C:\Users\dchinnasam\OneDrive\Documents\00. Admin\test.docx")

Function CaptureAUTScreenShot
    ImageDir = "C:\Users\dchinnasam\OneDrive\Documents\00. Admin\" 
   Set BR = SAPGuiSession("guicomponenttype:=12").SAPGuiWindow("guicomponenttype:=21")
 
    If BR.Exist Then
 
        strTime = Split(Replace(Time,":","-")," ")
 
        ImageName = strTime(0) & " " & strTime(1)
 
        Set fso = CreateObject("Scripting.FileSystemObject")
 
        On Error Resume Next
 
        BR.CaptureBitmap ImageDir & ImageName & ".png"
 
        strCapImages = strCapImages & "," & ImageDir & ImageName & ".png"
 
        Set fso = Nothing
 
        If Err.Number > 0 Then
 
            Reporter.ReportEvent micFail,"Some error occured while capturing Screen shot",""
 
        End If
 
        On Error Goto 0
 
    End If
 
End Function


Function CopyImagesToWord(WordFileName)

    Const MOVE_SELECTION = 0

    Const END_OF_STORY = 6

    strCapImages = MID(strCapImages,2)

    If strCapImages <> Empty Then

        arrStrCapImages = Split(strCapImages,",")

    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(WordFileName) Then

        blnExistingFile = True

    Else

        blnExistingFile = False
    
    End If

    Set fso = Nothing

    Set objWord = CreateObject("Word.Application")

    If blnExistingFile = False Then

        Set objDoc = objWord.Documents.Add

    Else

        Set objDoc = objWord.Documents.Open(WordFileName) 
   
    End If

    Set objSelection = objWord.Selection

    objSelection.EndKey END_OF_STORY,MOVE_SELECTION

    objSelection.TypeParagraph

    objSelection.Font.Name = "Verdana"

    objSelection.Font.Size = 12

    objSelection.Font.Bold = True

    objSelection.ParagraphFormat.Alignment = wdAlignParagraphCenter

    objSelection.TypeText "Captured Screen Shots copied to word document on " & Now

    objSelection.TypeParagraph

    For intCnt = 1 to Ubound(arrStrCapImages)

            objSelection.EndKey END_OF_STORY,MOVE_SELECTION

            objSelection.TypeParagraph

            objSelection.Font.Name = "Verdana"

            objSelection.Font.Size = 8

            objSelection.InlineShapes.AddPicture arrStrCapImages(intCnt),true

            objSelection.EndKey END_OF_STORY,MOVE_SELECTION

            objSelection.TypeParagraph

            If Err.number > 0 Then

                Reporter.ReportEvent micWarning,"Invalid Image file path: " & arrStrCapImages(intCnt)

            End If

            On error Goto 0
    Next

        'Saving the word document
        objSelection.WholeStory

        ObjDoc.SaveAs(WordFileName)

        objWord.Quit(wdSaveChanges)

        OutputToWord = True

        If Err.number > 0 Then

            Reporter.ReportEvent micFail,"Unable to Save word document",""

        End If   
         
        On error Goto 0

        arrStrCapImages = Null
    
End Function

