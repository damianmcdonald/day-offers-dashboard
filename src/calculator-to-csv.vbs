' calculator-to-csv.vbs
' Author: Damian McDonald
' Creation date: 06/05/2020
' Purpose:
' 	Reads values from an instance of the Day Calculator and writes those values to a CSV dataset
' Structure of the CSV dataset is provided below:
' client,status,statusDate,cloud,greenfield,regions,accounts,applications,vpcs,subnets,hasConnectivity,hasPeerings,hasDirectoryService,hasAdvancedSecurity,hasAdvancedLogging,hasAdvancedMonitoring,hasAdvancedBackup,virtualMachines,buckets,databases,hasELB,hasAutoScripts,hasOtherServices,service1,service2,service3,service4,service5,phase1EstimatePre,phase1Estimate,phase1Deviation,phase2EstimatePre,phase2Estimate,phase2Deviation,phase3EstimatePre,phase3Estimate,phase3Deviation,phase4EstimatePre,phase4Estimate,phase4Deviation,totalPre,total,totalDeviation,travel,administered,geoLocation,isValid

if WScript.Arguments.Count <> 2 then
    WScript.Echo ">>> ERROR >>> Incorrect number of arguments provided to script."
    WScript.Echo ">>> INFO >>> USAGE: cscript calculator-to-csv.vbs PATH_TO_CALCULATOR.xlsm PATH_TO_DATASET.csv"	
	WScript.Quit 99
end if

Sub Main()

	'Grab the cli arguments
	calculatorFilePath = WScript.Arguments(0)
	datasetFilePath = WScript.Arguments(1)
	 
	'Define the worksheet positions
	WKS_VALIDACION = 1
	WKS_CALCULADORA = 2
	WKS_OUTPUT = 3
		
	'Open the Excel calculator in Read-Only mode
	Set objExcel = CreateObject("Excel.Application")
	Call objExcel.Workbooks.Open(calculatorFilePath, True, True)
	
	'Grab the relevant worksheets
	Set worksheetValidacion = objExcel.ActiveWorkbook.Sheets(WKS_VALIDACION)
	Set worksheetCalculadora = objExcel.ActiveWorkbook.Sheets(WKS_CALCULADORA)
	Set worksheetOutput = objExcel.ActiveWorkbook.Sheets(WKS_OUTPUT)

	'WORKSHEET VALIDACION
	
	'Validacion worksheet cell references
	REF_CLOUD = "C3"
	REF_GREENFIELD = "C4"
	REF_REGIONS = "C5"
	REF_ACCOUNTS = "C6"
	REF_APPS = "C7"
	REF_VPCS = "C9" 
	REF_SUBNETS = "C10"
	REF_VPN = "C11" 
	REF_PEERING = "C12"
	REF_DIRECTORY_SERVICE = "C14"
	REF_ADV_SECURITY = "C15"
	REF_ADV_LOGGING = "C17"
	REF_ADV_MONITORING = "C18"
	REF_ADV_BACKUP = "C19"
	REF_VMS = "C21"
	REF_BUCKETS = "C22"
	REF_DATABASES = "C23"
	REF_ELB = "C24"
	REF_AUTO_SCRIPTS = "C25"
	REF_OTHER_SERVICES = "C27"
	REF_SERVICE_1 = "Service1ComboBox"
	REF_SERVICE_2 = "Service2ComboBox" 
	REF_SERVICE_3 = "Service3ComboBox" 
	REF_SERVICE_4 = "Service4ComboBox"
	REF_SERVICE_5 = "Service5ComboBox"
	
	'Grab the values from the Validacion worksheet
	cloud = worksheetValidacion.Range(REF_CLOUD).Value
	greenfield = ConvertBoolean(worksheetValidacion.Range(REF_GREENFIELD).Value)
	regions = worksheetValidacion.Range(REF_REGIONS).Value
	accounts = worksheetValidacion.Range(REF_ACCOUNTS).Value
	apps = worksheetValidacion.Range(REF_APPS).Value
	vpcs = worksheetValidacion.Range(REF_VPCS).Value
	subnets = worksheetValidacion.Range(REF_SUBNETS).Value
	vpn = ConvertBoolean(worksheetValidacion.Range(REF_VPN).Value)
	peering = ConvertBoolean(worksheetValidacion.Range(REF_PEERING).Value)
	directoryService = ConvertBoolean(worksheetValidacion.Range(REF_DIRECTORY_SERVICE).Value)
	advSecurity = ConvertBoolean(worksheetValidacion.Range(REF_ADV_SECURITY).Value)
	advLogging = ConvertBoolean(worksheetValidacion.Range(REF_ADV_LOGGING).Value)
	advMonitoring = ConvertBoolean(worksheetValidacion.Range(REF_ADV_MONITORING).Value)
	advBackup = ConvertBoolean(worksheetValidacion.Range(REF_ADV_BACKUP).Value)
	vms = worksheetValidacion.Range(REF_VMS).Value
	buckets = worksheetValidacion.Range(REF_BUCKETS).Value
	databases = worksheetValidacion.Range(REF_DATABASES).Value
	elb = ConvertBoolean(worksheetValidacion.Range(REF_ELB).Value)
	autoScripts = ConvertBoolean(worksheetValidacion.Range(REF_AUTO_SCRIPTS).Value)
	otherServices = ConvertBoolean(worksheetValidacion.Range(REF_OTHER_SERVICES).Value)
	
	'Get the values from the other services combo boxes
	Dim service1ComboBox, service2ComboBox, service3ComboBox, service4ComboBox, service5ComboBox, OTHER_SERVICES_BLANK_TEXT
	OTHER_SERVICES_BLANK_TEXT = "Otros servicios no selecionado"
	
	Set service1ComboBox = worksheetValidacion.OLEObjects(REF_SERVICE_1)
	If service1ComboBox.Object.value = OTHER_SERVICES_BLANK_TEXT Then
		service1 = ""
	Else 
		service1 = service1ComboBox.Object.value
	End If
	
	Set service2ComboBox = worksheetValidacion.OLEObjects(REF_SERVICE_2)
	If service2ComboBox.Object.value = OTHER_SERVICES_BLANK_TEXT Then
		service2 = ""
	Else 
		service2 = service2ComboBox.Object.value
	End If
	
	Set service3ComboBox = worksheetValidacion.OLEObjects(REF_SERVICE_3)
	If service3ComboBox.Object.value = OTHER_SERVICES_BLANK_TEXT Then
		service3 = ""
	Else 
		service3 = service3ComboBox.Object.value
	End If
	
	Set service4ComboBox = worksheetValidacion.OLEObjects(REF_SERVICE_4)
	If service4ComboBox.Object.value = OTHER_SERVICES_BLANK_TEXT Then
		service4 = ""
	Else 
		service4 = service4ComboBox.Object.value
	End If
	
	Set service5ComboBox = worksheetValidacion.OLEObjects(REF_SERVICE_5)
	If service5ComboBox.Object.value = OTHER_SERVICES_BLANK_TEXT Then
		service5 = ""
	Else 
		service5 = service5ComboBox.Object.value
	End If
	
	
	'WORKSHEET CALCULADORA

	'Calculadora worksheet cell references
	REF_IS_ADMINISTERED = "F10"
	REF_TRAVEL = "F13"
	REF_CLIENT_NAME = "F16"
	REF_OFFER_DATE = "F17"
	REF_CLIENT_LOCATION = "F18"
	
	'Grab the values from the Calculadora worksheet
	isAdministered = ConvertBoolean(worksheetCalculadora.Range(REF_IS_ADMINISTERED).Value)
	travel = worksheetCalculadora.Range(REF_TRAVEL).Value
	clientName = worksheetCalculadora.Range(REF_CLIENT_NAME).Value
	offerDate = worksheetCalculadora.Range(REF_OFFER_DATE).Value
	clientLocation = worksheetCalculadora.Range(REF_CLIENT_LOCATION).Value


	'WORKSHEET OUTPUT
	
	'Output worksheet cell references
	REF_PHASE_1_PRE_ESTIMATE = "E4"
	REF_PHASE_2_PRE_ESTIMATE = "E6"
	REF_PHASE_3_PRE_ESTIMATE = "E7"
	REF_PHASE_3_S1_PRE_ESTIMATE = "E9"
	REF_PHASE_3_S2_PRE_ESTIMATE = "E10"
	REF_PHASE_3_S3_PRE_ESTIMATE = "E11"
	REF_PHASE_3_S4_PRE_ESTIMATE = "E12"
	REF_PHASE_3_S5_PRE_ESTIMATE = "E13"
	REF_PHASE_4_PRE_ESTIMATE = "E15"
	TOTAL_PRE_ESTIMATE = "E21"
	
	REF_PHASE_1_ESTIMATE = "H4"
	REF_PHASE_2_ESTIMATE = "H6"
	REF_PHASE_3_ESTIMATE = "H7"
	REF_PHASE_3_S1_ESTIMATE = "H9"
	REF_PHASE_3_S2_ESTIMATE = "H10"
	REF_PHASE_3_S3_ESTIMATE = "H11"
	REF_PHASE_3_S4_ESTIMATE = "H12"
	REF_PHASE_3_S5_ESTIMATE = "H13"
	REF_PHASE_4_ESTIMATE = "H15"
	TOTAL_ESTIMATE = "H21"
	
	IS_CALCULATOR_SCOPE_VALID = "K3"
	
	'Grab the values from the Output worksheet
	
	'Pre-estimations
	
	'Fase de recopilacion
	phase1EstimatePre = worksheetOutput.Range(REF_PHASE_1_PRE_ESTIMATE).Value
	'Fase de diseno
	phase2EstimatePre = worksheetOutput.Range(REF_PHASE_2_PRE_ESTIMATE).Value
	'Fase de implantacion
	phase3EstimatePre = worksheetOutput.Range(REF_PHASE_3_PRE_ESTIMATE).Value
	phase3EstimatePreS1 = worksheetOutput.Range(REF_PHASE_3_S1_PRE_ESTIMATE).Value
	phase3EstimatePreS2 = worksheetOutput.Range(REF_PHASE_3_S2_PRE_ESTIMATE).Value
	phase3EstimatePreS3 = worksheetOutput.Range(REF_PHASE_3_S3_PRE_ESTIMATE).Value
	phase3EstimatePreS4 = worksheetOutput.Range(REF_PHASE_3_S4_PRE_ESTIMATE).Value
	phase3EstimatePreS5 = worksheetOutput.Range(REF_PHASE_3_S5_PRE_ESTIMATE).Value
	'Aggregated value including estimates for any additional services
	phase3EstimatePreAggregate = SumEstimates(phase3EstimatePre, phase3EstimatePreS1, phase3EstimatePreS2, phase3EstimatePreS3, phase3EstimatePreS4, phase3EstimatePreS5)    
	'Fase de soporte
	phase4EstimatePre = worksheetOutput.Range(REF_PHASE_4_PRE_ESTIMATE).Value
	'Total pre-estimate
	totalEstimatePre = worksheetOutput.Range(TOTAL_PRE_ESTIMATE).Value
	
	' Estimations
	
	'Fase de recopilacion
	phase1Estimate = worksheetOutput.Range(REF_PHASE_1_ESTIMATE).Value
	'Fase de diseno
	phase2Estimate = worksheetOutput.Range(REF_PHASE_2_ESTIMATE).Value
	'Fase de implantacion
	phase3Estimate = worksheetOutput.Range(REF_PHASE_3_ESTIMATE).Value
	phase3EstimateS1 = worksheetOutput.Range(REF_PHASE_3_S1_ESTIMATE).Value
	phase3EstimateS2 = worksheetOutput.Range(REF_PHASE_3_S2_ESTIMATE).Value
	phase3EstimateS3 = worksheetOutput.Range(REF_PHASE_3_S3_ESTIMATE).Value
	phase3EstimateS4 = worksheetOutput.Range(REF_PHASE_3_S4_ESTIMATE).Value
	phase3EstimateS5 = worksheetOutput.Range(REF_PHASE_3_S5_ESTIMATE).Value
	'Aggregated value including estimates for any additional services
	phase3EstimateAggregate = SumEstimates(phase3Estimate, phase3EstimateS1, phase3EstimateS2, phase3EstimateS3, phase3EstimateS4, phase3EstimateS5)    
	'Fase de soporte
	phase4Estimate = worksheetOutput.Range(REF_PHASE_4_ESTIMATE).Value
	'Total pre-estimate
	totalEstimate = worksheetOutput.Range(TOTAL_ESTIMATE).Value
	
	'Calculate the degree of deviation between the pre-estimate and the  estimate
	phase1Deviation = phase1EstimatePre - phase1Estimate
	phase2Deviation = phase2EstimatePre - phase2Estimate
	phase3Deviation = phase3EstimatePreAggregate - phase3EstimateAggregate
	phase4Deviation = phase4EstimatePre - phase4Estimate
	totalDeviation = totalEstimatePre - totalEstimate
	
	isCalculatorScopeValid = UCase(worksheetOutput.Range(IS_CALCULATOR_SCOPE_VALID).Value)


	'CSV operations
	
	'CSV structure is described below
	'client,status,statusDate,cloud,greenfield,regions,accounts,applications,vpcs,subnets,hasConnectivity,hasPeerings,hasDirectoryService,hasAdvancedSecurity,hasAdvancedLogging,hasAdvancedMonitoring,hasAdvancedBackup,virtualMachines,buckets,databases,hasELB,hasAutoScripts,hasOtherServices,service1,service2,service3,service4,service5,phase1EstimatePre,phase1Estimate,phase1Deviation,phase2EstimatePre,phase2Estimate,phase2Deviation,phase3EstimatePre,phase3Estimate,phase3Deviation,phase4EstimatePre,phase4Estimate,phase4Deviation,totalPre,total,totalDeviation,travel,administered,geoLocation,isValid

	'Build the CSV string for insertion
	Set sb = new StringBuilder
	
	Call sb.AppendWithValidation(clientName, "Client Name", "String")
	Call sb.Append("OFFER_RECEIVED")
	Call sb.AppendWithValidation(offerDate, "Offer Date", "Date")
	Call sb.AppendWithValidation(cloud, "Cloud", "String")
	Call sb.AppendWithValidation(greenfield, "Greenfield", "Boolean")
	Call sb.AppendWithValidation(regions, "# Regions", "Number")
	Call sb.AppendWithValidation(accounts, "# Accounts", "Number")
	Call sb.AppendWithValidation(apps, "# Applications", "Number")
	Call sb.AppendWithValidation(vpcs, "# VPCs", "Number")
	Call sb.AppendWithValidation(subnets, "# Subnets", "Number")
	Call sb.AppendWithValidation(vpn, "Connectivity On-Premises", "Boolean")
	Call sb.AppendWithValidation(peering, "VPC Peering", "Boolean")
	Call sb.AppendWithValidation(directoryService, "Directory Services", "Boolean")
	Call sb.AppendWithValidation(advSecurity, "Advanced Security", "Boolean")
	Call sb.AppendWithValidation(advLogging, "Advanced Logging", "Boolean")	
	Call sb.AppendWithValidation(advMonitoring, "Advanced Monitoring", "Boolean")
	Call sb.AppendWithValidation(advBackup, "Advanced Backup", "Boolean")
	Call sb.AppendWithValidation(vms, "# Virtual Machines", "Number")
	Call sb.AppendWithValidation(buckets, "# Buckets", "Number")
	Call sb.AppendWithValidation(databases, "# Databases", "Number")
	Call sb.AppendWithValidation(elb, "Load Balancer", "Boolean")
	Call sb.AppendWithValidation(autoScripts, "Automation Scripts", "Boolean")
	Call sb.AppendWithValidation(otherServices, "Other Services", "Boolean")
	Call sb.Append(service1)
	Call sb.Append(service2)
	Call sb.Append(service3)
	Call sb.Append(service4)
	Call sb.Append(service5)	
	Call sb.AppendWithValidation(phase1EstimatePre, "Phase 1 Pre Estimate", "Number")
	Call sb.AppendWithValidation(phase1Estimate, "Phase 1  Estimate", "Number")
	Call sb.AppendWithValidation(phase1Deviation, "Phase 1 Deviation", "Number")
	Call sb.AppendWithValidation(phase2EstimatePre, "Phase 2 Pre Estimate", "Number")
	Call sb.AppendWithValidation(phase2Estimate, "Phase 2  Estimate", "Number")
	Call sb.AppendWithValidation(phase2Deviation, "Phase 2 Deviation", "Number")
	Call sb.AppendWithValidation(phase3EstimatePreAggregate, "Phase 3 Pre Estimate", "Number")
	Call sb.AppendWithValidation(phase3EstimateAggregate, "Phase 3  Estimate", "Number")
	Call sb.AppendWithValidation(phase3Deviation, "Phase 3 Deviation", "Number")
	Call sb.AppendWithValidation(phase4EstimatePre, "Phase 4 Pre Estimate", "Number")
	Call sb.AppendWithValidation(phase4Estimate, "Phase 4  Estimate", "Number")
	Call sb.AppendWithValidation(phase4Deviation, "Phase 4 Deviation", "Number")
	Call sb.AppendWithValidation(totalEstimatePre, "Total Pre Estimate", "Number")
	Call sb.AppendWithValidation(totalEstimate, "Total  Estimate", "Number")
	Call sb.AppendWithValidation(totalDeviation, "Total Deviation", "Number")
	Call sb.AppendWithValidation(travel, "# viajes", "Number")
	Call sb.AppendWithValidation(isAdministered, "Is administetred", "Boolean")
	Call sb.AppendWithValidation(clientLocation, "Client Location", "String")
	Call sb.AppendWithValidation(isCalculatorScopeValid, "Calculator Scope", "Boolean")
	
	WScript.Echo ">>> INFO >>> CSV data: "
	WScript.Echo sb.ToString()
	
	'Write to the CSV file
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set csvFile = objFSO.OpenTextFile(datasetFilePath, ForAppending, True)
	'Append the CSV string to the dataset starting with a new line
	csvFile.Write vbCrLf & sb.ToString()
	
	'Close the csv file when we are done
	csvFile.Close

	objExcel.ActiveWorkbook.Close(False)
	objExcel.Quit

 End Sub
 
 Function SumEstimates(ByVal base, ByVal s1, ByVal s2, ByVal s3, ByVal s4, ByVal s5)
	SumEstimates = base + s1 + s2 +s3 + s4 + s5
 End Function
 
 Function ConvertBoolean(ByVal inputValue)
	If inputValue = "No" OR inputValue = "Por defecto" OR inputValue = "Snapshots (1 diario)" Then
        ConvertBoolean = "FALSE"
	Else
        ConvertBoolean = "TRUE"
    End If
 End Function
 
 Function FormatDate(ByVal inputDate)
	dd = Right("00" & Day(inputDate), 2)
    mm = Right("00" & Month(inputDate), 2)
    yyyy = Year(inputDate)  
    FormatDate= yyyy & "-" & mm & "-" & dd
End Function
 
 Class StringBuilder
     
    Dim stringArray
     
    Private Sub Class_Initialize()
        Set stringArray = CreateObject("System.Collections.ArrayList")
    End Sub
     
    Public Sub Append(ByVal strValue)
        stringArray.Add strValue
    End Sub
    
	Public Sub AppendWithValidation(ByVal inputValue, ByVal fieldName, ByVal validationType)
		If inputValue = "" OR TypeName(inputValue) = "Empty" Then
			WScript.Echo ">>> ERROR >>> fieldName: " & fieldName & " is empty."
		End If
		If validationType = "String" Then
			If TypeName(inputValue) <> "String" Then
				WScript.Echo ">>> ERROR >>> fieldName: " & fieldName & " with value: " & inputValue & " is of type " & TypeName(inputValue) & " and not String as expected."
			End IF
		End If
		If validationType = "Boolean" Then
			If inputValue <> "TRUE" AND inputValue <> "FALSE" Then
				WScript.Echo ">>> ERROR >>> fieldName: " & fieldName & " with value: " & inputValue & " is not of type Boolean as expected. Expected a TRUE or FALSE value."
			End IF
		End If
		If validationType = "Date" Then
			If TypeName(inputValue) <> "Date" Then
				WScript.Echo ">>> ERROR >>> fieldName: " & fieldName & " with value: " & inputValue & " is of type " & TypeName(inputValue) & " and not Date as expected."
				'Even if the date is a String and not in Date format, still try to format the date
				On Error Resume Next
				Err.Clear      ' Clear any possible Error that previous code raised
				'Format the date, append it to the array and exit Sub
				stringArray.Add FormatDate(inputValue)
				If Err.Number <> 0 Then
					WScript.Echo ">>> ERROR >>>: " & Err.Number
					WScript.Echo ">>> ERROR (Hex) >>>: " & Hex(Err.Number)
					WScript.Echo ">>> ERROR (Source) >>>: " &  Err.Source
					WScript.Echo ">>> ERROR (Description) >>>: " &  Err.Description
					Err.Clear             ' Clear the Error
				Else
					WScript.Echo ">>> INFO >>> Successfully converted the String " & inputValue & " to a Date object."
				End If
				On Error Goto 0           ' Don't resume on Error
				Exit Sub
			Else
				'Format the date, append it to the array and exit Sub
				stringArray.Add FormatDate(inputValue)
				Exit Sub
			End IF
		End If
		
		If validationType = "Number" Then
			If TypeName(inputValue) <> "Integer" AND TypeName(inputValue) <> "Long" AND TypeName(inputValue) <> "Single" AND TypeName(inputValue) <> "Double" AND TypeName(inputValue) <> "Decimal" Then
				WScript.Echo ">>> ERROR >>> fieldName: " & fieldName & " with value: " & inputValue & " is of type " & TypeName(inputValue) & " and not a Number type as expected."
			End IF
		End If
        stringArray.Add inputValue
    End Sub
	
    Public Sub PrePend(ByVal strValue)
        stringArray.Insert 0, strValue
    End Sub
     
    Public Function ToString()
        ToString = Join(stringArray.ToArray(), ",")
    End Function
     
    Public Function Count()
        Count = stringArray.Count()
    End Function
     
    Public Sub Reset()
        stringArray.Clear()
        Class_Initialize
    End Sub
     
    Public Function Contains(ByVal strString)
        Contains = False
        If stringArray.Contains(strString) Then Contains = True
    End Function
     
End Class
 
 call Main