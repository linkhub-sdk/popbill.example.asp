<!--#include file="Popbill.asp"--> 
<!--#include file="FaxService.asp"--> 
<html>
<head>
	<title>ASP 참 그지같다.</title>
	<meta http-equiv="Content-Type" content="text/html;charset=euc-kr" />
</head>
<body>
<div>
<%
	
	set m_FaxService = new FaxService
	m_FaxService.Initialize "TESTER", "t4B19Ph5K2aIh9oNd91Q99Vwe9jST2/2IJbWjxhCgsA="
	m_FaxService.IsTest = True
	
	
	On Error Resume Next

	UnitCost = m_FaxService.GetUnitCost("1231212312")

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "Fax.GetUnitCost : " + CStr(UnitCost)
	End If

	On Error GoTo 0

	Dim receivers(1)

	Set receivers(0) = New FaxReceiver

	receivers(0).receiverNum = "00011112222"
	'receivers(0).receiverName = "수신자 명칭"

	Set receivers(1) = New FaxReceiver

	receivers(1).receiverNum = "00011112222"
	'receivers(1).receiverName = "수신자 명칭"

	FilePaths = Array("C:\Inetpub\wwwroot\Popbill\로고.gif","C:\Inetpub\wwwroot\Popbill\로고.gif")

	ReserveDT = "" '예약시간.
	UserID = ""  '팝빌아이디

	On Error Resume Next

	UnitCost = m_FaxService.SendFAX("1231212312","07075106766",receivers,FilePaths,ReserveDT , UserID)

	If Err.Number <> 0 then
		Response.Write("Error Number -> " & Err.Number)
		Response.write("<BR>Error Source -> " & Err.Source)
		Response.Write("<BR>Error Desc   -> " & Err.Description)
		Err.Clears
	Else
		Response.write "Fax.GetUnitCost : " + CStr(UnitCost)
	End If

	On Error GoTo 0


%>
</div>
</body>
</html>