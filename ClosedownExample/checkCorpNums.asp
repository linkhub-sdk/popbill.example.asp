<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="../Example.css" media="screen" />
		<title>휴폐업조회 API SDK ASP Example.</title>
	</head>
	<!--#include file="common.asp"--> 
	<%
		'**************************************************************
		' 다수의 사업자에 대한 휴폐업여부를 조회합니다. (최대 1000건)
		'**************************************************************

		'팝빌회원 사업자번호
		UserCorpNum = "1234567890"		

		'조회할 사업자번호 배열, 최대 1000건
		Dim CorpNumList(3)
		CorpNumList(0) = "1234567890"
		CorpNumList(1) = "4108600477"
		CorpNumList(2) = "110-04-45791"
						
		On Error Resume Next

		Set result = m_ClosedownService.checkCorpNums(UserCorpnum, CorpNumList)
		
		If Err.Number <> 0 Then
			code = Err.Number
			message = Err.Description
			Err.Clears
		End If

		On Error GoTo 0
	%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>휴폐업조회 - 대량</legend>
				<br/>
				<p class="info">> state (휴폐업상태) : null-알수없음, 0-등록되지 않은 사업자번호, 1-사업중, 2-폐업, 3-휴업</p>
				<p class="info">> type (사업 유형) : null-알수없음, 1-일반과세자, 2-면세과세자, 3-간이과세자, 4-비영리법인, 국가기관</p>
				<br/>
			<%
				If Not IsEmpty(result) Then  
					For i=0 To result.Count-1
			%>
					<fieldset class="fieldset2">
						<legend>휴폐업조회 - 대량</legend>
						<ul>
								<li>사업자번호(corpNum) : <%= result.Item(i).corpNum%></li>		
								<li>휴폐업상태(state) : <%= result.Item(i).state%></li>
								<li>사업자유형(type) : <%= result.Item(i).ctype%></li>	
								<li>휴폐업일자(stateDate) : <%= result.Item(i).stateDate%></li>	
								<li>국세청 확일일자(checkDate) : <%= result.Item(i).checkDate%></li>	
						</ul>
					</fieldset>
			<%
					Next
				End If 
				If Not IsEmpty(code) then
			%>
				<fieldset class="fieldset2">
					<legend>휴폐업조회 - 단건</legend>
					<ul>
						<li>Response.code : <%= code %> </li>
						<li>Response.message : <%= message %></li>
					</ul>
				</fieldset>
			<%
				End If
			%>		

			</fieldset>

		<script type ="text/javascript">
			 window.onload=function(){
				 document.getElementById('CorpNum').focus();
			 }
			 
			 function search(){
				document.getElementById('corpnum_form').submit();
			 }		 
		 </script>
	</body>
</html>