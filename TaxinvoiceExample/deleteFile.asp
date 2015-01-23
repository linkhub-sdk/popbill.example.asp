<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"	'회원 사업자번호, "-" 제외
	testUserID = "testkorea"    '회원 아이디
	KeyType= "SELL"             '발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	MgtKey = "20150121-07"      '연동관리번호 
	FileID = "5131AACD-9D35-4CCE-BAC7-4943653FB002.PBF "   '첨부파일 목록(getFiles) AttachedFile 값 참조. 

	On Error Resume Next
	
	Set Presponse = m_TaxinvoiceService.DeleteFile(testCorpNum, KeyType ,MgtKey, FileID, testUserID)
	
	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else 
		code = Presponse.code
		message = Presponse.message
	End If
	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>세금계산서 첨부파일 1개 삭제</legend>
					<ul>
						<li>Response.code : <%=code%> </li>
						<li>Response.message : <%=message%> </li>
					</ul>									
			</fieldset>
		 </div>
	</body>
</html>