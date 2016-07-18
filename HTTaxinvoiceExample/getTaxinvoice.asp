<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 

<%
	' 연동회원 사업자번호
	testCorpNum = "1234567890"

	' 전자세금계산서 국세청승인번호 
	NTSConfirmNum = "20160714410000290000083e"
	
	' 연동회원 아이디 
	UserID = "testkorea"				 
	
	On Error Resume Next

	Set result = m_HTTaxinvoiceService.GetTaxinvoice(testCorpNum, NTSConfirmNUm, UserID)
	
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
				<% If code = 0 Then %>
				<legend>상세정보 조회</legend>
				<ul>
					<li>ntsconfirmNum (국세청승인번호) : <%=result.ntsconfirmNum%> </li>
					<li>writeDate (작성일자) : <%=result.writeDate%> </li>
					<li>issueDT (발행일시) : <%=result.issueDT%> </li>
					<li>invoiceType (전자세금계산서 종류) : <%=result.invoiceType%> </li>
					<li>taxType (과세형태) : <%=result.taxType%> </li>
					<li>taxTotal (세액 합계) : <%=result.taxTotal%> </li>
					<li>supplyCostTotal (공급가액 합계) : <%=result.supplyCostTotal%> </li>
					<li>totalAmount (합계금액) : <%=result.totalAmount%> </li>
					<li>purposeType (영수/청구) : <%=result.purposeType%> </li>
					<li>serialNum (일련번호) : <%=result.serialNum%> </li>
					<li>cash (현금) : <%=result.cash%> </li>
					<li>chkBill (수표) : <%=result.chkBill%> </li>
					<li>credit (외상) : <%=result.credit%> </li>
					<li>note (어음) : <%=result.note%> </li>
					<li>remark1 (비고1) : <%=result.remark1%> </li>
					<li>remark2 (비고2) : <%=result.remark2%> </li>
					<li>remark3 (비고3) : <%=result.remark3%> </li>

					<li>modifyCode (수정 사유코드 ) : <%=result.modifyCode%> </li>
					<li>orgNTSConfirmNum (원본 전자세금계산서 국세청승인번호) : <%=result.orgNTSConfirmNum%> </li>

					<li>invoicerCorpNum (공급자 사업자번호) : <%=result.invoicerCorpNum%> </li>
					<li>invoicerMgtKey (공급자 관리번호) : <%=result.invoicerMgtKey%> </li>
					<li>invoicerTaxRegID (공급자 종사업장번호 ) : <%=result.invoicerTaxRegID%> </li>
					<li>invoicerCorpName (공급자 상호) : <%=result.invoicerCorpName%> </li>
					<li>invoicerCEOName (공급자 대표자성명) : <%=result.invoicerCEOName%> </li>
					<li>invoicerAddr (공급자 주소) : <%=result.invoicerAddr%> </li>
					<li>invoicerBizType (공급자 업태) : <%=result.invoicerBizType%> </li>
					<li>invoicerBizClass (공급자 종목) : <%=result.invoicerBizClass%> </li>
					<li>invoicerContactName (공급자 담당자 성명) : <%=result.invoicerContactName%> </li>
					<li>invoicerTEL (공급자 연락처) : <%=result.invoicerTEL%> </li>
					<li>invoicerEmail (공급자 이메일) : <%=result.invoicerEmail%> </li>

					<li>invoiceeCorpNum (공급받는자 사업자번호) : <%=result.invoiceeCorpNum%> </li>
					<li>invoiceeType (공급받는자 구분) : <%=result.invoiceeType%> </li>
					<li>invoiceeMgtKey (공급받는자 관리번호) : <%=result.invoiceeMgtKey%> </li>
					<li>invoiceeTaxRegID (공급받는자 종사업장번호) : <%=result.invoiceeTaxRegID%> </li>
					<li>invoiceeCorpName (공급받는자 상호) : <%=result.invoiceeCorpName%> </li>
					<li>invoiceeCEOName (공급받는자 대표자성명) : <%=result.invoiceeCEOName%> </li>
					<li>invoiceeAddr (공급받는자 주소) : <%=result.invoiceeAddr%> </li>
					<li>invoiceeBizType (공급받는자 업태) : <%=result.invoiceeBizType%> </li>
					<li>invoiceeBizClass (공급받는자 종목) : <%=result.invoiceeBizClass%> </li>
					<li>invoiceeContactName1 (공급받는자 담당자 성명) : <%=result.invoiceeContactName1%> </li>
					<li>invoiceeTEL1 (공급받는자 담당자 연락처) : <%=result.invoiceeTEL1%> </li>
					<li>invoiceeEmail1 (공급받는자 담당자 이메일) : <%=result.invoiceeEmail1%> </li>
				</ul>
				<fieldset class="fieldset2">	
				<%
					For i=0 To UBound(result.detailList) -1 
				%>
						<legend>품목정보 [<%=i+1%>]</legend>
						<ul>
							<li> serialNum (일련번호) : <%= result.detailList(i).serialNum %></li>
							<li> purchaseDT (거래일자) : <%= result.detailList(i).purchaseDT %></li>
							<li> itemName (품명) : <%= result.detailList(i).itemName %></li>
							<li> spec (규격) : <%= result.detailList(i).spec %></li>
							<li> qty (수량) : <%= result.detailList(i).qty %></li>
							<li> unitCost (단가) : <%= result.detailList(i).unitCost %></li>
							<li> supplyCost (공급가액) : <%= result.detailList(i).supplyCost %></li>
							<li> tax (세액) : <%= result.detailList(i).tax %></li>
							<li> remark (비고) : <%= result.detailList(i).remark %></li>
						</ul>
				<%
						Next					
				%>
				</fieldset>
				<%	Else  %>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	End If	%>
			</fieldset>
		 </div>
	</body>
</html>