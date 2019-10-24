<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 전자명세서 1건의 상세정보를 조회합니다.
	' - 응답항목에 대한 자세한 사항은 "[전자명세서 API 연동매뉴얼] > 4.1.
	'   전자명세서 구성" 을 참조하시기 바랍니다.
	'**************************************************************

	'팝빌 회원 사업자번호, "-"제외 10자리
	testCorpNum = "1234567890"

	'팝빌 회원 아이디
	userID = "testkorea"

	'명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서), 124(발주서), 125(입금표), 126(영수증)
	itemCode = "121"					

	'문서번호
	mgtKey = "20191024-023"				

	On Error Resume Next
	
	Set result = m_StatementService.GetDetailInfo(testCorpNum, itemCode, mgtKey, userID)
	
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
				<legend>전자명세서 상세정보</legend>
				<ul>
					<% If code = 0 Then %>
						 <li> itemCode(문서종류코드) : <%=result.itemCode%> </li>
						<li> mgtKey(관리번호) : <%=result.mgtKey%> </li>
						<li> invoiceNum(팝빌부여 문서고유번호) : <%=result.invoiceNum%> </li>
						<li> formCode(맞춤양식 코드) : <%=result.formCode%> </li>
						<li> writeDate(작성일자) : <%=result.writeDate%> </li>
						<li> taxType(세금형태) : <%=result.taxType %> </li>
						<li> senderCorpNum(발신자 사업자번호) : <%=result.senderCorpNum%> </li>
						<li> senderTaxRegID(발신자 종사업장번호) : <%=result.senderTaxRegID%> </li>
						<li> senderCorpName(발신자 상호) : <%=result.senderCEOName%> </li>
						<li> senderCEOName(발신자 대표자성명) : <%=result.senderCEOName%> </li>
						<li> senderAddr(발신자 주소) : <%=result.senderAddr%> </li>
						<li> senderBizClass(발신자 종목) : <%=result.senderBizClass%> </li>
						<li> senderBizType(발신자 업태) : <%=result.senderBizType%> </li>
						<li> senderContactName(발신자 담당자명) : <%=result.senderContactName%> </li>
						<li> senderTEL(발신자 연락처) : <%=result.senderTEL%> </li>
						<li> senderHP(발신자 휴대폰번호) : <%=result.senderHP%> </li>
						<li> senderEmail(발신자 메일주소) : <%=result.senderEmail%> </li>
						<li> receiverCorpNum(수신자 사업자번호) : <%=result.receiverCorpNum%> </li>
						<li> receiverTaxRegID(수신자 종사업장번호) : <%=result.receiverTaxRegID%> </li>
						<li> receiverCorpName(수신자 상호) : <%=result.receiverCorpName%> </li>
						<li> receiverCEOName(수신자 대표자성명) : <%=result.receiverCEOName%> </li>
						<li> receiverAddr(수신자 주소) : <%=result.receiverAddr%> </li>
						<li> receiverBizClass(수신자 종목) : <%=result.receiverBizClass%> </li>
						<li> receiverBizType(수신자 업태) : <%=result.receiverBizType%> </li>
						<li> receiverContactName(수신자 담당자명) : <%=result.receiverContactName%> </li>
						<li> receiverTEL(수신자 연락처) : <%=result.receiverTEL%> </li>
						<li> receiverHP(수신자 휴대폰번호) : <%=result.receiverHP%> </li>
						<li> receiverEmail(수신자 메일주소) : <%=result.receiverEmail%> </li>
						<li> totalAmount(합계금액) : <%=result.totalAmount%> </li>
						<li> supplyCostTotal(공급가액 합계) : <%=result.supplyCostTotal%> </li>
						<li> taxTotal(세액 합계) : <%=result.taxTotal%> </li>
						<li> purposeType(영수/청구) : <%=result.purposeType%> </li>
						<li> serialNum(기재상 일련번호) : <%=result.serialNum%> </li>
						<li> remark1(비고1) : <%=result.remark1%> </li>
						<li> remark2(비고2) : <%=result.remark2%> </li>
						<li> remark3(비고3) : <%=result.remark3%> </li>
						<li> businessLicenseYN(사업자등록증 첨부여부) : <%=result.businessLicenseYN%> </li>
						<li> bankBookYN(통장사본 첨부여부) : <%=result.bankBookYN%> </li>
						<li> smssendYN(알림문자 전송여부) : <%=result.smssendYN%> </li>
						<li> autoacceptYN(발행시 자동승인 여부) : <%=result.autoacceptYN%> </li>
						
						<!--기타 상세항목 생략-->

						<fieldset class="fieldset2">
							<legend>추가속성</legend>
							<ul>
							<% For Each propertyKey In result.propertyBag.keys() %>
								<li> <%=propertyKey%> : <%=result.propertyBag.get(propertyKey)%></li>
							<% Next %>
							</ul>
						</fieldset>
						<% For i=0 To Ubound(result.detailList)-1%>
								<fieldset class="fieldset2">
								<legend> 상세항목 <%=i+1%> </legend>
									<ul>
										<li> serialNum(일련번호) : <%=result.detailList(i).serialNum%> </li>
										<li> purchaseDT(거래일자) : <%=result.detailList(i).purchaseDT%> </li>
										<li> itemName(품목명) : <%=result.detailList(i).itemName%> </li>
										<li> spec(규격) : <%=result.detailList(i).spec%> </li>
										<li> qty(수량) : <%=result.detailList(i).qty%> </li>
										<li> unitCost(단가) : <%=result.detailList(i).unitCost%> </li>
										<li> supplyCost(공급가액) : <%=result.detailList(i).supplyCost%> </li>
										<li> tax(세액) : <%=result.detailList(i).tax%> </li>
										<li> remark(비고) : <%=result.detailList(i).remark%> </li>
										<li> spare1(여분1) : <%=result.detailList(i).spare1%> </li>
										<li> spare2(여분2) : <%=result.detailList(i).spare2%> </li>
										<li> spare3(여분3) : <%=result.detailList(i).spare3%> </li>
										<li> spare4(여분4) : <%=result.detailList(i).spare4%> </li>
										<li> spare5(여분5) : <%=result.detailList(i).spare5%> </li>
										<li> spare6(여분6) : <%=result.detailList(i).spare6%> </li>
										<li> spare7(여분7) : <%=result.detailList(i).spare7%> </li>
										<li> spare8(여분8) : <%=result.detailList(i).spare8%> </li>
										<li> spare9(여분9) : <%=result.detailList(i).spare9%> </li>
										<li> spare10(여분10) : <%=result.detailList(i).spare10%> </li>
										<li> spare11(여분11) : <%=result.detailList(i).spare11%> </li>
										<li> spare12(여분12) : <%=result.detailList(i).spare12%> </li>
										<li> spare13(여분13) : <%=result.detailList(i).spare13%> </li>
										<li> spare14(여분14) : <%=result.detailList(i).spare14%> </li>
										<li> spare15(여분15) : <%=result.detailList(i).spare15%> </li>
										<li> spare16(여분16) : <%=result.detailList(i).spare16%> </li>
										<li> spare17(여분17) : <%=result.detailList(i).spare17%> </li>
										<li> spare18(여분18) : <%=result.detailList(i).spare18%> </li>
										<li> spare19(여분19) : <%=result.detailList(i).spare19%> </li>
										<li> spare20(여분20) : <%=result.detailList(i).spare20%> </li>
									</ul>
								</fieldset>
							<% 
								Next
								Else
							%>
		
							<li>Response.code : <%=code%> </li>
							<li>Response.message: <%=message%> </li>
						<% 
							End If
						%>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>