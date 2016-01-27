<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%

	testCorpNum = "1234567890"		 ' 팝빌 회원 사업자번호
	userID = "testkorea"					 ' 팝빌 회원 아이디
	memo = "즉시발행 메모"			 ' 메모 
	mgtKey = "20160126-10"			 ' 관리번호 

	Set newStatement = New Statement

    newStatement.writeDate = "20160126"             '필수, 기재상 작성일자
    newStatement.purposeType = "영수"               '필수, {영수, 청구}
    newStatement.taxType = "과세"                   '필수, {과세, 영세, 면세}
    newStatement.formCode = ""						'맞춤양식코드(기본값 "")
    
    newStatement.itemCode = "121"					'명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서) 124(발주서), 125(입금표), 126(영수증)
    
    newStatement.mgtKey = mgtKey
    
    newStatement.senderCorpNum = testCorpNum
    newStatement.senderTaxRegID = "" '종사업자 식별번호. 필요시 기재. 형식은 숫자 4자리.
    newStatement.senderCorpName = "공급자 상호"
    newStatement.senderCEOName = "공급자"" 대표자 성명"
    newStatement.senderAddr = "공급자 주소"
    newStatement.senderBizClass = "공급자 업종"
    newStatement.senderBizType = "공급자 업태,업태2"
    newStatement.senderContactName = "공급자 담당자명"
    newStatement.senderEmail = "test@test.com"
    newStatement.senderTEL = "070-7070-0707"
    newStatement.senderHP = "010-000-2222"
    
    newStatement.receiverCorpNum = "8888888888"
    newStatement.receiverCorpName = "공급받는자 상호"
    newStatement.receiverCEOName = "공급받는자 대표자 성명"
    newStatement.receiverAddr = "공급받는자 주소"
    newStatement.receiverBizClass = "공급받는자 업종"
    newStatement.receiverBizType = "공급받는자 업태"
    newStatement.receiverContactName = "공급받는자 담당자명"
    newStatement.receiverEmail = "test@receiver.com"
    
    newStatement.supplyCostTotal = "100000"      '필수 공급가액 합계
    newStatement.taxTotal = "10000"                  '필수 세액 합계
    newStatement.totalAmount = "110000"             '필수 합계금액.  공급가액 + 세액
    
    newStatement.serialNum = "123"
    newStatement.remark1 = "비고1"
    newStatement.remark2 = "비고2"
    newStatement.remark3 = "비고3"
    
    newStatement.businessLicenseYN = False		'사업자등록증 이미지 첨부시 설정.
    newStatement.bankBookYN = False				'통장사본 이미지 첨부시 설정.
    newStatement.faxsendYN = False				'발행시 Fax발송시 설정.
    newStatement.smssendYN = True				'발행시 문자발송기능 사용시 활용
	

	Set newDetail = New StatementDetail

    newDetail.serialNum = "1"             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20150110"   '거래일자  yyyyMMdd
    newDetail.itemName = "품명"
    newDetail.spec = "규격"
    newDetail.unit = "단위"
    newDetail.qty = "1" '수량           ' 소숫점 2자리까지 문자열로 기재가능
    newDetail.unitCost = "100000"       ' 소숫점 2자리까지 문자열로 기재가능
    newDetail.supplyCost = "100000"
    newDetail.tax = "10000"
    newDetail.remark = "비고"
    newDetail.spare1 = "spare1"
    newDetail.spare2 = "spare2"
    newDetail.spare3 = "spare3"
    newDetail.spare4 = "spare4"
    newDetail.spare5 = "spare5"

	newStatement.AddDetail newDetail
	
	Set newDetail = New StatementDetail

    newDetail.serialNum = "2"             '일련번호 1부터 순차 기재
    newDetail.purchaseDT = "20150112"   '거래일자  yyyyMMdd
    newDetail.itemName = "품명"
    newDetail.spec = "규격"
    newDetail.unit = "단위"
    newDetail.qty = "1" '수량           ' 소숫점 2자리까지 문자열로 기재가능
    newDetail.unitCost = "100000"       ' 소숫점 2자리까지 문자열로 기재가능
    newDetail.supplyCost = "100000"
    newDetail.tax = "10000"
    newDetail.remark = "비고"
    newDetail.spare1 = "spare1"
    newDetail.spare2 = "spare2"
    newDetail.spare3 = "spare3"
    newDetail.spare4 = "spare4"
    newDetail.spare5 = "spare5"

	newStatement.AddDetail newDetail
	

	'추가속성, 자세한사항은 전자명세서 API 연동매뉴얼 [5.부록 > 5.2 기본양식 추가속성 테이블] 참조.
	newStatement.propertyBag.Set "Balance", "150000"
	newStatement.propertyBag.Set "CBalance", "100000"

	On Error Resume Next

	Set result = m_StatementService.RegistIssue(testCorpNum, newStatement, memo, userID)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else
		code = result.code
		message = result.message
	End If

	On Error GoTo 0
	
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>전자명세서 즉시발행</legend>
				<ul>
					<li>Response.code : <%=code%> </li>
					<li>Response.message: <%=message%> </li>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>