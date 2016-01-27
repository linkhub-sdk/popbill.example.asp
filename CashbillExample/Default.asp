<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		
		<title>팝빌 SDK ASP Example.</title>
	</head>
	<body>
		<div id="content">
			
			<p class="heading1">팝빌 현금영수증 SDK ASP Example.</p>

			<br/>

			<fieldset class="fieldset1">
				<legend>팝빌 기본 API</legend>

				<fieldset class="fieldset2">
					<legend>회원정보</legend>
					<ul>
						<li><a href="checkIsMember.asp">checkCorpIsMember</a> - 연동회원사 가입 여부 확인</li>
						<li><a href="checkID.asp">checkID</a> - 아이디 중복확인</li>
						<li><a href="joinMember.asp">joinMember</a> - 연동회원사 가입 요청</li>
						<li><a href="getBalance.asp">getBalance</a> - 연동회원사 잔여포인트 확인</li>
						<li><a href="getPartnerBalance.asp">getPartnerBalance</a> - 파트너 잔여포인트 확인</li>
						<li><a href="getPopbillURL.asp">getPopbillURL</a> - 팝빌 SSO URL 요청</li>
						<li><a href="listContact.asp">listContact</a> - 담당자 목록 조회</li>
						<li><a href="updateContact.asp">updateContact</a> - 담당자 정보 수정</li>
						<li><a href="registContact.asp">registContact</a> - 담당자 추가</li>
						<li><a href="updateCorpInfo.asp">updateCorpInfo</a> - 회사정보 수정</li>
						<li><a href="getCorpInfo.asp">getCorpInfo</a> - 회사정보 확인</li>
					</ul>
				</fieldset>

			</fieldset>
			
			<br />
			
			<fieldset class="fieldset1">
				<legend>현금영수증 관련 API</legend>
				
				<fieldset class="fieldset2">
					<legend>등록/수정/발행/삭제</legend>
					<ul>
						<li><a href="checkMgtKeyInUse.asp">checkMgtKeyInUse</a> - 연동관리번호 사용여부 확인</li>
						<li><a href="registIssue.asp">registIssue</a> - 현금영수증 즉시발행</li>
						<li><a href="register.asp">register</a> - 현금영수증 임시저장</li>
						<li><a href="update.asp">update</a> - 현금영수증 수정</li>
						<li><a href="issue.asp">issue</a> - 현금영수증 발행</li>
						<li><a href="cancelIssue.asp">cancelIssue</a> - 현금영수증 발행취소</li>
						<li><a href="delete.asp">delete</a> - 현금영수증 삭제</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>정보 확인</legend>
					<ul>
						<li><a href="getInfo.asp">getInfo</a> - 현금영수증 상태확인</li>
						<li><a href="getInfos.asp">getInfos</a> - 현금영수증 상태 대량 확인</li>
						<li><a href="getLogs.asp">getLogs</a> - 현금영수증 이력 확인</li>
						<li><a href="getDetailInfo.asp">getDetailInfo</a> - 현금영수증 상세정보 확인</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>부가기능</legend>
					<ul>
						<li><a href="sendEmail.asp">sendEmail</a> - 알림메일 재전송</li>
						<li><a href="sendSMS.asp">sendSMS</a> - 알림문자 재전송</li>
						<li><a href="sendFAX.asp">sendFAX</a> - 팩스 전송</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>팝빌 현금영수증 SSO URL 기능</legend>
					<ul>
						<li><a href="getURL.asp">getURL</a> - 현금영수증 관련 SSO URL 확인</li>
						<li><a href="getPopUpURL.asp">getPopUpURL</a> - 해당 현금영수증의 팝빌 화면을 표시하는 URL 확인</li>
						<li><a href="getPrintURL.asp">getPrintURL</a> - 해당 현금영수증의 팝빌 인쇄 화면을 표시하는 URL 확인</li>
						<li><a href="getEPrintURL.asp">getEPrintURL</a> - 해당 현금영수증의 팝빌 인쇄 화면을 표시하는 URL 확인(공급받는자용)</li>
						<li><a href="getMassPrintURL.asp">getMassPrintURL</a> - 다량(최대100건)의 현금영수증 인쇄 화면을 표시하는 URL 확인 </li>
						<li><a href="getMailURL.asp">getMailURL</a> - 해당 현금영수증의 전송메일상의 "보기" 버튼에 해당하는 URL 확인 </li>
					</ul>
				</fieldset>
				<fieldset class="fieldset2">
					<legend>기타</legend>
					<ul>
						<li><a href="getUnitCost.asp">getUnitCost</a> - 현금영수증 발행단가 확인</li>
					</ul>
				</fieldset>

			</fieldset>
		 </div>
	</body>
</html>