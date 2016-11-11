<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		
		<title>팝빌 SDK ASP Example.</title>
	</head>

	<body>

		<div id="content">

			<p class="heading1">팝빌 홈택스 현금영수증 연계 API SDK Example.</p>
			
			<br/>

			<fieldset class="fieldset1">
				<legend>팝빌 기본 API</legend>

				<fieldset class="fieldset2">
					<legend>회원사 정보</legend>
					<ul>					
						<li><a href="checkIsMember.asp">checkIsMember</a> - 연동회원 가입여부 확인</li>
						<li><a href="checkID.asp">checkID</a> - 아이디 중복확인</li>
						<li><a href="joinMember.asp">joinMember</a> - 연동회원 가입 요청</li>
						<li><a href="getChargeInfo.asp">getChargeInfo</a> - 과금정보 확인</li>
						<li><a href="getBalance.asp">getBalance</a> - 연동회원 잔여포인트 확인</li>
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
				<legend>홈택스 연계 관련 API</legend>
				
				<fieldset class="fieldset2">
					<legend>매출/매입 내역 수집</legend>
					<ul>
						<li><a href="requestJob.asp">requestJob</a> - 수집 요청</li>
						<li><a href="getJobState.asp">getJobState</a> - 수집 상태 확인</li>
						<li><a href="listActiveJob.asp">listActiveJob</a> - 수집 상태 목록 확인</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>매출/매입 수집 결과 조회</legend>
					<ul>
						<li><a href="search.asp">search</a> - 수집 결과 조회</li>
						<li><a href="summary.asp">summary</a> - 수집 결과 요약정보 조회</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>부가 기능</legend>
					<ul>
						<li><a href="getFlatRatePopUpURL.asp">getFlatRatePopUpURL</a> - 정액제 서비스 신청 URL</li>
						<li><a href="getFlatRateState.asp">getFlatRateState</a> - 정액제 서비스 상태 확인</li>
						<li><a href="getCertificatePopUpURL.asp">getCertificatePopUpURL</a> - 홈택스연계 공인인증서 등록 URL</li>
						<li><a href="getCertificateExpireDate.asp">getCertificateExpireDate</a> - 홈택스연계 공인인증서 만료일자 확인</li>
					</ul>
				</fieldset>				
			</fieldset>
		 </div>
	</body>
</html>