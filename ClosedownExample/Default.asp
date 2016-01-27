<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		
		<title>팝빌 SDK ASP Example.</title>
	</head>
	<body>
		<div id="content">
			
			<p class="heading1">팝빌 휴폐업조회 SDK ASP Example.</p>

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
				<legend>휴폐업조회 관련 API</legend>
				
				<fieldset class="fieldset2">
					<legend>단건/대량 조회</legend>
					<ul>
						<li><a href="checkCorpNum.asp">checkCorpNum</a> -  휴폐업조회(단건)</li>
						<li><a href="checkCorpNums.asp">checkCorpNums</a> - 휴폐업조회(대량)</li>	
					</ul>
				</fieldset>
			</fieldset>
		 </div>
	</body>
</html>