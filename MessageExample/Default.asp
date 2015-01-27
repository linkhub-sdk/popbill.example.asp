<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		
		<title>팝빌 SDK ASP Example.</title>
	</head>
	<body>
		<div id="content">
			
			<p class="heading1">팝빌 문자메시지 API SDK SpringMVC Example.</p>

			<br/>

			<fieldset class="fieldset1">
				<legend>팝빌 기본 API</legend>

				<fieldset class="fieldset2">
					<legend>회원사 정보</legend>
					<ul>
						<li><a href="checkIsMember.asp">checkIsMember</a> - 연동회원사 가입 여부 확인</li>
						<li><a href="joinMember.asp">joinMember</a> - 연동회원사 가입 요청</li>
						<li><a href="getBalance.asp">getBalance</a> - 연동회원사 잔여포인트 확인</li>
						<li><a href="getPartnerBalance.asp">getPartnerBalance</a> - 파트너 잔여포인트 확인</li>
						<li><a href="getPopbillURL.asp">getPopbillURL</a> - 팝빌 SSO URL 요청</li>
					</ul>
				</fieldset>

			</fieldset>
			
			<br />
			
			<fieldset class="fieldset1">
				<legend>문자메시지 관련 API</legend>
				
				<fieldset class="fieldset2">
					<legend>단문 문자 전송</legend>
					<ul>
						<li><a href="sendSMS.asp">sendSMS</a> - 단문 문자메시지 1건 전송</li>
						<li><a href="sendSMS_Multi.asp">sendSMS</a> - 단문 문자메시지 다량(최대1000건) 전송</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>장문 문자 전송</legend>
					<ul>
						<li><a href="sendLMS.asp">sendLMS</a> - 장문 문자메시지 1건 전송</li>
						<li><a href="sendLMS_Multi.asp">sendLMS</a> - 장문 문자메시지 다량(최대1000건) 전송</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>단/장문 문자 자동 전송</legend>
					<ul>
						<li><a href="sendXMS.asp">sendXMS</a> - 단/장문 문자메시지 1건 전송</li>
						<li><a href="sendXMS_Multi.asp">sendXMS</a> - 단/장문 문자메시지 다량(최대1000건) 전송</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>전송결과/예약취소</legend>
					<ul>
						<li><a href="getMessages.asp">getMessages</a> - 접수번호에 해당하는 문자메시지 전송결과 확인</li>
						<li><a href="cancelReserve.asp">cancelReserve</a> - 예약문자메시지의 예약 취소. 예약시간 10분전까지만 가능.</li>
					</ul>
				</fieldset>
				
				<fieldset class="fieldset2">
					<legend>기타</legend>
					<ul>
						<li><a href="getURL.asp">getURL</a> - 문자메시지 관련 SSO URL 확인</li>
						<li><a href="getUnitCost.asp">getUnitCost</a> - 세금계산서 발행 단가 확인</li>
					</ul>
				</fieldset>

			</fieldset>
		 </div>
	</body>
</html>