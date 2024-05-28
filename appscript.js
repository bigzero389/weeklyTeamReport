function reqWeeklyTeamReport() {
  const senderMailAddress = "dyheo@hist.co.kr";
  const teamMailAddress = "selhkt@hist.co.kr";
  // const teamMailAddress = senderMailAddress;

  const today = new Date(); // 오늘
  const tomorrow = new Date(today.setDate(today.getDate() + 1)); // 내일
  const dayAfterTomorrow = new Date(today.setDate(today.getDate() + 2)); // 모레

  console.log("Scheduler START : ${today}");

  // 오늘부터 뒤에 2일동안 휴일을 카운트해서 근무일이 1일 이상인 경우만 메일발송
  const isAfterdayOff = isHoliday(tomorrow)
    ? 1
    : 0 + isHoliday(dayAfterTomorrow)
      ? 1
      : 0;
  if (isHoliday(today)) {
    // 수요일이 공휴일이면
    if (isAfterdayOff == 2) {
      // 목금 둘다 공휴일이 이면
      sendWeeklyReportEmail(
        tomorrow,
        senderMailAddress,
        "",
        "(목금요일 공휴일 이므로 주간보고 의사결정필요)",
      );
    }
    if (isAfterdayOff < 2) {
      // 목금 중 업무일자 1일 이상
      sendWeeklyReportEmail(tomorrow, teamMailAddress, "", "");
    }
  } else if (!isHoliday(today)) {
    // 수요일이 공휴일이 아니면
    if (isAfterdayOff == 2) {
      // 목금 둘다 공휴일이 이면
      sendWeeklyReportEmail(
        tomorrow,
        teamMailAddress,
        "",
        " 목/금요일 공휴일 이므로 금일중으로 작성요망",
      );
    }
    if (isAfterdayOff < 2) {
      // 목금 중 업무일자 1일 이상
      sendWeeklyReportEmail(tomorrow, teamMailAddress, "", "");
    }
  } else {
    console.info("It's not correct day");
  }
}

// 지난주 주간업무보고 백업, 일자수정 및 작성요청 메일발송
const sendWeeklyReportEmail = (
  targetDate,
  mailAddress,
  commentAdded,
  subjectAdded,
) => {
  //selhkt@hist.co.kr

  // 지난주 주간업무보고 백업
  const weeklyReportFile = SpreadsheetApp.getActiveSpreadsheet(); // 현재 공유된 클라우드팀 주간업무보고 파일
  const files = DriveApp.getFilesByName(weeklyReportFile.getName());
  while (files.hasNext()) {
    const file = files.next();
    const destination = DriveApp.getFolderById(
      "1i7gwUjOoL1Jn5jH5FR19C9RYgoOQkB9Z",
    ); // 클라우드팀 주간업무보고
    const copiedFile = file.makeCopy(
      weeklyReportFile.getName() + "(복사본)",
      destination,
    );
  }

  // 스프레드시트의 제일 앞에 시트 복사해서 맨뒤시트로 백업
  const backupSheetName = weeklyReportFile.getSheets()[0].getName() + "_temp";
  const newSheetId = SpreadsheetApp.openById(weeklyReportFile.getId());
  const newSheet = weeklyReportFile.getSheets()[0].copyTo(newSheetId);
  newSheet.setName(backupSheetName);

  // 스프레드시트 맨앞에 시트 이름변경 및 백업시트이름 _temp 제거
  const targetDateStrAddWeekdayStr =
    getDateStrKOR(targetDate, "MM/dd") +
    "(" +
    weekStr(targetDate.getDay()) +
    ")"; // MM/dd(weekday)
  weeklyReportFile.getSheets()[0].setName(targetDateStrAddWeekdayStr);
  newSheet.setName(backupSheetName.replace("_temp", ""));

  // 주간업무보고 제목 변경
  weeklyReportFile.rename(`[${targetDateStrAddWeekdayStr}] DT기술팀주간보고`);

  // 클라우드팀 메일 발송, 추가메일 발송이 필요한 경우 대비 배열 구성
  const emailAdressList = [mailAddress]; // selhkt@hist.co.kr, DT기술팀
  for (i = 0, j = emailAdressList.length; i < j; i++) {
    const subject =
      "[DT기술팀] " +
      targetDateStrAddWeekdayStr +
      " 주간업무보고 작성 요청건" +
      subjectAdded; // // 이메일을 제목
    const message =
      "허대영입니다. \n" +
      "이번주[" +
      targetDateStrAddWeekdayStr +
      "] 주간업무보고 작성바랍니다.  \n" +
      "문서 : https://docs.google.com/spreadsheets/d/1CUsisWBMvk04j92M5kxeAzFB2k3pcOdWJvY2zsD2U-8/edit?usp=sharing \n" +
      "\n " +
      commentAdded +
      "\n" +
      ">>이 메일은 AppScript 스케줄러에 의해서 발송되었습니다.<<";

    // 이메일을 보내는 영역
    MailApp.sendEmail(emailAdressList[i], subject, message);
  }
};

// 요일명 리턴
const weekStr = (weekNum) => {
  if (weekNum === 0) {
    return "일";
  } else if (weekNum === 1) {
    return "월";
  } else if (weekNum === 2) {
    return "화";
  } else if (weekNum === 3) {
    return "수";
  } else if (weekNum === 4) {
    return "목";
  } else if (weekNum === 5) {
    return "금";
  } else if (weekNum === 6) {
    return "토";
  } else {
    throw new Error("!!!WeekString Error");
  }
};

// 2024
const holidayArray = [
  "01-01",
  "02-09",
  "02-10",
  "02-11",
  "02-12",
  "03-01",
  "04-10",
  "05-01",
  "05-05",
  "05-06",
  "05-16",
  "06-06",
  "08-15",
  "09-16",
  "09-17",
  "09-18",
  "10-03",
  "10-09",
  "12-25",
];

const getDateStrKOR = (date, formatStr) => {
  const timeZone = Session.getScriptTimeZone();
  return Utilities.formatDate(date, timeZone, formatStr);
};

const isHoliday = (date) => {
  if (date) {
    // dateStr 기준으로 변경
    const dateStr = getDateStrKOR(date, "MM-dd");
    for (let i = 0, j = holidayArray.length; i < j; i++) {
      if (dateStr === holidayArray[i]) {
        return true;
      }
    }
  }
  return false;
};

// 휴가일자
/*
const isDayOff = (dateStr) => {
  switch(dateStr) {
    case '2023-01-01':
      return true;
    case '2023-12-25':
      return true;
    default:
      return false;
  }
}
*/
