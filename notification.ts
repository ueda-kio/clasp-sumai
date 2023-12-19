const settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('G会進行担当')!;
const SLACK_API_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_API_TOKEN')!;
//! MEMO: 離任時にスクリプトプロパティを変更すること
const MAIL_ADDRESS = PropertiesService.getScriptProperties().getProperty('MAIL_ADDRESS')!;

function getAssignees() {
  const { startRow, nameColumn, assigneeColumn } = GLOBAL_SETTINGS;
  const endRow = settingSheet.getLastRow();
  const dataList = settingSheet.getRange(`${nameColumn}${startRow}:${assigneeColumn}${endRow}`).getValues();
  const assignees = dataList
    .map((data) => {
      const [name, _, assignee] = data;
      return assignee ? (name as string) : false;
    })
    .filter((name): name is Exclude<typeof name, false> => typeof name === 'string');
  return assignees;
}

function createMessage(assignees: string[]) {
  return `
  <@here>
  本日G会あります！
  ****************************
  本日のG会のナレシェア担当者は
  ${assignees.join('・')}
  です！よろしくお願いします！
  ****************************
  `.trim();
}

function postToSlack(text: string) {
  const { channelId } = GLOBAL_SETTINGS;
  UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    headers: { Authorization: `Bearer ${SLACK_API_TOKEN}` },
    payload: {
      channel: channelId,
      text,
    },
  });
}

function isHoldMeeting() {
  const { meetingTitle } = GLOBAL_SETTINGS;
  const calendar = CalendarApp.getCalendarById(MAIL_ADDRESS);
  const todaysEvents = calendar.getEventsForDay(new Date());
  return todaysEvents.some((event) => event.getTitle().includes(meetingTitle));
}

function notification() {
  if (!SLACK_API_TOKEN || !MAIL_ADDRESS) {
    throw new Error('SLACK_API_TOKEN or MAIL_ADDRESS is not defined');
  }

  const isMeeting = isHoldMeeting();
  if (!isMeeting) {
    Logger.log('本日は開催日ではありません');
    return;
  }

  const assignees = getAssignees();
  const message = createMessage(assignees);
  postToSlack(message);
}
