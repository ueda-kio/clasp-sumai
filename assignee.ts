const settingSheet = SpreadsheetApp.getActiveSpreadsheet()!;
const GLOBAL_SETTINGS = {
  nameColumn: 'I', // メンバーの名前が入っている列
  assigneeColumn: 'K', // 今日の担当者に★をつける列
  dateColumn: 'L', // 最後に担当した日付が入っている列
  startRow: 14, // 一人目のメンバーの行
} as const;

type Data = { name: string; row: number; date: Date };

function test(value: string | number | Date, num = 37) {
  settingSheet.getRange(`G${num}`).setValue(value);
}

const isInit = false;

function getDate() {
  const date = new Date();
  date.setMinutes(0);
  date.setSeconds(0);
  date.setMilliseconds(0);
  return date;
}

function getAssignees(sortedDateList: Data[]) {
  const uniqueDate = Array.from(new Set(sortedDateList.map((obj) => obj.date)));

  const chosen: Data[] = [];
  for (const date of uniqueDate) {
    const sameNumObjects = sortedDateList.filter((obj) => obj.date === date);
    while (chosen.length < 2 && sameNumObjects.length > 0) {
      const randomIndex = Math.floor(Math.random() * sameNumObjects.length);
      chosen.push(sameNumObjects[randomIndex]);
      sameNumObjects.splice(randomIndex, 1);
    }
    if (chosen.length === 2) break;
  }

  return chosen;
}

function getList() {
  const { nameColumn, assigneeColumn, dateColumn, startRow } = GLOBAL_SETTINGS;
  const endRow = settingSheet.getLastRow();
  const dateList: Data[] = [];
  for (let i = startRow; i <= endRow; i++) {
    const nameCell = settingSheet.getRange(`${nameColumn}${i}`);
    const nameValue: string = nameCell.getValue();
    if (!nameValue) break; // 名前が入っていない場合は終了

    const assigneeCell = settingSheet.getRange(`${assigneeColumn}${i}`);
    const staredValue: string = assigneeCell.getValue();

    const dateCell = settingSheet.getRange(`${dateColumn}${i}`);
    let dateValue: Date | '' = dateCell.getValue();

    if (staredValue === '★') {
      const date = getDate();
      dateCell.setValue(date);
      dateValue = date;
      assigneeCell.setValue('');
    }

    if (dateValue === '') {
      const date = getDate();
      dateCell.setValue(date);
      dateValue = date;
    }
    dateList.push({ name: nameValue, row: i, date: dateValue });
  }

  // Dateでソートして返す
  return dateList.sort((a, b) => a.date.getTime() - b.date.getTime());
}

function setAssignee(assigneesList: Data[]) {
  const { assigneeColumn } = GLOBAL_SETTINGS;

  const assignMemberNames = assigneesList.map((data) => {
    settingSheet.getRange(`${assigneeColumn}${data.row}`).setValue('★').setFontColor('#e69138');
    return data.name;
  });

  return assignMemberNames;
}

function init() {
  const { nameColumn, dateColumn, startRow } = GLOBAL_SETTINGS;
  const date = getDate();
  const endRow = 30;
  const range = settingSheet.getRange(`${dateColumn}${startRow}:${dateColumn}${endRow}`);
  const values: Date[][] = [];
  for (var i = 0; i < range.getHeight(); i++) {
    values.push([date]);
  }

  range.setValues(values);

  const names: string[][] = settingSheet.getRange(`${nameColumn}${startRow}:${nameColumn}${endRow}`).getValues();
  const data: Data[] = names.map((name, index) => {
    return { name: name[0], row: index + startRow, date };
  }, {});

  const assignees = getAssignees(data);
  const assignMemberNames = setAssignee(assignees);
}

function run() {
  if (!settingSheet) throw new Error('シートが見つかりません');
  if (isInit) {
    init();
  } else {
    const sortedDateList = getList();
    const assignees = getAssignees(sortedDateList);
    const assignMemberNames = setAssignee(assignees);
  }
}
