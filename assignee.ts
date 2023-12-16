const activeSheet = SpreadsheetApp.getActiveSpreadsheet()!;

type Data = { name: string; row: number; date: Date };

function test(value: string | number | Date, num = 37) {
  activeSheet.getRange(`G${num}`).setValue(value);
}

function getDate() {
  const date = new Date();
  date.setMinutes(0);
  date.setSeconds(0);
  date.setMilliseconds(0);
  return date;
}

/**
 * - リストの中から最も古い日付の要素2つを返す
 * - 日付が同じ場合はランダムで選択
 */
function getOldTow(sortedDataList: Data[]) {
  const uniqueDate = Array.from(new Set(sortedDataList.map((obj) => obj.date)));

  const chosen: Data[] = [];
  for (const date of uniqueDate) {
    const sameNumObjects = sortedDataList.filter((obj) => obj.date.getTime() === date.getTime());
    while (chosen.length < 2 && sameNumObjects.length > 0) {
      const randomIndex = Math.floor(Math.random() * sameNumObjects.length);
      chosen.push(sameNumObjects[randomIndex]);
      sameNumObjects.splice(randomIndex, 1);
    }
    if (chosen.length === 2) break;
  }

  return chosen;
}

function getDataList() {
  const { nameColumn, assigneeColumn, dateColumn, startRow } = GLOBAL_SETTINGS;
  const endRow = activeSheet.getLastRow();
  const dateList: Data[] = [];
  for (let i = startRow; i <= endRow; i++) {
    const nameCell = activeSheet.getRange(`${nameColumn}${i}`);
    const nameValue: string = nameCell.getValue();
    if (!nameValue) break; // 名前が入っていない場合は終了

    const assigneeCell = activeSheet.getRange(`${assigneeColumn}${i}`);
    const staredValue: string = assigneeCell.getValue();

    const dateCell = activeSheet.getRange(`${dateColumn}${i}`);
    let dateValue: Date | '' = dateCell.getValue();

    // '★'がある場合直近担当者のため現在日時を挿入
    if (staredValue === '★') {
      const date = getDate();
      dateCell.setValue(date);
      dateValue = date;
      assigneeCell.setValue('');
    }

    // 日付がない場合現在日時を挿入（新規参画者想定）
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
    activeSheet.getRange(`${assigneeColumn}${data.row}`).setValue('★').setFontColor('#e69138');
    return data.name;
  });

  return assignMemberNames;
}

function run() {
  if (!activeSheet) throw new Error('シートが見つかりません');
  const sortedDataList = getDataList();
  const assignees = getOldTow(sortedDataList);
  const assignMemberNames = setAssignee(assignees);
}
