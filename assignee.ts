// ***「担当者決めボタン」を押すとナレシェア担当者を2名抽選するスクリプト ***
const activeSheet = SpreadsheetApp.getActiveSpreadsheet()!;

type Data = { name: string; row: number; date: Date };

/** 現在時刻を返す。GAS上でズレが生じることを避けるため分、秒、ミリ秒に 0 を設定 */
function getDate() {
  const date = new Date();
  date.setMinutes(0);
  date.setSeconds(0);
  date.setMilliseconds(0);
  return date;
}

/**
 * ナレシェア担当者の最終担当日リストを取得
 * - 直近のナレシェア担当者から「★」を外し現在日時を設定
 * - 新規参画者（現在日時なしのメンバー）には現在日時を設定
 */
function getDataList() {
  const { nameColumn, assigneeColumn, dateColumn, startRow, assignedMark } = GLOBAL_SETTINGS;
  const endRow = activeSheet.getLastRow();
  const dateList: Data[] = [];
  for (let i = startRow; i <= endRow; i++) {
    const nameValue: string = activeSheet.getRange(`${nameColumn}${i}`).getValue();
    if (!nameValue) break; // 名前が入っていない場合は終了

    const assigneeCell = activeSheet.getRange(`${assigneeColumn}${i}`);
    const staredValue: string = assigneeCell.getValue();

    const dateCell = activeSheet.getRange(`${dateColumn}${i}`);
    let dateValue: Date | '' = dateCell.getValue();

    // '★'がある場合直近担当者のため現在日時を挿入
    if (staredValue === assignedMark) {
      const now = getDate();
      dateCell.setValue(now);
      dateValue = now;
      assigneeCell.setValue('');
    }

    // 日付がない場合現在日時を挿入（新規参画者想定）
    if (dateValue === '') {
      const now = getDate();
      dateCell.setValue(now);
      dateValue = now;
    }
    dateList.push({ name: nameValue, row: i, date: dateValue });
  }

  // Dateでソートして返す
  return dateList.sort((a, b) => a.date.getTime() - b.date.getTime());
}

/** リストの中から最も古い日付の要素2つを返す。日付が同じ場合はランダムで選択 */
function getOldTow(sortedDataList: Data[]) {
  const uniqueDate = Array.from(new Set(sortedDataList.map((obj) => obj.date)));

  const assignees: Data[] = [];
  for (const date of uniqueDate) {
    const sameNumObjects = sortedDataList.filter((obj) => obj.date.getTime() === date.getTime());
    while (assignees.length < 2 && sameNumObjects.length > 0) {
      const randomIndex = Math.floor(Math.random() * sameNumObjects.length);
      assignees.push(sameNumObjects[randomIndex]);
      sameNumObjects.splice(randomIndex, 1);
    }
    if (assignees.length === 2) break;
  }

  return assignees;
}

/** 抽選された担当者に「★」マークをつける */
function setAssignee(assigneesList: Data[]) {
  const { assigneeColumn, assignedMark } = GLOBAL_SETTINGS;

  const assignMemberNames = assigneesList.map((data) => {
    activeSheet.getRange(`${assigneeColumn}${data.row}`).setValue(assignedMark).setFontColor('#e69138');
    return data.name;
  });

  // return assignMemberNames;
}

function run() {
  if (!activeSheet) throw new Error('シートが見つかりません');
  const sortedDataList = getDataList();
  const assignees = getOldTow(sortedDataList);
  setAssignee(assignees);
}
