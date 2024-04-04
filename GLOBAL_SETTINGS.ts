const GLOBAL_SETTINGS = {
  nameColumn: 'I', // メンバーの名前が入っている列
  facilitatorColumn: 'J', // ファシリテーターに★をつける列
  assigneeColumn: 'K', // ナレシェア担当者に★をつける列
  dateColumn: 'L', // 最後に担当した日付が入っている列
  startRow: 14, // 一人目のメンバーの行
  assignedMark: '★', // 担当者に★をつける
  channelId: 'C04ERSPG47Q', // general
  meetingTitle: '住まいFEG会', //! MEMO: スケジュールのタイトルが変更された場合反映すること

  // ここからtest
  // channelId: 'C06AAQ34DK5', // botお試しチャンネル
} as const;
