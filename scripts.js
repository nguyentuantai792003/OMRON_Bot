// ======================================================================

// ======================================================
// CONFIG
// ======================================================

// ---------- Timezone ----------
const TIMEZONE = "Asia/Ho_Chi_Minh";


// ---------- Webhook URLs ----------
const WEBHOOK_URL_DEV = "https://chat.googleapis.com/v1/spaces/AAQAPKxUNbc/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=VTs4B8GwruElSSbhWOdLevsbqXrpeV2TEVxoX7K0i7g";
const WEBHOOK_URL_TEST = "https://chat.googleapis.com/v1/spaces/AAQA52lgp7c/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=HM3mZ0kDneCLrBIAZDb4jUJr1ziFJa40MswXAT51pOs";
const WEBHOOK_URL_ALL = "https://chat.googleapis.com/v1/spaces/AAQArnDwCvM/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=bz8KW8Z7QebfYw4zzCPR8gDIKWpV6kcV1CJBULZ1Xs8";
const WEBHOOK_URL_DEV_APP = "https://chat.googleapis.com/v1/spaces/AAQAmi4Ldn0/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=qgmhuD7QHrN5CNezGSGAjWumMuJHSEq2uz3_eHn6Oiw";


// ---------- Google Drive File IDs ----------
const QNA_FILE_ID = "11zTOcemaCTOqnLBdu9Baqbht8pzdee4g";                // file Q&A
const DAILY_REPORT_FILE_ID = "1kZqsdlQh766lVGgQLcyDlW0fU-9KzOCj";       // file daily report


// ---------- Sheet Names ----------
const QNA_SHEET_NAME = "質問票";
const TODO_SHEET_NAME = "Todo list";


// ---------- Google Sheet URLs ----------
const QNA_SHEET_URL = "https://docs.google.com/spreadsheets/d/11zTOcemaCTOqnLBdu9Baqbht8pzdee4g/edit?gid=1600679394#gid=1600679394";
const DAILY_REPORT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1kZqsdlQh766lVGgQLcyDlW0fU-9KzOCj/edit?gid=1648614001#gid=1648614001";


// ---------- Backlog ----------
const BACKLOG_API_KEY = "VbKwUKrcUBP0yvxQnwRS3cUvSyxwSOd2gmPdYfSr5rbkEPbyWbXVaYKOoH24EBea";
const BACKLOG_SPACE = "vti-corp";
const BACKLOG_PROJECT_ID = 763680;
const BACKLOG_ISSUE_TYPE_ID = 4108322;
const BACKLOG_PRIORITY_ID = 3;


// ---------- Trigger Settings ----------
const BACKLOG_TRIGGER_HOUR = 17;
const BACKLOG_TRIGGER_START_MINUTE = 0;
const BACKLOG_TRIGGER_END_MINUTE = 20;


// ---------- User Mapping ----------
const USER_MAP = {
  "PHUONG To Xuan (VTI.D8)": "users/112000589756579847344",
  "TAI Nguyen Tuan (VTI.D8)": "users/102503072226812970864",
  // "ANH Nguyen Thi (VTI.D8)": "users/112574408672368417117",
  "AN Nguyen Viet (VTI.D8)": "users/105071698862055191046",
  "DAT Vu Tien 2 (VTI.D8)": "users/110916778349951870683",
  // "THIEN Nguyen Xuan 1 (VTI.D8)": "users/xxxx",
  // "HUONG Do Thi Thu (VTI.D8)": "users/xxxx",
  // "CONG Nguyen Thanh (VTI.D8)": "users/xxxx"
};


// ---------- Excluded Users ----------
const EXCLUDED_USERS = [
  "THIEN Nguyen Xuan 1 (VTI.D8)",
  "HUONG Do Thi Thu (VTI.D8)",
  "CONG Nguyen Thanh (VTI.D8)",
  "ANH Nguyen Thi (VTI.D8)"
];


// ======================================================
// COMMON HELPERS
// ======================================================

function createTempSpreadsheetFromExcel(fileId, tempName) {
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();

  const tempFile = Drive.Files.create(
    {
      name: tempName,
      mimeType: "application/vnd.google-apps.spreadsheet"
    },
    blob
  );

  return tempFile.id;
}

function deleteTempFile(fileId) {
  DriveApp.getFileById(fileId).setTrashed(true);
}

function isSameDay(d1, d2) {
  return d1 &&
    d2 &&
    d1.getFullYear() === d2.getFullYear() &&
    d1.getMonth() === d2.getMonth() &&
    d1.getDate() === d2.getDate();
}

function parseDate(value) {
  if (!value) return null;

  if (value instanceof Date) {
    return value;
  }

  const parts = value.toString().split("/");

  if (parts.length !== 3) {
    return null;
  }

  return new Date(
    parts[2],
    parts[0] - 1,
    parts[1]
  );
}

function sendGoogleChat(webhookUrl, message) {
  UrlFetchApp.fetch(webhookUrl, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      text: message
    })
  });
}


// ======================================================
// DAILY Q&A REPORT
// ======================================================

function dailyQnAReport() {
  // ======================================================
  // 1. Skip weekend
  // ======================================================

  const now = new Date();
  const day = Number(
    Utilities.formatDate(now, TIMEZONE, "u")
  );

  // 1 = Monday ... 7 = Sunday
  if (day === 6 || day === 7) {
    console.log("Skip weekend (VN timezone)");
    return;
  }


  // ======================================================
  // 2. Open Q&A file
  // ======================================================

  const tempFileId = createTempSpreadsheetFromExcel(
    QNA_FILE_ID,
    "temp_qna"
  );

  const ss = SpreadsheetApp.openById(tempFileId);
  const sheet = ss.getSheetByName(QNA_SHEET_NAME);

  if (!sheet) {
    throw new Error(
      "Không tìm thấy sheet: " + QNA_SHEET_NAME
    );
  }


  // ======================================================
  // 3. Read data
  // ======================================================

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const col = {};

  headers.forEach((header, index) => {
    col[header] = index;
  });


  // ======================================================
  // 4. Group by status
  // ======================================================

  const stats = {
    notAnswered: [],   // 未回答
    received: [],      // 受付
    inProgress: [],    // 回答作成中
    needMoreInfo: [],  // 差戻し/追加質問
    answered: []       // 回答済
  };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const status = row[col["ステータス"]];
    const date = row[col["質問日"]];

    const item = {
      id: row[0],
      question: row[col["質問内容"]],
      date: new Date(date),
      rawDate: date
    };

    switch (status) {
      case "未回答":
        stats.notAnswered.push(item);
        break;

      case "受付":
        stats.received.push(item);
        break;

      case "回答作成中":
        stats.inProgress.push(item);
        break;

      case "差戻し/追加質問":
        stats.needMoreInfo.push(item);
        break;

      case "回答済":
      case "クローズ":
      case "キャンセル":
        stats.answered.push(item);
        break;

      default:
        stats.notAnswered.push(item);
    }
  }


  // ======================================================
  // 5. Calculate pending items
  // ======================================================

  const pending = [
    ...stats.notAnswered,
    ...stats.received,
    ...stats.inProgress,
    ...stats.needMoreInfo
  ];

  // Oldest first
  pending.sort((a, b) => a.date - b.date);

  const today = new Date();

  pending.forEach((item) => {
    item.days = Math.floor(
      (today - item.date) / (1000 * 60 * 60 * 24)
    );
  });

  const oldest =
    pending.length > 0
      ? pending[0]
      : null;


  // ======================================================
  // 6. Build message
  // ======================================================

  let msg = `📊 BÁO CÁO Q&A HÀNG NGÀY\n\n`;

  msg += `📌 Thống kê theo trạng thái:\n`;
  msg += `- ❗ Chưa trả lời (未回答): ${stats.notAnswered.length}\n`;
  msg += `- 📥 Đã nhận (受付): ${stats.received.length}\n`;
  msg += `- ⚙️ Đang xử lý (回答作成中): ${stats.inProgress.length}\n`;
  msg += `- ❓ Câu hỏi bổ sung (差戻し/追加質問): ${stats.needMoreInfo.length}\n`;
  msg += `- ✅ Đã trả lời (回答済): ${stats.answered.length}\n`;

  if (oldest) {
    msg += `\n🔴 Câu hỏi tồn lâu nhất (${oldest.days} ngày):\n`;
    msg += `- ${oldest.id}: ${oldest.question}\n`;
  }

  msg += `\n📋 Top 5 câu hỏi tồn lâu nhất:\n`;

  pending.slice(0, 5).forEach((item) => {
    const flag = item.days > 3 ? "🔥 " : "";
    msg += `- ${flag}${item.id} (${item.days} ngày)\n`;
  });

  msg += `\n🔗 Mở file Q&A:\n${QNA_SHEET_URL}`;


  // ======================================================
  // 7. Send Google Chat
  // ======================================================

  sendGoogleChat(WEBHOOK_URL_ALL, msg);
  sendGoogleChat(WEBHOOK_URL_TEST, msg);


  // ======================================================
  // 8. Delete temp file
  // ======================================================

  deleteTempFile(tempFileId);
}


// ======================================================
// DAILY DEV REPORT
// ======================================================

function getDailyReportData() {
  // ======================================================
  // 1. Get today
  // ======================================================

  const now = new Date();

  const today = new Date(
    Utilities.formatDate(
      now,
      TIMEZONE,
      "yyyy-MM-dd"
    )
  );


  // ======================================================
  // 2. Detect current month sheet
  // ======================================================

  const month = Number(
    Utilities.formatDate(
      now,
      TIMEZONE,
      "M"
    )
  );

  const year = Utilities.formatDate(
    now,
    TIMEZONE,
    "yyyy"
  );

  const sheetName = `T${month}${year}`;


  // ======================================================
  // 3. Open daily report file
  // ======================================================

  const tempFileId = createTempSpreadsheetFromExcel(
    DAILY_REPORT_FILE_ID,
    "temp_dev_report"
  );

  const ss = SpreadsheetApp.openById(tempFileId);
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(
      "Không tìm thấy sheet: " + sheetName
    );
  }


  // ======================================================
  // 4. Read data
  // ======================================================

  const data = sheet.getDataRange().getValues();


  // ======================================================
  // 5. Filter only today's records
  // ======================================================

  let currentDate = null;
  const reportByUser = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Column A = Date
    if (row[0]) {
      currentDate = parseDate(row[0]);
    }

    if (!isSameDay(currentDate, today)) {
      continue;
    }

    const user = row[2];       // Column C
    const reportNote = row[3]; // Column D

    if (!user) {
      continue;
    }


    // ======================================================
    // 6. Exclude specific users
    // ======================================================

    if (EXCLUDED_USERS.includes(user)) {
      continue;
    }

    reportByUser[user] = {
      note: reportNote
    };
  }


  // ======================================================
  // 7. Delete temp file
  // ======================================================

  deleteTempFile(tempFileId);


  // ======================================================
  // 8. Return result
  // ======================================================

  return {
    reportByUser,
    today
  };
}


// ======================================================
// CREATE BACKLOG ISSUE
// ======================================================

function createBacklogIssue(reportByUser, today) {
  // ======================================================
  // 1. Prepare Backlog API URL
  // ======================================================
  
  const url =
    `https://${BACKLOG_SPACE}.backlog.com/api/v2/issues` +
    `?apiKey=${BACKLOG_API_KEY}`;

  const formattedDate = Utilities.formatDate(
    today,
    TIMEZONE,
    "dd.MM.yyyy"
  );


  // ======================================================
  // 2. Prepare date labels
  // ======================================================

  const todayStr = Utilities.formatDate(
    today,
    TIMEZONE,
    "dd/MM"
  );

  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  const tomorrowStr = Utilities.formatDate(
    tomorrow,
    TIMEZONE,
    "dd/MM"
  );


  // ======================================================
  // 3. Build description from daily report
  // ======================================================

  let description = "";

  Object.keys(reportByUser).forEach((user) => {
    const userData = reportByUser[user];

    const parsed = parseNote(
      userData.note
    );

    description += `${user}\n`;

    // DONE
    description +=
      `Công việc đã done trong ngày ${todayStr}:\n`;

    if (parsed.today) {
      description += `${parsed.today}\n`;
    }

    // NEXT PLAN
    description +=
      `Task ưu tiên trong ngày ${tomorrowStr}:\n`;

    if (parsed.next) {
      description += `${parsed.next}\n`;
    }

    description += `\n`;
  });


  // ======================================================
  // 4. Append TODO list
  // ======================================================

  const todos = getTodoInProgress();

  if (todos.length > 0) {
    description += `TODO\n\n`;

    todos.forEach((todo) => {
      description += `${todo.description}\n`;

      if (todo.note) {
        description += `--> ${todo.note}\n`;
      }

      description += `\n`;
    });
  }


  // ======================================================
  // 5. Create payload
  // ======================================================

  const payload = {
    projectId: BACKLOG_PROJECT_ID,
    summary: `[DAILY MEETING] ${formattedDate}`,
    description: description,
    issueTypeId: BACKLOG_ISSUE_TYPE_ID,
    priorityId: BACKLOG_PRIORITY_ID
  };


  // ======================================================
  // 6. Call Backlog API
  // ======================================================

  UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  });
}


// ======================================================
// SEND DAILY REPORT MESSAGE
// ======================================================

function sendDailyChatReport() {
  // ======================================================
  // 1. Skip weekend
  // ======================================================

  const now = new Date();
  const day = Number(
    Utilities.formatDate(now, TIMEZONE, "u")
  );

  // 1 = Monday ... 7 = Sunday
  if (day === 6 || day === 7) {
    console.log("Skip weekend (VN timezone)");
    return;
  }


  // ======================================================
  // 2. Get report data
  // ======================================================

  const {
    reportByUser,
    today
  } = getDailyReportData();

  if (Object.keys(reportByUser).length === 0) {
    console.log("No data today");
    return;
  }


  // ======================================================
  // 3. Check who has reported
  // ======================================================

  function isReported(note) {
    if (!note) {
      return false;
    }

    const cleaned = note
      .toString()
      .replace(/Today:/gi, "")
      .replace(/Next plan:/gi, "")
      .replace(/Issue:/gi, "")
      .trim();

    return cleaned.length > 0;
  }


  // ======================================================
  // 4. Build message header
  // ======================================================

  const todayStr = Utilities.formatDate(
    today,
    TIMEZONE,
    "dd/MM/yyyy"
  );

  let msg =
    `📊 DAILY TASK DEV REPORT (${todayStr})\n\n`;

  let mentionText = "";


  // ======================================================
  // 5. Build report by user
  // ======================================================

  Object.keys(reportByUser).forEach((user) => {
    const userData = reportByUser[user];

    msg += `👤 ${user}\n`;

    if (!isReported(userData.note)) {
      msg += `⚠️ Chưa report\n`;

      if (USER_MAP[user]) {
        mentionText += `<${USER_MAP[user]}> `;
      }

    } else {
      msg += `${userData.note}\n`;
    }

    msg += `\n`;
  });


  // ======================================================
  // 6. Add mention section
  // ======================================================

  msg += `📣 Chưa report: ${mentionText}\n`;


  // ======================================================
  // 7. Add file link
  // ======================================================

  msg += `\n🔗 Mở file daily report:\n`;
  msg += `${DAILY_REPORT_SHEET_URL}`;


  // ======================================================
  // 8. Send Google Chat
  // ======================================================

  sendGoogleChat(WEBHOOK_URL_TEST, msg);
  sendGoogleChat(WEBHOOK_URL_DEV_APP, msg);
}


// ======================================================
// BACKLOG REPORT
// ======================================================

function sendBacklogReport() {
  // ======================================================
  // 1. Skip weekend
  // ======================================================
  const now = new Date();
  const day = Number(
    Utilities.formatDate(now, TIMEZONE, "u")
  );

  // 1 = Monday ... 7 = Sunday
  if (day === 6 || day === 7) {
    console.log("Skip weekend (VN timezone)");
    return;
  }
  
  // ======================================================
  // 2. Check trigger time
  // ======================================================
  
  const hour = now.getHours();
  const minute = now.getMinutes();

  // Run only during configured time range
  if (
    !(
      hour === BACKLOG_TRIGGER_HOUR &&
      minute >= BACKLOG_TRIGGER_START_MINUTE &&
      minute <= BACKLOG_TRIGGER_END_MINUTE
    )
  ) {
    return;
  }


  // ======================================================
  // 3. Prevent duplicate execution
  // ======================================================

  const key = Utilities.formatDate(
    now,
    TIMEZONE,
    `yyyy-MM-dd-${BACKLOG_TRIGGER_HOUR}`
  );

  const scriptProperties =
    PropertiesService.getScriptProperties();

  if (
    scriptProperties.getProperty("lastRun") === key
  ) {
    return;
  }

  scriptProperties.setProperty("lastRun", key);


  // ======================================================
  // 4. Get daily report data
  // ======================================================

  const {
    reportByUser,
    today
  } = getDailyReportData();


  // ======================================================
  // 5. Create Backlog issue
  // ======================================================
  
  createBacklogIssue(
    reportByUser,
    today
  );
}


// ======================================================
// FORMAT USER TASK
// ======================================================

function parseNote(note) {
  if (!note) {
    return {
      today: "",
      next: ""
    };
  }

  let text = note.toString();

  // Remove Issue completely
  text = text.replace(
    /Issue:.*$/is,
    ""
  );

  const todayMatch = text.match(
    /Today:(.*?)(Next plan:|$)/is
  );

  const nextMatch = text.match(
    /Next plan:(.*)/is
  );

  return {
    today: todayMatch
      ? todayMatch[1].trim()
      : "",

    next: nextMatch
      ? nextMatch[1].trim()
      : ""
  };
}


// ======================================================
// TODO LIST
// ======================================================

function getTodoInProgress() {
  // ======================================================
  // 1. Open Todo List file
  // ======================================================

  const tempFileId = createTempSpreadsheetFromExcel(
    DAILY_REPORT_FILE_ID,
    "temp_todo"
  );

  const ss = SpreadsheetApp.openById(tempFileId);
  const sheet = ss.getSheetByName(TODO_SHEET_NAME);

  if (!sheet) {
    throw new Error(
      "Không tìm thấy sheet: " + TODO_SHEET_NAME
    );
  }


  // ======================================================
  // 2. Read data
  // ======================================================

  const data = sheet.getDataRange().getValues();

  const todos = [];


  // ======================================================
  // 3. Filter status = progress
  // ======================================================

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const stt = row[0];          // Column A
    const description = row[1];  // Column B
    const pic = row[2];          // Column C
    const deadline = row[3];     // Column D
    const status = row[4];       // Column E
    const note = row[5];         // Column F

    if (!status) {
      continue;
    }

    const statusText = status
      .toString()
      .toLowerCase();

    if (statusText.includes("progress")) {
      todos.push({
        stt,
        description,
        pic,
        deadline,
        note
      });
    }
  }


  // ======================================================
  // 4. Delete temp file
  // ======================================================

  deleteTempFile(tempFileId);


  // ======================================================
  // 5. Return todo list
  // ======================================================

  return todos;
}