// CODE.GS - SERVER SIDE SCRIPT

// USER CONFIGURATION
const SPREADSHEET_ID = '14BI-24jotB8qrglwXn2oKe2cfwfNwbmLBFKqYp-zG-M';
const CV_FOLDER_ID = '1xhqxLSXYQLQ0qSYhnQjAXkwI-EbUtv_M';
const CANDIDATE_SHEET_NAME = 'DATA ỨNG VIÊN';
const SETTINGS_SHEET_NAME = 'CẤU HÌNH HỆ THỐNG';
const COMPANY_CONFIG_SHEET_NAME = 'CẤU HÌNH CÔNG TY';
const EVALUATION_SHEET_NAME = 'EVALUATIONS';
const USERS_SHEET_NAME = 'Users';
const SOURCE_SHEET_NAME = 'Survey_Data';
const PROJECTS_SHEET_NAME = 'PROJECTS';
const TICKETS_SHEET_NAME = 'RECRUITMENT_TICKETS';
const DOCUMENT_FOLDER_ID = ''; // Will be auto-created if empty

// ====================== CẤU HÌNH GEMINI CV PARSER (2026) ======================
const GEMINI_API_KEY = 'AIzaSyCOIQiwucbtKWpFUF-65ICt8QNJygUvZL4';
const MODEL = 'gemini-2.5-flash';   // Khuyến nghị: nhanh + PDF native cực mạnh

const CV_SCHEMA = {
  "type": "object",
  "properties": {
    "Name": {
      "type": "string",
      "description": "Họ và tên đầy đủ của ứng viên"
    },
    "Phone": {
      "type": "string",
      "description": "Số điện thoại (có thể kèm mã vùng)"
    },
    "Email": {
      "type": "string",
      "description": "Địa chỉ email"
    },
    "Birth_Year": {
      "type": "integer",
      "description": "Năm sinh (chỉ số năm, ví dụ: 1995). Nếu không có thì null"
    },
    "Experience": {
      "type": "string",
      "description": "Số năm kinh nghiệm hoặc tóm tắt kinh nghiệm làm việc"
    },
    "School": {
      "type": "string",
      "description": "Tên trường đại học/cao đẳng chính"
    },
    "Major": {
      "type": "string",
      "description": "Ngành học chính"
    },
    "Position": {
      "type": "string",
      "description": "Vị trí ứng tuyển phù hợp nhất dựa trên toàn bộ CV"
    }
  },
  "required": ["Name", "Email"]
};

/**
 * Hàm chính nhận file CV → trả về JSON sạch
 */
function apiParseCV(fileData) {
  try {
    return callGeminiAI(fileData);
  } catch (error) {
    console.error("Lỗi Gemini CV Parser:", error);
    return { 
      error: "Lỗi xử lý CV: " + error.toString(),
      raw: error.message 
    };
  }
}

/**
 * Gọi Gemini với Structured JSON Output
 */
/**
 * Gọi Gemini với Structured JSON Output
 * Chấp nhận: fileData (CV PDF) HOẶC textOnly + prompt + schema
 */
function callGeminiAI(input) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${GEMINI_API_KEY}`;
  let parts = [];

  if (input.text_only) {
    parts.push({ "text": input.custom_prompt || "Hãy xử lý yêu cầu sau và trả về JSON." });
  } else {
    const promptText = input.custom_prompt || "Bạn là chuyên gia tuyển dụng. Hãy phân tích CV đính kèm và trích xuất thông tin chính xác theo JSON Schema đã định nghĩa.";
    parts.push({ "text": promptText });

    if (input.base64) {
      parts.push({
        "inline_data": {
          "mime_type": input.type || "application/pdf",
          "data": input.base64
        }
      });
    }
  }

  const payload = {
    "contents": [{ "parts": parts }],
    "generationConfig": {
      "responseMimeType": "application/json",
      "responseJsonSchema": input.schema || (input.text_only ? undefined : CV_SCHEMA),
      "temperature": 0.2,
      "maxOutputTokens": 2048
    }
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    throw new Error(`Gemini API Error (${responseCode}): ${responseText}`);
  }

  const json = JSON.parse(responseText);
  if (!json.candidates?.[0]?.content?.parts?.[0]?.text) {
    throw new Error("Không nhận được kết quả từ Gemini");
  }

  const resultText = json.candidates[0].content.parts[0].text.trim();
  try {
    return JSON.parse(resultText);
  } catch (e) {
    const cleaned = resultText.replace(/```json|```/g, "").trim();
    return JSON.parse(cleaned);
  }
}

/**
 * AI: Tạo JD công việc tự động
 */
function apiGenerateJD(jobTitle, coreRequirements) {
  try {
    const prompt = `Bạn là một chuyên gia HR. Hãy soạn thảo một bản mô tả công việc (JD) chi tiết cho vị trí: ${jobTitle}. 
    Các yêu cầu chính bao gồm: ${coreRequirements}. 
    Bản JD cần có các mục rõ ràng (Giới thiệu, Trách nhiệm, Yêu cầu, Quyền lợi). 
    Ngôn ngữ: Tiếng Việt, chuyên nghiệp. Trả về JSON: {"jd": "nội dung JD dạng HTML hoặc Markdown nhẹ"}`;

    const res = callGeminiAI({ 
      text_only: true, 
      custom_prompt: prompt,
      schema: {
        type: "object",
        properties: {
          jd: { type: "string" }
        },
        required: ["jd"]
      }
    });
    return { success: true, jd: res.jd };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * AI: Phân tích mức độ phù hợp giữa Ứng viên và Ticket JD
 */
function apiAnalyzeCandidateMatching(candidateId, ticketId) {
  try {
    // 1. Lấy dữ liệu ứng viên
    const candSheet = getSheetByName(CANDIDATE_SHEET_NAME);
    const candData = candSheet.getDataRange().getValues();
    const candHeaders = candData[0];
    const candidateRow = candData.find(row => String(row[candHeaders.indexOf('ID')]) === String(candidateId));

    if (!candidateRow) throw new Error("Không tìm thấy ứng viên.");
    
    const candidate = {};
    candHeaders.forEach((h, i) => candidate[h] = candidateRow[i]);

    // 2. Lấy dữ liệu Ticket/JD
    const ticketSheet = getSheetByName(TICKETS_SHEET_NAME);
    let ticket = null;
    if (ticketSheet) {
        const tickData = ticketSheet.getDataRange().getValues();
        const tickHeaders = tickData[0];
        const ticketRow = tickData.find(row => String(row[tickHeaders.indexOf('TicketID') || tickHeaders.indexOf('Mã Ticket')]) === String(ticketId));
        if (ticketRow) {
            ticket = {};
            tickHeaders.forEach((h, i) => ticket[h] = ticketRow[i]);
        }
    }

    const jdText = ticket ? `Vị trí: ${ticket['Position'] || ticket['Vị trí']}. Yêu cầu: ${ticket['Job_Description'] || ticket['Mô tả'] || ''}` : "Chưa rõ JD cụ thể.";
    const candidateText = `Học vấn: ${candidate['School']}. Chuyên ngành: ${candidate['Major']}. Kinh nghiệm: ${candidate['Experience']}. Ghi chú: ${candidate['Notes']}`;

    const prompt = `Hãy so sánh ứng viên với JD sau. Chấm điểm 0-100 và nêu ưu/nhược điểm.
    JD: ${jdText}
    Candidate: ${candidateText}
    Trả về JSON: {"score": number, "pros": ["..."], "cons": ["..."], "summary": "..."}`;

    const analysis = callGeminiAI({ 
        text_only: true, 
        custom_prompt: prompt,
        schema: {
            type: "object",
            properties: {
                score: { type: "number" },
                pros: { type: "array", items: { type: "string" } },
                cons: { type: "array", items: { type: "string" } },
                summary: { type: "string" }
            },
            required: ["score", "pros", "cons", "summary"]
        }
    });

    return { success: true, analysis: analysis };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * API Chatbot: Trả lời câu hỏi dựa trên ngữ cảnh hệ thống
 */
function apiChatWithGemini(userMessage) {
  try {
    const url = "https://generativelanguage.googleapis.com/v1beta/models/" + MODEL + ":generateContent?key=" + GEMINI_API_KEY;
    
    // Tạo ngữ cảnh hệ thống để AI hiểu vai trò
    const systemContext = `Bạn là "Trợ lý Tuyển dụng Trung Khánh" - một AI tích hợp trong hệ thống ATS.
    Nhiệm vụ:
    1. Tra cứu dữ liệu (Ứng viên, Ticket, Lịch trình).
    2. Giải đáp quy trình tuyển dụng.
    3. Hỗ trợ soạn thảo (Email, JD, Ghi chú).
    
    Phong cách: Chuyên nghiệp, thân thiện, trả lời ngắn gọn.
    Định dạng: Sử dụng Markdown nhẹ, xuống dòng rõ ràng.
    
    Dữ liệu ngữ cảnh hiện tại:
    - Người phụ trách: Admin Trung Khánh.
    - Hệ thống: ATS (Applicant Tracking System).`;

    const payload = {
      "contents": [
        { 
          "role": "user", 
          "parts": [{ "text": systemContext + "\n\nNgười dùng hỏi: " + userMessage }] 
        }
      ],
      "generationConfig": { 
        "temperature": 0.7,
        "maxOutputTokens": 1024
      }
    };

    const options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      throw new Error(`Gemini Error: ${responseText}`);
    }

    const json = JSON.parse(responseText);
    const answer = json.candidates?.[0]?.content?.parts?.[0]?.text || "Em chưa tìm ra câu trả lời phù hợp, anh thử hỏi cách khác nhé.";
    
    return { success: true, answer: answer };

  } catch (error) {
    Logger.log("Chatbot Error: " + error.toString());
    return { success: false, message: error.toString(), answer: "Xin lỗi, em đang gặp sự cố kết nối. Vui lòng thử lại sau." };
  }
}

/**
 * Hàm tạo lịch phỏng vấn trên Google Calendar
 */
function apiCreateInterviewSchedule(scheduleData) {
    try {
        const { candidateName, candidateEmail, managerEmail, startTime, endTime, location, note } = scheduleData;
        
        const start = new Date(startTime);
        const end = new Date(endTime);
        
        const eventTitle = `[Phỏng vấn] ${candidateName} - Phụ trách: ${managerEmail}`;
        const description = `Ghi chú: ${note}\n\nỨng viên: ${candidateName}\nEmail: ${candidateEmail}`;
        
        const event = CalendarApp.getDefaultCalendar().createEvent(eventTitle, start, end, {
            location: location || 'Văn phòng Công ty / Online (Google Meet)',
            description: description,
            guests: `${candidateEmail},${managerEmail}`, 
            sendInvites: true
        });

        return { 
            success: true, 
            eventId: event.getId(),
            message: 'Đã tạo lịch và gửi thư mời thành công!'
        };
        
    } catch (error) {
        return { success: false, message: 'Lỗi tạo lịch: ' + error.toString() };
    }
}


// Shared Candidate Aliases for mapping Frontend keys to Sheet Headers
const CANDIDATE_ALIASES = {
  'ID': ['ID', 'Mã', 'Mã ứng viên'],
  'Name': ['Name', 'Họ và Tên', 'Họ tên', 'Họ tên ứng viên'],
  'Gender': ['Gender', 'Giới tính'],
  'Birth_Year': ['Birth_Year', 'Năm sinh', 'Birth Year', 'Namsinh'],
  'Phone': ['Phone', 'Số điện thoại', 'SĐT', 'SDT'],
  'Email': ['Email'],
  'Department': ['Department', 'Phòng ban', 'Bộ phận'],
  'Position': ['Position', 'Vị trí', 'Vị trí ứng tuyển', 'Chức danh'],
  'Experience': ['Experience', 'Kinh nghiệm'],
  'School': ['School', 'Học vấn', 'Trường học', 'Tên trường', 'Đại học'],
  'Education_Level': ['Education_Level', 'Trình độ', 'Trình độ học vấn', 'Education Level'],
  'Major': ['Major', 'Chuyên ngành', 'Ngành học'],
  'Salary_Expectation': ['Salary_Expectation', 'Mức lương mong muốn', 'Expected Salary', 'Expected_Salary', 'Salary', 'Lương đề xuất', 'Expected_Salary_TicketID'],
  'Source': ['Source', 'Nguồn', 'Nguồn ứng viên', 'Data Sources'],
  'Recruiter': ['Recruiter', 'Người phụ trách', 'Nhân viên tuyển dụng', 'Người tuyển dụng'],
  'Stage': ['Stage', 'Trạng thái', 'Trạng thái tuyển dụng', 'Giai đoạn'],
  'Status': ['Status', 'Tình trạng liên hệ', 'Trạng thái liên lạc', 'Contact_Status'],
  'Notes': ['Notes', 'Ghi chú'],
  'CV_Link': ['CV_Link', 'Link CV', 'CV Link'],
  'Applied_Date': ['Applied_Date', 'Ngày ứng tuyển', 'Ngày Apply', 'Applied Date'],
  'User': ['User', 'Người cập nhật', 'Creator'],
  'TicketID': ['TicketID', 'Mã Ticket', 'Ticket ID', 'Mã yêu cầu'],
  'Rejection_Type': ['Rejection_Type', 'Loại từ chối'],
  'Rejection_Reason': ['Rejection_Reason', 'Lý do từ chối'],
  'Rejection_Source': ['Rejection_Source', 'Đối tượng từ chối'],
  'Hire_Date': ['Hire_Date', 'Ngày nhận việc']
};

// DEBUG EMAIL PERMISSION
function debugSendsEmail() {
  const email = Session.getActiveUser().getEmail();
  Logger.log("Current User: " + email);
  if (email) {
    MailApp.sendEmail({
      to: email,
      subject: "Test Permission [ATS]",
      htmlBody: "<p>Nếu bạn nhận được email này, quyền gửi email đã được cấp thành công!</p>"
    });
    Logger.log("Đã gửi email test đến: " + email);
  } else {
    Logger.log("Không lấy được email người dùng hiện tại. Hãy chạy bằng tài khoản chủ sở hữu.");
  }
}

// HELPER: Log to Sheet for Debugging
function logToSheet(message) {
  Logger.log(message);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Debug_Logs');
    if (!sheet) {
      sheet = ss.insertSheet('Debug_Logs');
      sheet.appendRow(['Timestamp', 'Message']);
    }
    sheet.appendRow([new Date(), message]);
  } catch (e) {
    Logger.log('Log Error: ' + e.toString());
  }
}

/**
 * Initialize new sheets for Projects and Tickets
 */
function initializeProjectSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. Projects Sheet
  let projectSheet = ss.getSheetByName(PROJECTS_SHEET_NAME);
  if (!projectSheet) {
    projectSheet = ss.insertSheet(PROJECTS_SHEET_NAME);
    const headers = ['Mã Dự án', 'Tên Dự án', 'Quy trình (Workflow)', 'Người quản lý', 'Ngày bắt đầu', 'Ngày kết thúc', 'Trạng thái', 'Chỉ tiêu', 'Ngân sách', 'Ghi chú'];
    projectSheet.appendRow(headers);
    projectSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
    
    // Add a default project
    projectSheet.appendRow(['PROJ001', 'Dự án Tuyển dụng Tiêu chuẩn', 'Ứng tuyển, Xét duyệt hồ sơ, Sơ vấn, Phỏng vấn, Phê duyệt nhận việc, Mời nhận việc, Đã nhận việc, Từ chối', 'Admin', Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy"), '', 'Active']);
  } else {
    // Sync headers if needed
    const headers = ['Mã Dự án', 'Tên Dự án', 'Quy trình (Workflow)', 'Người quản lý', 'Ngày bắt đầu', 'Ngày kết thúc', 'Trạng thái', 'Chỉ tiêu', 'Ngân sách', 'Ghi chú'];
    const currentHeaders = projectSheet.getRange(1, 1, 1, projectSheet.getLastColumn()).getValues()[0];
    if (currentHeaders.length < headers.length) {
      projectSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  
  // 2. Tickets Sheet
  let ticketSheet = ss.getSheetByName(TICKETS_SHEET_NAME);
  if (!ticketSheet) {
    ticketSheet = ss.insertSheet(TICKETS_SHEET_NAME);
  }
  
  const ticketHeaders = [
    'Mã Ticket', 'Mã Dự án', 'Vị trí cần tuyển', 'Số lượng', 'Phòng ban', 'Loại hình làm việc', 
    'Giới tính', 'Tuổi', 'Học vấn', 'Chuyên môn', 'Kinh nghiệm', 
    'Ngày bắt đầu', 'Deadline', 'Quản lý trực tiếp', 'Văn phòng làm việc', 
    'Trạng thái Phê duyệt', 'Chi phí tuyển dụng', 'Recruiter phụ trách', 'Ngày phê duyệt', 'Creator'
  ];
  ticketSheet.getRange(1, 1, 1, ticketHeaders.length).setValues([ticketHeaders]);
  ticketSheet.getRange(1, 1, 1, ticketHeaders.length).setFontWeight('bold').setBackground('#f3f3f3');
  
  return "Initialization successful";
}

function sheetToObjects(sheetName) {
  try {
    const sheet = getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Sheet not found: ${sheetName}`);
      return [];
    }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    // Filter headers to avoid empty/undefined keys
    const headers = data[0].map(h => String(h).trim()).filter(h => h !== '');
    
    return data.slice(1).map(row => {
      const obj = {};
      headers.forEach((header, i) => {
        let val = row[i];
        // Clean values for reliable JSON serialization
        if (val instanceof Date) {
          obj[header] = val.toISOString();
        } else if (val === undefined || val === null) {
          obj[header] = '';
        } else if (typeof val === 'object') {
          // Flatten objects or stringify them
          try { obj[header] = JSON.stringify(val); } catch(e) { obj[header] = String(val); }
        } else {
          obj[header] = val;
        }
      });
      return obj;
    });
  } catch (e) {
    Logger.log(`Error reading sheet ${sheetName}: ${e.toString()}`);
    return [];
  }
}

function apiGetProjects() {
  return { success: true, data: sheetToObjects(PROJECTS_SHEET_NAME) };
}

function apiGetTickets() {
  return { success: true, data: sheetToObjects(TICKETS_SHEET_NAME) };
}

/**
 * 4. API: PROJECT MANAGEMENT
 */
function apiSaveProject(projectData) {
  try {
    let sheet = getSheetByName(PROJECTS_SHEET_NAME);
    if (!sheet) {
      initializeProjectSheets();
      sheet = getSheetByName(PROJECTS_SHEET_NAME);
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const findIdx = (key) => {
      let lowerHeaders = headers.map(h => h.toString().trim().toLowerCase());
      let idx = lowerHeaders.indexOf(key.toLowerCase());
      if (idx !== -1) return idx;
      // You can expand this if needed, but for Projects we use direct match for now
      return -1;
    };

    const quotaIdx = findIdx('Chỉ tiêu');
    const budgetIdx = findIdx('Ngân sách');
    const codeIdx = findIdx('Mã Dự án');
    
    let rowIndex = -1;
    if (projectData.code && codeIdx !== -1) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][codeIdx]).trim() === String(projectData.code).trim()) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    if (rowIndex !== -1) {
      // Update specific columns to avoid overwriting other fields if headers changed
      const updateData = {
        'Tên Dự án': projectData.name || '',
        'Quy trình (Workflow)': projectData.workflow || '',
        'Người quản lý': projectData.manager || '',
        'Ngày bắt đầu': projectData.startDate || '',
        'Ngày kết thúc': projectData.endDate || '',
        'Chỉ tiêu': parseFloat(projectData.quota) || 0,
        'Ngân sách': parseFloat(projectData.budget) || 0
      };

      Object.keys(updateData).forEach(key => {
        const idx = findIdx(key);
        if (idx !== -1) {
          sheet.getRange(rowIndex, idx + 1).setValue(updateData[key]);
        }
      });

      return { success: true, message: 'Cập nhật dự án thành công.' };
    } else {
      // For new project, construct row carefully
      const newRow = new Array(headers.length).fill('');
      const mapping = {
        'Mã Dự án': projectData.code || ('PROJ_' + new Date().getTime()),
        'Tên Dự án': projectData.name || '',
        'Quy trình (Workflow)': projectData.workflow || '',
        'Người quản lý': projectData.manager || '',
        'Ngày bắt đầu': projectData.startDate || '',
        'Ngày kết thúc': projectData.endDate || '',
        'Trạng thái': 'Active',
        'Chỉ tiêu': parseFloat(projectData.quota) || 0,
        'Ngân sách': parseFloat(projectData.budget) || 0
      };
      
      headers.forEach((h, i) => {
        if (mapping[h]) newRow[i] = mapping[h];
      });

      sheet.appendRow(newRow);
      return { success: true, message: 'Lưu dự án mới thành công.' };
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * 5. API: RECRUITMENT TICKETS
 */
function apiSaveTicket(ticketData, currentUser) {
  try {
    let sheet = getSheetByName(TICKETS_SHEET_NAME);
    if (!sheet) {
      initializeProjectSheets();
      sheet = getSheetByName(TICKETS_SHEET_NAME);
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const ticketCode = ticketData.code || ('TICK_' + new Date().getTime());
    
    // Check for existing ticket to update (support Admin edit and Manager edit request)
    let rowIndex = -1;
    if (ticketData.code) {
      const codeIdx = headers.indexOf('Mã Ticket');
      if (codeIdx !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][codeIdx]).trim() === String(ticketData.code).trim()) {
            rowIndex = i + 1;
            break;
          }
        }
      }
    }

    // Role-based status logic
    let status = ticketData.approvalStatus || 'Pending';
    if (currentUser && currentUser.role === 'Admin') {
       // Admins can approve immediately if they filled Recruiter
       if (ticketData.recruiterEmail) status = 'Approved';
    } else if (status === 'Approved') {
       // If a Manager edits an Approved ticket -> set back to Pending
       status = 'Pending';
    }
    
    const row = [
      ticketCode,
      ticketData.projectCode || '',
      ticketData.position || '',
      ticketData.quantity || 1,
      ticketData.department || '',
      ticketData.workType || '',
      ticketData.gender || '',
      ticketData.age || '',
      ticketData.education || '',
      ticketData.major || '',
      ticketData.experience || '',
      ticketData.startDate || '',
      ticketData.deadline || '',
      ticketData.directManager || '',
      ticketData.office || '',
      status, 
      ticketData.costs ? JSON.stringify(ticketData.costs) : '', // Chi phí
      ticketData.recruiterEmail || '', // Recruiter
      status === 'Approved' ? Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd") : '', // Ngày duyệt
      currentUser ? (currentUser.username || currentUser.email) : '' // Creator column
    ];
    
    if (rowIndex > 0) {
      // Update existing
      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
      
      const userDisplay = currentUser ? (currentUser.username || currentUser.email) : 'System';
      logActivity(userDisplay, 'Cập nhật Ticket', `Đã cập nhật phiếu yêu cầu tuyển dụng: ${ticketCode}. Trạng thái: ${status}`, ticketCode);
      
      if (currentUser && currentUser.role === 'Manager' && ticketData.approvalStatus === 'Approved') {
         createNotification('admin', 'Ticket', `Manager ${userDisplay} yêu cầu chỉnh sửa ticket đã duyệt: ${ticketCode}. Trạng thái đã quay về Chờ duyệt.`, ticketCode);
      }
    } else {
      // New row
      sheet.appendRow(row);
      const userDisplay = currentUser ? (currentUser.username || currentUser.email) : 'System';
      logActivity(userDisplay, 'Tạo Ticket', `Đã tạo phiếu yêu cầu tuyển dụng: ${ticketCode}`, ticketCode);

      if (currentUser && currentUser.role === 'Manager') {
         createNotification('admin', 'Ticket', `Manager ${userDisplay} vừa tạo một yêu cầu tuyển dụng mới: ${ticketCode}`, ticketCode);
      }
    }

    return { success: true, message: rowIndex > 0 ? 'Cập nhật phiếu thành công.' : 'Tạo phiếu yêu cầu thành công.' };
  } catch (e) {
    Logger.log('ERROR apiSaveTicket: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function apiApproveTicket(ticketCode, approvalData, adminUser) {
  try {
    const sheet = getSheetByName(TICKETS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const ticketIdx = headers.indexOf('Mã Ticket');
    const statusIdx = headers.indexOf('Trạng thái Phê duyệt');
    const costsIdx = headers.indexOf('Chi phí tuyển dụng');
    const recruiterIdx = headers.indexOf('Recruiter phụ trách');
    const dateIdx = headers.indexOf('Ngày phê duyệt');
    const creatorIdx = headers.indexOf('Creator') !== -1 ? headers.indexOf('Creator') : (headers.length - 1);
    
    for (let i = 1; i < data.length; i++) {
       if (String(data[i][ticketIdx]).trim() === String(ticketCode).trim()) {
         const newStatus = approvalData.status || 'Approved';
         sheet.getRange(i + 1, statusIdx + 1).setValue(newStatus);
         sheet.getRange(i + 1, costsIdx + 1).setValue(JSON.stringify(approvalData.costs || []));
         sheet.getRange(i + 1, recruiterIdx + 1).setValue(approvalData.recruiterEmail || '');
         sheet.getRange(i + 1, dateIdx + 1).setValue(new Date());

         // Notification to creator
         const creator = data[i][creatorIdx];
         if (creator) {
           const actionText = newStatus === 'Approved' ? 'đã được duyệt' : 'thay đổi trạng thái';
           createNotification(creator, 'Ticket', `Phiếu tuyển dụng ${ticketCode} của bạn ${actionText}.`, ticketCode);
         }

         logActivity(adminUser || 'Admin', 'Phê duyệt Ticket', `Đã duyệt/cập nhật Ticket ${ticketCode} với trạng thái ${newStatus}`, ticketCode);

         return { success: true, message: 'Cập nhật Ticket thành công.' };
       }
    }
    return { success: false, message: 'Không tìm thấy Ticket.' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// DEBUG: Inspect sheet structure and data
function debugSheetData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = ss.getSheets();
    let report = [];
    sheets.forEach(s => {
      const name = s.getName();
      const data = s.getDataRange().getValues();
      report.push(`--- Sheet: ${name} (Rows: ${data.length}) ---`);
      if (data.length > 0) report.push(`Headers: ${JSON.stringify(data[0])}`);
      if (data.length > 1) report.push(`Sample (R2): ${JSON.stringify(data[1]).substring(0, 200)}...`);
    });
    return report.join('\n');
  } catch (e) {
    return 'Debug Error: ' + e.toString();
  }
}

// 1. SETUP & ROUTING
function doGet(e) {
  if (!e) e = { parameter: {} };
  
  if (e.parameter && e.parameter.view === 'career') {
      return HtmlService.createTemplateFromFile('CareerPage')
          .evaluate()
          .setTitle('Tuyển Dụng - HR Recruit')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
          .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (e.parameter && e.parameter.view === 'news') {
      return HtmlService.createTemplateFromFile('NewsPage')
          .evaluate()
          .setTitle('Tin Tức - HR Recruit')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
          .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const isDbSetup = !!SPREADSHEET_ID;
  const t = HtmlService.createTemplateFromFile('Index');
  t.isDbSetup = isDbSetup;
  return t.evaluate()
      .setTitle('Recruitment ATS System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

// 2. AUTHENTICATION
function apiLogin(username, password) {
  try {
      // Simple check for config existence
      if (!SPREADSHEET_ID) {
        return { success: false, message: 'ChÆ°a cáº¥u hÃ¬nh Spreadsheet ID.' };
      }

      // Normalize inputs
      const cleanUser = username ? username.toString().trim() : '';
      const cleanPass = password ? password.toString().trim() : '';

      const users = apiGetTableData('Users');
      
      // 1. Check DB Users (Case insensitive for username)
      const user = users.find(u => {
          return u.Username && u.Username.toString().toLowerCase() === cleanUser.toLowerCase() && u.Password == cleanPass;
      });
      
      if (user) {
        return { 
            success: true, 
            user: { 
                username: user.Username, 
                role: String(user.Role || '').trim(), 
                name: user.Full_Name,
                department: user.Department || '',
                email: String(user.Email || '').toLowerCase().trim()
            } 
        };
      }
      
      // 2. Fallback Default Admin (if not found in DB)
      // Allow 'admin' (case insensitive) with specific passwords
      if (cleanUser.toLowerCase() === 'admin' && (cleanPass === 'admin' || cleanPass === '123456')) {
           return { 
               success: true, 
               message: 'Login via Default Admin', 
               user: { 
                   username: 'admin', 
                   role: 'Admin', 
                   name: 'System Admin',
                   department: 'All',
                   email: 'admin@hr.com'
               } 
           };
      }

      return { success: false, message: 'Sai tÃªn Ä‘Äƒng nháº­p hoáº·c máº­t kháº©u. (Input: ' + cleanUser + ')' };
  } catch (e) {
      return { success: false, message: 'Server Error: ' + e.toString() };
  }
}

// 3. API: GET INITIAL DATA
function apiGetInitialData(requestingUsername) {
  Logger.log('=== apiGetInitialData Called === User: ' + requestingUsername);
  
  const result = {
    candidates: [],
    stages: [],
    departments: [],
    recruiters: [],
    users: [],
    emailTemplates: [],
    companyInfo: {},
    projects: [],
    tickets: [],
    aliases: CANDIDATE_ALIASES
  };

  try {
    ensureSheetStructure(); // Auto-update structure if missing

    if (!SPREADSHEET_ID) {
      Logger.log('ERROR: SPREADSHEET_ID is missing');
      return result;
    }

    // 1. Determine User Role & Scope
    let userRole = 'Viewer';
    let userDept = '';
    
    try {
      if (requestingUsername) {
          const users = apiGetTableData('Users');
          const user = users.find(u => u.Username && u.Username.toString().toLowerCase() === requestingUsername.toString().toLowerCase());
          if (user) {
              userRole = user.Role;
              userDept = user.Department;
          } else if (requestingUsername.toLowerCase() === 'admin') {
              userRole = 'Admin';
              userDept = 'All';
          }
      }
    } catch (e) {
      Logger.log('Warning: Error getting user role: ' + e.toString());
    }
    
    Logger.log(`User: ${requestingUsername}, Role: ${userRole}, Dept: ${userDept}`);

    // 2. Load Candidates
    try {
      let candidates = apiGetTableData(CANDIDATE_SHEET_NAME) || [];
      
      // APPLY SCOPING RULES
      if (userRole === 'Manager' && userDept && userDept !== 'All') {
          candidates = candidates.filter(c => {
              const cDept = (c.Department || '').toString().trim().toLowerCase();
              const uDept = userDept.toString().trim().toLowerCase();
              return cDept === uDept;
          });
      }
      
      // Clean candidates data with Smart Mapping
      result.candidates = candidates.map(function(c) {
        let obj = {};
        
        // Prepare lowercase keys for case-insensitive lookup
        const cLower = {};
        Object.keys(c).forEach(k => {
          cLower[k.toString().toLowerCase().trim()] = c[k];
        });

        Object.keys(CANDIDATE_ALIASES).forEach(key => {
            const aliases = CANDIDATE_ALIASES[key] || [];
            const searchPool = [key, ...aliases].map(s => s.toLowerCase().trim());
            
            let val = '';
            for (let s of searchPool) {
                if (cLower[s] !== undefined) {
                    val = cLower[s];
                    break;
                }
            }

            // Date Special Handling
            if (val instanceof Date) {
                obj[key] = val.toISOString();
            } else {
                obj[key] = val === null || val === undefined ? '' : String(val);
            }
        });
        
        return obj;
      });
    } catch (e) {
      Logger.log('Error loading candidates: ' + e.toString());
    }

    // 3. Load other components with safety
    try {
      // Disabled global stages loading - transitioning to project-specific workflows
      // Transitioning to Support both English and Vietnamese for Stages
      result.stages = [
           {ID: 'S1', Stage_Name: 'Ứng tuyển', Order: 1, Color: '#0d6efd'},
           {ID: 'S2', Stage_Name: 'Sơ vấn', Order: 2, Color: '#6610f2'},
           {ID: 'S3', Stage_Name: 'Phỏng vấn', Order: 3, Color: '#fd7e14'},
           {ID: 'S4', Stage_Name: 'Xét duyệt hồ sơ', Order: 4, Color: '#0dcaf0'},
           {ID: 'S5', Stage_Name: 'Phê duyệt nhận việc', Order: 5, Color: '#20c997'},
           {ID: 'S6', Stage_Name: 'Mời nhận việc', Order: 6, Color: '#198754'},
           {ID: 'S7', Stage_Name: 'Đã nhận việc', Order: 7, Color: '#28a745'},
           {ID: 'S8', Stage_Name: 'Từ chối', Order: 8, Color: '#dc3545'}
      ];
    } catch (e) { Logger.log('Error loading stages fallback: ' + e.toString()); }

    try {
      const deptResult = apiGetDepartments();
      result.departments = deptResult.success ? deptResult.departments : [];
    } catch (e) { Logger.log('Error loading departments: ' + e.toString()); }

    try {
      const recResult = apiGetRecruiters();
      result.recruiters = recResult.success ? recResult.recruiters : [];
    } catch (e) { Logger.log('Error loading recruiters: ' + e.toString()); }

    try {
      result.users = apiGetTableData(USERS_SHEET_NAME) || [];
    } catch (e) { Logger.log('Error loading users: ' + e.toString()); }

    try {
      result.emailTemplates = apiGetEmailTemplates() || [];
    } catch (e) { Logger.log('Error loading templates: ' + e.toString()); }

    try {
      const compInfoResult = apiGetCompanyInfo();
      result.companyInfo = compInfoResult.data || {};
    } catch (e) { Logger.log('Error loading company info: ' + e.toString()); }

    try {
      const projResult = apiGetProjects();
      result.projects = projResult.data || [];
    } catch (e) { Logger.log('Error loading projects: ' + e.toString()); }

    try {
      const tickResult = apiGetTickets();
      result.tickets = tickResult.data || [];
    } catch (e) { Logger.log('Error loading tickets: ' + e.toString()); }

    try {
      result.rejectionReasons = apiGetTableData('RejectionReasons') || [];
    } catch (e) { Logger.log('Error loading rejection reasons: ' + e.toString()); }

    Logger.log('apiGetInitialData completed successfully');
    return result;
    
  } catch(e) {
    Logger.log('FATAL ERROR in apiGetInitialData: ' + e.toString());
    return result; // Always return the structured object
  }
}

/**
 * 3.1 API: COMPANY CONFIGURATION
 */
function apiGetCompanyInfo() {
  try {
    const sheet = getSheetByName(COMPANY_CONFIG_SHEET_NAME);
    if (!sheet) return { success: true, data: {} };
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: {} };
    
    const config = {};
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];
      const value = data[i][1];
      if (key) {
        try {
          config[key] = JSON.parse(value);
        } catch (e) {
          config[key] = value;
        }
      }
    }
    return { success: true, data: config };
  } catch (e) {
    Logger.log('ERROR apiGetCompanyInfo: ' + e.toString());
    return { success: false, message: e.toString(), data: {} };
  }
}

function apiSaveCompanyInfo(config) {
  try {
    let sheet = getSheetByName(COMPANY_CONFIG_SHEET_NAME);
    if (!sheet) {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      sheet = ss.insertSheet(COMPANY_CONFIG_SHEET_NAME);
      sheet.getRange('A1:B1').setValues([['Cài đặt', 'Giá trị']]).setFontWeight('bold');
    }
    
    const data = [['Cài đặt', 'Giá trị']];
    for (const key in config) {
      let val = config[key];
      if (typeof val === 'object') {
        val = JSON.stringify(val);
      }
      data.push([key, val]);
    }
    
    sheet.clearContents();
    if (data.length > 0) {
      sheet.getRange(1, 1, data.length, 2).setValues(data);
    }
    return { success: true, message: 'Đã lưu thông tin công ty thành công' };
  } catch (e) {
    Logger.log('ERROR apiSaveCompanyInfo: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}


// 4. DATABASE HELPERS
function getSheetByName(name) {
  if (!SPREADSHEET_ID) return null;
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
  } catch (e) {
    return null;
  }
}

function apiGetTableData(sheetName) {
  Logger.log('--- apiGetTableData called for: ' + sheetName);
  const sheet = getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('ERROR: Sheet not found: ' + sheetName);
    return [];
  }
  
  const startRow = 2;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  Logger.log('Sheet found. lastRow: ' + lastRow + ', lastCol: ' + lastCol);
  
  if (lastRow < startRow) {
    Logger.log('No data rows (lastRow < 2)');
    return [];
  }
  
  // Get Headers
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  Logger.log('Headers: ' + JSON.stringify(headers));
  
  // Get Data
  const data = sheet.getRange(startRow, 1, lastRow - 1, lastCol).getValues();
  Logger.log('Data rows retrieved: ' + data.length);
  
  return data.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      const key = header.toString().trim();
      if(key) obj[key] = row[index];
    });
    return obj;
  });
}

/**
 * Robust Column Index Finder with Aliases
 * @param {Array} headers The first row of the sheet
 * @param {string} name Internal key name to look for
 * @returns {number} 0-based index or -1
 */
function findColumnIndex(headers, name) {
  if (!headers || !name) return -1;
  const lower = name.toLowerCase().replace(/_/g, ' ');
  // Common Aliases for recruitment sheets
  const aliasesMap = {
      'ID': ['Mã', 'EvaID', 'CandidateID'],
      'Candidate_ID': ['Mã ứng viên'],
      'Candidate_Name': ['Ứng viên', 'Tên ứng viên', 'Candidate', 'Họ và tên'],
      'Position': ['Vị trí', 'Vị trí tuyển dụng', 'Job'],
      'Department': ['Phòng ban', 'Bộ phận', 'Dept'],
      'Recruiter_Email': ['Recruiter', 'Người phụ trách', 'Email người tuyển dụng', 'Sales phụ trách'],
      'Manager_Email': ['Manager', 'Người đánh giá', 'Người chấm điểm', 'Email người đánh giá', 'Người phỏng vấn'],
      'Created_At': ['Ngày tạo', 'Thời gian gửi', 'Ngày gửi'],
      'Status': ['Trạng thái', 'Tình trạng'],
      'Final_Result': ['Kết quả', 'Kết quả cuối cùng', 'Đánh giá cuối'],
      'Manager_Comment': ['Nhận xét', 'Ghi chú manager', 'Comment', 'Nhận xét chi tiết'],
      'Completed_At': ['Ngày hoàn thành', 'Thời gian hoàn thành'],
      'Batch_ID': ['Mã nhóm', 'BatchID'],
      'Criteria_Config': ['Cấu hình tiêu chí', 'Tiêu chí'],
      'Scores_JSON': ['Dữ liệu điểm chi tiết', 'Scores JSON'],
      'Manager_Name': ['Tên người đánh giá'],
      'Manager_Department': ['Bộ phận người đánh giá'],
      'Manager_Position': ['Vị trí người đánh giá'],
      'Total_Score': ['Tổng điểm', 'Điểm trung bình'],
      'Proposed_Salary': ['Lương đề xuất', 'Salary'],
      'Signature_Status': ['Trạng thái chữ ký', 'Ký tên']
  };
  
  const aliases = aliasesMap[name] || [];
  return headers.findIndex(h => {
    const s = h.toString().trim().toLowerCase().replace(/_/g, ' ');
    if (s === lower || s === name.toLowerCase().replace(/_/g, ' ')) return true;
    return aliases.some(a => s === a.toLowerCase().replace(/_/g, ' '));
  });
}

/**
 * Fetch survey data from 'Survey_Data' sheet
 * Linked by Email
 */
function apiGetSurveyData(email) {
  Logger.log('ðŸ” apiGetSurveyData for: ' + email);
  if (!email) return {};
  
  const data = apiGetTableData('Survey_Data');
  if (data.length === 0) {
    Logger.log('âš ï¸ No data found in Survey_Data sheet');
    return {};
  }
  
  // Find record by email (case insensitive)
  const record = data.find(r => {
    // 1. Try common email column names
    let rEmail = r['Email'] || r['Äá»‹a chá»‰ email'] || r['email'] || r['EMAIL'];
    
    // 2. If not found, look for any column header that contains "email"
    if (!rEmail) {
      const emailKey = Object.keys(r).find(k => k.toLowerCase().includes('email'));
      if (emailKey) rEmail = r[emailKey];
    }
    
    return rEmail && rEmail.toString().toLowerCase().trim() === email.toString().toLowerCase().trim();
  });
  
  if (!record) {
    Logger.log('âŒ No survey record found for: ' + email);
    return {};
  }
  
  Logger.log('âœ… Found survey record for: ' + email);

  // Normalize mapping (Sheet Header -> Internal Placeholder)
  const mapping = {
    'CCCD': ['Sá»‘ CCCD', 'CCCD', 'Sá»‘ chá»©ng minh nhÃ¢n dÃ¢n', 'CMND'],
    'CCCD_Date': ['NgÃ y cáº¥p', 'NgÃ y cáº¥p CCCD', 'NgÃ y cáº¥p CMND'],
    'CCCD_Place': ['NÆ¡i cáº¥p', 'NÆ¡i cáº¥p CCCD', 'NÆ¡i cáº¥p CMND'],
    'HK_Address': ['Há»™ kháº©u thÆ°á»ng trÃº', 'Äá»‹a chá»‰ thÆ°á»ng trÃº', 'Há»™ kháº©u', 'HK_Address'],
    'Current_Address': ['Chá»— á»Ÿ hiá»‡n nay', 'Äá»‹a chá»‰ hiá»‡n táº¡i', 'NÆ¡i á»Ÿ hiá»‡n nay', 'Current_Address'],
    'Bank_Account': ['Sá»‘ tÃ i khoáº£n', 'Sá»‘ tÃ i khoáº£n ngÃ¢n hÃ ng', 'STK', 'Bank_Account'],
    'Bank_Name': ['NgÃ¢n hÃ ng', 'TÃªn ngÃ¢n hÃ ng', 'Má»Ÿ táº¡i ngÃ¢n hÃ ng', 'Bank_Name'],
    'Bank_Branch': ['Chi nhÃ¡nh', 'Chi nhÃ¡nh ngÃ¢n hÃ ng', 'Bank_Branch'],
    'DOB': ['NgÃ y sinh', 'Sinh ngÃ y', 'DOB']
  };

  let normalizedData = {};
  Object.keys(mapping).forEach(internalKey => {
    const sheetKeys = mapping[internalKey];
    for (const sKey of sheetKeys) {
      if (record[sKey] !== undefined && record[sKey] !== null) {
        normalizedData[internalKey] = record[sKey];
        break;
      }
    }
  });

  // Also include original keys just in case
  return { ...record, ...normalizedData };
}

// 5. INITIAL SETUP (Legacy support, mostly manual now)
// 5. INITIAL SETUP & STRUCTURE AUTO-UPDATE
function ensureSheetStructure() {
  if (!SPREADSHEET_ID) return;
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. CANDIDATE SHEET
    let cSheet = ss.getSheetByName(CANDIDATE_SHEET_NAME);
    if (!cSheet) {
      cSheet = ss.insertSheet(CANDIDATE_SHEET_NAME);
      cSheet.appendRow(['ID', 'Name', 'Phone', 'Email', 'Position', 'Source', 'Stage', 'CV_Link', 'Applied_Date', 'Status', 'School', 'Education_Level', 'Major', 'Birth_Year', 'Recruiter', 'Notes', 'Experience', 'Expected_Salary', 'TicketID', 'Rejection_Type', 'Rejection_Reason', 'Hire_Date']);
    } else {
      // Ensure all columns exist
      const currentHeaders = cSheet.getRange(1, 1, 1, Math.max(cSheet.getLastColumn(), 1)).getValues()[0].map(h => String(h).trim().toLowerCase());
      const expectedHeaders = [
        'ID', 'Name', 'Phone', 'Email', 'Position', 'Department', 'Stage', 'Status', 
        'Source', 'Recruiter', 'Notes', 'CV_Link', 'Applied_Date', 'TicketID',
        'Rejection_Type', 'Rejection_Reason', 'Hire_Date'
      ];
      expectedHeaders.forEach(hName => {
        if (currentHeaders.indexOf(hName.toLowerCase()) === -1) {
          cSheet.getRange(1, cSheet.getLastColumn() + 1).setValue(hName);
          currentHeaders.push(hName.toLowerCase());
        }
      });
    }

    // 2. STAGES (Legacy - Disabled)
    /*
    let sSheet = ss.getSheetByName('Stages');
    if (!sSheet) {
      sSheet = ss.insertSheet('Stages');
      sSheet.appendRow(['ID', 'Stage_Name', 'Order', 'Color']);
      sSheet.appendRow(['S1', 'Apply', 1, '#0d6efd']);
      sSheet.appendRow(['S2', 'Interview', 2, '#fd7e14']);
      sSheet.appendRow(['S3', 'Offer', 3, '#198754']);
      sSheet.appendRow(['S4', 'Rejected', 4, '#dc3545']);
    }
    */

    // 3. DEPARTMENTS (Management Tags)
    let dSheet = ss.getSheetByName('Departments');
    if (!dSheet) {
      dSheet = ss.insertSheet('Departments');
      dSheet.appendRow(['Name', 'Manager_Email']);
      dSheet.appendRow(['IT', '']);
      dSheet.appendRow(['HR', '']);
      dSheet.appendRow(['Sales', '']);
    }

    // 4. USERS
    let uSheet = ss.getSheetByName('Users');
    if (!uSheet) {
      uSheet = ss.insertSheet('Users');
      uSheet.appendRow(['Username', 'Password', 'Role', 'Full_Name', 'Email', 'Department']);
      // Default Admin
      uSheet.appendRow(['admin', 'admin', 'Admin', 'System Admin', 'admin@hr.com', 'All']);
    }
    
    // 5. JOBS
    let jSheet = ss.getSheetByName('Jobs');
    if (!jSheet) {
       jSheet = ss.insertSheet('Jobs');
       jSheet.appendRow(['ID', 'Title', 'Department', 'Location', 'Type', 'Status', 'Created_Date', 'Description']);
    }

    // 5.1 REJECTION REASONS
    let rrSheet = ss.getSheetByName('RejectionReasons');
    if (!rrSheet) {
      rrSheet = ss.insertSheet('RejectionReasons');
      rrSheet.appendRow(['ID', 'Type', 'Reason', 'Order']);
      rrSheet.appendRow(['R1', 'Company', 'Kỹ năng chưa phù hợp', 1]);
      rrSheet.appendRow(['R2', 'Company', 'Kỳ vọng lương quá cao', 2]);
      rrSheet.appendRow(['R3', 'Candidate', 'Đã nhận việc nơi khác', 1]);
      rrSheet.appendRow(['R4', 'Candidate', 'Thay đổi định hướng', 2]);
    }
    
    // 6. SOURCES
    let srcSheet = ss.getSheetByName('DATA_SOURCES');
    if (!srcSheet) {
        srcSheet = ss.insertSheet('DATA_SOURCES');
        srcSheet.appendRow(['Source_Name', 'Created_At']);
        srcSheet.appendRow(['Website', new Date()]);
        srcSheet.appendRow(['LinkedIn', new Date()]);
        srcSheet.appendRow(['Facebook', new Date()]);
        srcSheet.appendRow(['Referral', new Date()]);
    }
    
    // 7. EMAIL TEMPLATES
    let etSheet = ss.getSheetByName('EmailTemplates');
    if(!etSheet) {
        etSheet = ss.insertSheet('EmailTemplates');
        etSheet.appendRow(['ID', 'Name', 'Subject', 'Body', 'Last_Updated']);
    }
    
    // Check/Reset EVALUATIONS sheet
    let evSheet = ss.getSheetByName('EVALUATIONS'); // Assuming EVALUATION_SHEET_NAME is defined elsewhere
    if (evSheet && forceReset) {
      ss.deleteSheet(evSheet);
      evSheet = null;
    }
    
    if (!evSheet) {
        evSheet = ss.insertSheet('EVALUATIONS');
        evSheet.appendRow([
          'ID', 'Candidate_ID', 'Candidate_Name', 'Position', 'Department', 
          'Recruiter_Email', 'Manager_Email', 'Created_At', 'Status', 
          'Score_Professional', 'Score_Soft_Skills', 'Score_Culture', 
          'Final_Result', 'Manager_Comment', 'Completed_At',
          'Batch_ID', 'Criteria_Config', 'Scores_JSON',
          'Manager_Name', 'Manager_Department', 'Manager_Position', 
          'Total_Score', 'Proposed_Salary', 'Signature_Status'
        ]);
        // Format header
        evSheet.getRange(1, 1, 1, 24).setBackground('#f1f3f5').setFontWeight('bold');
        // Assuming logToSheet is defined elsewhere
        // logToSheet('EVALUATIONS sheet created/reset with 24 columns.');
    } else {
        // Safe update for existing sheet (ensure it has ALL 24 columns)
        const currentHeaders = evSheet.getRange(1, 1, 1, Math.max(evSheet.getLastColumn(), 1)).getValues()[0].map(h => String(h).trim().toLowerCase());
        const expectedHeaders = [
          'ID', 'Candidate_ID', 'Candidate_Name', 'Position', 'Department', 
          'Recruiter_Email', 'Manager_Email', 'Created_At', 'Status', 
          'Score_Professional', 'Score_Soft_Skills', 'Score_Culture', 
          'Final_Result', 'Manager_Comment', 'Completed_At',
          'Batch_ID', 'Criteria_Config', 'Scores_JSON',
          'Manager_Name', 'Manager_Department', 'Manager_Position', 
          'Total_Score', 'Proposed_Salary', 'Signature_Status'
        ];
        
        expectedHeaders.forEach(hName => {
          if (currentHeaders.indexOf(hName.toLowerCase()) === -1) {
            evSheet.getRange(1, evSheet.getLastColumn() + 1).setValue(hName);
            currentHeaders.push(hName.toLowerCase());
          }
        });
    }
    
  } catch (e) {
    Logger.log('Structure Init Error: ' + e.toString());
  }
}

function apiResetEvaluationSheet() {
  // Assuming logToSheet is defined elsewhere
  // logToSheet('Manual Reset of EVALUATIONS sheet requested.');
  ensureSheetStructure(true);
  return { success: true, message: 'Đã reset bảng EVALUATIONS thành công. Hệ thống đã tạo lại 24 cột chuẩn.' };
}

function apiSetupDatabase() {
    ensureSheetStructure();
    return { success: true, url: 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID };
}

// 6. API: CANDIDATE MANAGEMENT
function apiCreateCandidate(formObject, fileData) {
  try {
    const sheet = getSheetByName(CANDIDATE_SHEET_NAME);
    if (!sheet) return { success: false, message: 'KhÃ´ng tÃ¬m tháº¥y Sheet: ' + CANDIDATE_SHEET_NAME };

    const newId = 'C' + new Date().getTime(); 
    const appliedDate = new Date().toISOString().slice(0, 10); 
    
    // 1. Handle File Upload
    let cvLink = '';
    if(fileData && fileData.data && fileData.name) {
        try {
            const folder = DriveApp.getFolderById(CV_FOLDER_ID);
            const contentType = fileData.type || 'application/pdf';
            const blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), contentType, fileData.name);
            
            if(formObject.Phone || formObject.phone) {
                const ext = fileData.name.split('.').pop();
                blob.setName((formObject.Phone || formObject.phone) + '.' + ext);
            }
            
            const file = folder.createFile(blob);
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            cvLink = file.getUrl();
        } catch(err) {
            Logger.log('Upload Error: ' + err);
            cvLink = 'Error Uploading: ' + err.toString(); 
        }
    }

    // 2. Ensure all columns exist (Case-Insensitive)
    const requiredKeys = Object.keys(CANDIDATE_ALIASES);
    let headers = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
    let lowerHeaders = headers.map(h => h.toString().trim().toLowerCase());
    
    requiredKeys.forEach(key => {
        let exists = lowerHeaders.includes(key.toLowerCase());
        if (!exists) {
            // Check aliases
            if (CANDIDATE_ALIASES[key]) {
                exists = CANDIDATE_ALIASES[key].some(alias => lowerHeaders.includes(alias.toString().toLowerCase()));
            }
        }
        if (!exists) {
            const newHeader = key; 
            sheet.getRange(1, sheet.getLastColumn() + 1).setValue(newHeader);
            headers.push(newHeader);
            lowerHeaders.push(newHeader.toLowerCase());
        }
    });

    // 3. Map Data to Row
    const row = new Array(headers.length).fill('');
    
    // Normalize formObject keys to PascalCase (Internal consistency)
    const normalizedData = {
        ID: newId,
        Applied_Date: appliedDate,
        CV_Link: cvLink || formObject.CV_Link || formObject.cv_link || ''
    };
    
    if (cvLink) normalizedData.CV_Link = cvLink; // Force priority for uploaded link

    // Merge from formObject
    Object.keys(formObject).forEach(k => {
        let normalizedKey = k;
        // Simple mapping for common lowercase variants
        if (k === 'name' || k === 'FullName') normalizedKey = 'Name';
        if (k === 'gender') normalizedKey = 'Gender';
        if (k === 'phone') normalizedKey = 'Phone';
        if (k === 'email') normalizedKey = 'Email';
        if (k === 'department') normalizedKey = 'Department';
        if (k === 'position') normalizedKey = 'Position';
        if (k === 'experience') normalizedKey = 'Experience';
        if (k === 'school') normalizedKey = 'School';
        if (k === 'education_level') normalizedKey = 'Education_Level';
        if (k === 'major') normalizedKey = 'Major';
        if (k === 'dob') normalizedKey = 'Birth_Year';
        if (k === 'expected_salary' || k === 'salary' || k === 'Salary_Expectation') normalizedKey = 'Salary_Expectation';
        if (k === 'source') normalizedKey = 'Source';
        if (k === 'recruiter') normalizedKey = 'Recruiter';
        if (k === 'contact_status' || k === 'status') normalizedKey = 'Status';
        if (k === 'stage') normalizedKey = 'Stage';
        if (k === 'ticket_id' || k === 'TicketID') normalizedKey = 'TicketID';
        if (k === 'notes' || k === 'NewNote') normalizedKey = 'NewNote';

        // PROTECT system fields & Preserve CV_Link if already set by upload
        if (normalizedKey !== 'ID' && normalizedKey !== 'Applied_Date' && normalizedKey !== 'CV_Link') {
            normalizedData[normalizedKey] = formObject[k];
        }
    });

    if (cvLink) normalizedData.CV_Link = cvLink; // Force priority for uploaded link (Override anything from form)

    // Map to columns using Robust findColIndex logic
    Object.keys(CANDIDATE_ALIASES).forEach(key => {
        let colIndex = findColIndexInternal(headers, key);
        if (colIndex !== -1) {
            let val = normalizedData[key] || '';
            // Handle specific overrides if needed
            if (key === 'Status' && !val) val = normalizedData['Stage']; 
            row[colIndex] = val;
        }
    });

    sheet.appendRow(row);
    SpreadsheetApp.flush(); 

    // Helper for Internal Mapping
    function findColIndexInternal(headerList, key) {
        let cleanHeaders = headerList.map(h => h.toString().trim().toLowerCase());
        
        // 1. Try exact key
        let idx = cleanHeaders.indexOf(key.toLowerCase());
        if (idx !== -1) return idx;
        
        // 2. Try aliases
        const aliases = CANDIDATE_ALIASES[key] || [];
        for (let alias of aliases) {
            let aIdx = cleanHeaders.indexOf(alias.toString().trim().toLowerCase());
            if (aIdx !== -1) return aIdx;
        }
        return -1;
    }
    // Handle initial note for new candidate
    if (normalizedData.NewNote) {
        // We reuse apiUpdateCandidate logic for note appending if we want consistent formatting
        // but for a new row it's easier to just put it in 'Notes' column
        let notesCol = headers.findIndex(h => CANDIDATE_ALIASES['Notes'].includes(h.toString()) || h.toString() === 'Notes');
        if (notesCol !== -1) {
            const user = normalizedData.User || 'Admin';
            const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm dd/MM/yyyy");
            const entry = `[${time}] ${user}: ${normalizedData.NewNote}`;
            sheet.getRange(sheet.getLastRow(), notesCol + 1).setValue(entry);
        }
    }
    
    // LOG ACTIVITY
    const userEmail = normalizedData.User || Session.getActiveUser().getEmail() || 'Admin';
    logActivity(userEmail, 'ThÃªm á»©ng viÃªn', `ThÃªm á»©ng viÃªn má»›i: ${normalizedData.Name}`, newId);
    
    return { success: true, message: 'ThÃªm há»“ sÆ¡ thÃ nh cÃ´ng!', data: apiGetInitialData(), candidateId: newId };
  } catch (e) {
    Logger.log('apiCreateCandidate Error: ' + e.toString());
    return { success: false, message: 'Lá»—i server: ' + e.toString() };
  }
}

function apiUpdateCandidateStatus(candidateId, newStage, rejectionData = null) {
    Logger.log('--- apiUpdateCandidateStatus (Kanban Move) ---');
    try {
        const sheet = getSheetByName(CANDIDATE_SHEET_NAME);
        if (!sheet) return { success: false, message: 'Sheet not found' };
        
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        
        // Helper to find column index with CASE-INSENSITIVE aliases
        const findIdx = (key) => {
            let lowerHeaders = headers.map(h => h.toString().trim().toLowerCase());
            let idx = lowerHeaders.indexOf(key.toLowerCase());
            if (idx !== -1) return idx;
            const aliases = CANDIDATE_ALIASES[key] || [];
            for (let alias of aliases) {
                let aIdx = lowerHeaders.indexOf(alias.toString().toLowerCase());
                if (aIdx !== -1) return aIdx;
            }
            return -1;
        };

        const idIndex = findIdx('ID');
        const stageIndex = findIdx('Stage'); // We want to update STAGE (VÃ²ng tuyá»ƒn dá»¥ng)
        
        const hireDateIndex = findIdx('Hire_Date');
        const rejectTypeIndex = findIdx('Rejection_Type');
        const rejectReasonIndex = findIdx('Rejection_Reason');

        if (idIndex === -1 || stageIndex === -1) {
            return { success: false, message: 'Missing required columns (ID or Stage)' };
        }
        
        for (let i = 1; i < data.length; i++) {
            if (data[i][idIndex] == candidateId) {
                // Update Stage
                sheet.getRange(i + 1, stageIndex + 1).setValue(newStage);
                
                // Automate Hire Date
                if (hireDateIndex !== -1) {
                    if (newStage === 'Hired' || newStage === 'Đã tuyển' || newStage === 'Nhận việc' || newStage === 'Official') {
                        if (!data[i][hireDateIndex]) {
                           sheet.getRange(i + 1, hireDateIndex + 1).setValue(new Date());
                        }
                    }
                }

                // Update Rejection Data
                if (rejectionData) {
                    if (rejectTypeIndex !== -1) sheet.getRange(i + 1, rejectTypeIndex + 1).setValue(rejectionData.type || '');
                    if (rejectReasonIndex !== -1) sheet.getRange(i + 1, rejectReasonIndex + 1).setValue(rejectionData.reason || '');
                } else if (!newStage.toLowerCase().includes('loại') && !newStage.toLowerCase().includes('reject')) {
                    // Reset rejection info if moved out of rejected
                    if (rejectTypeIndex !== -1) sheet.getRange(i + 1, rejectTypeIndex + 1).setValue('');
                    if (rejectReasonIndex !== -1) sheet.getRange(i + 1, rejectReasonIndex + 1).setValue('');
                }
                
                // LOG ACTIVITY
                const userEmail = Session.getActiveUser().getEmail() || 'Admin'; 
                const candidateName = data[i][findIdx('Name')] || candidateId;
                logActivity(userEmail, 'Chuyá»ƒn vÃ²ng', `Chuyá»ƒn á»©ng viÃªn **${candidateName}** sang vÃ²ng **${newStage}**`, candidateId);
                
                return { success: true };
            }
        }
        return { success: false, message: 'Candidate not found' };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// UPDATE FULL CANDIDATE DETAILS
function apiUpdateCandidate(candidateData, fileData) {
  Logger.log('=== UPDATE CANDIDATE (FULL) ===');
  Logger.log('Data: ' + JSON.stringify(candidateData));
  
  try {
    const sheet = getSheetByName(CANDIDATE_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    let headers = data[0];
    const idColIndex = headers.findIndex(h => h.toString().toLowerCase() === 'id');
    
    if (idColIndex === -1) {
      return { success: false, message: 'ID column not found' };
    }

    // Helper to find column index with CASE-INSENSITIVE aliases
    const findColIndex = (key) => {
      let cleanHeaders = headers.map(h => h.toString().trim().toLowerCase());
      let searchKey = key.toString().trim().toLowerCase();
      
      let idx = cleanHeaders.indexOf(searchKey);
      if (idx !== -1) return idx;
      
      const aliases = CANDIDATE_ALIASES[key] || [];
      for (let alias of aliases) {
        let aIdx = cleanHeaders.indexOf(alias.toString().trim().toLowerCase());
        if (aIdx !== -1) return aIdx;
      }
      return -1;
    };
    
    // 1. Handle File Upload (If new file is provided during update)
    let cvLink = '';
    if(fileData && fileData.data && fileData.name) {
        try {
            const folder = DriveApp.getFolderById(CV_FOLDER_ID);
            const contentType = fileData.type || 'application/pdf';
            const blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), contentType, fileData.name);
            
            if(candidateData.Phone || candidateData.phone) {
                const ext = fileData.name.split('.').pop();
                blob.setName((candidateData.Phone || candidateData.phone) + '.' + ext);
            }
            
            const file = folder.createFile(blob);
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            cvLink = file.getUrl();
            candidateData.CV_Link = cvLink; // Override link with the new one
        } catch(err) {
            Logger.log('Upload Error during update: ' + err);
        }
    }
    
    // 2. AUTO-CREATE MISSING COLUMNS (Using alias awareness)
    const missingColumns = [];
    Object.keys(candidateData).forEach(key => {
      if (key === 'ID' || key === 'NewNote' || key === 'User') return;
      
      // Check if this key or its aliases exist already
      let colIndex = findColIndex(key);
      if (colIndex === -1) {
        missingColumns.push(key);
      }
    });
    
    if (missingColumns.length > 0) {
      Logger.log('Creating missing columns during update: ' + missingColumns.join(', '));
      const lastCol = sheet.getLastColumn();
      missingColumns.forEach((colName, index) => {
        sheet.getRange(1, lastCol + index + 1).setValue(colName);
        headers.push(colName); // Add to local headers array for immediate use
      });
      SpreadsheetApp.flush();
    }
    
    // Normalize Data (In case of lowercase keys from form)
    const normalizedData = {};
    Object.keys(candidateData).forEach(k => {
        let normalizedKey = k;
        if (k === 'name') normalizedKey = 'Name';
        if (k === 'gender') normalizedKey = 'Gender';
        if (k === 'phone') normalizedKey = 'Phone';
        if (k === 'email') normalizedKey = 'Email';
        if (k === 'department') normalizedKey = 'Department';
        if (k === 'position') normalizedKey = 'Position';
        if (k === 'experience') normalizedKey = 'Experience';
        if (k === 'school') normalizedKey = 'School';
        if (k === 'education_level') normalizedKey = 'Education_Level';
        if (k === 'major') normalizedKey = 'Major';
        if (k === 'dob') normalizedKey = 'Birth_Year';
        if (k === 'expected_salary' || k === 'salary') normalizedKey = 'Salary_Expectation';
        if (k === 'source') normalizedKey = 'Source';
        if (k === 'recruiter') normalizedKey = 'Recruiter';
        if (k === 'contact_status' || k === 'status') normalizedKey = 'Status';
        if (k === 'stage') normalizedKey = 'Stage';
        if (k === 'ticket_id' || k === 'TicketID') normalizedKey = 'TicketID';
        if (k === 'notes' || k === 'NewNote') normalizedKey = 'NewNote';
        if (k === 'Rejection_Type') normalizedKey = 'Rejection_Type';
        if (k === 'Rejection_Reason') normalizedKey = 'Rejection_Reason';
        if (k === 'Rejection_Source') normalizedKey = 'Rejection_Source';
        
        normalizedData[normalizedKey] = candidateData[k];
    });

    // 3. Find row and update
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] == candidateData.ID) {
        Logger.log('Found candidate at row ' + (i + 1));
        
        // Update fields based on CANDIDATE_ALIASES keys
        Object.keys(CANDIDATE_ALIASES).forEach(key => {
          if (key === 'ID' || key === 'NewNote' || key === 'User' || key === 'Notes') return;

          let colIndex = findColIndex(key);
          if (colIndex !== -1) {
            let value = normalizedData[key];
            
            // If the field is missing from the update data, don't overwrite with empty
            // UNLESS it's explicitly provided as an empty string (meaning clear field)
            if (value !== undefined) {
                // Special handling for Hire_Date if it's a string from frontend
                if (key === 'Hire_Date' && value && typeof value === 'string') {
                   value = new Date(value);
                }
                sheet.getRange(i + 1, colIndex + 1).setValue(value || '');
            }
          }
        });
        
        // Auto Hire Date logic
        const stageIdx = findColIndex('Stage');
        const hireDateIdx = findColIndex('Hire_Date');
        if (stageIdx !== -1 && hireDateIdx !== -1) {
            const newStage = normalizedData.Stage;
            if (newStage === 'Hired' || newStage === 'Đã tuyển' || newStage === 'Nhận việc' || newStage === 'Official') {
                 if (!data[i][hireDateIdx]) {
                    sheet.getRange(i + 1, hireDateIdx + 1).setValue(new Date());
                 }
            }
        }


        // HANDLE NEW NOTE APPENDING
        if (candidateData.NewNote) {
            let notesIndex = findColIndex('Notes');

            if (notesIndex !== -1) {
                const currentNotes = data[i][notesIndex] || '';
                const user = candidateData.User || 'User';
                const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm dd/MM/yyyy");
                const newEntry = `[${time}] ${user}: ${candidateData.NewNote}`;
                const updatedNotes = currentNotes ? currentNotes + "\n" + newEntry : newEntry;
                
                sheet.getRange(i + 1, notesIndex + 1).setValue(updatedNotes);

                Logger.log('Appended note: ' + newEntry);

                // --- NOTIFICATION FOR MENTIONS ---
                try {
                    const mentionRegex = /@([a-zA-Z0-9_\u00C0-\u1EF9]+)/g; 
                    const mentions = [...candidateData.NewNote.matchAll(mentionRegex)].map(m => m[1]);
                    
                    if (mentions.length > 0) {
                        const users = apiGetTableData('Users');
                        const uniqueMentions = [...new Set(mentions)];
                        const sender = candidateData.User || 'Someone';
                        const candidateName = getCandidateNameById(candidateData.ID);

                        uniqueMentions.forEach(mention => {
                            const targetUser = users.find(u => {
                                const uName = String(u.Username).toLowerCase();
                                const fName = u.Full_Name ? String(u.Full_Name).replace(/\s+/g, '_').toLowerCase() : '';
                                const search = mention.toLowerCase();
                                return uName === search || fName === search;
                            });

                            if (targetUser) {
                                createNotification(
                                    targetUser.Username, 
                                    'Mention',
                                    `${sender} đã nhắc đến bạn trong hồ sơ ${candidateName}`,
                                    candidateData.ID
                                );
                            }
                        });
                    }
                } catch (err) {
                    Logger.log('Error processing mentions: ' + err.toString());
                }
                // ---------------------------------
            }
        }
        
        if (candidateData.NewNote) {
            const userEmail = candidateData.User || Session.getActiveUser().getEmail() || 'Admin';
            const candidateName = getCandidateNameById(candidateData.ID);
            logActivity(userEmail, 'ThÃªm ghi chÃº', `ThÃªm highlight cho á»©ng viÃªn **${candidateName}**: ${candidateData.NewNote}`, candidateData.ID);
        } else {
             const userEmail = candidateData.User || Session.getActiveUser().getEmail() || 'Admin';
             const candidateName = getCandidateNameById(candidateData.ID);
             logActivity(userEmail, 'Cáº­p nháº­t há»“ sÆ¡', `Cáº­p nháº­t thÃ´ng tin chi tiáº¿t cho á»©ng viÃªn **${candidateName}**`, candidateData.ID);
        }
        
        Logger.log('Successfully updated candidate at row ' + (i + 1));
        
        // Return updated data
        return {
          success: true,
          message: 'Cáº­p nháº­t thÃ nh cÃ´ng!',
          data: apiGetInitialData() // Refresh all data
        };
      }
    }
    
    return { success: false, message: 'Candidate ID not found: ' + candidateData.ID };
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// DELETE CANDIDATE
function apiDeleteCandidate(candidateId) {
  Logger.log('=== DELETE CANDIDATE ===');
  Logger.log('ID: ' + candidateId);
  
  try {
    const sheet = getSheetByName(CANDIDATE_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.findIndex(h => h.toString().toLowerCase() === 'id');
    
    if (idColIndex === -1) {
      return { success: false, message: 'ID column not found' };
    }
    
    // Find and delete row with matching ID
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] == candidateId) {
        Logger.log('Found candidate at row ' + (i + 1) + ', deleting...');
        sheet.deleteRow(i + 1);
        Logger.log('Successfully deleted candidate');
        
        return {
          success: true,
          message: 'ÄÃ£ xÃ³a á»©ng viÃªn thÃ nh cÃ´ng!'
        };
      }
    }
    
    return { success: false, message: 'Candidate ID not found: ' + candidateId };
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// 7. API: JOB MANAGEMENT
// 7. API: JOB MANAGEMENT
// 7. API: JOB MANAGEMENT
// 7. API: JOB MANAGEMENT
function apiGetJobs() {
    if (!SPREADSHEET_ID) return [{ID: 'ERR', Title: 'No SPREADSHEET_ID'}];
    try {
        const jobs = apiGetTableData('Jobs');
        
        if (!jobs || jobs.length === 0) {
            return [{ID: 'WARN', Title: 'Debug: apiGetTableData returned 0 rows', Status: 'Debug'}];
        }

        // Map keys for Admin UI consistency
        const mappedJobs = jobs.map(j => {
            return {
                ID: j.ID || j.id,
                Title: j.Title || j['Vá»‹ trÃ­'] || j['TiÃªu Ä‘á»'] || '',
                Department: j.Department || j['PhÃ²ng ban'] || '',
                Location: j.Location || j['Äá»‹a Ä‘iá»ƒm'] || '',
                Type: j.Type || j['Loáº¡i hÃ¬nh'] || '',
                Status: j.Status || j['Tráº¡ng thÃ¡i'] || '',
                Description: j.Description || j['MÃ´ táº£'] || j['MÃ´ táº£ cÃ´ng viá»‡c'] || '',
                Created_Date: j.Created_Date || j['NgÃ y táº¡o'] || ''
            };
        });
        return mappedJobs.reverse(); // Newest first
    } catch (e) {
        return [{ID: 'CRITICAL', Title: 'Error: ' + e.toString(), Status: 'Error'}];
    }
}

// DEBUG FUNCTION V2 (Trace Mode)
function apiDebugJobsV2() {
    let trace = "1. Start. ";
    try {
        if (!SPREADSHEET_ID) return { trace: trace + "Error: SPREADSHEET_ID missing" };
        trace += "2. ID: " + SPREADSHEET_ID.substring(0,5) + "... ";
        
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        trace += "3. SS Opened. ";
        
        const sheet = ss.getSheetByName('Jobs');
        trace += "4. Sheet 'Jobs': " + (sheet ? "Found" : "Not Found") + ". ";
        
        if (!sheet) {
             const allSheets = ss.getSheets().map(s => s.getName());
             return { trace: trace + "Available Sheets: " + allSheets.join(', ') };
        }
        
        const data = sheet.getDataRange().getValues();
        trace += "5. Data Rows: " + data.length + ". ";
        
        const headers = data.length > 0 ? data[0].map(String) : [];
        return { 
            success: true, 
            trace: trace + "Done.",
            headers: headers,
            rowCount: data.length
        };
    } catch (e) {
        return { trace: trace + "EXCEPTION: " + e.toString() };
    }
}

function apiCreateJob(formObject) {
    try {
        let sheet = getSheetByName('Jobs');
        const headers = ['ID', 'Title', 'Department', 'Position', 'Location', 'Type', 'Status', 'Created_Date', 'Description', 'TicketID'];
        
        if (!sheet) {
             const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
             sheet = ss.insertSheet('Jobs');
             sheet.appendRow(headers);
        }
        
        // Handle empty sheet case
        if (sheet.getLastColumn() === 0) {
             sheet.appendRow(headers);
        }

        const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const newId = 'J' + new Date().getTime();
        const createdDate = new Date().toISOString().slice(0, 10);

        // Map data to headers dynamically
        const rowData = currentHeaders.map(h => {
            const head = h.toString().trim();
            if (head === 'ID') return newId;
            if (head === 'Title') return formObject.title || '';
            if (head === 'Department') return formObject.department || '';
            if (head === 'Position') return formObject.position || '';
            if (head === 'Location') return formObject.location || '';
            if (head === 'Type') return formObject.type || '';
            if (head === 'Status') return 'Open';
            if (head === 'Created_Date') return createdDate;
            if (head === 'Description') return formObject.description || '';
            if (head === 'TicketID') return formObject.ticketId || '';
            return '';
        });
        
        sheet.appendRow(rowData);
        
        return { success: true, message: 'Đã tạo tin tuyển dụng thành công!' };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function apiDeleteJob(jobId) {
    return apiDeleteRow('Jobs', jobId);
}

function apiUpdateJobStatus(jobId, newStatus) {
    try {
        const sheet = getSheetByName('Jobs');
        if (!sheet) return { success: false, message: 'Sheet Jobs not found' };
        
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        
        // Find ID and Status columns (English or Vietnamese)
        const idIndex = headers.findIndex(h => h.toString().trim().toLowerCase() === 'id');
        let statusIndex = headers.findIndex(h => h.toString().trim().toLowerCase() === 'status');
        if (statusIndex === -1) {
             statusIndex = headers.findIndex(h => h.toString().trim().toLowerCase() === 'tráº¡ng thÃ¡i');
        }
        
        if (idIndex === -1) return { success: false, message: 'ID column not found' };
        if (statusIndex === -1) return { success: false, message: 'Status column not found' };
        
        for (let i = 1; i < data.length; i++) {
            if (data[i][idIndex] == jobId) {
                sheet.getRange(i + 1, statusIndex + 1).setValue(newStatus);
                return { success: true, message: 'Cáº­p nháº­t tráº¡ng thÃ¡i thÃ nh cÃ´ng' };
            }
        }
        return { success: false, message: 'Job not found' };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// Reuse generic delete if available, or implement simple one
function apiDeleteRow(sheetName, id) {
    try {
        const sheet = getSheetByName(sheetName);
        if (!sheet) return { success: false, message: 'Sheet not found' };
        
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const idIndex = headers.findIndex(h => h.toString().trim().toLowerCase() === 'id');
        
        if (idIndex === -1) return { success: false, message: 'ID column not found' };
        
        for (let i = 1; i < data.length; i++) {
            if (data[i][idIndex] == id) {
                sheet.deleteRow(i + 1);
                return { success: true, message: 'ÄÃ£ xÃ³a thÃ nh cÃ´ng' };
            }
        }
        return { success: false, message: 'ID not found' };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// [DELETED DUPLICATE apiGetOpenJobs]

// 8. API: SETTINGS & ADMINISTRATION
// 8. API: SETTINGS & ADMINISTRATION
function apiGetSettings() {
    // STAGES are now project-specific or defaulted.
    const defaultStages = [
         {Stage_Name: 'Apply', Order: 1, Color: '#0d6efd'},
         {Stage_Name: 'Interview', Order: 2, Color: '#fd7e14'},
         {Stage_Name: 'Offer', Order: 3, Color: '#198754'},
         {Stage_Name: 'Rejected', Order: 4, Color: '#dc3545'}
    ];

    return {
        users: apiGetTableData('Users'),
        stages: defaultStages, // Return defaults for compatibility
        departments: apiGetTableData('Departments'),
        companyInfo: apiGetCompanyInfo().data
    };
}

function apiSaveStages(stagesArray) {
    return { success: false, message: 'Cấu hình quy trình chung đã bị bãi bỏ. Vui lòng cấu hình quy trình theo từng Dự án.' };
}

function apiCreateUser(user) {
    try {
        let sheet = getSheetByName(USERS_SHEET_NAME);
        if (!sheet) {
             const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
             sheet = ss.insertSheet(USERS_SHEET_NAME);
             // Username, Password, Full_Name, Role, Email, Department, Phone
             sheet.appendRow(['Username', 'Password', 'Full_Name', 'Role', 'Email', 'Department', 'Phone']);
        }
        
        // Check duplicate
        const data = sheet.getDataRange().getValues();
        for(let i=1; i<data.length; i++) {
             if(data[i][0] == user.username) {
                  return { success: false, message: 'TÃªn Ä‘Äƒng nháº­p Ä‘Ã£ tá»“n táº¡i' };
             }
        }
        
        sheet.appendRow([
            user.username, 
            user.password, 
            user.fullname, 
            user.role, 
            user.email || '',       // Col 5
            user.department,        // Col 6
            user.phone || ''        // Col 7
        ]);
        return { success: true, message: 'ÄÃ£ thÃªm ngÆ°á»i dÃ¹ng' };
    } catch(e) {
        return { success: false, message: e.toString() };
    }
}

function apiEditUser(user) {
    try {
        const sheet = getSheetByName(USERS_SHEET_NAME);
        if(!sheet) return { success: false, message: 'Users sheet not found' };
        
        const data = sheet.getDataRange().getValues();
        for(let i=1; i<data.length; i++) {
            if(data[i][0] == user.username) {
                // Determine layout based on headers
                // Assuming standard: Username(0), Password(1), Full_Name(2), Role(3), Department(4), Email(5), Phone(6)
                
                // Update fields (except username)
                if(user.password) sheet.getRange(i+1, 2).setValue(user.password);
                sheet.getRange(i+1, 3).setValue(user.fullname);
                sheet.getRange(i+1, 4).setValue(user.role);
                sheet.getRange(i+1, 5).setValue(user.email || '');  // Col 5: Email
                sheet.getRange(i+1, 6).setValue(user.department);   // Col 6: Department
                // Check if Phone column exists (Col 7)
                if (data[0].length >= 7) {
                     sheet.getRange(i+1, 7).setValue(user.phone || '');
                } else if (user.phone) {
                     // If column missing but data provided, might need to handle, but for now just skip or simple append?
                     // Safer to assume column exists or was added by getTableData logic
                }
                
                return { success: true, message: 'ÄÃ£ cáº­p nháº­t thÃ´ng tin user' };
            }
        }
        return { success: false, message: 'User not found' };
    } catch(e) {
        return { success: false, message: e.toString() };
    }
}

function apiDeleteUser(username) {
    try {
        const sheet = getSheetByName(USERS_SHEET_NAME);
        if(!sheet) return { success: false, message: 'Users sheet not found' };
        
        const data = sheet.getDataRange().getValues();
        for(let i=1; i<data.length; i++) {
            if(data[i][0] == username) {
                sheet.deleteRow(i+1);
                return { success: true, message: 'ÄÃ£ xÃ³a user' };
            }
        }
        return { success: false, message: 'User not found' };
    } catch(e) {
        return { success: false, message: e.toString() };
    }
}

function apiGetEmailTemplates() {
    try {
        let sheet = getSheetByName('Email_Templates');
        if (!sheet) {
            const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
            sheet = ss.insertSheet('Email_Templates');
            sheet.appendRow(['ID', 'Name', 'Subject', 'Body']);
            // Add defaults
            sheet.appendRow(['1', 'Offer Email', 'Má»i nháº­n viá»‡c - [Candidate Name]', 'ChÃ o [Name],\n\nChÃºc má»«ng báº¡n Ä‘Ã£ trÃºng tuyá»ƒn...']);
            sheet.appendRow(['2', 'Reject Email', 'ThÃ´ng bÃ¡o káº¿t quáº£ phá»ng váº¥n', 'ChÃ o [Name],\n\nCáº£m Æ¡n báº¡n Ä‘Ã£ quan tÃ¢m...']);
            sheet.appendRow(['3', 'Interview Invite', 'Má»i phá»ng váº¥n', 'ChÃ o [Name],\n\nChÃºng tÃ´i muá»‘n má»i báº¡n tham gia phá»ng váº¥n...']);
        }
        
        // Handle empty/new sheet
        if (sheet.getLastColumn() === 0) {
            sheet.appendRow(['ID', 'Name', 'Subject', 'Body']);
            // Add defaults
            sheet.appendRow(['1', 'Offer Email', 'Má»i nháº­n viá»‡c - [Candidate Name]', 'ChÃ o [Name],\n\nChÃºc má»«ng báº¡n Ä‘Ã£ trÃºng tuyá»ƒn...']);
            sheet.appendRow(['2', 'Reject Email', 'ThÃ´ng bÃ¡o káº¿t quáº£ phá»ng váº¥n', 'ChÃ o [Name],\n\nCáº£m Æ¡n báº¡n Ä‘Ã£ quan tÃ¢m...']);
            sheet.appendRow(['3', 'Interview Invite', 'Má»i phá»ng váº¥n', 'ChÃ o [Name],\n\nChÃºng tÃ´i muá»‘n má»i báº¡n tham gia phá»ng váº¥n...']);
        }
        
        return apiGetTableData('Email_Templates');
    } catch (e) {
        return [];
    }
}

// Redundant functions removed for cleanup



// 6. RECRUITER MANAGEMENT
const RECRUITER_SHEET_NAME = 'DATA_RECRUITERS';

function apiGetRecruiters() {
  Logger.log('=== GET RECRUITERS ===');
  try {
    let sheet = getSheetByName(RECRUITER_SHEET_NAME);
    
    // Check/Create Sheet
    if (!sheet) {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        sheet = ss.insertSheet(RECRUITER_SHEET_NAME);
        // New Headers: ID, Name, Email, Position, JoinDate
        sheet.getRange('A1:E1').setValues([['ID', 'TÃªn', 'Email', 'Chá»©c vá»¥', 'NgÃ y tham gia']]);
        // Add sample data
        const sampleId = 'REC' + new Date().getTime();
        const today = new Date().toISOString().slice(0, 10);
        sheet.getRange('A2:E2').setValues([[sampleId, 'Admin', 'admin@example.com', 'Quáº£n trá»‹ viÃªn', today]]);
        return { 
            success: true, 
            recruiters: [{
                id: sampleId,
                name: 'Admin', 
                email: 'admin@example.com',
                position: 'Quáº£n trá»‹ viÃªn',
                joinDate: today,
                totalCandidates: 0
            }] 
        };
    }
    
    // 1. Get Recruiters Data
    const data = sheet.getDataRange().getValues();
    const recruiters = [];
    
    // 2. Get Candidate Counts
    const candidateSheet = getSheetByName(CANDIDATE_SHEET_NAME);
    const candidateCounts = {}; // name -> count
    
    if (candidateSheet && candidateSheet.getLastRow() > 1) {
        const candData = candidateSheet.getDataRange().getValues();
        const headers = candData[0].map(h => h.toString().trim());
        
        // Find 'Recruiter' OR 'NgÆ°á»i phá»¥ trÃ¡ch' column
        let recColIndex = headers.findIndex(h => h === 'Recruiter' || h === 'NgÆ°á»i phá»¥ trÃ¡ch' || h.toLowerCase() === 'recruiter');
        
        if (recColIndex > -1) {
            for (let i = 1; i < candData.length; i++) {
                let recName = candData[i][recColIndex];
                if (recName) {
                    recName = recName.toString().trim();
                    candidateCounts[recName] = (candidateCounts[recName] || 0) + 1;
                }
            }
        }
    }

    // 3. Map Recruiters
    // Skip header (Row 1)
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        let id = row[0];
        let name = row[1];
        let email = row[2];
        let position = row[3] || '';
        let joinDate = row[4] || '';
        
        if (name || id) {
             const cleanName = name ? name.toString().trim() : '';
             recruiters.push({
                 id: id,
                 name: name,
                 email: email,
                 position: position,
                 joinDate: joinDate instanceof Date ? joinDate.toISOString().slice(0,10) : joinDate,
                 totalCandidates: candidateCounts[cleanName] || 0
             });
        }
    }
    return { success: true, recruiters: recruiters };
  } catch (e) {
    Logger.log('Error getting recruiters: ' + e);
    return { success: false, message: e.toString(), recruiters: [] };
  }
}

function apiAddRecruiter(data) {
    Logger.log('=== ADD RECRUITER ===');
    try {
        let sheet = getSheetByName(RECRUITER_SHEET_NAME);
        if (!sheet) {
            // Should be created by getRecruiters, but safe to check
             return { success: false, message: 'Vui lÃ²ng táº£i láº¡i trang Ä‘á»ƒ khá»Ÿi táº¡o dá»¯ liá»‡u.' };
        }
        
        // Data: { name, email, position, joinDate }
        const id = 'REC' + new Date().getTime();
        const row = [
            id,
            data.name,
            data.email,
            data.position,
            data.joinDate
        ];
        
        // Check duplicate NAME (optional, but good for linking)
        const currentData = sheet.getDataRange().getValues();
        for(let i=1; i<currentData.length; i++) {
            if(currentData[i][1] === data.name) { // Column 1 is Name
                return { success: false, message: 'TÃªn nhÃ¢n viÃªn Ä‘Ã£ tá»“n táº¡i: ' + data.name };
            }
        }
        
        sheet.appendRow(row);
        return { success: true, message: 'ÄÃ£ thÃªm thÃ nh cÃ´ng' };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function apiEditRecruiter(data) {
    Logger.log('=== EDIT RECRUITER: ' + data.id + ' ===');
    try {
        const sheet = getSheetByName(RECRUITER_SHEET_NAME);
        if (!sheet) return { success: false, message: 'Sheet not found' };
        
        const rows = sheet.getDataRange().getValues();
        // Row 1 is header.
        for (let i = 1; i < rows.length; i++) {
            if (String(rows[i][0]) === String(data.id)) { // Column 0 is ID
                // Found! Update content
                // Columns: ID(0), Name(1), Email(2), Position(3), JoinDate(4)
                
                // Only update Name if it doesn't conflict? 
                // For simplicity, allow update.
                
                sheet.getRange(i + 1, 2).setValue(data.name);    // Name (Col B)
                sheet.getRange(i + 1, 3).setValue(data.email);   // Email (Col C)
                sheet.getRange(i + 1, 4).setValue(data.position);// Position (Col D)
                sheet.getRange(i + 1, 5).setValue(data.joinDate);// JoinDate (Col E)
                
                return { success: true, message: 'ÄÃ£ cáº­p nháº­t thÃ´ng tin' };
            }
        }
        return { success: false, message: 'KhÃ´ng tÃ¬m tháº¥y ID nhÃ¢n viÃªn: ' + data.id };
    } catch(e) {
         return { success: false, message: 'Lá»—i: ' + e.toString() };
    }
}

function apiDeleteRecruiter(id) {
    Logger.log('=== DELETE RECRUITER ID: ' + id + ' ===');
    try {
        const sheet = getSheetByName(RECRUITER_SHEET_NAME);
        if (!sheet) return { success: false, message: 'Sheet not found' };
        
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            // Check ID first (Column 0)
            if (String(data[i][0]) === String(id)) {
                sheet.deleteRow(i + 1);
                return { success: true, message: 'ÄÃ£ xÃ³a nhÃ¢n viÃªn' };
            }
            // Fallback: Check Name (Column 1) for backward compatibility
             if (String(data[i][1]) === String(id)) {
                sheet.deleteRow(i + 1);
                return { success: true, message: 'ÄÃ£ xÃ³a nhÃ¢n viÃªn (theo tÃªn)' };
            }
        }
        return { success: false, message: 'KhÃ´ng tÃ¬m tháº¥y nhÃ¢n viÃªn' };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function apiSaveEmailTemplate(data) {
    Logger.log('=== SAVE EMAIL TEMPLATE ===');
    try {
        const sheet = getSheetByName('Email_Templates');
        if (!sheet) return { success: false, message: 'Sheet not found' };
        
        const rows = sheet.getDataRange().getValues();
        // Check if ID exists (Column A is ID)
        for (let i = 1; i < rows.length; i++) {
            if (String(rows[i][0]) === String(data.id)) {
                // Update
                // Only update Name if provided, otherwise keep existing
                if (data.name) sheet.getRange(i + 1, 2).setValue(data.name); 
                sheet.getRange(i + 1, 3).setValue(data.subject);
                sheet.getRange(i + 1, 4).setValue(data.body);
                return { success: true, message: 'ÄÃ£ cáº­p nháº­t máº«u email' };
            }
        }
        
        // Add New
        const newId = data.id || ('TPL' + new Date().getTime());
        // Default name if missing
        const newName = data.name || 'Máº«u má»›i (' + newId + ')';
        sheet.appendRow([newId, newName, data.subject, data.body]);
        
        return { success: true, message: 'ÄÃ£ lÆ°u máº«u email má»›i' };
    } catch(e) {
        return { success: false, message: e.toString() };
    }
}

function apiDeleteEmailTemplate(id) {
    try {
        const sheet = getSheetByName('Email_Templates');
        if (!sheet) return { success: false, message: 'Sheet not found' };
        
        const rows = sheet.getDataRange().getValues();
        for (let i = 1; i < rows.length; i++) {
            if (String(rows[i][0]) === String(id)) {
                sheet.deleteRow(i + 1);
                return { success: true, message: 'ÄÃ£ xÃ³a máº«u email' };
            }
        }
        return { success: false, message: 'KhÃ´ng tÃ¬m tháº¥y ID máº«u' };
    } catch(e) {
        return { success: false, message: e.toString() };
    }
}

function apiSendCustomEmail(to, cc, bcc, subject, body, candidateId) {
    Logger.log(`=== SEND EMAIL === To: ${to}, CC: ${cc}, BCC: ${bcc}`);
    try {
        if (!to) return { success: false, message: 'Thiáº¿u Ä‘á»‹a chá»‰ ngÆ°á»i nháº­n' };
        
        const options = {
            to: to,
            subject: subject,
            htmlBody: body.replace(/\n/g, '<br>')
        };
        
        if (cc) {
            options.cc = cc;
        }
        if (bcc) {
            options.bcc = bcc;
        }
        
        MailApp.sendEmail(options);
        
        // LOG ACTIVITY
        const userEmail = Session.getActiveUser().getEmail() || 'Admin';
        let detailMsg = `Gá»­i email tá»›i: ${to} - ${subject}`;
        if (candidateId) {
             const candidateName = getCandidateNameById(candidateId);
             detailMsg = `ÄÃ£ gá»­i email "${subject}" Ä‘áº¿n á»©ng viÃªn **${candidateName}**`;
        }
        
        logActivity(userEmail, 'Gá»­i Email', detailMsg, candidateId || '');
        
        // NOTIFY CC & BCC USERS
        if (candidateId) {
            const candidateName = getCandidateNameById(candidateId);
            const notifyUsers = [];

            if (cc) notifyUsers.push(...cc.split(/[,;]/));
            if (bcc) notifyUsers.push(...bcc.split(/[,;]/));

            notifyUsers.forEach(email => {
                const cleanEmail = email.trim().toLowerCase();
                if (cleanEmail) {
                    createNotification(
                        cleanEmail,
                        'Email',
                        `Báº¡n Ä‘Æ°á»£c thÃªm vÃ o email: "${subject}" gá»­i cho ${candidateName}`,
                        candidateId
                    );
                }
            });
        }
        
        return { success: true, message: 'Email Ä‘Ã£ Ä‘Æ°á»£c gá»­i!' };
    } catch(e) {
        Logger.log('Email Error: ' + e.toString());
        return { success: false, message: 'Gá»­i mail tháº¥t báº¡i: ' + e.toString() };
    }
}

// CHáº Y HÃ€M NÃ€Y Má»˜T Láº¦N TRONG TRÃŒNH BIÃŠN Táº¬P Äá»‚ Cáº¤P QUYá»€N Gá»¬I EMAIL
function testPermissions() {
    console.log("Kiá»ƒm tra quyá»n...");
    // DÃ²ng nÃ y chá»‰ Ä‘á»ƒ kÃ­ch hoáº¡t há»™p thoáº¡i cáº¥p quyá»n
    var quota = MailApp.getRemainingDailyQuota();
    console.log("Email Quota cÃ²n láº¡i: " + quota);
}

function apiCheckDuplicateCandidate(phone, email) {
    try {
        const sheet = getSheetByName(CANDIDATE_SHEET_NAME);
        if (!sheet) return { success: false };
        
        const data = sheet.getDataRange().getValues();
        const headers = data[0].map(h => h.toString().toLowerCase().trim());
        
        // Find columns
        const phoneIndex = headers.findIndex(h => h === 'phone' || h === 'sá»‘ Ä‘iá»‡n thoáº¡i' || h === 'sÄ‘t');
        const emailIndex = headers.findIndex(h => h === 'email');
        const nameIndex = headers.findIndex(h => h === 'name' || h === 'há» vÃ  tÃªn' || h === 'há» tÃªn');
        const dateIndex = headers.findIndex(h => h === 'applied_date' || h === 'ngÃ y á»©ng tuyá»ƒn');
        const posIndex = headers.findIndex(h => h === 'position' || h === 'vá»‹ trÃ­');
        
        if (phoneIndex === -1 && emailIndex === -1) return { success: false };
        
        // Clean inputs
        const cleanPhone = phone ? phone.toString().trim() : '';
        const cleanEmail = email ? email.toString().trim().toLowerCase() : '';
        
        if (!cleanPhone && !cleanEmail) return { success: false };
        
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const rowPhone = phoneIndex > -1 ? String(row[phoneIndex] || '').trim() : '';
            const rowEmail = emailIndex > -1 ? String(row[emailIndex] || '').trim().toLowerCase() : '';
            
            let match = false;
            let matchType = '';
            
            if (cleanPhone && rowPhone && cleanPhone === rowPhone) {
                match = true;
                matchType = 'Sá»‘ Ä‘iá»‡n thoáº¡i';
            } else if (cleanEmail && rowEmail && cleanEmail === rowEmail) {
                match = true;
                matchType = 'Email';
            }
            
            if (match) {
                const sheetId = sheet.getSheetId();
                const rowIndex = i + 1;
                const link = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit#gid=${sheetId}&range=A${rowIndex}`;
                
                return {
                    success: true,
                    found: true,
                    matchType: matchType,
                    name: nameIndex > -1 ? row[nameIndex] : '',
                    date: dateIndex > -1 ? (row[dateIndex] instanceof Date ? row[dateIndex].toISOString().slice(0,10) : row[dateIndex]) : '',
                    position: posIndex > -1 ? row[posIndex] : '',
                    link: link
                };
            }
        }
        
        return { success: true, found: false };
        
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

/* =================================================================================
   9. API: CAREER PORTAL
   ================================================================================= */

// DEBUG FUNCTION
function apiDebugJobs() {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('Jobs');
        if (!sheet) return { error: 'Sheet "Jobs" not found. Available: ' + ss.getSheets().map(s => s.getName()).join(', ') };
        
        const data = sheet.getDataRange().getValues();
        return {
            success: true,
            headers: data.length > 0 ? data[0] : [],
            rowCount: data.length,
            firstRow: data.length > 1 ? data[1] : null,
            message: 'Sheet found. Rows: ' + data.length
        };
    } catch (e) {
        return { error: 'Error: ' + e.toString() };
    }
}

function apiGetOpenJobs() {
    if (!SPREADSHEET_ID) return [];
    try {
        const jobs = apiGetTableData('Jobs');
        
        // Map keys to ensure standardized output for Career Portal
        const mappedJobs = jobs.map(j => {
            return {
                ID: j.ID || j.id,
                Title: j.Title || j['Vị trí'] || j['Tiêu đề'] || '',
                Department: j.Department || j['Phòng ban'] || '',
                Location: j.Location || j['Địa điểm'] || '',
                Type: j.Type || j['Loại hình'] || '',
                Status: j.Status || j['Trạng thái'] || '',
                Description: j.Description || j['Mô tả'] || j['Mô tả công việc'] || '',
                Created_Date: j.Created_Date || j['Ngày tạo'] || ''
            };
        });

        // Filter: Status == 'Open' (case insensitive)
        const openJobs = mappedJobs.filter(j => j.Status && j.Status.toString().toLowerCase() === 'open');
        return openJobs.reverse(); // Newest first
    } catch (e) {
        Logger.log('Error apiGetOpenJobs: ' + e.toString());
        return [];
    }
}


function apiTrackApplication(phone) {
    try {
        if (!phone) return { success: false, message: 'Vui lÃ²ng nháº­p sá»‘ Ä‘iá»‡n thoáº¡i' };
        
        const candidates = apiGetTableData(CANDIDATE_SHEET_NAME);
        const cleanInput = String(phone).replace(/\D/g, '');
        
        const history = candidates.filter(c => {
             const rowPhone = String(c.Phone || '').replace(/\D/g, '');
             // Match last 9 digits or exact full match
             return rowPhone === cleanInput || (cleanInput.length >= 9 && rowPhone.endsWith(cleanInput));
        }).map(c => ({
             Position: c.Position,
             Department: c.Department,
             Status: c.Status,
             Applied_Date: c.Applied_Date ? (c.Applied_Date instanceof Date ? c.Applied_Date.toLocaleDateString('vi-VN') : String(c.Applied_Date).substring(0,10)) : ''
        }));
        
        return { success: true, data: history };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function apiSubmitApplicationPublic(data) {
    Logger.log('=== PUBLIC APPLICATION ===');
    try {
        let sheet = getSheetByName(CANDIDATE_SHEET_NAME);
        if (!sheet) {
             const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
             sheet = ss.insertSheet(CANDIDATE_SHEET_NAME);
             // ID, Name, Phone, Email, Position, Source, Stage, Status, CV_Link, Applied_Date, Department, Contact_Status, Recruiter, Experience, Education, Expected_Salary, Notes
             sheet.appendRow(['ID', 'Name', 'Phone', 'Email', 'Position', 'Source', 'Stage', 'Status', 'CV_Link', 'Applied_Date', 'Department', 'Contact_Status', 'Recruiter', 'Experience', 'Education', 'Expected_Salary', 'Notes']);
        }
        
        const newId = 'C' + new Date().getTime();
        const appliedDate = new Date();
        
        // Append
        sheet.appendRow([
            newId,
            data.name,
            data.phone,
            data.email,
            data.position,
            'Website', // Source
            'Apply',   // Stage
            'New',     // Status
            data.cv_link || '',
            appliedDate,
            data.department,
            'New',     // Contact_Status
            '',        // Recruiter (Unassigned)
            '',        // Experience
            '',        // Education
            '',        // Salary
            'á»¨ng tuyá»ƒn tá»« Website' // Notes
        ]);
        
        SpreadsheetApp.flush(); // Ensure data is written before reading in apiSendApplicationReceivedEmail
        
        // NOTIFICATION
        try {
            const users = apiGetTableData('Users');
            const admins = users.filter(u => u.Role === 'Admin' && u.Email);
            if (admins.length > 0) {
                const recipients = admins.map(u => u.Email).join(',');
                MailApp.sendEmail({
                    to: recipients,
                    subject: '[ATS] á»¨ng viÃªn má»›i: ' + data.name + ' - ' + data.position,
                    htmlBody: `
                        <h3>CÃ³ á»©ng viÃªn má»›i á»©ng tuyá»ƒn tá»« Website</h3>
                        <p><strong>Há» tÃªn:</strong> ${data.name}</p>
                        <p><strong>Vá»‹ trÃ­:</strong> ${data.position}</p>
                        <p><strong>Sá»‘ Ä‘iá»‡n thoáº¡i:</strong> ${data.phone}</p>
                        <p><strong>CV Link:</strong> <a href="${data.cv_link}">${data.cv_link}</a></p>
                        <p>Vui lÃ²ng Ä‘Äƒng nháº­p há»‡ thá»‘ng ATS Ä‘á»ƒ xem chi tiáº¿t.</p>
                    `
                });
            }
        } catch (emailErr) {
            Logger.log('Failed to send notification email: ' + emailErr);
        }

        // AUTO-REPLY EMAIL (PUBLIC)
        try {
            apiSendApplicationReceivedEmail(newId);
        } catch (autoReplyErr) {
            logToSheet('Failed to send auto-reply email (caught in Code.gs): ' + autoReplyErr);
            Logger.log('Failed to send auto-reply email: ' + autoReplyErr);
        }

        return { success: true, message: 'Ná»™p há»“ sÆ¡ thÃ nh cÃ´ng! Email xÃ¡c nháº­n Ä‘Ã£ Ä‘Æ°á»£c gá»­i Ä‘áº¿n báº¡n.' };

    } catch(e) {
        return { success: false, message: e.toString() };
    }
}

// DEBUG FUNCTION V2 (Forced Rename to avoid cache/shadowing issues)
// 7. API: JOB MANAGEMENT V3 (Renamed to fix missing update issues)
// 7. API: JOB MANAGEMENT V3 (Clean & Robust)
function apiGetJobsV3() {
    try {
        if (!SPREADSHEET_ID) return [];
        
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('Jobs');
        if (!sheet) return [];
        
        const lastRow = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();
        if (lastRow < 2) return []; // Only headers or empty

        const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
        
        // Headers are known: ID, Title, Department, Location, Type, Status, Created_Date, Description
        // Indices: 0, 1, 2, 3, 4, 5, 6, 7
        
        const jobs = data.map(row => ({
            ID: row[0] ? String(row[0]) : '',
            Title: row[1] ? String(row[1]) : '',
            Department: row[2] ? String(row[2]) : '',
            Location: row[3] ? String(row[3]) : '',
            Type: row[4] ? String(row[4]) : '',
            Status: row[5] ? String(row[5]) : '',
            Created_Date: row[6] ? String(row[6]) : '',
            Description: row[7] ? String(row[7]) : ''
        }));
        
        return jobs.reverse(); // Newest first
    } catch (e) {
        return [{ID: 'CRITICAL', Title: 'Error V3: ' + e.toString(), Status: 'Error'}];
    }
}

function apiUpdateJobV3(jobData) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('Jobs');
        if (!sheet) return { success: false, message: 'Sheet not found' };
        
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        
        let rowIndex = -1;
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]) === String(jobData.id)) {
                rowIndex = i + 1; // 1-based row index
                break;
            }
        }
        
        if (rowIndex === -1) return { success: false, message: 'Job ID not found: ' + jobData.id };
        
        const setVal = (colName, val) => {
            const idx = findColumnIndex(headers, colName);
            if (idx !== -1) sheet.getRange(rowIndex, idx + 1).setValue(val || '');
        };

        setVal('Title', jobData.title);
        setVal('Department', jobData.department);
        setVal('Position', jobData.position);
        setVal('Location', jobData.location);
        setVal('Type', jobData.type);
        setVal('Description', jobData.description);
        setVal('TicketID', jobData.ticketId);
        
        return { success: true, message: 'Cập nhật thành công!' };
    } catch (e) {
        return { success: false, message: 'Lỗi update: ' + e.toString() };
    }
}

function apiGetOpenJobsV3() {
    try {
        const jobs = apiGetJobsV3();
        // Filter for "Open" status (case insensitive, handle clean output)
        return jobs.filter(j => {
            const s = j.Status.toLowerCase();
            return s === 'open' || s === 'mở' || s === 'đang tuyển';
        });
    } catch (e) {
        return [];
    }
}

// ============================================
// ACTIVITY LOGGING SYSTEM
// ============================================

function logActivity(user, action, details, refId) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let sheet = ss.getSheetByName('DATA_HOAT_DONG');
        if (!sheet) {
            sheet = ss.insertSheet('DATA_HOAT_DONG');
            sheet.appendRow(['Timestamp', 'User', 'Action', 'Details', 'Ref_ID']);
            sheet.getRange("A:A").setNumberFormat("HH:mm dd/MM/yyyy");
        }
        
        const timestamp = new Date();
        sheet.appendRow([timestamp, user, action, details, refId || '']);
        
    } catch (e) {
        Logger.log('Error logging activity: ' + e.toString());
    }
}

function getCandidateNameById(id) {
    try {
        const sheet = getSheetByName(CANDIDATE_SHEET_NAME);
        if (!sheet) return 'Unknown';
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const idIndex = headers.findIndex(h => h.toString().toLowerCase() === 'id');
        const nameIndex = headers.findIndex(h => h.toString().toLowerCase() === 'name' || h.toString().toLowerCase() === 'há» vÃ  tÃªn');
        
        if (idIndex === -1 || nameIndex === -1) return 'Unknown';
        
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][idIndex]) === String(id)) {
                return data[i][nameIndex];
            }
        }
    } catch (e) {
        return 'Unknown';
    }
    return 'Unknown';
}

function apiGetActivityLogs(limit) {
    if (!limit) limit = 10;
    // Logger.log('=== GET ACTIVITY LOGS === Limit: ' + limit);
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let sheet = ss.getSheetByName('DATA_HOAT_DONG');
        if (!sheet) return [];
        
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return []; 
        
        // Fetch more rows to allow for filtering
        // We want 'limit' VALID rows. System logs might be many.
        // Let's fetch last 100 rows (arbitrary safe number for performance)
        const fetchLimit = Math.max(limit * 5, 100);
        const numRows = Math.min(fetchLimit, lastRow - 1);
        const startRow = lastRow - numRows + 1;
        
        const data = sheet.getRange(startRow, 1, numRows, 5).getValues();
        
        // Reverse to show newest first
        const reversed = data.reverse();
        
        const validLogs = [];
        for (const row of reversed) {
            const user = String(row[1]).trim();
            const action = String(row[2]).trim();
            
            // FILTER: Hide System/Debug logs
            if (user.toLowerCase() === 'system') continue;
            if (action.toLowerCase() === 'debug') continue;
            
            let ts = row[0];
            if (ts instanceof Date) {
                ts = ts.toISOString();
            } else {
                ts = String(ts);
            }
            
            validLogs.push({
                timestamp: ts,
                user: user,
                action: action,
                details: row[3],
                refId: row[4]
            });
            
            if (validLogs.length >= limit) break;
        }
        
        return validLogs;
    } catch (e) {
        Logger.log('Error getting activity logs: ' + e.toString());
        return [];
    }
}

// ============================================
// DEPARTMENT & POSITION MANAGEMENT APIs
// ============================================

// Get all departments and their positions
function apiGetDepartments() {
  Logger.log('=== GET DEPARTMENTS ===');
  
  try {
    const sheet = getSheetByName(SETTINGS_SHEET_NAME);
    
    // If sheet doesn't exist, create it with sample data
    if (!sheet) {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const newSheet = ss.insertSheet(SETTINGS_SHEET_NAME);
      
      // Set headers
      newSheet.getRange('A1').setValue('PhÃ²ng ban');
      newSheet.getRange('B1').setValue('Vá»‹ trÃ­ 1');
      newSheet.getRange('C1').setValue('Vá»‹ trÃ­ 2');
      newSheet.getRange('D1').setValue('Vá»‹ trÃ­ 3');
      
      // Add sample data
      newSheet.getRange('A2').setValue('PhÃ²ng nhÃ¢n sá»±');
      newSheet.getRange('B2').setValue('Tuyá»ƒn dá»¥ng');
      newSheet.getRange('C2').setValue('ÄÃ o táº¡o');
      
      newSheet.getRange('A3').setValue('PhÃ²ng káº¿ toÃ¡n');
      newSheet.getRange('B3').setValue('Káº¿ toÃ¡n trÆ°á»Ÿng');
      newSheet.getRange('C3').setValue('Káº¿ toÃ¡n thuáº¿');
      
      Logger.log('Created new settings sheet with sample data');
      
      return {
        success: true,
        departments: [
          { name: 'PhÃ²ng nhÃ¢n sá»±', positions: ['Tuyá»ƒn dá»¥ng', 'ÄÃ o táº¡o'] },
          { name: 'PhÃ²ng káº¿ toÃ¡n', positions: ['Káº¿ toÃ¡n trÆ°á»Ÿng', 'Káº¿ toÃ¡n thuáº¿'] }
        ]
      };
    }
    
    const data = sheet.getDataRange().getValues();
    const departments = [];
    
    // Skip header row (row 0)
    for (let i = 1; i < data.length; i++) {
      const deptName = data[i][0];
      if (!deptName) continue; // Skip empty rows
      
      const positions = [];
      // Collect all non-empty positions from columns B onwards
      for (let j = 1; j < data[i].length; j++) {
        const position = data[i][j];
        if (position && position.toString().trim()) {
          positions.push(position.toString().trim());
        }
      }
      
      departments.push({
        name: deptName.toString().trim(),
        positions: positions
      });
    }
    
    Logger.log('Found ' + departments.length + ' departments');
    return { 
        success: true, 
        departments: departments,
        debug: {
            rows: data.length,
            sheetName: SETTINGS_SHEET_NAME
        }
    };
    
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, message: e.toString(), departments: [], stack: e.stack };
  }
}

// Add new department
function apiAddDepartment(deptName) {
  Logger.log('=== ADD DEPARTMENT: ' + deptName + ' ===');
  
  try {
    let sheet = getSheetByName(SETTINGS_SHEET_NAME);
    
    // Create sheet if doesn't exist
    if (!sheet) {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
      sheet.getRange('A1').setValue('PhÃ²ng ban');
    }
    
    // Check if department already exists
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toLowerCase() === deptName.toLowerCase()) {
        return { success: false, message: 'PhÃ²ng ban Ä‘Ã£ tá»“n táº¡i' };
      }
    }
    
    // Add to next empty row
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1).setValue(deptName);
    
    Logger.log('Added department at row ' + (lastRow + 1));
    return { success: true, message: 'ÄÃ£ thÃªm phÃ²ng ban' };
    
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// Add position to department
function apiAddPosition(deptName, position) {
  Logger.log('=== ADD POSITION: ' + position + ' to ' + deptName + ' ===');
  
  try {
    const sheet = getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Find department row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === deptName) {
        // Find next empty column in this row
        let emptyCol = -1;
        for (let j = 1; j < data[i].length + 5; j++) { // Check a few extra columns
          if (!data[i][j] || !data[i][j].toString().trim()) {
            emptyCol = j + 1; // +1 for 1-indexed
            break;
          }
        }
        
        if (emptyCol === -1) {
          // No empty column found in existing data, add to next column
          emptyCol = data[i].length + 1;
        }
        
        sheet.getRange(i + 1, emptyCol).setValue(position);
        Logger.log('Added position at row ' + (i + 1) + ', col ' + emptyCol);
        return { success: true, message: 'ÄÃ£ thÃªm vá»‹ trÃ­' };
      }
    }
    
    return { success: false, message: 'PhÃ²ng ban khÃ´ng tá»“n táº¡i' };
    
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// Delete department
function apiDeleteDepartment(deptName) {
  Logger.log('=== DELETE DEPARTMENT: ' + deptName + ' ===');
  
  try {
    const sheet = getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Find and delete department row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === deptName) {
        sheet.deleteRow(i + 1);
        Logger.log('Deleted department at row ' + (i + 1));
        return { success: true, message: 'ÄÃ£ xÃ³a phÃ²ng ban' };
      }
    }
    
    return { success: false, message: 'PhÃ²ng ban khÃ´ng tá»“n táº¡i' };
    
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// Delete position from department
function apiDeletePosition(deptName, position) {
  Logger.log('=== DELETE POSITION: ' + position + ' from ' + deptName + ' ===');
  
  try {
    const sheet = getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Find department row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === deptName) {
        // Find position column
        for (let j = 1; j < data[i].length; j++) {
          if (data[i][j] && data[i][j].toString().trim() === position) {
            sheet.getRange(i + 1, j + 1).setValue('');
            Logger.log('Deleted position at row ' + (i + 1) + ', col ' + (j + 1));
            return { success: true, message: 'ÄÃ£ xÃ³a vá»‹ trÃ­' };
          }
        }
        return { success: false, message: 'Vá»‹ trÃ­ khÃ´ng tá»“n táº¡i' };
      }
    }
    
    return { success: false, message: 'PhÃ²ng ban khÃ´ng tá»“n táº¡i' };
    
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}


// EMAIL APIS
// EMAIL APIS

function apiSendEmail(recipient, subject, body, cc) {
  try {
    const options = {
        htmlBody: body
    };
    if (cc) options.cc = cc;
    
    MailApp.sendEmail(recipient, subject, '', options);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// Send "Application Received" Auto-Reply
function apiSendApplicationReceivedEmail(candidateId) {
    logToSheet('apiSendApplicationReceivedEmail Called for ID: ' + candidateId);
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const candidateSheet = ss.getSheetByName(CANDIDATE_SHEET_NAME);
        // const templateSheet = ss.getSheetByName('EMAIL_TEMPLATES'); 
        
        if (!candidateSheet) {
            logToSheet('Error: Candidate sheet missing');
            return { success: false, message: 'Candidate sheet missing' };
        }

        // 1. Get Candidate Data
        const cData = candidateSheet.getDataRange().getValues();
        logToSheet('Searching candidate in ' + cData.length + ' rows.');
        
        let candidate = null;
        for(let i=1; i<cData.length; i++) {
            // Log for debugging first few rows if needed, or just matches
            if(String(cData[i][0]) === String(candidateId)) {
                logToSheet('Found Candidate at Row: ' + (i+1));
                // Map columns: ID=0, Name=1, Phone=2, Email=3, Position=4
                candidate = {
                    Name: cData[i][1],
                    Email: cData[i][3],
                    Position: cData[i][4]
                };
                break;
            }
        }
        
        if (!candidate) {
            logToSheet('Error: Candidate ID not found: ' + candidateId);
            return { success: false, message: 'Candidate not found' };
        }
        
        if (!candidate.Email) {
            logToSheet('Error: Candidate has no email. Name: ' + candidate.Name);
            return { success: false, message: 'No email address' };
        }

        // 2. Get/Create Template
        let subject = `XÃ¡c nháº­n Ä‘Ã£ nháº­n há»“ sÆ¡ á»©ng tuyá»ƒn - ${candidate.Name} - ${candidate.Position}`;
        let body = `
            <p>ChÃ o ${candidate.Name},</p>
            <p>ABC Holding xin chÃ¢n thÃ nh cáº£m Æ¡n báº¡n Ä‘Ã£ quan tÃ¢m vÃ  gá»­i há»“ sÆ¡ á»©ng tuyá»ƒn cho vá»‹ trÃ­ <strong>${candidate.Position}</strong>.</p>
            <p>ChÃºng tÃ´i xÃ¡c nháº­n Ä‘Ã£ nháº­n Ä‘Æ°á»£c CV vÃ  há»“ sÆ¡ cá»§a báº¡n. Bá»™ pháº­n Tuyá»ƒn dá»¥ng sáº½ tiáº¿n hÃ nh Ä‘Ã¡nh giÃ¡ vÃ  pháº£n há»“i káº¿t quáº£ vÃ²ng loáº¡i há»“ sÆ¡ tá»›i báº¡n trong vÃ²ng 3-5 ngÃ y lÃ m viá»‡c tá»›i.</p>
            <p>Náº¿u há»“ sÆ¡ phÃ¹ há»£p, chÃºng tÃ´i sáº½ liÃªn há»‡ Ä‘á»ƒ sáº¯p xáº¿p buá»•i phá»ng váº¥n.</p>
            <br>
            <p>TrÃ¢n trá»ng,</p>
            <p><strong>Bá»™ pháº­n Tuyá»ƒn dá»¥ng cÃ´ng ty ABC Holding</strong></p>
        `;

        logToSheet('Sending email to: ' + candidate.Email);

        // 3. Send Email
        MailApp.sendEmail({
            to: candidate.Email,
            subject: subject,
            htmlBody: body
        });
        
        logToSheet('Email sent successfully to: ' + candidate.Email);
        return { success: true, message: 'Email sent' };

    } catch (e) {
        logToSheet('Exception in apiSendApplicationReceivedEmail: ' + e.toString());
        return { success: false, message: e.toString() };
    }
}


// NEWS API

function apiGetNews() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('News');
    if (!sheet) {
      // Create if not exists
      sheet = ss.insertSheet('News');
      sheet.appendRow(['ID', 'Title', 'Image', 'Content', 'Date', 'Status']);
      return []; // Return empty first time
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const news = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // ID=0, Title=1, Image=2, Content=3, Date=4, Status=5
      news.push({
        ID: row[0],
        Title: row[1],
        Image: row[2],
        Content: row[3],
        Date: row[4] instanceof Date ? row[4].toISOString() : String(row[4]),
        Status: row[5]
      });
    }
    return news.reverse(); // Newest first
  } catch (e) {
    Logger.log('Error getting news: ' + e.toString());
    return [];
  }
}

function apiSaveNews(payload) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('News');
    if (!sheet) {
      sheet = ss.insertSheet('News');
      sheet.appendRow(['ID', 'Title', 'Image', 'Content', 'Date', 'Status']);
    }

    // 1. Upload New Files to Drive
    let newImageUrls = [];
    if (payload.newFiles && payload.newFiles.length > 0) {
        // Use CV_FOLDER_ID or a hardcoded one for simplicity, or create if not exists
        // Ideally we should have a separate folder, but let's reuse or find/create
        // For robustness, let's just use CV_FOLDER_ID if available, or root
        const folderId = (typeof CV_FOLDER_ID !== 'undefined' && CV_FOLDER_ID) ? CV_FOLDER_ID : '';
        let folder;
        if(folderId) {
             folder = DriveApp.getFolderById(folderId);
        } else {
             folder = DriveApp.getRootFolder();
        }

        payload.newFiles.forEach(fileData => {
            const blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), fileData.type, fileData.name);
            const file = folder.createFile(blob);
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            // Use thumbnail link which is more reliable for embedding
            // sz=w4000 requests a large version (up to 4000px width)
            newImageUrls.push(`https://drive.google.com/thumbnail?id=${file.getId()}&sz=w4000`);
        });
    }

    // 2. Combine with Existing Images
    const currentUrls = payload.currentImages ? payload.currentImages.split(/[\n,;]+/).map(s => s.trim()).filter(s => s) : [];
    const finalImageString = [...currentUrls, ...newImageUrls].join('\n');

    // 3. Save to Sheet
    if (payload.id) {
      // Update
      const data = sheet.getDataRange().getValues();
      let rowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(payload.id)) {
          rowIndex = i + 1;
          break;
        }
      }

      if (rowIndex !== -1) {
        sheet.getRange(rowIndex, 2).setValue(payload.title);
        sheet.getRange(rowIndex, 3).setValue(finalImageString);
        sheet.getRange(rowIndex, 4).setValue(payload.content);
        sheet.getRange(rowIndex, 5).setValue(new Date()); // Update Date
        sheet.getRange(rowIndex, 6).setValue(payload.status);
      } else {
        return { success: false, message: 'News not found to update' };
      }
    } else {
      // Create
      const newId = 'NEWS_' + new Date().getTime();
      const date = new Date();
      sheet.appendRow([
        newId,
        payload.title,
        finalImageString,
        payload.content,
        date,
        payload.status
      ]);
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function apiDeleteNews(id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('News');
    if (!sheet) return { success: false, message: 'Sheet not found' };

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(id)) {
            rowIndex = i + 1;
            break;
        }
    }

    if (rowIndex !== -1) {
        sheet.deleteRow(rowIndex);
        return { success: true };
    } else {
        return { success: false, message: 'ID not found' };
    }
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}


// ============================================
// INTERVIEW EVALUATION APIs
// ============================================

function apiCreateEvaluationRequest(candidateId, managerEmails, recruiterEmail, criteria) {
  logToSheet('apiCreateEvaluationRequest: ' + candidateId + ', Managers: ' + JSON.stringify(managerEmails));
  logActivity(recruiterEmail, 'Tạo yêu cầu đánh giá', `Đã tạo yêu cầu đánh giá cho ứng viên ${candidateId}`, candidateId);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(EVALUATION_SHEET_NAME);
    
    // Create sheet if not exists
    if (!sheet) {
      sheet = ss.insertSheet(EVALUATION_SHEET_NAME);
      // Headers: 18 legacy + 6 new = 24 cols
        sheet.appendRow([
          'ID', 'Candidate_ID', 'Candidate_Name', 'Position', 'Department', 
          'Recruiter_Email', 'Manager_Email', 'Created_At', 'Status', 
          'Score_Professional', 'Score_Soft_Skills', 'Score_Culture', 
          'Final_Result', 'Manager_Comment', 'Completed_At',
          'Batch_ID', 'Criteria_Config', 'Scores_JSON',
          'Manager_Name', 'Manager_Department', 'Manager_Position', 
          'Total_Score', 'Proposed_Salary', 'Signature_Status'
        ]);
    }

    // Unified logic moved to ensureSheetStructure
    ensureSheetStructure(false);

    // Get Candidate Details
    const cSheet = ss.getSheetByName(CANDIDATE_SHEET_NAME);
    const cData = cSheet.getDataRange().getValues();
    const cHeaders = cData[0];
    
    // Use robust column finding
    const getCCol = (name) => {
        // We can reuse the findColumnIndex helper if it's available globally
        // Or define a local one if needed. Code.gs has it.
        return findColumnIndex(cHeaders, name);
    };

    const cIdx = {
        id: getCCol('ID'),
        name: getCCol('Name'),
        pos: getCCol('Position'),
        dept: getCCol('Department'),
        recruiter: getCCol('Recruiter')
    };

    let candidate = null;
    for(let i=1; i<cData.length; i++) {
        if(cIdx.id !== -1 && String(cData[i][cIdx.id]) === String(candidateId)) {
            candidate = {
                Name: cIdx.name !== -1 ? cData[i][cIdx.name] : 'Unknown',
                Recruiter: cIdx.recruiter !== -1 ? cData[i][cIdx.recruiter] : '', 
                Position: cIdx.pos !== -1 ? cData[i][cIdx.pos] : '',
                Department: cIdx.dept !== -1 ? cData[i][cIdx.dept] : ''
            };
            break;
        }
    }

    if (!candidate) return { success: false, message: 'Không tìm thấy ứng viên trong hệ thống.' };

    // Batch Info
    const batchId = 'BATCH_' + new Date().getTime();
    const criteriaJson = JSON.stringify(criteria || []); 
    
    // Normalize managerEmails
    let managers = [];
    if (Array.isArray(managerEmails)) {
        managers = managerEmails;
    } else {
        managers = [managerEmails];
    }
    
    const createdAt = new Date();

    const headers = sheet.getDataRange().getValues()[0];
    const getCol = (name) => findColumnIndex(headers, name);

    managers.forEach(mgrEmail => {
        if(!mgrEmail) return;
        
        const evalId = 'EV_' + new Date().getTime() + '_' + Math.floor(Math.random() * 10000);
        
        // Build the row array based on headers
        const rowData = new Array(headers.length).fill('');
        
        const mapping = {
          'ID': evalId,
          'Candidate_ID': candidateId,
          'Candidate_Name': candidate.Name,
          'Position': candidate.Position,
          'Department': candidate.Department,
          'Recruiter_Email': recruiterEmail,
          'Manager_Email': mgrEmail,
          'Created_At': createdAt,
          'Status': 'Pending',
          'Batch_ID': batchId,
          'Criteria_Config': criteriaJson,
          'Scores_JSON': ''
        };

        Object.keys(mapping).forEach(key => {
          const colIdx = getCol(key);
          if (colIdx !== -1) rowData[colIdx] = mapping[key];
        });

        sheet.appendRow(rowData);
        
        // Send Email to Manager
        // Use dynamic URL
        const appUrl = ScriptApp.getService().getUrl();
        const subject = 'Yêu cầu đánh giá phỏng vấn: ' + candidate.Name + ' - ' + candidate.Position;
        const body = '<p>Chào anh/chị,</p>' +
                     '<p>Bộ phận Tuyển dụng mời anh/chị thực hiện đánh giá ứng viên <strong>' + candidate.Name + '</strong>.</p>' +
                     '<p>Vui lòng truy cập hệ thống <a href="' + appUrl + '">HR Recruit</a> để thực hiện.</p>' +
                     '<p>Trân trọng.</p>';
        
        if (MailApp.getRemainingDailyQuota() > 0) {
            try {
                MailApp.sendEmail({ to: mgrEmail, subject: subject, htmlBody: body });
            } catch(e) { console.log('Email error: ' + e.toString()); }
        }
        
        // NOTIFICATION
        let usernameToNotify = mgrEmail;
        try {
             // Basic lookup or just use email
             // Ideally we cache users map for performance but this is fine
             const users = apiGetTableData('Users');
             const u = users.find(x => (x.Email && x.Email.toLowerCase() === mgrEmail.toLowerCase()) || (x.Username && x.Username.toLowerCase() === mgrEmail.toLowerCase()));
             if (u) usernameToNotify = u.Username;
        } catch(e) {}
        
        createNotification(
            usernameToNotify, 
            'Evaluation', 
            `Bạn có yêu cầu đánh giá mới cho ứng viên ${candidate.Name} - ${candidate.Position}`, 
            candidateId
        );
    });

    return { success: true, message: 'Đã gửi yêu cầu đánh giá cho ' + managers.length + ' quản lý.' };
  } catch (e) {
    logToSheet('Error apiCreateEvaluationRequest: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function apiSubmitEvaluation(evalId, scores, result, comment, additionalData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(EVALUATION_SHEET_NAME);
    if (!sheet) return { success: false, message: 'Evaluation sheet not found' };

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let evalData = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(evalId)) {
        rowIndex = i + 1;
        const headers = data[0];
        const getCol = (name) => findColumnIndex(headers, name);
        const idx = {
          rec: getCol('Recruiter_Email'),
          name: getCol('Candidate_Name'),
          pos: getCol('Position')
        };
        evalData = {
          Recruiter_Email: idx.rec !== -1 ? data[i][idx.rec] : '',
          Candidate_Name: idx.name !== -1 ? data[i][idx.name] : '',
          Position: idx.pos !== -1 ? data[i][idx.pos] : ''
        };
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: 'Evaluation not found' };

    const headers = data[0];
    const getCol = (name) => findColumnIndex(headers, name);

    const updateCell = (colName, val) => {
      const cIdx = getCol(colName);
      if (cIdx !== -1) sheet.getRange(rowIndex, cIdx + 1).setValue(val);
    };

    // Update Row
    const completedAt = new Date();
    
    updateCell('Status', 'Completed');
    updateCell('Final_Result', result);
    updateCell('Manager_Comment', comment);
    updateCell('Completed_At', completedAt);
    updateCell('Scores_JSON', JSON.stringify(scores));

    // Additional data
    if (additionalData) {
      if (additionalData.managerName) updateCell('Manager_Name', additionalData.managerName);
      if (additionalData.managerDept) updateCell('Manager_Department', additionalData.managerDept);
      if (additionalData.managerPos) updateCell('Manager_Position', additionalData.managerPos);
      if (additionalData.totalScore) updateCell('Total_Score', additionalData.totalScore);
      if (additionalData.proposedSalary) updateCell('Proposed_Salary', additionalData.proposedSalary);
      if (additionalData.signatureStatus) updateCell('Signature_Status', additionalData.signatureStatus);
    }

    // Notify Recruiter
    if (evalData && evalData.Recruiter_Email) {
      if (MailApp.getRemainingDailyQuota() > 0) {
        MailApp.sendEmail({
          to: evalData.Recruiter_Email,
          subject: 'Kết quả đánh giá: ' + evalData.Candidate_Name + ' - ' + result,
          htmlBody: '<p>Manager đã hoàn thành đánh giá cho ứng viên <strong>' + evalData.Candidate_Name + '</strong>.</p><p>Kết quả: <strong>' + result + '</strong></p><p>Vui lòng kiểm tra chi tiết trên hệ thống.</p>'
        });
      }
    }
    
    logActivity(Session.getActiveUser().getEmail(), 'Hoàn thành đánh giá', `Đã chấm điểm cho ứng viên ${evalData ? evalData.Candidate_Name : evalId} - Kết quả: ${result}`, evalData ? evalData.Candidate_Name : '');

    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * XUẤT PHIẾU ĐÁNH GIÁ PHỎNG VẤN SANG PDF
 * @param {string} evalId
 */
// [REDUNDANT apiExportEvaluationPDF REMOVED]



function apiGetPendingEvaluations(managerKey) {
  // managerKey can be email or username
  const all = apiGetEvaluationsList(managerKey, 'Manager');
  return all.filter(e => {
    const s = (e.Status || '').toString().trim().toLowerCase();
    return s === 'pending' || s === 'chờ đánh giá';
  });
}

function apiGetEvaluationsList(searchKey, role) {
  logToSheet(`apiGetEvaluationsList started. User: ${searchKey}, Role: ${role}`);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(EVALUATION_SHEET_NAME);
    if (!sheet) {
        // Robust sheet find
        const allSheets = ss.getSheets();
        sheet = allSheets.find(s => s.getName().toUpperCase() === EVALUATION_SHEET_NAME.toUpperCase());
    }
    
    if (!sheet) {
        logToSheet('ERROR: EVALUATIONS sheet not found');
        return [];
    }

    const data = sheet.getDataRange().getValues();
    logToSheet(`Found ${data.length} rows (incl. header)`);
    if (data.length < 2) return [];

    const headers = data[0];
    const getCol = (name) => findColumnIndex(headers, name);

    // Indices mapping
    const idx = {
      id: getCol('ID'),
      canId: getCol('Candidate_ID'),
      name: getCol('Candidate_Name'),
      pos: getCol('Position'),
      dept: getCol('Department'),
      recEmail: getCol('Recruiter_Email'),
      mgrEmail: getCol('Manager_Email'),
      created: getCol('Created_At'),
      status: getCol('Status'),
      result: getCol('Final_Result'),
      comment: getCol('Manager_Comment'),
      completed: getCol('Completed_At'),
      batch: getCol('Batch_ID'),
      criteria: getCol('Criteria_Config'),
      scoresJson: getCol('Scores_JSON'),
      mgrName: getCol('Manager_Name'),
      mgrDept: getCol('Manager_Department'),
      mgrPos: getCol('Manager_Position'),
      total: getCol('Total_Score'),
      salary: getCol('Proposed_Salary'),
      sig: getCol('Signature_Status')
    };
    logToSheet(`Column indices: ${JSON.stringify(idx)}`);

    const roleNorm = String(role || '').toLowerCase().trim();
    const searchVal = String(searchKey || '').toLowerCase().trim();
    
    logToSheet(`Processing evaluations for: ${searchVal} (${roleNorm})`);

    // Lookup extra user details for matching (matching by Email, Username, or Name)
    let userDetails = { emails: [searchVal], names: [] };
    try {
        const users = apiGetTableData('Users');
        const userFound = users.find(u => {
            const uE = String(u.Email || '').toLowerCase().trim();
            const uU = String(u.Username || '').toLowerCase().trim();
            const uF = String(u.Full_Name || '').toLowerCase().trim();
            return uE === searchVal || uU === searchVal || uF === searchVal;
        });
        if (userFound) {
            if (userFound.Email) userDetails.emails.push(String(userFound.Email).toLowerCase().trim());
            if (userFound.Username) userDetails.emails.push(String(userFound.Username).toLowerCase().trim());
            if (userFound.Full_Name) userDetails.names.push(String(userFound.Full_Name).toLowerCase().trim());
        }
    } catch (e) {}

    const list = [];
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (idx.id !== -1 && !row[idx.id] && idx.name !== -1 && !row[idx.name]) continue;

        let include = false;
        if (roleNorm === 'admin') {
            include = true;
        } else {
            const recEmail = idx.recEmail !== -1 ? String(row[idx.recEmail] || '').toLowerCase().trim() : '';
            const mgrEmail = idx.mgrEmail !== -1 ? String(row[idx.mgrEmail] || '').toLowerCase().trim() : '';
            
            if (roleNorm === 'manager') {
                include = userDetails.emails.includes(mgrEmail) || userDetails.names.includes(mgrEmail);
            } else if (roleNorm === 'recruiter') {
                include = userDetails.emails.includes(recEmail) || userDetails.names.includes(recEmail);
            } else {
                // Secondary check for User/Viewer/Recruiter (default to email overlap)
                include = userDetails.emails.includes(recEmail) || userDetails.emails.includes(mgrEmail);
            }
        }

        if (include) {
            let scores = {};
            try { if (idx.scoresJson !== -1 && row[idx.scoresJson]) scores = JSON.parse(row[idx.scoresJson]); } catch(e) {}
            
            let criteria = [];
            try { if (idx.criteria !== -1 && row[idx.criteria]) criteria = JSON.parse(row[idx.criteria]); } catch(e) {}

            list.push({
                ID: idx.id !== -1 ? String(row[idx.id]) : '',
                Candidate_ID: idx.canId !== -1 ? String(row[idx.canId]) : '',
                Candidate_Name: idx.name !== -1 ? String(row[idx.name]) : '',
                Position: idx.pos !== -1 ? String(row[idx.pos]) : '',
                Department: idx.dept !== -1 ? String(row[idx.dept]) : '',
                Recruiter_Email: idx.recEmail !== -1 ? String(row[idx.recEmail]) : '',
                Manager_Email: idx.mgrEmail !== -1 ? String(row[idx.mgrEmail]) : '',
                Created_At: (idx.created !== -1 && row[idx.created] instanceof Date) ? row[idx.created].toISOString() : String(row[idx.created] || ''),
                Status: idx.status !== -1 ? String(row[idx.status]) : '',
                Final_Result: idx.result !== -1 ? String(row[idx.result]) : '',
                Manager_Comment: idx.comment !== -1 ? String(row[idx.comment]) : '',
                Completed_At: (idx.completed !== -1 && row[idx.completed] instanceof Date) ? row[idx.completed].toISOString() : String(row[idx.completed] || ''),
                Batch_ID: idx.batch !== -1 ? String(row[idx.batch]) : '',
                Criteria_Config: Array.isArray(criteria) ? criteria : [],
                Scores_JSON: scores,
                Manager_Name: idx.mgrName !== -1 ? String(row[idx.mgrName]) : '',
                Manager_Department: idx.mgrDept !== -1 ? String(row[idx.mgrDept]) : '',
                Manager_Position: idx.mgrPos !== -1 ? String(row[idx.mgrPos]) : '',
                Total_Score: idx.total !== -1 ? (row[idx.total] instanceof Date ? '0' : String(row[idx.total] || '0')) : '0',
                Proposed_Salary: idx.salary !== -1 ? String(row[idx.salary] || '') : '',
                Signature_Status: idx.sig !== -1 ? String(row[idx.sig] || '') : ''
            });
        }
    }
    logToSheet(`Finished apiGetEvaluationsList. Found: ${list.length}`);
    const finalResult = (list && list.length > 0) ? list.reverse() : [];
    return finalResult;
  } catch (e) {
    logToSheet(`CRITICAL ERROR in apiGetEvaluationsList: ${e.toString()}`);
    return [];
  }
}


// BULK UPLOAD HANDLER
function apiBulkUploadCandidates(filesData, source, stage) {
    // filesData is array of { name, type, data (base64) }
    const results = {
        success: [],
        errors: []
    };

    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = getSheetByName(CANDIDATE_SHEET_NAME);
        const cvFolder = DriveApp.getFolderById(CV_FOLDER_ID);
        
        // Default values
        if (!source) source = 'Bulk Import';
        if (!stage) stage = 'New'; // Or 'Má»›i' / 'á»¨ng tuyá»ƒn' depending on config

        filesData.forEach(file => {
            try {
                // 1. Decode and Save to Drive
                const decoded = Utilities.base64Decode(file.data);
                const blob = Utilities.newBlob(decoded, file.type, file.name);
                const driveFile = cvFolder.createFile(blob);
                const fileUrl = driveFile.getUrl();

                // 2. Prepare Candidate Data
                // Use Filename as Name, remove extension
                const candidateName = file.name.replace(/\.[^/.]+$/, "");
                const newId = 'C' + new Date().getTime() + Math.floor(Math.random() * 1000);
                const appliedDate = new Date();

                // 3. Append to Sheet
                // Order must match `apiCreateCandidate` or Sheet Headers
                // Headers: [ID, Name, Email, Phone, Position, Dept, Source, Stage, LinkCV, Date, ...]
                // We'll trust the order or use a safer mappping if possible. 
                // Currently relying on appendRow generally works if columns are standard.
                // Let's look at apiCreateCandidate to match column structure.
                
                // Based on previous code analysis, structure likely:
                // ID, Name, Position, Department, Recruiter, Manager, Created_At, Status (Stage), ...
                // Wait, need to be careful with column order. 
                // Let's look at `ensureSheetStructure` or `apiCreateCandidate` logic to be sure.
                
                // RE-READING apiCreateCandidate structure from memory/context:
                // It uses `setCol` to map keys. `appendRow` is dangerous if we don't know the exact order.
                // Better approach: Read headers first, then map.
                // For performance in bulk, reading headers once is fine.
                
                // Simplified assumptions based on standard 'DATA á»¨NG VIÃŠN'
                // [ID, Name, Email, Phone, Position, Department, Source, Recruiter, Manager, Status, CV_Link, Applied_Date, Notes, ...]
                
                // Let's try to reuse `apiCreateCandidate` logic internally? 
                // No, that takes a FormObject.
                
                // Let's do a safe append using header mapping.
                const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
                const rowData = new Array(headers.length).fill('');
                
                const setVal = (headerName, val) => {
                    const idx = headers.findIndex(h => h.toString().toLowerCase() === headerName.toLowerCase());
                    if (idx > -1) rowData[idx] = val;
                };

                setVal('ID', newId);
                setVal('Name', candidateName); // Name from filename
                setVal('Status', stage);
                setVal('Source', source);
                setVal('CV_Link', fileUrl);
                setVal('Applied_Date', appliedDate);
                // setVal('Recruiter', currentUser); // Difficult to get session from here without passing it

                sheet.appendRow(rowData);
                results.success.push(file.name);

            } catch (e) {
                results.errors.push(`${file.name}: ${e.toString()}`);
            }
        });

    } catch (e) {
        return { success: false, message: e.toString() };
    }

    return results;
}

function getSourceSheet() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SOURCE_SHEET_NAME);
    if (!sheet) {
        sheet = ss.insertSheet(SOURCE_SHEET_NAME);
        sheet.appendRow(['Source_Name', 'Created_At']);
        // Add default sources
        const defaults = ['Website', 'LinkedIn', 'Facebook', 'Referral', 'Job Portal'];
        defaults.forEach(d => sheet.appendRow([d, new Date()]));
    }
    return sheet;
}

function apiGetCandidateSources() {
    try {
        const sheet = getSourceSheet();
        const data = sheet.getDataRange().getValues();
        const sources = [];
        
        // Skip header
        for (let i = 1; i < data.length; i++) {
            if (data[i][0]) {
                sources.push(String(data[i][0]).trim());
            }
        }
        return sources;
    } catch (e) {
        return ['Website', 'LinkedIn', 'Facebook', 'Referral', 'Job Portal']; // Fallback
    }
}

function apiAddCandidateSource(sourceName) {
    if (!sourceName) return { success: false, message: 'TÃƒÂªn nguÃ¡Â»â€œn khÃƒÂ´ng Ã„â€˜Ã†Â°Ã¡Â»Â£c Ã„â€˜Ã¡Â»Æ’ trÃ¡Â»â€˜ng' };
    try {
        const sheet = getSourceSheet();
        const data = sheet.getDataRange().getValues();
        
        // Check duplicate
        const cleanName = sourceName.trim().toLowerCase();
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]).trim().toLowerCase() === cleanName) {
                return { success: false, message: 'NguÃ¡Â»â€œn nÃƒÂ y Ã„â€˜ÃƒÂ£ tÃ¡Â»â€œn tÃ¡ÂºÂ¡i' };
            }
        }
        
        sheet.appendRow([sourceName.trim(), new Date()]);
        return { success: true, message: 'Ã„ÂÃƒÂ£ thÃƒÂªm nguÃ¡Â»â€œn mÃ¡Â»â€ºi' };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function apiDeleteCandidateSource(sourceName) {
    try {
        const sheet = getSourceSheet();
        const data = sheet.getDataRange().getValues();
        
        const cleanName = sourceName.trim().toLowerCase();
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]).trim().toLowerCase() === cleanName) {
                sheet.deleteRow(i + 1);
                return { success: true, message: 'Ã„ÂÃƒÂ£ xÃƒÂ³a nguÃ¡Â»â€œn' };
            }
        }
        return { success: false, message: 'KhÃƒÂ´ng tÃƒÂ¬m thÃ¡ÂºÂ¥y nguÃ¡Â»â€œn Ã„â€˜Ã¡Â»Æ’ xÃƒÂ³a' };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function apiEditCandidateSource(oldName, newName) {
    if (!oldName || !newName) return { success: false, message: 'TÃƒÂªn nguÃ¡Â»â€œn khÃƒÂ´ng Ã„â€˜Ã†Â°Ã¡Â»Â£c Ã„â€˜Ã¡Â»Æ’ trÃ¡Â»â€˜ng' };
    try {
        const sheet = getSourceSheet();
        const data = sheet.getDataRange().getValues();
        const cleanOld = oldName.trim().toLowerCase();
        const cleanNew = newName.trim().toLowerCase();
        
        // Check if new name already exists (and is not the old name)
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]).trim().toLowerCase() === cleanNew) {
                return { success: false, message: 'TÃƒÂªn nguÃ¡Â»â€œn mÃ¡Â»â€ºi Ã„â€˜ÃƒÂ£ tÃ¡Â»â€œn tÃ¡ÂºÂ¡i' };
            }
        }
        
        // Find and update
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]).trim().toLowerCase() === cleanOld) {
                sheet.getRange(i + 1, 1).setValue(newName.trim());
                return { success: true, message: 'Ã„ÂÃƒÂ£ cÃ¡ÂºÂ­p nhÃ¡ÂºÂ­t nguÃ¡Â»â€œn' };
            }
        }
        return { success: false, message: 'KhÃƒÂ´ng tÃƒÂ¬m thÃ¡ÂºÂ¥y nguÃ¡Â»â€œn cÃ¡ÂºÂ§n sÃ¡Â»Â­a' };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// ============================================
// BULK IMPORT FROM SHEET / EXCEL
// ============================================

function apiImportFromSheet(importType, data, source, stage) {
    // importType: 'URL' or 'FILE'
    // data: URL string OR { name, base64 }
    
    const results = { success: 0, errors: [], message: '' };
    
    try {
        let textValues = [];
        let srcSS = null; 
        
        // 1. GET DATA SOURCE
        if (importType === 'URL') {
            try {
               srcSS = SpreadsheetApp.openByUrl(data);
            } catch(e) {
               return { success: false, message: 'KhÃƒÂ´ng thÃ¡Â»Æ’ mÃ¡Â»Å¸ link Google Sheet. HÃƒÂ£y Ã„â€˜Ã¡ÂºÂ£m bÃ¡ÂºÂ£o quyeÃŒÂ£`n truy cÃ¡ÂºÂ­p.' };
            }
        } else if (importType === 'FILE') {
            // Check file type based on name or mime
            const isCsv = data.name.toLowerCase().endsWith('.csv');
            
            if (isCsv) {
                 // CSV HANDLING (Native, works without Advanced Service)
                 try {
                     const decoded = Utilities.base64Decode(data.base64);
                     const csvString = Utilities.newBlob(decoded).getDataAsString();
                     const parsedCsv = Utilities.parseCsv(csvString);
                     
                     // Mock a "Sheet" object structure or just use array direct
                     srcData = parsedCsv;
                 } catch (e) {
                     return { success: false, message: 'Lá»—i Ä‘á»c file CSV: ' + e.toString() };
                 }
            } else {
                // EXCEL (.xlsx) Handling
                try {
                    const decoded = Utilities.base64Decode(data.base64);
                    const blob = Utilities.newBlob(decoded, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', data.name);
                    
                    const file = { title: data.name, mimeType: 'application/vnd.google-apps.spreadsheet' };
                    // This LINE requires Advanced Drive Service
                    const newFile = Drive.Files.insert(file, blob, { convert: true });
                    srcSS = SpreadsheetApp.openById(newFile.id);
                } catch (e) {
                    if (e.toString().includes('Drive is not defined')) {
                        return { success: false, message: 'Lá»—i: Há»‡ thá»‘ng chÆ°a báº­t "Advanced Drive Service" Ä‘á»ƒ Ä‘á»c file Excel (.xlsx). \n\nGIáº¢I PHÃP:\n1. LÆ°u file Excel thÃ nh Ä‘uÃ´i .csv (File > Save As > CSV).\n2. Hoáº·c dÃ¡n Link Google Sheet.' };
                    }
                    return { success: false, message: 'Lá»—i xá»­ lÃ½ file Excel: ' + e.toString() };
                }
            }
        }
        
        // If we have srcSS (from URL or Excel convert), read it. If we used CSV, srcData is already set.
        if (srcSS) {
            const srcSheet = srcSS.getSheets()[0];
            srcData = srcSheet.getDataRange().getValues();
            // Cleanup temp Excel file if needed
             if (importType === 'FILE') {
                 DriveApp.getFileById(srcSS.getId()).setTrashed(true);
            }
        }
        
        if (!srcData || srcData.length < 2) return { success: false, message: 'File khÃ´ng cÃ³ dá»¯ liá»‡u (Ã­t nháº¥t pháº£i cÃ³ 1 dÃ²ng tiÃªu Ä‘á» vÃ  1 dÃ²ng dá»¯ liá»‡u).' };
        
        const srcHeaders = srcData[0].map(h => String(h).toLowerCase().trim());
        
        // 2. DESTINATION
        const destSheet = getSheetByName(CANDIDATE_SHEET_NAME);
        const destHeaders = destSheet.getRange(1, 1, 1, destSheet.getLastColumn()).getValues()[0]; // Real headers
        
        // 3. MAP COLUMNS
        // We look for flexible keywords
        // Helper map function
        const findCol = (keywords) => {
             return srcHeaders.findIndex(h => keywords.some(k => h === k || h.includes(k)));
        };
        
        const map = {};
        map['Name'] = findCol(['name', 'tÃªn', 'há» vÃ  tÃªn', 'há» tÃªn']);
        map['Email'] = findCol(['email', 'thÆ°', 'mail']);
        map['Phone'] = findCol(['phone', 'sá»‘ Ä‘iá»‡n thoáº¡i', 'sÄ‘t', 'tel', 'mobile', 'di Ä‘á»™ng']);
        map['Position'] = findCol(['position', 'vá»‹ trÃ­', 'chá»©c danh', 'job']);
        map['Department'] = findCol(['department', 'phÃ²ng', 'phÃ²ng ban', 'bá»™ pháº­n']);
        map['Stage'] = findCol(['stage', 'giai Ä‘oáº¡n', 'status code']); 
        map['Status'] = findCol(['status', 'tráº¡ng thÃ¡i', 'tÃ¬nh tráº¡ng']);
        map['CV_Link'] = findCol(['cv', 'link', 'há»“ sÆ¡', 'drive']);
        map['Applied_Date'] = findCol(['applied_date', 'ngÃ y á»©ng tuyá»ƒn', 'date']);
        
        // Extended Fields
        map['Gender'] = findCol(['gender', 'giá»›i tÃ­nh']);
        map['Birth_Year'] = findCol(['birth_year', 'nÄƒm sinh', 'dob', 'yob']);
        map['School'] = findCol(['school', 'trÆ°á»ng', 'há»c váº¥n', 'tÃªn trÆ°á»ng']);
        map['Education_Level'] = findCol(['education_level', 'trÃ¬nh Ä‘á»™', 'trÃ¬nh Ä‘á»™ há»c váº¥n']);
        map['Major'] = findCol(['major', 'chuyÃªn ngÃ nh']);
        map['Experience'] = findCol(['experience', 'kinh nghiá»‡m']);
        map['Salary_Expectation'] = findCol(['salary', 'lÆ°Æ¡ng', 'má»©c lÆ°Æ¡ng', 'expectation']);
        map['Recruiter'] = findCol(['recruiter', 'ngÆ°á»i phá»¥ trÃ¡ch', 'nhÃ¢n viÃªn tuyá»ƒn dá»¥ng']);
        map['Source'] = findCol(['source', 'nguá»“n']);
        map['Notes'] = findCol(['notes', 'ghi chÃº']);
        
        if (map['Name'] === -1 && map['Email'] === -1 && map['Phone'] === -1) {
             console.log('Headers found:', srcHeaders);
             return { success: false, message: 'KhÃ´ng tÃ¬m tháº¥y cá»™t TÃªn, Email hoáº·c Sá»‘ Ä‘iá»‡n thoáº¡i trong file. Vui lÃ²ng kiá»ƒm tra tiÃªu Ä‘á» cá»™t.' };
        }

        // 4. PROCESS ROWS
        const candidatesToAdd = [];
        
        for (let i = 1; i < srcData.length; i++) {
            const row = srcData[i];
            
            // Validate row has at least some info
            const name = map['Name'] > -1 ? row[map['Name']] : '';
            const email = map['Email'] > -1 ? row[map['Email']] : '';
            
            if (!name && !email) continue; // Skip empty rows
            
            // Prepare Dest Row
            const destRow = new Array(destHeaders.length).fill('');
            
            // Helper to set dest value
            const setDest = (headerKey, val) => {
                 const idx = destHeaders.findIndex(h => h.toString().toLowerCase().trim() === headerKey.toLowerCase());
                 if (idx > -1) destRow[idx] = val;
            };
            
            // Generate IDs
            const newId = 'C' + new Date().getTime() + Math.floor(Math.random() * 10000) + i;
            
            setDest('ID', newId);
            setDest('Name', name || 'Unknown');
            setDest('Email', email);
            setDest('Phone', map['Phone'] > -1 ? row[map['Phone']] : '');
            setDest('Position', map['Position'] > -1 ? row[map['Position']] : '');
            setDest('Department', map['Department'] > -1 ? row[map['Department']] : '');
            setDest('CV_Link', map['CV_Link'] > -1 ? row[map['CV_Link']] : '');
            
            // New Fields
            setDest('Gender', map['Gender'] > -1 ? row[map['Gender']] : '');
            setDest('Birth_Year', map['Birth_Year'] > -1 ? row[map['Birth_Year']] : '');
            setDest('School', map['School'] > -1 ? row[map['School']] : '');
            setDest('Education_Level', map['Education_Level'] > -1 ? row[map['Education_Level']] : '');
            setDest('Major', map['Major'] > -1 ? row[map['Major']] : '');
            setDest('Experience', map['Experience'] > -1 ? row[map['Experience']] : '');
            setDest('Salary_Expectation', map['Salary_Expectation'] > -1 ? row[map['Salary_Expectation']] : '');
            setDest('Recruiter', map['Recruiter'] > -1 ? row[map['Recruiter']] : '');
            setDest('Source', map['Source'] > -1 ? row[map['Source']] : (source || 'Import'));
            setDest('Notes', map['Notes'] > -1 ? row[map['Notes']] : '');

            // Status Logic: Priority 1: File's Status/Tráº¡ng thÃ¡i, Priority 2: File's Stage/Giai Ä‘oáº¡n, Priority 3: Modal's Selection, Priority 4: 'New'
            let status = stage || 'New';
            if (map['Status'] > -1 && row[map['Status']]) {
                status = row[map['Status']];
            } else if (map['Stage'] > -1 && row[map['Stage']]) {
                status = row[map['Stage']];
            }
            setDest('Status', status);
            
            // Date Logic
            let date = new Date();
            if (map['Applied_Date'] > -1 && row[map['Applied_Date']]) {
                const d = row[map['Applied_Date']];
                if (d instanceof Date) date = d;
                else if (typeof d === 'string') date = new Date(d); // Basic parse
            }
            setDest('Applied_Date', date);
            
            // Add to batch
            candidatesToAdd.push(destRow);
        }

        
        // 5. BATCH APPEND
        if (candidatesToAdd.length > 0) {
            destSheet.getRange(destSheet.getLastRow() + 1, 1, candidatesToAdd.length, destHeaders.length).setValues(candidatesToAdd);
            results.success = candidatesToAdd.length;
        }
        
        // Cleanup Temp File if Excel
        if (importType === 'FILE' && srcSS) {
             DriveApp.getFileById(srcSS.getId()).setTrashed(true);
        }
        
        return { success: true, count: results.success, message: `Ã„ ÃƒÂ£ import thÃƒÂ nh cÃƒÂ´ng ${results.success} Ã¡Â»Â©ng viÃƒÂªn.` };

    } catch (e) {
        if (e.toString().includes('Drive is not defined')) {
             return { success: false, message: 'HÃ¡Â»â€¡ thÃ¡Â»â€˜ng chÃ†Â°a bÃ¡ÂºÂ­t Advanced Drive API Ã„â€˜Ã¡Â»Æ’ Ã„â€˜Ã¡Â»\x8dc Excel. Vui lÃƒÂ²ng dÃƒÂ¹ng Link Google Sheet.' };
        }
        return { success: false, message: 'LÃ¡Â»â€”i hÃ¡Â»â€¡ thÃ¡Â»â€˜ng: ' + e.toString() };
    }
}


// ============================================
// NOTIFICATION SYSTEM
// ============================================

function createNotification(toUser, type, message, relatedId) {
  logToSheet(`Creating notification for [${toUser}]: ${message}`);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Notifications');
    if (!sheet) {
        sheet = ss.insertSheet('Notifications');
        sheet.appendRow(['ID', 'ToUser', 'Type', 'Message', 'RelatedId', 'IsRead', 'CreatedAt']);
    }

    const newId = 'NOTIF_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
    const createdAt = new Date();

    const cleanUser = String(toUser).toLowerCase().trim();

    sheet.appendRow([
      newId,
      cleanUser,      // toUser (Lowercase)
      type,           // 'Time', 'Evaluation', 'Mention'
      message,
      relatedId,      // CandidateID or EvalID
      false,          // IsRead
      createdAt
    ]);
  } catch (e) {
    Logger.log('Error creating notification: ' + e.toString());
    logToSheet('Error creating notification: ' + e.toString());
  }
}

function apiGetNotifications(username, email) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Notifications');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    // Row structure: ['ID', 'ToUser', 'Type', 'Message', 'RelatedId', 'IsRead', 'CreatedAt']
    
    const notifications = [];
    const validUser = username ? String(username).toLowerCase().trim() : '';
    const validEmail = email ? String(email).toLowerCase().trim() : '';

    logToSheet(`Fetching notifications for User: [${validUser}], Email: [${validEmail}]`);

    // Loop from end (newest) to beginning
    for (let i = data.length - 1; i >= 1; i--) {
        const row = data[i];
        const toUser = String(row[1]).toLowerCase().trim(); 
        const isRead = row[5]; 

        // Check if ToUser matches Username OR Email
        const match = (validUser && toUser === validUser) || (validEmail && toUser === validEmail);
        
        const isUnread = isRead === false || isRead === 'false' || isRead === '';

        // Include if Unread OR if it's one of the recent 20 (regardless of read status)
        // Since we loop backwards, we just check match.
        
        if (match) {
             let createdStr = '';
             if (row[6] instanceof Date) {
                 createdStr = row[6].toISOString();
             } else {
                 createdStr = String(row[6]);
             }

             notifications.push({
                ID: row[0],
                ToUser: row[1],
                Type: row[2],
                Message: row[3],
                RelatedId: row[4],
                IsRead: !isUnread, // True if read
                CreatedAt: createdStr
            });
        }
        
        // Limit to 20 most recent notifications for performance and UI cleaniness
        if (notifications.length >= 20) break;
    }
    
    return notifications;
  } catch (e) {
    Logger.log('Error fetching notifications: ' + e.toString());
    return [];
  }
}

// [REDUNDANT apiMarkAllNotificationsRead REMOVED]
 

// RE-IMPLEMENTING apiMarkAllNotificationsRead correctly
function apiMarkAllNotificationsRead(username, email) {
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('Notifications');
        if (!sheet) return { success: false };
    
        const data = sheet.getDataRange().getValues();
        const validUser = username ? String(username).toLowerCase() : '';
        const validEmail = email ? String(email).toLowerCase() : '';

        for (let i = 1; i < data.length; i++) {
          const toUser = String(data[i][1]).toLowerCase();
          const isRead = data[i][5];
          
          if (isRead === false) {
              if ((validUser && toUser === validUser) || (validEmail && toUser === validEmail)) {
                  sheet.getRange(i + 1, 6).setValue(true);
              }
          }
        }
        return { success: true };
      } catch (e) {
        return { success: false, message: e.toString() };
      }
}

function apiMarkNotificationRead(id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Notifications');
    if (!sheet) return { success: false };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 6).setValue(true); // Column 6 is IsRead (1-based index)
        return { success: true };
      }
    }
    return { success: false, message: 'Not found' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// [REDUNDANT apiMarkAllNotificationsRead REMOVED]


/**
 * ONE-TIME CLEANUP: Merge redundant columns and delete aliases
 * Standardize on CANDIDATE_ALIASES keys.
 */
function apiCleanupCandidateSheet() {
  Logger.log('=== STARTING SHEET CLEANUP ===');
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CANDIDATE_SHEET_NAME);
    if (!sheet) return { success: false, message: 'Sheet not found' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Define mappings for merging (Target Key -> List of Alias Headers to capture data from)
    const mergeMap = {
      'Status': ['Contact_Status', 'TÃ¬nh tráº¡ng liÃªn há»‡', 'Tráº¡ng thÃ¡i liÃªn láº¡c'],
      'Salary_Expectation': ['Expected_Salary', 'Salary_Expectation_1', 'Salary', 'Expected Salary'],
      'Notes': ['Newnote', 'New Note', 'Ghi chÃº']
    };

    // 1. Identify Target and Source indices
    const columnsToProcess = {}; 
    
    Object.keys(mergeMap).forEach(key => {
        let targetIdx = headers.indexOf(key);
        if (targetIdx === -1) {
            // Try to find if any of the aliases is actually being used as the primary one
            for (let a of CANDIDATE_ALIASES[key] || []) {
                let idx = headers.indexOf(a);
                if (idx !== -1) {
                    targetIdx = idx;
                    break;
                }
            }
        }
        
        if (targetIdx !== -1) {
            const sources = [];
            mergeMap[key].forEach(alias => {
                const idx = headers.indexOf(alias);
                if (idx !== -1 && idx !== targetIdx) {
                    sources.push({idx: idx, name: alias});
                }
            });
            if (sources.length > 0) {
                columnsToProcess[key] = { targetIdx: targetIdx, sources: sources };
            }
        }
    });

    if (Object.keys(columnsToProcess).length === 0) {
      return { success: true, message: 'KhÃ´ng tÃ¬m tháº¥y cá»™t dÆ° thá»«a nÃ o cáº§n dá»n dáº¹p.' };
    }

    // 2. Perform Data Merge
    for (let i = 1; i < data.length; i++) {
        Object.keys(columnsToProcess).forEach(key => {
            const config = columnsToProcess[key];
            let targetVal = data[i][config.targetIdx];
            
            config.sources.forEach(source => {
                const sourceVal = data[i][source.idx];
                if (sourceVal && sourceVal.toString().trim() !== '') {
                    if (key === 'Notes') {
                        const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm dd/MM/yyyy");
                        const appendStr = `[${time}] Migration: ${sourceVal}`;
                        targetVal = targetVal ? targetVal + "\n" + appendStr : appendStr;
                    } else if (!targetVal || targetVal.toString().trim() === '') {
                        targetVal = sourceVal;
                    }
                }
            });
            
            if (targetVal !== data[i][config.targetIdx]) {
                sheet.getRange(i + 1, config.targetIdx + 1).setValue(targetVal);
            }
        });
    }

    // 3. Delete Redundant Columns
    const allSourceIndices = [];
    Object.keys(columnsToProcess).forEach(key => {
        columnsToProcess[key].sources.forEach(s => allSourceIndices.push(s.idx));
    });
    
    const sortedIndices = [...new Set(allSourceIndices)].sort((a, b) => b - a);
    const deletedNames = sortedIndices.map(idx => headers[idx]);
    
    sortedIndices.forEach(idx => {
        sheet.deleteColumn(idx + 1);
    });

    return { 
        success: true, 
        message: 'Dá»n dáº¹p hoÃ n táº¥t. ÄÃ£ xÃ³a ' + sortedIndices.length + ' cá»™t dÆ° thá»«a.',
        details: deletedNames.join(', ')
    };
  } catch (e) {
    Logger.log('Cleanup Error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * GENERATE PDF DOCUMENT (Offer, Contract, Profile)
 * @param {string} candidateId 
 * @param {string} templateName - The name of the HTML file (e.g. 'TemplateOffer')
 * @param {Object} extraData - Additional data from form (e.g. salary, startDate)
 */
function apiGenerateDocument(candidateId, templateName, extraData) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(CANDIDATE_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let candidateRow = -1;
    
    // 1. Find Candidate
    const idIdx = headers.indexOf('ID');
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(candidateId)) {
        candidateRow = i;
        break;
      }
    }
    
    if (candidateRow === -1) return { success: false, message: 'KhÃ´ng tÃ¬m tháº¥y á»©ng viÃªn' };
    
    const rawCandidate = {};
    headers.forEach((h, idx) => {
      rawCandidate[h] = data[candidateRow][idx];
    });

    // Normalize candidate data using CANDIDATE_ALIASES
    const candidate = {};
    Object.keys(CANDIDATE_ALIASES).forEach(key => {
      const aliases = CANDIDATE_ALIASES[key] || [];
      // Try key itself first
      if (rawCandidate[key] !== undefined) {
        candidate[key] = rawCandidate[key];
      } else {
        // Try aliases
        for (let a of aliases) {
          if (rawCandidate[a] !== undefined) {
            candidate[key] = rawCandidate[a];
            break;
          }
        }
      }
      if (candidate[key] === undefined) candidate[key] = '';
    });

    // Also keep raw fields just in case template uses exact headers
    const candidateMerged = { ...rawCandidate, ...candidate };

    // 2. Prepare Template Content
    let htmlContent = HtmlService.createTemplateFromFile(templateName).getRawContent();
    
    // 3. Replace Variables (Candidate Data)
    Object.keys(candidateMerged).forEach(key => {
      let val = candidateMerged[key] || '';
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      const regex = new RegExp(`{{${key}}}`, 'g');
      htmlContent = htmlContent.replace(regex, val);
    });

    // 3.0 Merge Survey Data (if available by email)
    const candidateEmail = candidate['Email'];
    Logger.log('ðŸ“§ Candidate Email from Data: ' + candidateEmail);
    if (candidateEmail) {
      const surveyData = apiGetSurveyData(candidateEmail);
      Logger.log('ðŸ“Š Survey Data Keys Found: ' + Object.keys(surveyData).join(', '));
      Object.keys(surveyData).forEach(key => {
        let val = surveyData[key] || '';
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
        }
        const regex = new RegExp(`{{${key}}}`, 'g');
        htmlContent = htmlContent.replace(regex, val);
      });
    }

    // 3.1 Load and Replace Company Information
    // 3.1 Load Company Information as Defaults
    let replacementVars = {};
    const companyRes = apiGetCompanyInfo();
    if (companyRes.success && companyRes.data) {
      const info = companyRes.data;
      replacementVars = {
        'CompanyName': info.name || '',
        'CompanyEmail': info.email || '',
        'CompanyPhone': info.phone || '',
        'CompanyWorkTime': info.worktime || '',
        'CompanyTaxCode': info.taxcode || '',
        'CompanyAddress': (info.addresses && info.addresses.length > 0) ? info.addresses[0] : '',
        'CompanySignerName': (info.signers && info.signers.length > 0) ? (typeof info.signers[0] === 'object' ? info.signers[0].name : info.signers[0]) : '',
        'CompanySignerPosition': (info.signers && info.signers.length > 0 && typeof info.signers[0] === 'object') ? info.signers[0].position : '',
        'CompanyNameShort': (info.name || '').split(' ').map(w => w[0]).join('').toUpperCase(),
        'CompanyCity': (info.addresses && info.addresses.length > 0) ? info.addresses[0].split(',')[0].trim() : ''
      };
    }

    // 4. Override with Extra Variables (from HR form)
    if (extraData) {
      Object.keys(extraData).forEach(key => {
        if (extraData[key] !== undefined && extraData[key] !== null && extraData[key] !== '') {
            replacementVars[key] = extraData[key];
        }
      });
    }

    // 4.1 Perform Batch Replacement
    Object.keys(replacementVars).forEach(key => {
        const regex = new RegExp(`{{${key}}}`, 'g');
        htmlContent = htmlContent.replace(regex, replacementVars[key]);
    });

    // 5. Add current date
    const now = new Date();
    const today = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
    const todayFull = `ngÃ y ${now.getDate()} thÃ¡ng ${now.getMonth() + 1} nÄƒm ${now.getFullYear()}`;
    
    htmlContent = htmlContent.replace(/{{Today}}/g, today);
    htmlContent = htmlContent.replace(/{{TodayFull}}/g, todayFull);
    htmlContent = htmlContent.replace(/{{Today}}/g, today);

    // 6. Create PDF
    const blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf');
    const fileName = `${templateName}_${candidate.Name}_${Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd")}.pdf`;
    
    // 7. Save to Drive
    const targetFolder = getDocumentFolder(templateName);
    const file = targetFolder.createFile(blob).setName(fileName);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // 8. Log Activity
    const userEmail = Session.getActiveUser().getEmail() || 'Admin';
    logActivity(userEmail, 'Táº¡o vÄƒn báº£n', `ÄÃ£ xuáº¥t file **${fileName}** tá»« máº«u **${templateName}** cho á»©ng viÃªn **${candidate.Name}**`, candidateId);

    return { 
      success: true, 
      url: file.getUrl(),
      downloadUrl: file.getDownloadUrl().replace("?e=download", ""),
      fileId: file.getId()
    };
    
  } catch (e) {
    Logger.log('PDF Generation Error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Helper to get or create folder for documents
 */
function getDocumentFolder(subfolderName) {
  let rootId = DOCUMENT_FOLDER_ID;
  let rootFolder;

  if (!rootId) {
    // Try to find in parent of CV folder
    try {
      const cvFolder = DriveApp.getFolderById(CV_FOLDER_ID);
      const parent = cvFolder.getParents().next();
      const folders = parent.getFoldersByName('ATS_Generated_Documents');
      if (folders.hasNext()) {
        rootFolder = folders.next();
      } else {
        rootFolder = parent.createFolder('ATS_Generated_Documents');
      }
    } catch (e) {
      rootFolder = DriveApp.createFolder('ATS_Generated_Documents');
    }
  } else {
    rootFolder = DriveApp.getFolderById(rootId);
  }

  // Create subfolder (e.g. "OfferLetters", "Contracts")
  const subFolders = rootFolder.getFoldersByName(subfolderName);
  if (subFolders.hasNext()) {
    return subFolders.next();
  } else {
    return rootFolder.createFolder(subfolderName);
  }
}

function apiDebugEvaluations() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(EVALUATION_SHEET_NAME);
    if (!sheet) return 'Sheet EVALUATIONS not found';
    const data = sheet.getDataRange().getValues();
    return {
       totalRows: data.length,
       headers: data[0],
       sampleData: data.length > 1 ? data[1] : 'No data rows'
    };
  } catch (e) {
    return 'Debug Error: ' + e.toString();
  }
}

/**
 * Xuất Phiếu Đánh Giá PV sang PDF (Mẫu chuyên nghiệp)
 * Hỗ trợ so sánh nhiều người phỏng vấn nếu thuộc cùng một Batch
 */
function apiExportEvaluationPDF(id, overrideResult, officialSalary) {
  try {
    logToSheet(`Starting PDF Export for Evaluation ID: ${id}. Override: ${overrideResult}, Salary: ${officialSalary}`);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(EVALUATION_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const getCol = (name) => findColumnIndex(headers, name);
    const idx = {
      id: getCol('ID'),
      canId: getCol('Candidate_ID'),
      name: getCol('Candidate_Name'),
      pos: getCol('Position'),
      dept: getCol('Department'),
      recEmail: getCol('Recruiter_Email'),
      mgrEmail: getCol('Manager_Email'),
      created: getCol('Created_At'),
      status: getCol('Status'),
      result: getCol('Final_Result'),
      comment: getCol('Manager_Comment'),
      batch: getCol('Batch_ID'),
      scoresJson: getCol('Scores_JSON'),
      mgrName: getCol('Manager_Name'),
      mgrPos: getCol('Manager_Position'),
      mgrDept: getCol('Manager_Department'),
      total: getCol('Total_Score'),
      salary: getCol('Proposed_Salary'),
      sig: getCol('Signature_Status')
    };

    // Find the main record
    const mainRow = data.find(r => r[idx.id] == id);
    if (!mainRow) throw new Error("Không tìm thấy dữ liệu đánh giá.");

    // Find all related evaluations in the same batch
    const batchId = mainRow[idx.batch];
    const relatedRows = batchId ? data.filter(r => r[idx.batch] === batchId && r[idx.status] === 'Completed') : [mainRow];
    
    // Sort related rows to have a consistent order
    relatedRows.sort((a, b) => String(a[idx.mgrEmail]).localeCompare(String(b[idx.mgrEmail])));

    // Prepare Template Data
    let htmlContent = HtmlService.createHtmlOutputFromFile('TemplateEvaluationProfessional').getContent();
    
    const candidateName = mainRow[idx.name] || '---';
    const candidateID = mainRow[idx.canId] || '---';
    const position = mainRow[idx.pos] || '---';
    const department = mainRow[idx.dept] || '---';
    const today = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm");

    // Fetch Company Info
    let companyName = "[TEN CONG TY]";
    const companyRes = apiGetCompanyInfo();
    if (companyRes.success && companyRes.data && companyRes.data.name) {
      companyName = companyRes.data.name;
    }
    
    // Build Table Headers for Managers
    let mgrHeaders = "";
    let mgrSubHeaders = "";
    relatedRows.forEach(row => {
      const name = row[idx.mgrName] || row[idx.mgrEmail];
      const pos = row[idx.mgrPos] || '';
      const dept = row[idx.mgrDept] || '';
      mgrHeaders += `
      <th colspan="1" class="text-center" style="min-width: 120px;">
        <div class="fw-bold">PV: ${name}</div>
        <div style="font-weight: normal; font-size: 10px; color: #555;">${pos} <br> ${dept}</div>
      </th>`;
      mgrSubHeaders += `<th class="text-center">Số điểm</th>`;
    });

    // Collect all unique criteria
    let allCriteria = new Set();
    relatedRows.forEach(row => {
      const scores = row[idx.scoresJson] ? JSON.parse(row[idx.scoresJson]) : {};
      Object.keys(scores).forEach(k => {
        if (!['professional', 'softSkills', 'culture'].includes(k)) allCriteria.add(k);
      });
    });
    if (allCriteria.size === 0) {
      allCriteria.add("Chuyên môn");
      allCriteria.add("Kỹ năng mềm");
      allCriteria.add("Văn hóa & Thái độ");
    }

    // Build Criteria Rows
    let criteriaHtml = "";
    allCriteria.forEach(crit => {
      criteriaHtml += `<tr><td class="fw-bold">${crit}</td>`;
      relatedRows.forEach(row => {
        const scores = row[idx.scoresJson] ? JSON.parse(row[idx.scoresJson]) : {};
        let score = scores[crit];
        if (score === undefined || score === null || score === "") {
           if (crit === "Chuyên môn") score = row[getCol('Score_Professional')];
           if (crit === "Kỹ năng mềm") score = row[getCol('Score_Soft_Skills')];
           if (crit === "Văn hóa & Thái độ") score = row[getCol('Score_Culture')];
        }
        if (score instanceof Date) score = "-";
        criteriaHtml += `<td class="text-center fw-bold">${score || '-'}</td>`;
      });
      criteriaHtml += `</tr>`;
    });

    // Build Total Score Row
    let totalScoreRow = "";
    let commentRows = "";
    relatedRows.forEach(row => {
      let personSum = 0;
      let count = 0;
      const scores = row[idx.scoresJson] ? JSON.parse(row[idx.scoresJson]) : {};
      allCriteria.forEach(crit => {
        let sc = scores[crit];
        if (sc === undefined || sc === null || sc === "") {
           if (crit === "Chuyên môn") sc = row[getCol('Score_Professional')];
           if (crit === "Kỹ năng mềm") sc = row[getCol('Score_Soft_Skills')];
           if (crit === "Văn hóa & Thái độ") sc = row[getCol('Score_Culture')];
        }
        const sn = parseFloat(sc);
        if (!isNaN(sn)) {
          personSum += sn;
          count++;
        }
      });
      const avg = count > 0 ? (personSum / count).toFixed(1) : '0.0';
      totalScoreRow += `<td class="text-center score-val text-danger">${avg}</td>`;
      const comment = row[idx.comment] || 'N/A';
      commentRows += `<td style="font-size: 11px; vertical-align: top; padding: 10px;">${comment}</td>`;
    });

    let salaryRow = "<tr><td class='fw-bold bg-gray'>Mức lương đề xuất</td>";
    relatedRows.forEach(row => {
      salaryRow += `<td class='text-center small'>${row[idx.salary] || '-'}</td>`;
    });
    salaryRow += "</tr>";

    const finalResult = overrideResult || mainRow[idx.result] || "Pending";
    const officialSalaryFull = officialSalary || mainRow[idx.salary] || "Theo thỏa thuận";
    const resultClass = (finalResult === 'Pass' ? 'pass-text' : (finalResult === 'Reject' ? 'reject-text' : 'pending-text'));
    const signatureStatus = mainRow[idx.sig] || "Đã xác nhận trên hệ thống ATS Recruit";
    const batchID = mainRow[idx.batch] || 'N/A';

    htmlContent = htmlContent
      .replace(/{{CompanyName}}/gi, companyName)
      .replace(/{{CandidateName}}/gi, candidateName)
      .replace(/{{CandidateID}}/gi, candidateID)
      .replace(/{{Position}}/gi, position)
      .replace(/{{Department}}/gi, department)
      .replace(/{{Today}}/gi, today)
      .replace(/{{ManagerHeaders}}/gi, mgrHeaders)
      .replace(/{{ManagerSubHeaders}}/gi, mgrSubHeaders)
      .replace(/{{CriteriaRows}}/gi, criteriaHtml)
      .replace(/{{TotalScoreRow}}/gi, totalScoreRow)
      .replace(/{{CommentRows}}/gi, commentRows)
      .replace(/{{SalaryRow}}/gi, salaryRow)
      .replace(/{{FinalResult}}/gi, finalResult)
      .replace(/{{OfficialSalary}}/gi, officialSalaryFull)
      .replace(/{{ResultClass}}/gi, resultClass)
      .replace(/{{SignatureStatus}}/gi, signatureStatus)
      .replace(/{{BatchID}}/gi, batchID);

    const blob = Utilities.newBlob(htmlContent, 'text/html', `Phieu_Danh_Gia_${candidateName}.html`);
    let folder;
    if (DOCUMENT_FOLDER_ID) {
      folder = DriveApp.getFolderById(DOCUMENT_FOLDER_ID);
    } else {
      const folders = DriveApp.getFoldersByName("ATS_Documents");
      folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("ATS_Documents");
    }
    
    const tempFile = folder.createFile(blob);
    const pdfBlob = tempFile.getAs('application/pdf').setName(`Phieu_Danh_Gia_${candidateName}.pdf`);
    tempFile.setTrashed(true);

    return {
      success: true,
      data: Utilities.base64Encode(pdfBlob.getBytes()),
      filename: `Phieu_Danh_Gia_${candidateName}.pdf`
    };
  } catch (e) {
    logToSheet(`PDF Export Error: ${e.toString()}`);
    return { success: false, message: e.toString() };
  }
}

/**
 * 3.2 API: REJECTION REASONS MANAGEMENT
 */
function apiSaveRejectionReasons(reasons) {
    try {
        const sheet = getSheetByName('RejectionReasons');
        if (!sheet) {
            SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet('RejectionReasons').appendRow(['ID', 'Type', 'Reason', 'Order']);
        }
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const rrSheet = ss.getSheetByName('RejectionReasons');
        
        // Clear existing (except header)
        if (rrSheet.getLastRow() > 1) {
            rrSheet.getRange(2, 1, rrSheet.getLastRow() - 1, rrSheet.getLastColumn()).clearContent();
        }

        if (reasons && reasons.length > 0) {
            const values = reasons.map((r, i) => [
                r.ID || 'R' + (i + 1),
                r.Type || 'Company',
                r.Reason || '',
                r.Order || (i + 1)
            ]);
            rrSheet.getRange(2, 1, values.length, 4).setValues(values);
        }
        
        return { success: true, message: 'Lưu lý do từ chối thành công' };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}
