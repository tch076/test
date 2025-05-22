function doGet(e)
{
  var output = HtmlService.createTemplateFromFile('dashboard_user');
  output = output.evaluate()
    .setTitle('Web Crud App')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  return output;
}
const SHEET_ID = '1RA8mZZkaxF_vfeMWqXKGS33oIoUEcFAVLfwzgCXTjlw'; // <--- *** ใส่ GOOGLE SHEET ID ของคุณที่นี่ ***
const SHEET_NAME_CUSTOMER = 'customer';
const SHEET_NAME_USERS = 'users';
const SHEET_NAME_FEEDBACK = 'feedback';
const ADMIN_USERNAME = 'admin';
const ADMIN_PASSWORD = '1234'; // Hashed version would be better in a real scenario

// --- Utility Functions to get Sheets ---
function getCustomerSheet() {
  try {
    return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME_CUSTOMER);
  } catch (e) {
    Logger.log('Error getting customer sheet: ' + e.message);
    throw new Error('ไม่สามารถเข้าถึงชีทข้อมูลการจองได้: ' + e.message);
  }
}

function getUsersSheet() {
  try {
    return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME_USERS);
  } catch (e) {
    Logger.log('Error getting users sheet: ' + e.message);
    throw new Error('ไม่สามารถเข้าถึงชีทข้อมูลผู้ใช้งานได้: ' + e.message);
  }
}

function getFeedbackSheet() {
  try {
    return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME_FEEDBACK);
  } catch (e) {
    Logger.log('Error getting feedback sheet: ' + e.message);
    throw new Error('ไม่สามารถเข้าถึงชีทข้อมูลข้อเสนอแนะได้: ' + e.message);
  }
}

// --- Hashing (Simple example, consider more robust libraries for production) ---
function simpleHash(password) {
  // This is a very basic hash for demonstration.
  // For production, use robust hashing like bcrypt or Argon2 (not natively available in GAS).
  // Utilities.computeDigest is an option for stronger hashing like SHA-256.
  var hashed = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return hashed.map(function(byte) {
    return ('0' + (byte & 0xFF).toString(16)).slice(-2);
  }).join('');
}


// --- Routing for Web App ---
function doGet(e) {
  if (e.parameter.page === 'register') {
    return HtmlService.createTemplateFromFile('register').evaluate().setTitle('สร้างบัญชีผู้ใช้ - ระบบจองห้องประชุม');
  } else if (e.parameter.page === 'dashboard_user') {
    // Basic check, improve with session management
    if (isUserLoggedIn()) {
      let template = HtmlService.createTemplateFromFile('dashboard_user');
      template.user = getSessionUser(); // Pass user data to template
      return template.evaluate().setTitle('Dashboard ผู้ใช้งาน - ระบบจองห้องประชุม');
    }
    return HtmlService.createTemplateFromFile('login').evaluate().setTitle('เข้าสู่ระบบ - ระบบจองห้องประชุม');
  } else if (e.parameter.page === 'dashboard_admin') {
    // Basic check, improve with session management
    if (isAdminLoggedIn()) {
       let template = HtmlService.createTemplateFromFile('dashboard_admin');
       return template.evaluate().setTitle('Dashboard ผู้ดูแลระบบ - ระบบจองห้องประชุม');
    }
    return HtmlService.createTemplateFromFile('login').evaluate().setTitle('เข้าสู่ระบบ - ระบบจองห้องประชุม');
  }
  // Default to login page
  return HtmlService.createTemplateFromFile('login').evaluate().setTitle('เข้าสู่ระบบ - ระบบจองห้องประชุม');
}

// --- User Session Management (Basic Example using PropertiesService) ---
function setUserSession(username, department, phone) {
  PropertiesService.getUserProperties().setProperty('loggedInUser', JSON.stringify({username: username, department: department, phone: phone}));
  PropertiesService.getUserProperties().setProperty('isUserLoggedIn', 'true');
}

function setAdminSession() {
  PropertiesService.getUserProperties().setProperty('isAdminLoggedIn', 'true');
}

function clearSession() {
  PropertiesService.getUserProperties().deleteProperty('loggedInUser');
  PropertiesService.getUserProperties().deleteProperty('isUserLoggedIn');
  PropertiesService.getUserProperties().deleteProperty('isAdminLoggedIn');
}

function isUserLoggedIn() {
  return PropertiesService.getUserProperties().getProperty('isUserLoggedIn') === 'true';
}

function isAdminLoggedIn() {
  return PropertiesService.getUserProperties().getProperty('isAdminLoggedIn') === 'true';
}

function getSessionUser() {
  const userJson = PropertiesService.getUserProperties().getProperty('loggedInUser');
  return userJson ? JSON.parse(userJson) : null;
}

// --- API functions callable from Frontend (google.script.run) ---

function userLogin(username, password) {
  try {
    const sheet = getUsersSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const usernameIndex = headers.indexOf('ชื่อ');
    const passwordIndex = headers.indexOf('รหัสผ่าน');
    const departmentIndex = headers.indexOf('กลุ่มสาระ');
    const phoneIndex = headers.indexOf('เบอร์โทรศัพท์');

    if (usernameIndex === -1 || passwordIndex === -1 || departmentIndex === -1 || phoneIndex === -1) {
        return { status: 'error', message: 'โครงสร้างชีทผู้ใช้ไม่ถูกต้อง (ไม่พบ Header ที่ต้องการ)' };
    }

    const hashedPassword = simpleHash(password);

    for (let i = 1; i < data.length; i++) {
      if (data[i][usernameIndex] === username && data[i][passwordIndex] === hashedPassword) {
        setUserSession(data[i][usernameIndex], data[i][departmentIndex], data[i][phoneIndex]);
        return { status: 'success', message: 'เข้าสู่ระบบสำเร็จ!', userType: 'user', userData: { name: data[i][usernameIndex], department: data[i][departmentIndex], phone: data[i][phoneIndex] } };
      }
    }
    return { status: 'error', message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  } catch (e) {
    Logger.log('Login Error: ' + e.toString());
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการเข้าสู่ระบบ: ' + e.message };
  }
}

function adminLogin(username, password) {
  if (username === ADMIN_USERNAME && password === ADMIN_PASSWORD) { // In real app, hash admin password too
    setAdminSession();
    return { status: 'success', message: 'เข้าสู่ระบบผู้ดูแลสำเร็จ!', userType: 'admin' };
  }
  return { status: 'error', message: 'ชื่อผู้ใช้หรือรหัสผ่านผู้ดูแลไม่ถูกต้อง' };
}

function logout() {
  clearSession();
  return { status: 'success', message: 'ออกจากระบบสำเร็จ' };
}


function registerUser(userData) {
  try {
    const sheet = getUsersSheet();
    // Check if user already exists
    const data = sheet.getDataRange().getValues();
    const usernameIndex = data[0].indexOf('ชื่อ');
    for(let i=1; i<data.length; i++){
        if(data[i][usernameIndex] === userData.name){
            return { status: 'error', message: 'ชื่อผู้ใช้นี้มีอยู่ในระบบแล้ว' };
        }
    }

    const hashedPassword = simpleHash(userData.password);
    sheet.appendRow([
      userData.name,
      userData.department,
      userData.phone,
      hashedPassword,
      new Date() // Registration Timestamp
    ]);
    return { status: 'success', message: 'สร้างบัญชีผู้ใช้สำเร็จ!' };
  } catch (e) {
    Logger.log('Registration Error: ' + e.toString());
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการสร้างบัญชี: ' + e.message };
  }
}

function submitBooking(bookingData) {
  try {
    const sheet = getCustomerSheet();
    const lock = LockService.getScriptLock();
    lock.tryLock(30000); // Lock for 30 seconds

    // --- Check for booking conflicts ---
    const bookings = sheet.getDataRange().getValues();
    const startTimeIndex = bookings[0].indexOf('วันที่และเวลาที่เริ่มต้น');
    const endTimeIndex = bookings[0].indexOf('วันที่และเวลาที่สิ้นสุด');
    const roomIndex = bookings[0].indexOf('ชื่อห้องประชุม');

    if (startTimeIndex === -1 || endTimeIndex === -1 || roomIndex === -1) {
      return { status: 'error', message: 'โครงสร้างชีทข้อมูลการจองไม่ถูกต้อง (ไม่พบ Header ที่ต้องการ)' };
    }

    const newBookingStart = new Date(bookingData.startTime);
    const newBookingEnd = new Date(bookingData.endTime);

    for (let i = 1; i < bookings.length; i++) {
      if (bookings[i][roomIndex] === bookingData.room) {
        const existingStart = new Date(bookings[i][startTimeIndex]);
        const existingEnd = new Date(bookings[i][endTimeIndex]);
        // Check for overlap: (StartA < EndB) and (StartB < EndA)
        if (newBookingStart < existingEnd && existingStart < newBookingEnd) {
          lock.releaseLock();
          return { status: 'error', message: 'ห้องประชุมนี้ถูกจองในช่วงเวลาที่เลือกแล้ว กรุณาเลือกเวลาอื่น' };
        }
      }
    }
    // --- End conflict check ---

    sheet.appendRow([
      bookingData.name,
      bookingData.phone,
      bookingData.department,
      bookingData.purpose,
      new Date(bookingData.startTime), // Store as Date objects
      new Date(bookingData.endTime),   // Store as Date objects
      bookingData.room,
      bookingData.equipment.join(', '),
      bookingData.recording,
      bookingData.details,
      new Date() // Booking Timestamp
    ]);
    lock.releaseLock();
    return { status: 'success', message: 'จองห้องประชุมสำเร็จ!' };
  } catch (e) {
    Logger.log('Booking Error: ' + e.toString());
    if (lock.hasLock()) lock.releaseLock();
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการจองห้องประชุม: ' + e.message };
  }
}

function getBookingsForCalendar(month, year) {
  try {
    const sheet = getCustomerSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    // Ensure headers exist before trying to find their index
    const requiredHeaders = ['วันที่และเวลาที่เริ่มต้น', 'วันที่และเวลาที่สิ้นสุด', 'ชื่อห้องประชุม', 'ชื่อผู้จอง'];
    for (const header of requiredHeaders) {
        if (headers.indexOf(header) === -1) {
            Logger.log(`Header not found in 'customer' sheet: ${header}`);
            return { status: 'error', message: `ไม่พบ Header '${header}' ในชีทข้อมูลการจอง` };
        }
    }

    const bookings = [];
    const targetMonth = parseInt(month); // month is 0-indexed for JS Date
    const targetYear = parseInt(year);

    for (let i = 1; i < data.length; i++) {
      const record = data[i];
      const startTime = new Date(record[headers.indexOf('วันที่และเวลาที่เริ่มต้น')]);
      if (startTime.getFullYear() === targetYear && startTime.getMonth() === targetMonth) {
        bookings.push({
          start: startTime.toISOString(),
          end: new Date(record[headers.indexOf('วันที่และเวลาที่สิ้นสุด')]).toISOString(),
          title: record[headers.indexOf('ชื่อห้องประชุม')] + ' (' + record[headers.indexOf('ชื่อผู้จอง')] + ')',
          room: record[headers.indexOf('ชื่อห้องประชุม')],
          booker: record[headers.indexOf('ชื่อผู้จอง')]
          // Add other details if needed for calendar display
        });
      }
    }
    return { status: 'success', data: bookings };
  } catch (e) {
    Logger.log('Error fetching bookings for calendar: ' + e.toString());
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการดึงข้อมูลการจอง: ' + e.message };
  }
}


function getRoomStatuses(selectedDate) {
  try {
    const sheet = getCustomerSheet();
    const allBookings = sheet.getDataRange().getValues();
    const headers = allBookings[0];
    const startTimeIndex = headers.indexOf('วันที่และเวลาที่เริ่มต้น');
    const endTimeIndex = headers.indexOf('วันที่และเวลาที่สิ้นสุด');
    const roomIndex = headers.indexOf('ชื่อห้องประชุม');
    const bookerIndex = headers.indexOf('ชื่อผู้จอง');
    const purposeIndex = headers.indexOf('วัตถุประสงค์ในการจอง');

    if ([startTimeIndex, endTimeIndex, roomIndex, bookerIndex, purposeIndex].includes(-1)) {
        return { status: 'error', message: 'โครงสร้างชีทข้อมูลการจองไม่ถูกต้อง (ไม่พบ Header ที่ต้องการสำหรับสถานะห้อง)' };
    }

    const rooms = ['ห้องเกียรติยศ', 'ห้องประชุมอาคาร 3', 'ห้องประชุมอาคาร 5', 'โดมอาคารเอนกประสงค์']; // Define your rooms
    const roomStatuses = {};
    const targetDate = new Date(selectedDate);
    targetDate.setHours(0,0,0,0); // Normalize to start of day

    rooms.forEach(room => {
      roomStatuses[room] = { status: 'ว่าง', bookingsToday: [] };
    });

    for (let i = 1; i < allBookings.length; i++) {
      const booking = allBookings[i];
      const bookingStart = new Date(booking[startTimeIndex]);
      const bookingEnd = new Date(booking[endTimeIndex]);

      // Check if booking is on the selectedDate
      if (bookingStart.getFullYear() === targetDate.getFullYear() &&
          bookingStart.getMonth() === targetDate.getMonth() &&
          bookingStart.getDate() === targetDate.getDate()) {

        const roomName = booking[roomIndex];
        if (roomStatuses[roomName]) {
          roomStatuses[roomName].status = 'ไม่ว่าง (มีการจอง)'; // Or more nuanced if checking specific times
          roomStatuses[roomName].bookingsToday.push({
            booker: booking[bookerIndex],
            purpose: booking[purposeIndex],
            startTime: bookingStart.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' }),
            endTime: bookingEnd.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' })
          });
        }
      }
    }
    return { status: 'success', data: roomStatuses };
  } catch (e) {
    Logger.log('Error getting room statuses: ' + e.toString());
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการดึงสถานะห้อง: ' + e.message };
  }
}


function submitFeedback(feedbackData) {
  try {
    const sheet = getFeedbackSheet();
    sheet.appendRow([
      feedbackData.name,
      feedbackData.department,
      feedbackData.feedbackText,
      new Date() // Timestamp
    ]);
    return { status: 'success', message: 'ส่งข้อเสนอแนะสำเร็จ ขอบคุณครับ/ค่ะ' };
  } catch (e) {
    Logger.log('Feedback Error: ' + e.toString());
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการส่งข้อเสนอแนะ: ' + e.message };
  }
}


// --- Admin Functions ---
function getAllUsers() {
  if (!isAdminLoggedIn()) return { status: 'error', message: 'ไม่ได้รับอนุญาต' };
  try {
    const sheet = getUsersSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const users = [];
    // Skip password column for security when displaying to admin
    const nameIndex = headers.indexOf('ชื่อ');
    const deptIndex = headers.indexOf('กลุ่มสาระ');
    const phoneIndex = headers.indexOf('เบอร์โทรศัพท์');
    const regDateIndex = headers.indexOf('Timestamp การลงทะเบียน'); // Assuming you add this

    if ([nameIndex, deptIndex, phoneIndex].includes(-1)) {
        return { status: 'error', message: 'โครงสร้างชีทผู้ใช้ไม่ถูกต้อง (ไม่พบ Header ที่ต้องการสำหรับแสดงรายชื่อ)' };
    }


    for (let i = 1; i < data.length; i++) {
      users.push({
        name: data[i][nameIndex],
        department: data[i][deptIndex],
        phone: data[i][phoneIndex],
        registered: regDateIndex !== -1 && data[i][regDateIndex] ? new Date(data[i][regDateIndex]).toLocaleDateString('th-TH') : 'N/A'
      });
    }
    return { status: 'success', data: users };
  } catch (e) {
    Logger.log('Error getting all users: ' + e.toString());
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการดึงข้อมูลผู้ใช้: ' + e.message };
  }
}

function adminDeleteUser(username) {
  if (!isAdminLoggedIn()) return { status: 'error', message: 'ไม่ได้รับอนุญาต' };
  try {
    const sheet = getUsersSheet();
    const data = sheet.getDataRange().getValues();
    const nameIndex = data[0].indexOf('ชื่อ');
    for (let i = 1; i < data.length; i++) {
      if (data[i][nameIndex] === username) {
        sheet.deleteRow(i + 1); // Rows are 1-indexed
        return { status: 'success', message: `ผู้ใช้ ${username} ถูกลบแล้ว` };
      }
    }
    return { status: 'error', message: `ไม่พบผู้ใช้ ${username}` };
  } catch (e) {
    Logger.log('Error deleting user: ' + e.toString());
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการลบผู้ใช้: ' + e.message };
  }
}

function adminResetUserPassword(username, newPassword) {
  if (!isAdminLoggedIn()) return { status: 'error', message: 'ไม่ได้รับอนุญาต' };
  try {
    const sheet = getUsersSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameIndex = headers.indexOf('ชื่อ');
    const passwordIndex = headers.indexOf('รหัสผ่าน');

    if (nameIndex === -1 || passwordIndex === -1) {
       return { status: 'error', message: 'โครงสร้างชีทผู้ใช้ไม่ถูกต้อง (ไม่พบ Header ชื่อหรือรหัสผ่าน)' };
    }


    for (let i = 1; i < data.length; i++) {
      if (data[i][nameIndex] === username) {
        const hashedNewPassword = simpleHash(newPassword);
        sheet.getRange(i + 1, passwordIndex + 1).setValue(hashedNewPassword);
        // Log this action for security audit
        Logger.log(`Admin reset password for user: ${username}`);
        return { status: 'success', message: `รหัสผ่านสำหรับผู้ใช้ ${username} ถูกรีเซ็ตแล้ว (รหัสผ่านใหม่: ${newPassword})` }; // For dev; don't show new pass in prod
      }
    }
    return { status: 'error', message: `ไม่พบผู้ใช้ ${username}` };
  } catch (e) {
    Logger.log('Error resetting password: ' + e.toString());
    return { status: 'error', message: 'เกิดข้อผิดพลาดในการรีเซ็ตรหัสผ่าน: ' + e.message };
  }
}

function getAdminDashboardStats(month, year) {
    if (!isAdminLoggedIn()) return { status: 'error', message: 'ไม่ได้รับอนุญาต' };
    try {
        const sheet = getCustomerSheet();
        const data = sheet.getDataRange().getValues();
        const headers = data[0];

        const dateStartIndex = headers.indexOf('วันที่และเวลาที่เริ่มต้น');
        const roomNameIndex = headers.indexOf('ชื่อห้องประชุม');
        const purposeIndex = headers.indexOf('วัตถุประสงค์ในการจอง');

        if ([dateStartIndex, roomNameIndex, purposeIndex].includes(-1)) {
            return { status: 'error', message: 'โครงสร้างชีทข้อมูลการจองไม่ถูกต้อง (ไม่พบ Header ที่ต้องการสำหรับสถิติ)' };
        }

        let totalBookingsThisMonth = 0;
        const roomCounts = {};
        const purposeCounts = {};
        const targetMonth = parseInt(month); // JS month is 0-indexed
        const targetYear = parseInt(year);

        for (let i = 1; i < data.length; i++) {
            const record = data[i];
            const bookingDate = new Date(record[dateStartIndex]);

            if (bookingDate.getFullYear() === targetYear && bookingDate.getMonth() === targetMonth) {
                totalBookingsThisMonth++;

                const roomName = record[roomNameIndex];
                roomCounts[roomName] = (roomCounts[roomName] || 0) + 1;

                const purpose = record[purposeIndex];
                purposeCounts[purpose] = (purposeCounts[purpose] || 0) + 1;
            }
        }

        let mostBookedRoom = 'N/A';
        let maxRoomBookings = 0;
        for (const room in roomCounts) {
            if (roomCounts[room] > maxRoomBookings) {
                mostBookedRoom = room;
                maxRoomBookings = roomCounts[room];
            }
        }

        let popularPurpose = 'N/A';
        let maxPurposeCount = 0;
        let totalPurposeInstances = 0;
        for (const purpose in purposeCounts) {
            totalPurposeInstances += purposeCounts[purpose];
            if (purposeCounts[purpose] > maxPurposeCount) {
                popularPurpose = purpose;
                maxPurposeCount = purposeCounts[purpose];
            }
        }
        const popularPurposePercentage = totalPurposeInstances > 0 ? ((maxPurposeCount / totalPurposeInstances) * 100).toFixed(1) : 0;


        return {
            status: 'success',
            data: {
                totalBookingsThisMonth: totalBookingsThisMonth,
                mostBookedRoom: { name: mostBookedRoom, count: maxRoomBookings },
                popularPurpose: { name: popularPurpose, percentage: popularPurposePercentage }
            }
        };

    } catch (e) {
        Logger.log('Error fetching admin stats: ' + e.toString());
        return { status: 'error', message: 'เกิดข้อผิดพลาดในการดึงข้อมูลสถิติ: ' + e.message };
    }
}

function exportDataToSheet() {
    if (!isAdminLoggedIn()) return { status: 'error', message: 'ไม่ได้รับอนุญาต' };
    try {
        const sourceSheet = getCustomerSheet();
        const data = sourceSheet.getDataRange().getValues();

        // Create a new spreadsheet for export
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
        const newSpreadsheetName = "Exported_Booking_Data_" + timestamp;
        const newSpreadsheet = SpreadsheetApp.create(newSpreadsheetName);
        const newSheet = newSpreadsheet.getSheets()[0]; // Get the first sheet of the new spreadsheet

        // Write data to the new sheet
        newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

        // Get the URL of the new spreadsheet
        const fileUrl = newSpreadsheet.getUrl();
        const fileId = newSpreadsheet.getId();

        // Optionally, move the file to a specific folder or set permissions
        // DriveApp.getFileById(fileId).moveTo(DriveApp.getFolderById("YOUR_FOLDER_ID"));

        return {
            status: 'success',
            message: 'ข้อมูลถูก Export ไปยัง Google Sheet ใหม่สำเร็จแล้ว',
            fileUrl: fileUrl,
            fileName: newSpreadsheetName
        };

    } catch (e) {
        Logger.log('Error exporting data: ' + e.toString());
        return { status: 'error', message: 'เกิดข้อผิดพลาดในการ Export ข้อมูล: ' + e.message };
    }
}


// Include function to serve CSS or other static assets if needed
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
