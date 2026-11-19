/**
 * Google Apps Script សម្រាប់ភ្ជាប់ជាមួយ "ប្រព័ន្ធគ្រប់គ្រងកម្មវិធីជំនួយសង្គម"
 * វិទ្យាល័យ ហ៊ុន សែន អង្គប្រីយ៍
 * 
 * របៀបដំឡើង (Setup Instructions):
 * 1. បង្កើត Google Sheet ថ្មីមួយនៅលើ Google Drive របស់អ្នក
 * 2. ចុច Extensions (ផ្នែកបន្ថែម) -> Apps Script
 * 3. លុបកូដចាស់ចេញ រួចចម្លង (Copy) កូដខាងក្រោមនេះទាំងស្រុងទៅបិទភ្ជាប់ (Paste)
 * 4. ចុចប៊ូតុង Deploy (ដាក់ឱ្យប្រើប្រាស់) -> New deployment (ការដាក់ឱ្យប្រើប្រាស់ថ្មី)
 * 5. ត្រង់ Select type ជ្រើសរើស "Web app"
 * 6. ត្រង់ Description ដាក់ "Social Assistance Sync"
 * 7. ត្រង់ Execute as ជ្រើសរើស "Me (អ៊ីមែលរបស់អ្នក)"
 * 8. ត្រង់ Who has access ជ្រើសរើស "Anyone" (សំខាន់ខ្លាំងណាស់ ដើម្បីឱ្យប្រព័ន្ធអាច Sync បាន)
 * 9. ចុច Deploy -> Review permissions (ផ្តល់សិទ្ធិឱ្យគណនីរបស់អ្នក) -> Advanced -> Go to Untitled project (unsafe) -> Allow
 * 10. ចម្លង (Copy) "Web app URL" យកទៅបិទភ្ជាប់ (Paste) ក្នុងទំព័រ Settings នៃកម្មវិធីជាការស្រេច!
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    if (action === "sync_all") {
      var students = data.students || [];
      sheet.clear();
      
      // Headers
      var headers = [
        "ល.រ", "អត្តលេខ", "សិស្សឈ្មោះ", "ភេទ", "ថ្ងៃខែឆ្នាំកំណើត", "ទីកន្លែងកំណើត",
        "ឪពុកឈ្មោះ", "មុខរបរឪពុក", "ម្តាយឈ្មោះ", "មុខរបរម្តាយ", "អាសយដ្ឋានបច្ចុប្បន្ន",
        "អាព្យាបាលឈ្មោះ", "មុខរបរអាព្យាបាល", "ត្រូវជា", "ថ្នាក់ទី", "ឆ្នាំសិក្សា",
        "ប្រភេទបណ្ណសមធម៌ក្រីក្រ", "លេខប័ណ្ណសមធម៌", "លេខទូរស័ព្ទ", "កំណត់សម្គាល់"
      ];
      
      // Title
      sheet.getRange(1, 1).setValue("បញ្ជីរាយនាមសិស្សទទួលកម្មវិធីជំនួយសង្គម - " + (data.school || "វិទ្យាល័យ ហ៊ុន សែន អង្គប្រីយ៍"));
      sheet.getRange(1, 1).setFontSize(14).setFontWeight("bold");
      
      // Write headers
      sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
      var headerRange = sheet.getRange(3, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#1F4E79");
      headerRange.setFontColor("#FFFFFF");
      
      if (students.length > 0) {
        var rows = [];
        for (var i = 0; i < students.length; i++) {
          var s = students[i];
          rows.push([
            i + 1,
            s.student_code || "",
            s.full_name || "",
            s.gender || "",
            s.dob || "",
            s.pob || "",
            s.father_name || "",
            s.father_job || "",
            s.mother_name || "",
            s.mother_job || "",
            s.current_address || "",
            s.guardian_name || "",
            s.guardian_job || "",
            s.guardian_relation || "",
            s.grade || "",
            s.academic_year || "",
            s.equity_card_type || "",
            s.equity_card_no || "",
            s.phone || "",
            s.note || ""
          ]);
        }
        sheet.getRange(4, 1, rows.length, headers.length).setValues(rows);
      }
      
      sheet.autoResizeColumns(1, headers.length);
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "បានធ្វើសមកាលកម្មទិន្នន័យចំនួន " + students.length + " នាក់ទៅ Google Sheet រួចរាល់!"
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "មិនស្គាល់ Action"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    
    if (lastRow < 4) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        students: []
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var data = sheet.getRange(4, 1, lastRow - 3, 20).getValues();
    var students = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (row[2]) { // If student name exists
        students.push({
          student_code: String(row[1] || ""),
          full_name: String(row[2] || ""),
          gender: String(row[3] || ""),
          dob: String(row[4] || ""),
          pob: String(row[5] || ""),
          father_name: String(row[6] || ""),
          father_job: String(row[7] || ""),
          mother_name: String(row[8] || ""),
          mother_job: String(row[9] || ""),
          current_address: String(row[10] || ""),
          guardian_name: String(row[11] || ""),
          guardian_job: String(row[12] || ""),
          guardian_relation: String(row[13] || ""),
          grade: String(row[14] || ""),
          academic_year: String(row[15] || ""),
          equity_card_type: String(row[16] || ""),
          equity_card_no: String(row[17] || ""),
          phone: String(row[18] || ""),
          note: String(row[19] || "")
        });
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      students: students
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
