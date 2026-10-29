const SHEET_ID = '1WQKsrBCiY6DrEnLSMq1jxH8fn29t0npvwKq1DdvYb3o';

function doGet(e) {
  if (e.parameter.action === 'getEventInfo') {
    const result = getEventInfo();
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid action" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  if (e.parameter.action === 'register') {
    try {
      const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("CongDangKyPrideTalk");
      if (!sheet) {
        return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Không tìm thấy Sheet lưu đăng ký.' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      
      const hoVaTen = e.parameter.HoVaTen;
      const soDienThoai = e.parameter.SoDienThoai;
      
      // Lấy thời gian hiện tại
      const time = new Date();
      const formattedTime = Utilities.formatDate(time, "GMT+7", "dd/MM/yyyy HH:mm:ss");

      // Lưu vào sheet CongDangKyPrideTalk
      sheet.appendRow([formattedTime, hoVaTen, soDienThoai]);

      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
}

function getEventInfo() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("LichPrideTalk");
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const colIdx = {
      MaSuKienTalk: headers.indexOf("MaSuKienTalk"),
      NgayToChuc: headers.indexOf("NgayToChuc"),
      TenChuDe: headers.indexOf("TenChuDe"),
      MoTa: headers.indexOf("MoTa"),
      DiaDiem: headers.indexOf("DiaDiem"),
      DangKyTu: headers.indexOf("DangKyTu"),
      DangKyDen: headers.indexOf("DangKyDen"),
      TrangThai: headers.indexOf("TrangThai")
    };

    let openEvents = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[colIdx.TrangThai] === "Mở đăng ký") {
        openEvents.push({
          MaSuKienTalk: row[colIdx.MaSuKienTalk],
          NgayToChuc: row[colIdx.NgayToChuc],
          TenChuDe: row[colIdx.TenChuDe],
          MoTa: row[colIdx.MoTa],
          DiaDiem: row[colIdx.DiaDiem],
          DangKyTu: row[colIdx.DangKyTu],
          DangKyDen: row[colIdx.DangKyDen]
        });
      }
    }

    if (openEvents.length === 0) {
      return { status: "not_found" };
    }

    // Lọc sự kiện có NgayToChuc lớn nhất
    openEvents.sort((a, b) => new Date(b.NgayToChuc).getTime() - new Date(a.NgayToChuc).getTime());
    const event = openEvents[0];

    const parseDate = (d) => (d instanceof Date) ? d.getTime() : (d ? new Date(d).getTime() : null);

    return {
      status: "success",
      data: {
        TenChuDe: event.TenChuDe,
        MoTa: event.MoTa,
        NgayToChuc: parseDate(event.NgayToChuc),
        DiaDiem: event.DiaDiem,
        DangKyTu: parseDate(event.DangKyTu),
        DangKyDen: parseDate(event.DangKyDen)
      }
    };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}