// ─── CONFIG ───────────────────────────────────────────────────────────────────
const SHEET_NAME    = 'Applications';
const PARENT_FOLDER = 'CP Applications';

// ─── POST HANDLER ─────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 1. Drive: get/create CP Applications → Firm Name folder
    const mainFolder = getOrCreateFolder(PARENT_FOLDER);
    const firmName   = (data.firm_name || 'Unknown').replace(/[\/\\:*?"<>|]/g, '').trim();
    const firmFolder = getOrCreateFolder(firmName, mainFolder);

    // 2. Decode uploaded files into blobs
    const fileKeys   = ['aadhaar_front','aadhaar_back','pan_card','cancelled_cheque','applicant_photo','firm_registration'];
    const fileLabels = {
      aadhaar_front:     'Aadhaar Front',
      aadhaar_back:      'Aadhaar Back',
      pan_card:          'PAN Card',
      cancelled_cheque:  'Cancelled Cheque',
      applicant_photo:   'Applicant Photo',
      firm_registration: 'Firm Registration'
    };
    const blobs = {};
    fileKeys.forEach(key => {
      const b64  = data[key + '_b64'];
      const name = data[key + '_name'];
      const mime = data[key + '_mime'] || 'application/octet-stream';
      if (b64 && name) {
        blobs[key] = Utilities.newBlob(Utilities.base64Decode(b64), mime, name);
      }
    });

    // 3. Auto-merge all docs into one PDF
    const kycPdfUrl = createKycPdf(firmName, data.pan || 'UNKNOWN', blobs, fileLabels, firmFolder);

    // 4. Write row to Google Sheet
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let   sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      const headers = [
        'Submitted At','Branch','RM Name','RM Ecode','BM Name','BM Ecode',
        'Sourcing Type','Firm Name','Proprietor Name','DOB','Constitution',
        'PAN','Contact','Email','Business Vintage',
        'Res Address','Res Pincode','Res City',
        'Biz Address','Biz Pincode','Biz City','Office Status',
        'GST Reg','GST Number',
        'Bank Name','Bank Location','Acc Type','Acc Number',
        'Ref1 Name','Ref1 Company','Ref1 Contact',
        'Ref2 Name','Ref2 Company','Ref2 Contact',
        'BM Recommendation','Business Potential',
        'Drive Folder','KYC Documents (PDF)',
        'Status'
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#1a6b3c')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      data.branch,        data.rm_name,        data.rm_ecode,
      data.bm_name,       data.bm_ecode,
      data.sourcing_type, data.firm_name,       data.proprietor_name,
      data.dob,           data.constitution,
      data.pan,           data.contact,         data.cp_email,
      data.business_vintage,
      data.res_address,   data.res_pincode,     data.res_city,
      data.biz_address,   data.biz_pincode,     data.biz_city,
      data.office_status,
      data.gst_reg,       data.gst_number || '',
      data.bank_name,     data.bank_location,   data.acc_type,  data.acc_number,
      data.ref1_name,     data.ref1_company,    data.ref1_contact,
      data.ref2_name,     data.ref2_company,    data.ref2_contact,
      data.bm_rec,        data.biz_potential,
      firmFolder.getUrl(), kycPdfUrl,
      'Pending'
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── CREATE MERGED KYC PDF ────────────────────────────────────────────────────
function createKycPdf(firmName, pan, blobs, fileLabels, firmFolder) {
  const pres = SlidesApp.create('__temp_kyc__');

  // Title slide
  const titleSlide = pres.getSlides()[0];
  titleSlide.getBackground().setSolidFill('#1a6b3c');
  const tb = titleSlide.insertTextBox('KYC Documents\n' + firmName + '\nPAN: ' + pan);
  tb.setWidth(560).setHeight(220).setLeft(80).setTop(90);
  tb.getText().getTextStyle().setForegroundColor('#ffffff').setFontSize(22);

  const keys = ['aadhaar_front','aadhaar_back','pan_card','cancelled_cheque','applicant_photo','firm_registration'];

  keys.forEach(function(key) {
    const blob = blobs[key];
    if (!blob) return;

    const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    const label = fileLabels[key];

    // Label header
    const hdr = slide.insertTextBox(label);
    hdr.setWidth(700).setHeight(28).setLeft(10).setTop(5);
    hdr.getText().getTextStyle().setFontSize(13).setBold(true).setForegroundColor('#1a6b3c');

    const mime = blob.getContentType() || '';
    if (mime.startsWith('image/')) {
      try {
        const img   = slide.insertImage(blob);
        const origW = img.getWidth();
        const origH = img.getHeight();
        const scale = Math.min(700 / origW, 355 / origH);
        const newW  = origW * scale;
        const newH  = origH * scale;
        img.setWidth(newW);
        img.setHeight(newH);
        img.setLeft((720 - newW) / 2);
        img.setTop(35 + (370 - newH) / 2);
      } catch (e) {
        slide.insertTextBox('Could not embed image:\n' + e.message)
          .setWidth(600).setHeight(100).setLeft(60).setTop(150);
      }
    } else {
      // PDF or other — note it (Slides can't embed PDF pages)
      slide.insertTextBox('Document: ' + blob.getName() + '\n(PDF — view via Drive folder link)')
        .setWidth(600).setHeight(100).setLeft(60).setTop(150);
    }
  });

  // Export as PDF and save
  const presId   = pres.getId();
  const presFile = DriveApp.getFileById(presId);
  const pdfBlob  = presFile.getAs('application/pdf');
  pdfBlob.setName('KYC_' + firmName.replace(/\s+/g, '_') + '_' + pan + '.pdf');

  const pdfFile = firmFolder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  presFile.setTrashed(true); // clean up temp Slides file

  return pdfFile.getUrl();
}

// ─── HELPER ───────────────────────────────────────────────────────────────────
function getOrCreateFolder(name, parent) {
  const iter = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  return iter.hasNext()
    ? iter.next()
    : (parent ? parent.createFolder(name) : DriveApp.createFolder(name));
}
