// ============================================================
//  Always On — Google Sheets Storage
//  Handles: Onboarding form + Company Knowledge submissions
//
//  SETUP:
//  1. Open a new Google Sheet
//  2. Extensions → Apps Script → paste this entire file
//  3. Save → Deploy → New deployment
//     · Type: Web app
//     · Execute as: Me
//     · Who has access: Anyone
//  4. Click Deploy → copy the Web App URL
//  5. Paste the URL into both HTML files where it says:
//     PASTE_YOUR_APPS_SCRIPT_URL_HERE
// ============================================================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.form_type === 'knowledge') {
      saveKnowledge(ss, data);
    } else if (data.form_type === 'onboarding') {
      saveOnboarding(ss, data);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── KNOWLEDGE SUBMISSIONS ──────────────────────────────────
function saveKnowledge(ss, d) {
  const TAB = 'Knowledge Submissions';
  const HEADERS = [
    'Timestamp', 'Company', 'Submitter Name', 'Submitter Email', 'Submitter Role',
    '1. Company Overview',
    '2. Who We Serve & Don\'t Serve',
    '3. Products & Services',
    '4. Pricing & MOQs',
    '5. Technical Knowledge & FAQs',
    '6. Sales Process',
    '7. Team & Territories',
    '8. Competitors & Positioning'
  ];

  const sheet = getOrCreateSheet(ss, TAB, HEADERS);

  sheet.appendRow([
    new Date(),
    d.company_name        || '',
    d.submitter_name      || '',
    d.submitter_email     || '',
    d.submitter_role      || '',

    // Section 1 — Company Overview
    joinFields([
      label('What the company does',    d.what_company_does),
      label('Industries / sectors',     d.sectors_served),
      label('Geographic coverage',      d.geographic_coverage),
      label('Company size',             d.company_size),
      label('Key differentiators',      d.key_differentiators),
    ]),

    // Section 2 — Who We Serve
    joinFields([
      label('Ideal customer',           d.ideal_customer),
      label('Sectors targeted',         d.sectors_targeted),
      label('Who we don\'t serve',      d.who_we_dont_serve),
      label('Minimum order value',      d.min_order_value),
      label('Restrictions',             d.restrictions),
    ]),

    // Section 3 — Products & Services
    joinFields([
      label('Core products',            d.core_products),
      label('Key brands',               d.key_brands),
      label('Popular products',         d.popular_products),
      label('Services offered',         d.services_offered),
      label('What we don\'t do',        d.what_we_dont_do),
    ]),

    // Section 4 — Pricing & MOQs
    joinFields([
      label('Pricing structure',        d.pricing_structure),
      label('Minimum order quantities', d.moq),
      label('Lead times',               d.lead_times),
      label('Credit terms',             d.credit_terms),
      label('Delivery charges',         d.delivery_charges),
    ]),

    // Section 5 — Technical Knowledge
    joinFields([
      label('Common questions',         d.common_questions),
      label('Key answers',              d.key_answers),
      label('Certifications',           d.certifications),
      label('Technical support',        d.technical_support),
    ]),

    // Section 6 — Sales Process
    joinFields([
      label('Enquiry handling',         d.enquiry_handling),
      label('Escalation triggers',      d.escalation_triggers),
      label('Qualifying questions',     d.qualifying_questions),
      label('Handoff style',            d.handoff_style),
    ]),

    // Section 7 — Team & Territories
    joinFields([
      label('Sales reps',               d.sales_reps),
      label('Internal contacts',        d.internal_contacts),
      label('Office hours',             d.office_hours),
      label('Out-of-hours',             d.out_of_hours),
    ]),

    // Section 8 — Competitors & Positioning
    joinFields([
      label('Main competitors',         d.main_competitors),
      label('Positioning',              d.competitive_positioning),
      label('Differentiators to use',   d.differentiators_to_emphasise),
    ]),
  ]);
}

// ── ONBOARDING SUBMISSIONS ─────────────────────────────────
function saveOnboarding(ss, d) {
  const TAB = 'Onboarding';
  const HEADERS = [
    'Timestamp', 'Company', 'Primary Contact', 'Email',
    'Website', 'Website Platform', 'Industry',
    'CRM', 'Geography', 'Company Size',
    'Full Submission (JSON)'
  ];

  const sheet = getOrCreateSheet(ss, TAB, HEADERS);

  sheet.appendRow([
    new Date(),
    d.company_name             || '',
    d.primary_contact_name     || '',
    d.primary_contact_email    || '',
    d.website                  || '',
    d.website_platform         || '',
    d.industry                 || '',
    d.crm                      || '',
    d.geography                || '',
    d.company_size             || '',
    JSON.stringify(d, null, 2),
  ]);
}

// ── HELPERS ───────────────────────────────────────────────
function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold')
               .setBackground('#43165c')
               .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    // Set wrap for all content columns
    sheet.getRange('A:Z').setWrap(true).setVerticalAlignment('top');
    // Widen columns
    sheet.setColumnWidth(1, 160);  // Timestamp
    sheet.setColumnWidth(2, 180);  // Company
    sheet.setColumnWidth(3, 160);  // Contact
    sheet.setColumnWidth(4, 200);  // Email
    for (let i = 5; i <= headers.length; i++) {
      sheet.setColumnWidth(i, 280);
    }
  }
  return sheet;
}

function label(name, value) {
  if (!value || !value.trim()) return null;
  return name + ':\n' + value.trim();
}

function joinFields(parts) {
  return parts.filter(Boolean).join('\n\n');
}
