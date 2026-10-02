/**
 * MAPSPILOT — Self-Hosted Lead Engine (Google Apps Script)
 * Runs 100% on your Google account. Free forever. No external services.
 *
 * INSTALL (5 minutes):
 * 1. Open the "Untitled form (Responses)" spreadsheet that is linked to your
 *    MapsPilot booking form (the one with columns: Timestamp, Business name,
 *    Business type, City / Area, Your WhatsApp number, What do you need?).
 * 2. Extensions → Apps Script. Delete anything there, paste ALL of this code.
 * 3. Deploy → New deployment → type: Web app.
 *    Execute as: Me.  Who has access: Anyone.
 *    Click Deploy, authorize, and copy the Web app URL (ends in /exec).
 * 4. Send that /exec URL to your assistant — the dashboard gets wired to it.
 *
 * WHAT IT DOES:
 * - Every new booking on the website lands in this Sheet → you get an instant
 *   email notification with the lead's details.
 * - The dashboard reads live leads from here (auto-refresh, real time).
 * - Dashboard status changes (new/contacted/won/lost) are written here instantly.
 * - Every lead automatically gets a ready-to-paste service kit.
 */

var OWNER_EMAIL = "somilsharma2000@gmail.com";   // ← where lead notifications go
var DASHBOARD_URL = "https://somilsharma2000.github.io/mapspilot/dashboard.html";

// ------------------------------------------------------------------
// WEB APP ENDPOINTS
// ------------------------------------------------------------------

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === "set" && p.id && p.status) {
    setLeadStatus(p.id, p.status);
    return output({ ok: true }, p);
  }
  // default: return all leads (with statuses + kits applied)
  return output({ leads: getLeads() }, p);
}

function doPost(e) {
  var p = (e && e.parameter) || {};
  if (p.action === "set" && p.id && p.status) {
    setLeadStatus(p.id, p.status);
    return output({ ok: true }, p);
  }
  return output({ leads: getLeads() }, p);
}

// Supports plain JSON and JSONP (callback=...) so any browser can read it.
function output(obj, p) {
  if (p && p.callback) {
    return ContentService
      .createTextOutput(p.callback + "(" + JSON.stringify(obj) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------------
// INSTANT EMAIL NOTIFICATION ON EVERY NEW BOOKING
// ------------------------------------------------------------------

function onFormSubmit(e) {
  try {
    var v = e.namedValues || {};
    var name = v["Business name"] || "New business";
    var subject = "🔥 MapsPilot lead: " + name;
    var body =
      "New booking just came in!\n\n" +
      "Business: " + (v["Business name"] || "-") + "\n" +
      "Type: " + (v["Business type"] || "-") + "\n" +
      "City: " + (v["City / Area"] || "-") + "\n" +
      "WhatsApp: " + (v["Your WhatsApp number"] || "-") + "\n" +
      "They need: " + (v["What do you need?"] || "-") + "\n\n" +
      "Open your Control Center: " + DASHBOARD_URL + "\n" +
      "Call them from the dashboard — the service kit is ready for copy-paste.";
    MailApp.sendEmail(OWNER_EMAIL, subject, body);
  } catch (err) {
    // never let a notification error block anything
  }
}

// ------------------------------------------------------------------
// LEADS
// ------------------------------------------------------------------

function getLeads() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheets()[0]; // "Form Responses 1"
  var data = formSheet.getDataRange().getValues();
  var statuses = getStatusMap(ss);

  var leads = [];
  for (var i = 1; i < data.length; i++) { // row 0 = headers
    var row = data[i];
    if (!row[1]) continue; // skip empty rows
    var id = "R" + (i + 1); // stable id (form responses only append)
    var ts = row[0] instanceof Date ? row[0] : new Date();
    var type = String(row[2] || "Other");
    var lead = {
      id: id,
      businessName: String(row[1]),
      businessType: type,
      city: String(row[3] || ""),
      phone: String(row[4] || ""),
      need: String(row[5] || ""),
      status: statuses[id] || "new",
      created: Utilities.formatDate(ts, "Asia/Calcutta", "yyyy-MM-dd HH:mm"),
      createdTs: ts.getTime(),
      kit: makeKit(String(row[1]), type, String(row[3] || "your area"))
    };
    leads.push(lead);
  }
  return leads;
}

function setLeadStatus(id, status) {
  var allowed = ["new", "contacted", "won", "lost"];
  if (allowed.indexOf(status) === -1) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Statuses");
  if (!sh) sh = ss.insertSheet("Statuses");
  sh.appendRow([id, status, new Date()]);
}

function getStatusMap(ss) {
  var sh = ss.getSheetByName("Statuses");
  var map = {};
  if (!sh) return map;
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) { // last status wins
    if (rows[i][0]) map[String(rows[i][0])] = String(rows[i][1]);
  }
  return map;
}

// ------------------------------------------------------------------
// SERVICE KITS (templates per business type, filled automatically)
// ------------------------------------------------------------------

function makeKit(name, type, city) {
  var t = KITS[type] || KITS["Other"];
  var f = function(s) {
    return String(s)
      .split("{{name}}").join(name)
      .split("{{city}}").join(city)
      .split("{{type}}").join(type);
  };
  var kit = {};
  for (var k in t) kit[k] = Array.isArray(t[k]) ? t[k].map(f) : f(t[k]);
  kit.review = t.review.map(f).join("\n");
  kit.audit = t.audit.map(f).join("\n");
  kit.categories = f(t.categories);
  return kit;
}

var KITS = {
  "Salon / Barber": {
    pitch: "Hi! This is MapsPilot. You asked for the free Google profile check for {{name}}. I've finished the first look — can I send you the 5-point report on WhatsApp today?",
    description: "{{name}} is a trusted salon in {{city}}. We offer haircuts, styling, colour, facials and grooming for men and women. Our stylists take time to understand the look you want, and we keep every tool clean and hygienic. Walk in for a quick trim or book a full makeover. Easy to find, friendly staff, honest prices.",
    categories: "1) Hair Salon  2) Beauty Salon  3) Barber Shop — pick the first one as primary.",
    welcome: "Welcome to {{name}} on Google! Find us in {{city}}. Walk in or call to book your slot — we're here all week.",
    offer: "Free hair consultation + 10% off your first colour or facial this month. Show this post at the counter.",
    review: [
      "1) Positive: \"Thank you so much for your kind words! We're glad you loved the result. See you at your next appointment.\"",
      "2) Mixed: \"Thank you for the feedback. We're sorry about the wait — we're adjusting booking slots. Call ahead and we'll fit you in faster.\"",
      "3) Negative: \"We're sorry your visit didn't go well. Please call us directly so we can understand and make it right.\""
    ],
    audit: [
      "1) Photos: add 10-15 clear photos (front board, inside, stylist at work, before/after).",
      "2) Hours: confirm opening days and times, add holiday hours.",
      "3) Category: primary category should be Hair Salon.",
      "4) Services: list haircut, colour, facial, bridal, kids cuts as services with prices.",
      "5) Reviews: ask every happy customer for a Google review using the review link."
    ]
  },
  "Clinic / Dentist": {
    pitch: "Hi! This is MapsPilot. You asked for the free Google profile check for {{name}}. I've finished the first look at the clinic's profile — can I send you the 5-point report on WhatsApp today?",
    description: "Trusted clinic in {{city}}. We offer caring, personal treatment for your whole family. Our doctors take time to listen, explain clearly, and treat with modern equipment. Walk in for general consultations, or call to book your visit. Easy parking nearby. Open six days a week for your convenience.",
    categories: "1) Medical Clinic  2) Doctor  3) Medical Center — pick the first one as primary.",
    welcome: "Welcome to {{name}} on Google! Find us in {{city}}. Call to book your visit — we're here six days a week.",
    offer: "Free consultation for first-time patients this month. Show this post at the counter. Call to book.",
    review: [
      "1) Positive: \"Thank you so much for your kind words! We're glad you had a good experience. See you at your next visit.\"",
      "2) Mixed: \"Thank you for the feedback. We're sorry about the wait — we're adjusting appointment slots. Please call ahead and we'll fit you in faster.\"",
      "3) Negative: \"We're sorry your visit didn't go well. Please call us directly so we can understand and make it right.\""
    ],
    audit: [
      "1) Photos: add 8-10 clear clinic photos (front board, waiting area, doctor, equipment).",
      "2) Hours: confirm opening days/times, add holiday hours.",
      "3) Category: check the primary category is set correctly.",
      "4) Services: list consultations, checkups, follow-ups as services.",
      "5) Reviews: ask every happy patient for a Google review using the review link."
    ]
  },
  "Gym / Trainer": {
    pitch: "Hi! This is MapsPilot. You asked for the free Google profile check for {{name}}. I've finished the first look — can I send you the 5-point report on WhatsApp today?",
    description: "{{name}} in {{city}} — the place to get strong, fit, and confident. Personal training, group classes, weight loss and muscle gain programs. Certified trainers, clean equipment, and a plan made for YOUR body and goals. First trial session free. Come see the gym before you join.",
    categories: "1) Gym  2) Physical Fitness Program  3) Personal Trainer — pick the first one as primary.",
    welcome: "Welcome to {{name}} on Google! Find us in {{city}}. Your first trial session is free — call to book it.",
    offer: "Free trial session + free fitness assessment for new members this month. Show this post at the desk.",
    review: [
      "1) Positive: \"Thank you for the great words! Great to see your progress — keep pushing. See you at your next session.\"",
      "2) Mixed: \"Thanks for the honest feedback. We're sorry the timing didn't work — ask us about off-peak slots, they're much freer.\"",
      "3) Negative: \"We're sorry your experience wasn't right. Please call us directly so we can fix it.\""
    ],
    audit: [
      "1) Photos: add 10+ photos (equipment, training floor, trainers in action, front board).",
      "2) Hours: confirm timings, separate peak/off-peak if relevant.",
      "3) Category: primary category should be Gym.",
      "4) Services: list personal training, group classes, diet plans as services.",
      "5) Reviews: ask every member after their free trial for a Google review."
    ]
  },
  "Restaurant / Cafe": {
    pitch: "Hi! This is MapsPilot. You asked for the free Google profile check for {{name}}. I've finished the first look — can I send you the 5-point report on WhatsApp today?",
    description: "{{name}} in {{city}} — fresh, flavourful food made with care. Come for a quick bite or a full family meal. Clean kitchen, friendly service, and honest prices. Our regulars say it best: once you try us, you come back. Dine in, take away, or call ahead for your order.",
    categories: "1) Restaurant  2) Cafe  3) Coffee Shop — pick the first one as primary.",
    welcome: "Welcome to {{name}} on Google! Find us in {{city}}. Dine in or call ahead for takeaway — fresh food all day.",
    offer: "Free soft drink with any family combo this month. Show this post when you order.",
    review: [
      "1) Positive: \"Thank you so much! We're glad you enjoyed the food. See you again soon.\"",
      "2) Mixed: \"Thank you for the feedback. We're sorry about the wait — call ahead for takeaway and it'll be ready when you arrive.\"",
      "3) Negative: \"We're sorry your visit wasn't good. Please call us directly so we can make it right.\""
    ],
    audit: [
      "1) Photos: add 12+ real food photos (menu favourites, interiors, front board).",
      "2) Hours: confirm opening hours, mark kitchen closing time.",
      "3) Category: primary category should be Restaurant.",
      "4) Menu: upload the menu; add popular dishes as products if possible.",
      "5) Reviews: ask happy diners for a Google review — reply to every one."
    ]
  },
  "Electrician / Plumber / AC Repair": {
    pitch: "Hi! This is MapsPilot. You asked for the free Google profile check for {{name}}. I've finished the first look — can I send you the 5-point report on WhatsApp today?",
    description: "{{name}} serves {{city}} with fast, reliable electrical, plumbing and AC work. Same-day visits, honest quotes before we start, and work done right the first time. No call-out drama, no surprise bills. Call anytime for repairs, installations, or a yearly service check.",
    categories: "1) Electrician  2) Plumber  3) Air Conditioning Repair Service — pick the first one as primary.",
    welcome: "Welcome to {{name}} on Google! Serving {{city}}. Call for same-day repair — honest quotes, guaranteed work.",
    offer: "Free inspection with any repair booked this month. Mention this post when you call.",
    review: [
      "1) Positive: \"Thank you for trusting us! Glad the job went smooth. Call anytime you need us again.\"",
      "2) Mixed: \"Thanks for the honest feedback. We're sorry about the timing — call ahead and we'll give you a proper slot.\"",
      "3) Negative: \"We're sorry something wasn't right. Please call us directly so we can come back and fix it.\""
    ],
    audit: [
      "1) Photos: add 8-10 photos (team at work, van, tools, finished jobs).",
      "2) Service area: add all areas you serve.",
      "3) Category: primary category should be Electrician.",
      "4) Services: list repairs, installations, AC service with starting prices.",
      "5) Reviews: after every completed job, send the customer the review link."
    ]
  },
  "Coaching / Tuition": {
    pitch: "Hi! This is MapsPilot. You asked for the free Google profile check for {{name}}. I've finished the first look — can I send you the 5-point report on WhatsApp today?",
    description: "{{name}} in {{city}} — small batches, personal attention, real results. Our teachers explain concepts clearly, track every student's progress, and keep parents updated. Regular tests, doubt sessions, and exam-focused practice. Book a free demo class and see the difference before you join.",
    categories: "1) Coaching Center  2) Tutoring Service  3) Educational Institution — pick the first one as primary.",
    welcome: "Welcome to {{name}} on Google! Find us in {{city}}. Book a free demo class — call today.",
    offer: "Free demo class + free assessment for new students this month. Mention this post when you call.",
    review: [
      "1) Positive: \"Thank you! Great to hear about the improvement. Hard work pays — see you in class.\"",
      "2) Mixed: \"Thank you for the feedback. We're sorry about the batch timing — ask us about our other slots.\"",
      "3) Negative: \"We're sorry your experience wasn't good. Please call us directly so we can sort it out.\""
    ],
    audit: [
      "1) Photos: add 8-10 photos (classroom, teachers, students in class, results wall).",
      "2) Hours: confirm batch timings on the profile.",
      "3) Category: primary category should be Coaching Center.",
      "4) Services: list subjects, classes, and boards you cover.",
      "5) Reviews: ask parents of improving students for a Google review."
    ]
  },
  "Retail Shop": {
    pitch: "Hi! This is MapsPilot. You asked for the free Google profile check for {{name}}. I've finished the first look — can I send you the 5-point report on WhatsApp today?",
    description: "{{name}} in {{city}} — everything you need, at fair prices, all in one place. Fresh stock, honest billing, and a shopkeeper who knows what you're looking for. Walk in today, or call to check availability before you come.",
    categories: "1) Store  2) General Store — pick the first one as primary (edit to your exact shop type).",
    welcome: "Welcome to {{name}} on Google! Find us in {{city}}. Walk in or call ahead to check stock.",
    offer: "Special opening offer this month — ask at the counter and show this post.",
    review: [
      "1) Positive: \"Thank you for shopping with us! See you again soon.\"",
      "2) Mixed: \"Thanks for the feedback — we're sorry about the stock issue. Call ahead and we'll keep it ready for you.\"",
      "3) Negative: \"We're sorry about your experience. Please call us directly so we can make it right.\""
    ],
    audit: [
      "1) Photos: add 8-10 photos (front board, inside, popular products).",
      "2) Hours: confirm opening hours, add festival hours.",
      "3) Category: set the exact shop type as primary category.",
      "4) Products: list your main product categories.",
      "5) Reviews: ask regular customers for a Google review."
    ]
  },
  "Other": {
    pitch: "Hi! This is MapsPilot. You asked for the free Google profile check for {{name}}. I've finished the first look — can I send you the 5-point report on WhatsApp today?",
    description: "{{name}} in {{city}} — trusted local service with happy customers. Clear communication, honest pricing, and work done properly. Call us to discuss what you need, or walk in today.",
    categories: "Pick the closest matching primary category in Google's list, then 2-3 related ones as secondary.",
    welcome: "Welcome to {{name}} on Google! Find us in {{city}}. Call or visit — we're happy to help.",
    offer: "Special first-time customer offer this month. Mention this post.",
    review: [
      "1) Positive: \"Thank you so much! We're glad you were happy. See you again.\"",
      "2) Mixed: \"Thanks for the honest feedback — we'll improve. Please call ahead next time.\"",
      "3) Negative: \"We're sorry it wasn't right. Please call us directly so we can make it right.\""
    ],
    audit: [
      "1) Photos: add 8-10 clear photos (front board, inside, work in progress).",
      "2) Hours: confirm opening days and times.",
      "3) Category: set the best-matching primary category.",
      "4) Services: list what you offer with starting prices.",
      "5) Reviews: ask every happy customer for a Google review."
    ]
  }
};
