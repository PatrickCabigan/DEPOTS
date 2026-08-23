const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, HeadingLevel,
  PageOrientation, VerticalAlign
} = require("docx");

const LABEL_W = 2200;
const CONTENT_W = 7160;
const TOTAL_W = LABEL_W + CONTENT_W;

const HEADER_FILL = "7B1113";
const LABEL_FILL = "D9D9D9";

const borders = {
  top: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
  left: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
  right: { style: BorderStyle.SINGLE, size: 4, color: "999999" },
};

function para(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, bold: !!opts.bold, color: opts.color, italics: !!opts.italics })],
    spacing: { after: opts.after !== undefined ? opts.after : 60 },
  });
}

function multiPara(lines, opts = {}) {
  if (!lines || lines.length === 0) return [para("None.", { after: 0 })];
  return lines.map((l, i) => para(l, { after: i === lines.length - 1 ? 0 : 60 }));
}

function labelCell(text) {
  return new TableCell({
    width: { size: LABEL_W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: LABEL_FILL },
    verticalAlign: VerticalAlign.TOP,
    borders,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: [para(text, { bold: true, after: 0 })],
  });
}

function contentCell(children) {
  return new TableCell({
    width: { size: CONTENT_W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: "FFFFFF" },
    verticalAlign: VerticalAlign.TOP,
    borders,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: Array.isArray(children) ? children : [para(children, { after: 0 })],
  });
}

function headerRow(id, name) {
  return new TableRow({
    children: [
      new TableCell({
        columnSpan: 2,
        width: { size: TOTAL_W, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: HEADER_FILL },
        borders,
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: `${id}: ${name}`, bold: true, color: "FFFFFF", size: 26, underline: {} }),
            ],
          }),
        ],
      }),
    ],
  });
}

function fieldRow(label, content) {
  return new TableRow({ children: [labelCell(label), contentCell(content)] });
}

function buildUseCaseTable(uc) {
  const rows = [
    headerRow(uc.id, uc.name),
    fieldRow("Use Case ID", uc.id),
    fieldRow("Use Case Name", uc.name),
    fieldRow("Primary Actor(s)", uc.primary),
    fieldRow("Secondary Actor(s)", uc.secondary),
    fieldRow("Description", multiPara(uc.description)),
    fieldRow("Trigger", uc.trigger),
    fieldRow("Preconditions", multiPara(uc.preconditions)),
    fieldRow("Postconditions", multiPara(uc.postconditions)),
    fieldRow("Main Success Scenario", multiPara(uc.mainSuccess)),
    fieldRow("Alternate Flows", multiPara(uc.alternate)),
    fieldRow("Exception Flows", multiPara(uc.exception)),
    fieldRow("Business Rules", multiPara(uc.businessRules)),
    fieldRow("Priority", uc.priority),
    fieldRow("Frequency of Use", uc.frequency),
  ];
  if (uc.includes) {
    rows.push(fieldRow("Includes", uc.includes));
  }
  return new Table({
    width: { size: TOTAL_W, type: WidthType.DXA },
    columnWidths: [LABEL_W, CONTENT_W],
    rows,
  });
}

const useCases = [
  {
    id: "UC-01", name: "Register Vehicle",
    primary: "Student / Faculty / Staff",
    secondary: "BMO Administrator",
    description: ["A registered APC user enrolls a vehicle at the BMO office and is issued an RFID card linked to their APC ID for automated parking access."],
    trigger: "User visits the BMO office to register a vehicle for parking access.",
    preconditions: ["1. User has a valid APC ID.", "2. User has valid vehicle registration documents (OR/CR)."],
    postconditions: ["1. A user record is created with name, APC ID, role, and plate number.", "2. A unique RFID card UID is generated and linked to the user record.", "3. User receives a physical RFID card."],
    mainSuccess: [
      "1. User presents a valid APC ID and vehicle OR/CR to BMO staff.",
      "2. BMO Administrator enters the user's name, APC ID number, role, and plate number into the system.",
      "3. System generates a unique RFID card UID and links it to the user record.",
      "4. BMO Administrator assigns the physical RFID card to the user by tapping it on a reader.",
      "5. User receives the physical RFID card (credit-card size).",
    ],
    alternate: [
      "A1 \u2013 Lost or stolen card: BMO Administrator deactivates the old card UID and issues a replacement linked to the same user record.",
      "A2 \u2013 Second vehicle: A registered user requesting access for a second vehicle is processed through the visitor temporary-card flow (UC-06) instead of a new permanent registration, per APC's one-car-per-user policy.",
    ],
    exception: [
      "E1 \u2013 Duplicate registration: If the APC ID already has an active card, the system flags the duplicate and BMO Administrator resolves it before issuing a new card.",
      "E2 \u2013 Invalid documents: If the OR/CR or APC ID cannot be verified, registration is declined until valid documents are presented.",
    ],
    businessRules: ["\u2022 One RFID card per user under APC's one-car-per-user policy.", "\u2022 Vehicle information must be validated against a presented OR/CR before a card is issued."],
    priority: "High",
    frequency: "Once per user (occasional \u2013 new enrollments and replacements)",
  },
  {
    id: "UC-02", name: "Enter Parking via RFID",
    primary: "Student / Faculty / Staff",
    secondary: "RFID Reader",
    description: ["A registered driver taps an RFID card at the entrance to gain automated access to the parking facility."],
    trigger: "Driver arrives at the entrance and taps their RFID card on the reader.",
    preconditions: ["1. Driver has an active, non-expired RFID card.", "2. RFID Reader and boom barrier are operational."],
    postconditions: ["1. Entry is logged with name, role, plate number, timestamp, and camera snapshot.", "2. Boom barrier opens and the driver proceeds to park."],
    mainSuccess: [
      "1. Driver taps the RFID card on the entrance reader (within \u22485 cm).",
      "2. RFID Reader sends the card UID to the backend system for validation.",
      "3. System validates the card and checks whether the lot is full.",
      "4. If the card is valid and the lot is not full, the boom barrier opens within 2 seconds.",
      "5. System logs the entry with the driver's name, role, plate number, timestamp, and a camera snapshot.",
    ],
    alternate: [],
    exception: [
      "E1 \u2013 Invalid or expired card: Red LED lights up at the entrance and the boom barrier remains closed.",
      "E2 \u2013 Lot full: RFID Reader rejects the tap, the red LED lights up, and the barrier stays closed until a slot becomes available.",
      "E3 \u2013 RFID or sensor failure: Security Guard performs manual verification at the entrance and records the movement as a fallback.",
    ],
    businessRules: ["\u2022 Average authentication time must be under 2 seconds per vehicle during peak periods."],
    priority: "High",
    frequency: "Every registered-vehicle entry (\u2248100\u2013120 vehicles on peak days)",
  },
  {
    id: "UC-03", name: "View Parking Availability",
    primary: "Student / Faculty / Staff",
    secondary: "None",
    description: ["A driver inside the parking facility uses LED slot indicators to locate an available parking space."],
    trigger: "Driver enters the basement parking area after passing the entrance.",
    preconditions: ["1. LED indicators are installed and functioning above each slot.", "2. Detect Occupancy Status (UC-04) is running and providing current slot data."],
    postconditions: ["Driver identifies and proceeds to an available (green) slot."],
    mainSuccess: [
      "1. Driver enters the parking area.",
      "2. Each slot's LED indicator displays its current status (green = free, red = occupied), continuously refreshed by the included Detect Occupancy Status use case.",
      "3. Driver visually scans the LEDs and proceeds to a free slot.",
    ],
    alternate: [],
    exception: ["E1 \u2013 All slots occupied: All LEDs show red and the system marks the lot as full, blocking further entries until a slot frees up."],
    businessRules: ["\u2022 LED status must update within 5 seconds of a vehicle parking or leaving a slot."],
    priority: "High",
    frequency: "Every entry",
    includes: "Detect Occupancy Status (UC-04)",
  },
  {
    id: "UC-04", name: "Detect Occupancy Status",
    primary: "AI Camera",
    secondary: "None",
    description: ["The AI-powered camera continuously monitors parking bays using YOLOv5 to determine which slots are occupied or free and reports status to the backend."],
    trigger: "Continuous, automatic scan cycle (every 10 seconds).",
    preconditions: ["1. AI Camera is powered on and correctly positioned.", "2. YOLOv5 detection model is loaded and operational."],
    postconditions: ["1. Slot occupancy status is updated in the backend, dashboard, and LED indicators.", "2. A \u201cFULL\u201d flag is set or cleared as appropriate."],
    mainSuccess: [
      "1. AI Camera scans each parking bay using YOLOv5, approximately every 10 seconds.",
      "2. Camera sends a status message (slot ID and occupied/free status) to the backend server via MQTT.",
      "3. Backend updates the corresponding LED indicator and the dashboard slot map.",
      "4. If all bays are occupied, backend sets a \u201cFULL\u201d flag; RFID Reader then blocks new entries.",
      "5. When a vehicle exits and a bay becomes free, the FULL flag clears within 10 seconds and normal entry resumes.",
    ],
    alternate: [],
    exception: ["E1 \u2013 Detection failure (poor lighting, camera obstruction): Slot status is marked \u201cunknown\u201d (grey) on the dashboard until detection resumes."],
    businessRules: ["\u2022 Minimum 95% occupancy detection accuracy, based on YOLOv5 benchmark performance.", "\u2022 Detection update interval must not exceed 10 seconds per slot."],
    priority: "High",
    frequency: "Continuous (every 10 seconds during operating hours)",
  },
  {
    id: "UC-05", name: "Exit Parking",
    primary: "Student / Faculty / Staff, Visitor",
    secondary: "Exit Sensor",
    description: ["A driver leaves the parking facility through a free-flow exit that requires no card tap or manual action."],
    trigger: "Vehicle approaches the exit gate.",
    preconditions: ["Vehicle is currently inside the parking facility."],
    postconditions: ["1. Exit is logged with a timestamp.", "2. Parking occupancy count is updated."],
    mainSuccess: [
      "1. Driver approaches the exit gate.",
      "2. Exit Sensor detects the departing vehicle.",
      "3. Exit gate opens automatically within 1 second.",
      "4. System logs the exit with a timestamp and updates the available slot count.",
    ],
    alternate: [],
    exception: ["E1 \u2013 Sensor failure: Security Guard performs manual verification at the exit and records the movement as a fallback."],
    businessRules: ["\u2022 Exit is a free-flow operation; no RFID tap or user action is required."],
    priority: "High",
    frequency: "Every exit",
  },
  {
    id: "UC-06", name: "Request Temporary Card",
    primary: "Visitor",
    secondary: "Security Guard",
    description: ["A visitor without a registered RFID card is issued a temporary card by the security guard to gain one-time parking access."],
    trigger: "A visitor (parent, delivery, guest) arrives at the entrance requesting parking access.",
    preconditions: ["Visitor can provide name, plate number, and purpose of visit."],
    postconditions: ["1. Visitor entry is logged digitally (via included Log Visitor Entry).", "2. Visitor is in possession of a temporary RFID card."],
    mainSuccess: [
      "1. Visitor provides their name, plate number, and purpose of visit to the security guard.",
      "2. Guard logs the visitor's details digitally on a tablet (includes UC-07 \u2013 Log Visitor Entry).",
      "3. Guard hands a temporary RFID card to the visitor.",
      "4. Visitor taps the temporary card at the entrance; the boom barrier opens automatically.",
      "5. On exit, the visitor returns the temporary card and the guard marks it \u201creturned\u201d on the tablet.",
    ],
    alternate: ["A1 \u2013 Registered user with a second vehicle: processed through this same flow instead of standard registration."],
    exception: ["E1 \u2013 Visitor cannot provide required details: Guard denies entry or escalates to BMO Administrator."],
    businessRules: ["\u2022 The temporary card stores no visitor data; it is used only to operate the boom barrier.", "\u2022 Visitor payment is handled directly between the visitor and the guard in cash; the system does not record or process any payment."],
    priority: "High",
    frequency: "As needed (variable, lower volume than registered-user entries)",
    includes: "Log Visitor Entry (UC-07)",
  },
  {
    id: "UC-07", name: "Log Visitor Entry",
    primary: "Security Guard",
    secondary: "None",
    description: ["The security guard records a visitor's details digitally at the point of entry, replacing the manual paper logbook."],
    trigger: "Included automatically whenever Request Temporary Card (UC-06) is performed.",
    preconditions: ["Guard has access to the visitor-logging tablet."],
    postconditions: ["Visitor's name, plate number, purpose, and timestamp are stored digitally."],
    mainSuccess: [
      "1. Guard opens the visitor registration screen on the tablet.",
      "2. Guard enters the visitor's name, plate number, and purpose of visit.",
      "3. Guard taps \u201cGive Temporary Card,\u201d and the system logs the entry digitally.",
      "4. The list of recently issued cards (visitor name, plate number, timestamp) is updated on the guard's screen.",
    ],
    alternate: [],
    exception: ["E1 \u2013 Tablet unavailable: Guard temporarily records details on paper and enters them digitally once the tablet is available."],
    businessRules: ["\u2022 No paper logbook entry is required under normal operation."],
    priority: "High",
    frequency: "Same as visitor requests",
  },
  {
    id: "UC-08", name: "Override Boom Barrier",
    primary: "Security Guard",
    secondary: "None",
    description: ["The security guard manually opens or closes the boom barrier outside of normal RFID validation, for use during emergencies or system malfunctions."],
    trigger: "An emergency, system malfunction, or exceptional situation requires manual gate control.",
    preconditions: ["Guard is on duty and authorized to operate the manual override control."],
    postconditions: ["Boom barrier state is changed manually, independent of RFID validation."],
    mainSuccess: [
      "1. Guard identifies a situation requiring manual control (e.g., system malfunction, emergency vehicle).",
      "2. Guard activates the manual override control at the guard booth.",
      "3. Boom barrier opens or closes immediately, bypassing RFID validation.",
      "4. Guard records the override event and reason for the audit trail.",
    ],
    alternate: [],
    exception: ["E1 \u2013 Override control unresponsive: Guard escalates to BMO Administrator or System Administrator for manual physical intervention."],
    businessRules: ["\u2022 Overrides should be logged with a reason to preserve audit accountability."],
    priority: "Medium",
    frequency: "Rare (emergency or exception situations only)",
  },
  {
    id: "UC-09", name: "Monitor Live Display",
    primary: "Security Guard",
    secondary: "None",
    description: ["The security guard views a live, always-on display screen at the guard booth showing current slot occupancy and recent entries, without needing to log in."],
    trigger: "Guard is on duty at the guard booth.",
    preconditions: ["Guard display screen is powered on and connected to the backend."],
    postconditions: ["Guard is informed of current parking status and recent activity."],
    mainSuccess: [
      "1. Guard views the display screen at the guard booth.",
      "2. Screen automatically shows the live, color-coded slot map (green = free, red = occupied).",
      "3. Screen shows a recent entries log (name, role, plate, timestamp) without requiring login.",
      "4. Guard monitors the screen throughout the shift for anomalies or a full lot.",
    ],
    alternate: [],
    exception: ["E1 \u2013 Display loses connection to backend: Screen shows a \u201cdata unavailable\u201d state; guard falls back to visual inspection until connection is restored."],
    businessRules: ["\u2022 No login or password is required to view the guard display.", "\u2022 The guard display shows only the slot map and recent entries \u2013 no counters or analytics charts."],
    priority: "High",
    frequency: "Continuous, throughout each guard shift",
  },
  {
    id: "UC-10", name: "Monitor Dashboard & Analytics",
    primary: "BMO Administrator",
    secondary: "None",
    description: ["The BMO administrator views a full web dashboard showing real-time occupancy, live counters, recent entries, and parking analytics to support data-driven decision-making."],
    trigger: "BMO Administrator logs into the DEPOTS web dashboard.",
    preconditions: ["Administrator has valid dashboard login credentials."],
    postconditions: ["Administrator has an up-to-date view of parking occupancy, activity, and usage trends."],
    mainSuccess: [
      "1. Administrator logs into the DEPOTS dashboard using their APC ID and role.",
      "2. Dashboard displays a color-coded slot map (green = free, red = occupied, grey = unknown) and live counters (total cars, students, faculty/staff, visitors), refreshing at least every 10 seconds.",
      "3. Administrator reviews the recent entries log (name, role, plate, timestamp, entry/exit status).",
      "4. Administrator views analytics such as peak usage hours and the student/faculty/staff/visitor breakdown.",
      "5. Administrator generates or exports reports as needed for parking planning.",
    ],
    alternate: [],
    exception: [
      "E1 \u2013 Lot reaches capacity: Dashboard displays a \u201cPARKING FULL\u201d alert.",
      "E2 \u2013 Data feed interrupted: Affected metrics are shown as unavailable until the connection to the backend is restored.",
    ],
    businessRules: ["\u2022 Dashboard must refresh within 5\u201310 seconds of a database update.", "\u2022 Dashboard must display at least four key metrics (total slots, available slots, daily entries, average entry time)."],
    priority: "High",
    frequency: "As needed, typically daily",
  },
  {
    id: "UC-11", name: "Manage System Configuration",
    primary: "System Administrator",
    secondary: "None",
    description: ["The system administrator configures, maintains, and troubleshoots DEPOTS, including RFID registrations, user accounts, and database settings."],
    trigger: "System requires configuration changes, maintenance, or troubleshooting.",
    preconditions: ["Administrator has system-level access credentials."],
    postconditions: ["System configuration, user accounts, or database settings are updated and continuous operation is maintained."],
    mainSuccess: [
      "1. System Administrator logs into the system configuration panel.",
      "2. Administrator manages RFID registrations, user accounts, and database settings as needed.",
      "3. Administrator applies system updates or troubleshoots reported issues.",
      "4. Administrator verifies that all system components (RFID reader, AI camera, dashboard, guard display) are operating normally after changes.",
    ],
    alternate: [],
    exception: ["E1 \u2013 Update causes a system fault: Administrator rolls back the change and restores the previous stable configuration."],
    businessRules: ["\u2022 Configuration changes should be logged for accountability and troubleshooting history."],
    priority: "Medium",
    frequency: "As needed (maintenance-driven, infrequent)",
  },
];

const children = [];
children.push(new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { after: 200 },
  children: [new TextRun({ text: "Use Case Full Description / Fully Dressed Use Case", bold: true })],
}));
children.push(new Paragraph({
  spacing: { after: 300 },
  children: [new TextRun({ text: "Digital Entry & Parking Occupancy Tracking System (DEPOTS)", italics: true, color: "555555" })],
}));

useCases.forEach((uc, idx) => {
  children.push(buildUseCaseTable(uc));
  children.push(new Paragraph({ text: "", spacing: { after: 300 } }));
});

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
      },
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync("/home/claude/DEPOTS_Fully_Dressed_Use_Cases.docx", buf);
  console.log("done");
});