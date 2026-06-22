const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
        LevelFormat, PageOrientation } = require('docx');
const fs = require('fs');

const NAVY = "1E3A5F";
const TEAL = "2C8A6E";
const GOLD = "C9941A";
const LIGHT_TEAL = "E1F5EE";
const LIGHT_NAVY = "EBF0F7";
const LIGHT_GOLD = "FEF9EC";
const LIGHT_PURPLE = "EEEDFE";
const LIGHT_CORAL = "FAECE7";
const LIGHT_GREEN = "EAF3DE";
const LIGHT_GRAY = "F1EFE8";
const WHITE = "FFFFFF";
const GRAY = "64748B";

const border = { style: BorderStyle.SINGLE, size: 1, color: "E2E8F0" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorders = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }
};

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 28, bold: true, color: NAVY })]
  });
}

function heading2(text, color = NAVY) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: TEAL, space: 4 } },
    children: [new TextRun({ text, font: "Arial", size: 22, bold: true, color })]
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 180, after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 18, bold: true, color: TEAL })]
  });
}

function body(text) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 20, color: "374151" })]
  });
}

function bulletItem(text, bold_prefix = "") {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [
      ...(bold_prefix ? [new TextRun({ text: bold_prefix + " ", font: "Arial", size: 20, bold: true, color: NAVY })] : []),
      new TextRun({ text, font: "Arial", size: 20, color: "374151" })
    ]
  });
}

function spacer(lines = 1) {
  return new Paragraph({ spacing: { before: 60 * lines, after: 0 }, children: [new TextRun("")] });
}

function labelRow(label, value, fill = LIGHT_NAVY) {
  return new TableRow({
    children: [
      new TableCell({
        width: { size: 2600, type: WidthType.DXA },
        borders,
        shading: { fill, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 80 },
        children: [new Paragraph({ children: [new TextRun({ text: label, font: "Arial", size: 18, bold: true, color: NAVY })] })]
      }),
      new TableCell({
        width: { size: 6760, type: WidthType.DXA },
        borders,
        shading: { fill: WHITE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 80 },
        children: [new Paragraph({ children: [new TextRun({ text: value, font: "Arial", size: 18, color: "374151" })] })]
      })
    ]
  });
}

function techTable(rows, headerColor = LIGHT_NAVY) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2600, 6760],
    rows: rows.map((r, i) => labelRow(r[0], r[1], i === 0 ? headerColor : (i % 2 === 0 ? LIGHT_GRAY : WHITE)))
  });
}

function qaBlock(q, a) {
  return [
    new Paragraph({
      spacing: { before: 120, after: 40 },
      children: [new TextRun({ text: "Q: " + q, font: "Arial", size: 20, bold: true, color: NAVY })]
    }),
    new Paragraph({
      spacing: { before: 0, after: 80 },
      indent: { left: 360 },
      children: [new TextRun({ text: "A: " + a, font: "Arial", size: 20, color: "374151" })]
    })
  ];
}

// ── DOCUMENT ─────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 18, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 180, after: 60 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children: [

      // ── COVER ─────────────────────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 720, after: 120 },
        children: [new TextRun({ text: "Revenue Recovery Engine", font: "Arial", size: 52, bold: true, color: NAVY })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "Complete Technology Stack — Interview Prep Guide", font: "Arial", size: 28, color: TEAL })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 480 },
        children: [new TextRun({ text: "Anjana Ramia  |  Bay Area RevOps  |  2026", font: "Arial", size: 20, color: GRAY, italics: true })]
      }),

      // Intro
      new Paragraph({
        spacing: { before: 0, after: 120 },
        children: [new TextRun({
          text: "This guide explains every technology used to build the Revenue Recovery Engine — what each tool does, why it was chosen, and how to answer interview questions about it. Technologies are grouped by layer so you can explain the full architecture in sequence.",
          font: "Arial", size: 20, color: "374151"
        })]
      }),
      spacer(),

      // ── SECTION 1: FRONTEND ───────────────────────────────────────────────
      heading1("1. Frontend & UI Layer"),
      body("The user-facing application. Realtors interact with this to upload their CRM data and get scored lead lists."),
      spacer(),

      heading3("Streamlit"),
      techTable([
        ["What it is", "A Python framework that turns Python scripts into interactive web apps — no HTML/CSS/JavaScript required."],
        ["Why used", "Zero frontend development needed. Builds data apps 10x faster than traditional web frameworks."],
        ["What you built", "5-tab layout: Upload & Clean, Score & Prioritize, Outreach Generator, Client Dashboard, Monthly Report."],
        ["Key features used", "st.tabs, st.sidebar, st.dataframe, st.file_uploader, st.session_state, st.secrets for secure password storage."],
        ["Hosted on", "Streamlit Community Cloud — free, auto-deploys from GitHub on every push."],
      ], LIGHT_TEAL),

      spacer(),

      heading3("st.session_state"),
      body("Streamlit reruns the entire script on every user interaction. session_state stores variables (uploaded dataframe, scored results, selected client) so data persists across tab switches without re-processing. Without it, every tab click would re-upload and re-score the CSV."),

      spacer(),

      heading3("st.secrets"),
      body("Stores sensitive values (ADMIN_PASSWORD) encrypted in Streamlit Cloud, not in code. Referenced in the app as st.secrets.get('ADMIN_PASSWORD'). This means the password never appears in your GitHub repo — a production security practice."),

      spacer(),
      ...qaBlock(
        "Why did you choose Streamlit over Flask or Django?",
        "Streamlit was the right tool for this stage because the primary users are realtors, not developers. The goal was a CSV-in, results-out experience in under 60 seconds. Flask or Django would have required building a separate frontend in HTML/JS, doubling the development time. Streamlit let me focus entirely on the scoring logic. When enterprise clients need a custom UI, the FastAPI layer is already built for that."
      ),
      ...qaBlock(
        "What are Streamlit's limitations?",
        "Streamlit reruns the entire script on every interaction, which can be slow for heavy computations — mitigated here with session_state caching. The filesystem resets periodically on Streamlit Cloud, which is why we migrated lead capture from CSV to SQLite. It also doesn't support concurrent users well at scale, but for a pilot with 1–10 realtors it's perfectly sufficient."
      ),
      spacer(2),

      // ── SECTION 2: SCORING ENGINE ─────────────────────────────────────────
      heading1("2. Python Scoring Engine"),
      body("The core business logic. Converts raw CRM data into prioritised, actionable lead lists."),
      spacer(),

      heading3("pandas"),
      techTable([
        ["What it is", "The standard Python library for data manipulation — think Excel but in code."],
        ["Why used", "Reading CSVs, cleaning columns, calculating days since last contact, filtering by temperature tier."],
        ["Key operations", "pd.read_csv(), DataFrame filtering, .apply() for row-by-row scoring, .sort_values() for priority ordering, .groupby() for buyer/seller splits."],
        ["Interview note", "pandas is industry-standard in data analytics, data science, and RevOps automation. Every hiring manager in this space recognises it."],
      ], LIGHT_PURPLE),

      spacer(),

      heading3("numpy"),
      techTable([
        ["What it is", "Numerical computing library. Handles mathematical operations efficiently."],
        ["Used for", "np.nan for missing values, np.isnan() for null checks in scoring, math.clamp operations for score bounding."],
      ], LIGHT_PURPLE),

      spacer(),

      heading3("Scoring formula"),
      body("The four-signal model behind every priority score:"),
      spacer(),
      techTable([
        ["Signal", "Weight — What it measures"],
        ["Recency", "50% — Days since last contact. Hot ≤30 days, Warm ≤90, Cold ≤180, Dormant 180+"],
        ["Lead Type", "25% — Past Client (1.0), Referral (0.95), Investor (0.85), Buyer/Seller (0.80), Renter (0.50)"],
        ["Contact Completeness", "25% — Both email + phone (1.0), either (0.6), neither (0.2)"],
        ["Source Multiplier", "Applied last — Closing probability by acquisition channel, capped at 1.0"],
      ], LIGHT_GOLD),

      spacer(),

      heading3("Lead Source Multiplier — 4 Tiers"),
      techTable([
        ["Tier", "Sources — Multiplier — Rationale"],
        ["High Probability", "Referral, Past Client, Open House, Sign Call — 1.25–1.35 — Kaggle-validated: referrals convert at 2x paid social"],
        ["Portal / Active Searcher", "Zillow, Redfin, Realtor.com, Trulia — 1.15–1.20 — Paid to be there AND actively searching"],
        ["Paid / Organic Social", "Instagram, Facebook, Google, Organic — 0.95–1.05 — Middle ground, intent varies"],
        ["Outbound / Low Signal", "Cold Call, Direct Mail, Door Knock — 0.80–0.85 — Spray-and-pray, score conservatively"],
      ], LIGHT_GREEN),

      spacer(),

      heading3("CPL / Sunk Cost Recovery"),
      body("A separate financial metric layer built on top of scoring. When a realtor enters their average Cost Per Lead (CPL) in the sidebar, the engine calculates:"),
      bulletItem("Spend at Risk = dormant leads × CPL"),
      bulletItem("Recoverable Spend = projected reactivations × CPL"),
      bulletItem("Recovery ROI = projected revenue ÷ spend at risk"),
      body("This surfaces a number no other lead scoring tool shows: 'You spent $30,000 generating these dormant leads. Recovering 5% returns $47,500 — a 12.5x ROI.' CPL is kept separate from the source multiplier because CPL measures financial exposure, not closing probability — conflating them would make the scoring wrong."),

      spacer(),
      ...qaBlock(
        "How did you validate your scoring weights?",
        "I ran exploratory data analysis against a Kaggle dataset of 9,240 real leads from a lead scoring study. Key findings that shaped the weights: referral leads converted at roughly 2x the rate of paid social leads — confirming the 1.35 multiplier for referrals. Website-visit leads had a sharp 70% drop-off after 48 hours — informing the aggressive recency decay for that source. The Kaggle data validated the intuition but I kept the domain-specific tiers (Zillow, Redfin etc.) separate because the dataset was from an education platform, not real estate."
      ),
      ...qaBlock(
        "Why is recency weighted at 50%?",
        "Real estate lead conversion is strongly time-sensitive. A lead that contacted you yesterday is fundamentally more likely to close than one from two years ago, regardless of source. The 50% weight reflects this. The dormant bonus (+0.15) partially counteracts pure recency decay for very old leads — because a past client from 3 years ago still deserves a higher priority than a cold call from last month."
      ),
      spacer(2),

      // ── SECTION 3: DATA CLEANING ──────────────────────────────────────────
      heading1("3. Data Cleaning Layer"),
      body("data_cleaner.py — the first thing that runs after a CSV is uploaded. Normalises inconsistent CRM exports into a standard format."),
      spacer(),

      heading3("Key functions"),
      techTable([
        ["standardize_columns()", "Maps 40+ CRM column name variants to internal standard. 'last_activity', 'date_contacted', 'last_contact_date' all map to Last_Contact_Date."],
        ["parse_date_flexible()", "Tries 8 different date formats (YYYY-MM-DD, MM/DD/YYYY, etc.) before giving up. Handles messy real-world CRM exports."],
        ["flag_missing_fields()", "Adds boolean Flag_ columns marking missing email, phone, name, date, lead type. Used in contact completeness scoring."],
        ["flag_duplicates()", "Identifies duplicate leads by (name + email) or (name + phone) combination. Keeps first occurrence."],
        ["compute_quality_score()", "Produces a 0–100 data quality score with weighted breakdown by field completeness. Shown to the user before scoring."],
      ], LIGHT_TEAL),

      spacer(),
      ...qaBlock(
        "Why build a custom data cleaner instead of using an existing library?",
        "The real estate CRM landscape is fragmented — Follow Up Boss, kvCORE, Lofty, and custom spreadsheets all export differently. An off-the-shelf cleaner wouldn't know that 'lead_origin', 'channel', 'acquisition_source', and 'source' all mean the same thing. The COLUMN_MAP dictionary encodes domain-specific knowledge that a generic library can't have. The quality score also gives realtors immediate feedback on their data hygiene — something you wouldn't get from a generic cleaning tool."
      ),
      spacer(2),

      // ── SECTION 4: REST API ───────────────────────────────────────────────
      heading1("4. REST API Layer"),
      body("The enterprise integration layer. Makes the scoring engine callable from Salesforce, Agentforce, or any other system — without sharing Python code."),
      spacer(),

      heading3("FastAPI"),
      techTable([
        ["What it is", "A modern Python framework for building REST APIs. Faster than Flask, auto-generates documentation."],
        ["Why not Flask", "FastAPI generates an OpenAPI 3.0 specification automatically from your Python code. This is the exact format Salesforce External Services requires. Flask needs manual documentation. FastAPI makes the Salesforce integration free."],
        ["Key endpoints", "POST /score-lead (the scoring engine), GET /health (liveness check for Salesforce Named Credentials)"],
        ["Auto-docs", "Swagger UI at /docs — interactive documentation generated from the Pydantic models. Anyone can test the API without writing any code."],
      ], LIGHT_CORAL),

      spacer(),

      heading3("Pydantic v2"),
      techTable([
        ["What it is", "Data validation library. Defines the exact shape of API request and response data using Python classes."],
        ["LeadInput model", "Defines 5 required inputs: lead_source (str), days_idle (int), lead_type (str), has_email (bool), has_phone (bool)"],
        ["LeadScore model", "Defines 4 outputs: score (int 1-10), temperature (str), next_action (str), source_tier (str)"],
        ["Alias workaround", "In the Salesforce integration, lead_source carries the Salesforce Record ID (not the actual source string) using alias='lead_source' in the Pydantic model. This was a practical workaround for Flow field mapping constraints."],
        ["Why it matters", "Pydantic catches type mismatches before they reach your code. A string sent where an int is expected raises a clear 422 error instead of a silent bug."],
      ], LIGHT_CORAL),

      spacer(),

      heading3("slowapi — Rate Limiting"),
      techTable([
        ["What it is", "Rate limiting middleware for FastAPI built on limits library."],
        ["Configuration", "30 requests per minute per IP address on /score-lead. 1000/minute on /health (to prevent Uptime Robot pings from triggering the limiter)."],
        ["Why it matters", "Free Render instances can be overwhelmed by abuse or misconfigured clients. Rate limiting prevents a single caller from crashing the service. Also signals production readiness to any technical interviewer."],
        ["Implementation", "@limiter.limit('30/minute') decorator on the endpoint function. Requires request: Request as a parameter."],
      ], LIGHT_CORAL),

      spacer(),

      heading3("OpenAPI 3.0.3 specification"),
      body("FastAPI auto-generates this from your Pydantic models. The spec describes every endpoint, input parameter, output field, and data type. Saved as api/openapi.json in GitHub. Used in Salesforce External Services registration to teach Salesforce what your API does and how to call it. Forced to version 3.0.3 (not 3.1) because Salesforce's parser has better compatibility with 3.0.3."),

      spacer(),
      ...qaBlock(
        "Explain the difference between a tool and infrastructure in the context of your build.",
        "The Streamlit app is a tool — a person opens it, clicks buttons, gets results. It's passive; it waits for a human. The FastAPI layer is infrastructure — it receives automated requests from Salesforce, processes them, and returns results without any human involvement. Same scoring logic, two delivery mechanisms. The Streamlit app serves solo realtors who upload a CSV. The API serves enterprise CRMs that need automatic, real-time scoring the moment a lead enters the system."
      ),
      ...qaBlock(
        "What is an OpenAPI spec and why does it matter for Salesforce?",
        "An OpenAPI specification is a structured JSON or YAML document that describes exactly what an API does — its endpoints, inputs, outputs, and data types. It's the API's instruction manual. Salesforce's External Services feature reads this spec and automatically makes your API endpoints available as native actions inside Flow Builder — no Apex required. FastAPI generates this spec automatically from your Pydantic models, which means the code IS the documentation. Any change to the scoring model automatically updates the spec."
      ),
      spacer(2),

      // ── SECTION 5: DATABASE ───────────────────────────────────────────────
      heading1("5. Database & Storage"),
      spacer(),

      heading3("SQLite"),
      techTable([
        ["What it is", "A file-based SQL database built into Python's standard library. No server required."],
        ["Why not PostgreSQL", "PostgreSQL requires a separate server. SQLite is a single file — appropriate for pilot scale (dozens of users). The upgrade path to PostgreSQL is straightforward when needed."],
        ["Lead capture table", "Stores email, name, timestamp for every person who signs up through the email gate. UNIQUE constraint on email prevents duplicates."],
        ["Client manager", "Stores realtor clients, processing runs, monthly tracking data. Multi-table schema with foreign keys."],
        ["Key improvement over CSV", "CSV files can corrupt when two users submit simultaneously (race condition). SQLite handles concurrent writes safely with locking."],
        ["INSERT OR IGNORE", "Silently skips duplicate emails instead of erroring. Keeps the email gate working smoothly even if someone submits twice."],
      ], LIGHT_GOLD),

      spacer(),
      ...qaBlock(
        "Why migrate from CSV to SQLite for lead capture?",
        "CSV files have a race condition problem — if two people submit their email simultaneously, both writes can interleave and corrupt the file. SQLite handles concurrent writes with file-level locking. It also persists better on Streamlit Cloud, which occasionally resets the filesystem. SQLite is the right tool for single-server, low-to-medium concurrent use. The UNIQUE constraint on email means duplicate submissions are handled at the database level, not in application code."
      ),
      spacer(2),

      // ── SECTION 6: DEPLOYMENT ─────────────────────────────────────────────
      heading1("6. Deployment & Infrastructure"),
      spacer(),

      heading3("Render (API hosting)"),
      techTable([
        ["What it is", "A cloud platform for deploying web services. Similar to Heroku."],
        ["Free tier limitation", "Instances spin down after 15 minutes of inactivity, causing a ~50 second cold start on the next request. Mitigated by Uptime Robot pings every 5 minutes."],
        ["Configuration", "Root Directory: api/, Build Command: pip install -r requirements.txt, Start Command: uvicorn main:app --host 0.0.0.0 --port $PORT"],
        ["Auto-deploy", "Connected to GitHub main branch. Every push triggers a new deployment."],
        ["Live URL", "https://revenue-recovery-engine-tjut.onrender.com"],
      ], LIGHT_TEAL),

      spacer(),

      heading3("Uptime Robot"),
      techTable([
        ["What it is", "Free monitoring service. Pings your URLs every 5 minutes and sends email alerts if they go down."],
        ["Two monitors", "App monitor: revenue-recovery-engine.streamlit.app. API monitor: /health endpoint on Render."],
        ["Dual purpose", "Alerts you to downtime AND keeps the Render free instance warm (prevents cold starts when Salesforce calls the API)."],
        ["405 on /health", "Uptime Robot sends HEAD requests; the /health endpoint only allows GET. Fixed by adding a high rate limit (1000/min) to the health endpoint so it responds to both methods."],
      ], LIGHT_TEAL),

      spacer(),

      heading3("GitHub"),
      techTable([
        ["Role", "Version control AND CI/CD trigger. Every push to main auto-deploys Streamlit and Render."],
        ["Three repos document the evolution", "Python-Database-Reactivation (first build), Reactivation-Engine (Streamlit), Revenue-Recovery-Engine (current, with API and Salesforce integration)"],
        ["MIT License", "Open source. Timestamps your authorship and signals transparency."],
        ["openapi.json in /api folder", "Critical for Salesforce External Services — the spec must be accessible to show interviewers and clients."],
      ], LIGHT_TEAL),

      spacer(2),

      // ── SECTION 7: SALESFORCE ─────────────────────────────────────────────
      heading1("7. Salesforce Integration"),
      body("The enterprise bridge. Connects the Python scoring engine to Salesforce CRM so leads are scored automatically when they enter the system — zero manual steps."),
      spacer(),

      heading3("Named Credentials"),
      techTable([
        ["What it is", "Salesforce's secure address book for external API endpoints."],
        ["Why it exists", "Prevents hardcoding API URLs and credentials in Flow logic. One central record — update the URL once, every Flow that uses it updates automatically."],
        ["Configuration", "Label: LeadRecoveryAPI, URL: Render endpoint, Enabled for Callouts: ON, Allow Formulas in HTTP Header/Body: checked."],
        ["Linked to", "External Credential (LeadRecoveryExternal) which handles the authentication protocol."],
        ["PD1 connection", "Understanding Named Credentials before writing a line of FastAPI code meant the API was built to the right spec from the start — no retrofitting required."],
      ], LIGHT_GREEN),

      spacer(),

      heading3("External Credentials"),
      techTable([
        ["What it is", "The authentication layer that Named Credentials delegate to. Defines HOW Salesforce authenticates with the external system."],
        ["Configuration", "Authentication Protocol: No Authentication (the API is open for pilot). Principals: LeadRecoveryPrincipal (Named Principal type)."],
        ["Key lesson", "Forgetting to add a Principal caused silent callout failures — Salesforce would attempt the request and receive no response. Empty Principals section = blocked callouts."],
      ], LIGHT_GREEN),

      spacer(),

      heading3("Remote Site Settings"),
      techTable([
        ["What it is", "Salesforce's URL whitelist. Every external URL a Flow or Apex class calls must be explicitly registered here."],
        ["Why it matters", "Without registering the Render URL, Salesforce silently blocks all outbound callouts. No error appears in the Flow — it just doesn't call the API. This was the root cause of the 'no POST requests in Render logs' problem."],
        ["Configuration", "Remote Site Name: LeadRecoveryAPI, URL: https://revenue-recovery-engine-tjut.onrender.com, Active: checked."],
      ], LIGHT_GREEN),

      spacer(),

      heading3("Record-Triggered Flow (LeadRecoveryUpdateNew)"),
      techTable([
        ["What it is", "Salesforce's low-code automation tool. No Apex required."],
        ["Trigger", "Lead object, A record is created or updated."],
        ["Entry condition", "Recovery_Score__c Is Null — only fires when a Lead has no score yet, preventing infinite loops."],
        ["When to run (updated records)", "Only when a record is updated to meet the condition requirements — required by Salesforce for async paths."],
        ["Async path", "External callouts (HTTP requests to outside systems) MUST run on the asynchronous path. Synchronous path = blocked by Salesforce."],
        ["Canvas elements", "Start → Call Scoring API (HTTP Callout Action) → Parse Score from Response (Assignment) → Write Score to Lead (Update Records)"],
      ], LIGHT_GREEN),

      spacer(),

      heading3("HTTP Callout Action"),
      techTable([
        ["What it is", "A built-in Flow action that makes an HTTP request to an external URL. Alternative to External Services that bypasses schema parsing."],
        ["Why used instead of External Services", "External Services requires a precisely formatted OpenAPI spec — Salesforce's parser rejected the Pydantic v2 schema with anyOf/nullable fields. HTTP Callout lets you define the request and response structure directly by pasting sample JSON."],
        ["Request body", "Apex-Defined Variable (RequestBody) containing the 5 fields mapped from Lead record fields."],
        ["Response parsing", "2XX response stored via 'Manually assign variables'. Score extracted using {!Call_Scoring_API.2XX.score} output variable."],
        ["Key insight", "HTTP Callout was introduced in Salesforce Flow Builder as a simpler alternative to External Services for exactly this kind of REST API integration."],
      ], LIGHT_GREEN),

      spacer(),

      heading3("Custom Fields on Lead Object"),
      techTable([
        ["Recovery_Score__c", "Number (2 digits, 0 decimals). Stores the 1–10 priority score written back by the Flow. Displayed on the Lead record page layout."],
        ["Has_Email__c", "Custom Boolean. Pre-validated email flag passed to the API as has_email input."],
        ["Has_Phone__c", "Custom Boolean. Pre-validated phone flag passed to the API as has_phone input."],
        ["Lead_Recovery_Score__c", "Duplicate field created during testing — exists but is not used. Can be deleted from Object Manager."],
        ["lead_id__c", "Formula field = Id. Captures the Salesforce Record ID as text for the API's lead_source alias workaround."],
      ], LIGHT_GREEN),

      spacer(),
      ...qaBlock(
        "Explain the architecture: how does a lead get scored in Salesforce?",
        "When a Lead record is created or updated in Salesforce with no existing Recovery Score, the Record-Triggered Flow fires on its asynchronous path. The Flow assembles a JSON request body from the Lead's fields — lead source mapped via the lead_id alias, days idle calculated from LastModifiedDate, email and phone from custom boolean fields. It sends a POST request to the FastAPI endpoint on Render via Named Credentials. The Python engine scores the lead using the four-signal model and returns a JSON response with score, temperature, next_action, and source_tier. The Flow extracts the score integer from the 2XX response and writes it to Recovery_Score__c on the Lead record. The whole process is asynchronous — the user doesn't wait for it. The score appears on the Lead record within 30–90 seconds of saving."
      ),
      ...qaBlock(
        "Why did you use HTTP Callout instead of External Services?",
        "External Services requires a valid OpenAPI 3.0 specification, and Salesforce's parser has compatibility issues with Pydantic v2's nullable field syntax — it generates 'anyOf' structures that the parser rejects. HTTP Callout bypasses this entirely: you paste sample JSON for the request and response, Salesforce infers the data structure, and you get the same result without schema registration headaches. For a production implementation with multiple endpoints, External Services would be the right choice. For a pilot proving the concept, HTTP Callout was the pragmatic path."
      ),
      ...qaBlock(
        "What is an asynchronous path in Salesforce Flow and why is it required for external callouts?",
        "Salesforce runs record-triggered flows synchronously by default — the flow executes within the same database transaction as the record save. Making an HTTP request to an external system during a database transaction violates Salesforce's governor limits and causes a timeout. The asynchronous path solves this: it decouples the external callout from the record save transaction. The record saves immediately. The async path runs as a separate queued job seconds later. This is why the score doesn't appear on the Lead record instantly — it appears 30–90 seconds after save, which is the async job processing time."
      ),
      spacer(2),

      // ── SECTION 8: AI TOOLS ───────────────────────────────────────────────
      heading1("8. AI Tools Used in the Build"),
      body("A deliberate multi-model workflow — different tools for different problem types. This is itself a differentiator worth explaining in interviews."),
      spacer(),

      techTable([
        ["Tool", "Role in the build"],
        ["Claude (Anthropic)", "Python architecture, Pydantic model design, FastAPI structure, scoring formula logic, error handling, database design, LinkedIn content strategy. Strongest at reasoning through business logic and Python implementation."],
        ["Gemini (Google)", "Salesforce Flow architecture, External Services schema mapping, Apex-Defined Variable configuration, Flow troubleshooting. Strongest at Salesforce-specific implementation details."],
        ["Antigravity", "Rapid UI prototyping — transformed architecture specs into working Streamlit code quickly. Best for fast iteration on frontend."],
        ["Perplexity", "Research validation — checking CAC statistics, real estate CPL benchmarks, Zillow data science claims before putting them in posts."],
        ["ChatGPT", "Supplementary queries when stuck. Quick syntax lookups."],
        ["NotebookLM", "Document synthesis — understanding Salesforce documentation, Flow Builder guides."],
        ["OpenAI API (optional)", "Powers the AI-generated outreach in the app when a user provides an API key. Falls back to built-in templates if no key provided."],
      ], LIGHT_PURPLE),

      spacer(),
      ...qaBlock(
        "How did you use AI tools in this build and what did you learn about using them effectively?",
        "I used a deliberate multi-model approach: Claude for Python and architecture, Gemini for Salesforce and Flow mechanics. The key insight was that different models have different strengths based on their training data. Gemini had better exposure to Salesforce Flow Builder specifics — probably from Salesforce documentation and developer forums. Claude was stronger at reasoning through Python architecture and explaining business trade-offs. Using them as specialists rather than generalists produced better results than relying on one tool for everything. The most important skill isn't knowing which tool to use — it's knowing how to frame the question precisely enough that the tool gives you something useful."
      ),
      spacer(2),

      // ── SECTION 9: SECURITY ───────────────────────────────────────────────
      heading1("9. Security & Reliability"),
      spacer(),

      techTable([
        ["Practice", "Implementation"],
        ["Secrets management", "ADMIN_PASSWORD stored in Streamlit secrets, not in code. Falls back to environment variable for local dev. Never committed to GitHub."],
        ["Error handling", "try/except blocks on all critical paths: CSV upload, scoring, outreach generation, PDF export, admin panel. Errors logged to app_errors.log, not surfaced as raw tracebacks to users."],
        ["logging module", "Python stdlib. Logs ERROR level and above to app_errors.log with timestamp, level, and message. Allows debugging production issues without seeing them in the UI."],
        ["Rate limiting", "slowapi on the FastAPI endpoint — 30 requests/minute per IP. Prevents abuse of the free Render instance."],
        ["Email gate", "Every app visitor must provide an email to access the scoring features. Captured in SQLite with UNIQUE constraint. Provides both access control and lead capture."],
        ["MIT License", "Open source license on GitHub. Timestamps authorship. Required for portfolio credibility."],
      ], LIGHT_GRAY),

      spacer(2),

      // ── SECTION 10: CERTIFICATIONS ────────────────────────────────────────
      heading1("10. Certifications Informing the Build"),
      body("These credentials are not just resume lines — each one directly shaped architectural decisions in the Revenue Recovery Engine."),
      spacer(),

      techTable([
        ["Certification", "How it shaped the build"],
        ["Salesforce PD1", "Knew to build the API before building the Flow. Understood Named Credentials, External Services, and async path requirements before writing a line of code. This is why the OpenAPI spec was built into the API design from day one."],
        ["SFMC Email Specialist", "Shaped the outreach generator — understanding email deliverability, personalisation tokens, and campaign structure. The email/voicemail/SMS triple output mirrors SFMC's multi-channel campaign thinking."],
        ["Google PM Cert", "Drove the product thinking: user story first (realtor uploads CSV, gets scored leads in 60 seconds), feature prioritisation (build what matters for the pilot, defer everything else), and the roadmap structure."],
        ["Digital Marketing Bootcamp", "Understanding CPL (cost per lead) and CAC (customer acquisition cost) led directly to the Sunk Cost Recovery dashboard — a feature no competing tool has because most builders don't think in marketing finance terms."],
        ["MBA", "The financial framing: 'You spent $30,000 on dormant leads. Recovering 5% returns $47,500 — a 12.5x ROI.' This is CFO language applied to a realtor tool. The business case is as important as the technology."],
      ], LIGHT_GOLD),

      spacer(2),

      // ── SECTION 11: QUICK INTERVIEW Q&A ──────────────────────────────────
      heading1("11. Quick-Fire Interview Q&A"),
      body("These are the most likely questions from RevOps Architect, Salesforce BA, and AI Solutions Engineer interviews."),
      spacer(),

      heading3("Architecture & Design"),
      ...qaBlock(
        "Walk me through your end-to-end architecture.",
        "A realtor uploads their CRM CSV to the Streamlit app. data_cleaner.py normalises 40+ column name variants and scores data quality 0–100. reactivation_engine.py scores every lead 1–10 using four signals: recency (50%), lead type (25%), contact completeness (25%), and a lead source multiplier based on closing probability. The results display in a prioritised table with reactivation emails generated automatically. Separately, a FastAPI endpoint exposes the same scoring logic as a REST API deployed on Render. In Salesforce, a Record-Triggered Flow calls this API via Named Credentials whenever a Lead is created or updated with no existing score. The score writes back to Recovery_Score__c automatically."
      ),
      ...qaBlock(
        "What is RevOps and how does this project demonstrate RevOps thinking?",
        "RevOps — Revenue Operations — is the practice of aligning sales, marketing, and customer success around shared data, processes, and technology to improve revenue outcomes. This project is RevOps thinking in action: I identified a revenue leak (dormant leads representing sunk acquisition cost), built a system to quantify it (the CPL dashboard showing spend at risk and recovery ROI), and created an automated process to address it (scored lead list with ready-to-send outreach). The Salesforce integration is specifically a RevOps pattern — connecting the scoring engine to the CRM where salespeople actually work, so the insight is actionable without any extra tools."
      ),
      ...qaBlock(
        "What would you build next?",
        "Three things in order. First: fix the lead_source mapping in the Salesforce Flow — currently it passes the Record ID as a workaround; the actual source string (Zillow, Referral etc.) should go through the source multiplier for fully accurate scoring. Second: build an outcome logging UI so realtors can log when a reactivated lead closes — this starts accumulating training data for the self-improving scoring model. Third: connect to Agentforce — define an Agent Action that calls the scoring API, identifies cold leads autonomously, and drafts the reactivation email without a human in the loop. That's the enterprise version of what the Streamlit outreach generator does manually today."
      ),
      ...qaBlock(
        "What's the difference between a Salesforce Admin and what you can do?",
        "A Salesforce Admin configures what exists in the platform. What I built does something different: it extends Salesforce to call external Python logic and write the results back — without Apex, without a developer. The Named Credentials + HTTP Callout + Flow combination means any Admin can trigger a custom AI scoring model from within Salesforce, using clicks not code. That's a Solutions Architect pattern — designing integrations that bridge what Salesforce does natively with what Python can do that Salesforce can't."
      ),

      spacer(2),

      // ── FINAL PAGE: CHEAT SHEET ───────────────────────────────────────────
      heading1("12. One-Line Tech Definitions (Cheat Sheet)"),
      body("Memorise these for quick-fire questions."),
      spacer(),

      techTable([
        ["Term", "One-line definition"],
        ["Streamlit", "Python framework that turns scripts into web apps — no HTML needed"],
        ["pandas", "Python library for data manipulation — reading, cleaning, filtering tabular data"],
        ["FastAPI", "Python framework for building REST APIs — auto-generates OpenAPI documentation"],
        ["Pydantic", "Data validation library — enforces types on API inputs and outputs"],
        ["REST API", "A URL that accepts data and returns a result — callable from any system"],
        ["OpenAPI spec", "A JSON document describing an API's endpoints, inputs, and outputs"],
        ["SQLite", "A file-based database — no server required, handles concurrent writes safely"],
        ["slowapi", "Rate limiting middleware — limits how many requests one IP can make per minute"],
        ["Render", "Cloud platform for deploying web services — like Heroku, free tier available"],
        ["Named Credentials", "Salesforce's secure address book for external API endpoints"],
        ["External Credentials", "Authentication layer behind Named Credentials — defines how SF connects to external systems"],
        ["Remote Site Settings", "Salesforce URL whitelist — external URLs must be registered before callouts work"],
        ["Record-Triggered Flow", "Salesforce automation that runs when a record is created, updated, or deleted"],
        ["Async path", "Flow lane for external callouts — decoupled from the record save transaction"],
        ["HTTP Callout Action", "Built-in Flow action for calling external REST APIs without Apex"],
        ["Apex-Defined Variable", "Structured object variable in Flow — maps to a custom class defining the data shape"],
        ["Recovery_Score__c", "Custom Salesforce field on Lead object — stores the Python-generated priority score"],
        ["CPL", "Cost Per Lead — average acquisition cost for one lead from paid channels"],
        ["CAC", "Customer Acquisition Cost — total cost to acquire one paying customer"],
        ["Sunk cost recovery", "Financial metric showing ROI of reactivating leads already paid for"],
        ["Uptime Robot", "Free monitoring service — pings URLs every 5 minutes, emails on downtime"],
        ["MIT License", "Open source license — allows anyone to use, modify, distribute the code"],
        ["session_state", "Streamlit mechanism for persisting data across script reruns"],
        ["st.secrets", "Streamlit's encrypted storage for sensitive values — keeps passwords out of code"],
        ["Lead source multiplier", "Score modifier based on acquisition channel closing probability"],
        ["Dormant bonus", "+0.15 score boost for leads idle 180+ days — primary reactivation targets"],
        ["Source tier", "Label grouping lead sources: High Probability / Portal / Social / Outbound"],
      ], LIGHT_NAVY),

      spacer(2),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 0 },
        children: [new TextRun({ text: "Revenue Recovery Engine  —  github.com/Anjanaramia/Revenue-Recovery-Engine", font: "Arial", size: 18, color: GRAY, italics: true })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 0 },
        children: [new TextRun({ text: "revenue-recovery-engine.streamlit.app  |  revenue-recovery-engine-tjut.onrender.com/docs", font: "Arial", size: 18, color: GRAY, italics: true })]
      }),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/mnt/user-data/outputs/RRE_Tech_Stack_Interview_Prep.docx', buffer);
  console.log('Done');
}).catch(e => { console.error(e); process.exit(1); });
