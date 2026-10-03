/**
 * =========================================================================
 * WORKFLOW & RECRUITMENT AUTOMATION ENGINE (GOOGLE APPS SCRIPT)
 * =========================================================================
 * An event-driven Google Apps Script engine for multi-division candidate
 * intake, dynamic role resolution, and lifecycle communications.
 *
 * Supported Lifecycle Stages:
 *  1. Interview Invite       (Triggered by 'Interview' column = 'Yes')
 *  2. Application Outcome    (Triggered by 'Outcome'   column = 'Offer' | 'Reject')
 *  3. Division Onboarding    (Triggered by 'Outcome'   column = 'Onboarding')
 *
 * @author Yong Xi Leow
 * @license MIT
 * =========================================================================
 */

// =========================================================================
// GLOBAL CONFIGURATION & MAPPINGS
// =========================================================================
const PIPELINE_CONFIG = {
  // Brand & Identity
  ORG_NAME: "Acme Global",
  PORTAL_URL: "https://example.com",
  CALENDLY_URL: "https://calendly.com/your-team/intro-call",
  CONTACT_NAME: "Recruitment Coordinator",
  CONTACT_PHONE: "+61 400 000 000",

  // Email Senders
  SENDERS: {
    CAREERS: {
      EMAIL: "careers@example.com",
      NAME: "Acme Careers Desk"
    },
    OPERATIONS: {
      EMAIL: "contact@example.com",
      NAME: "Acme Operations"
    }
  },

  // Dynamic Column Header Matchers
  HEADERS: {
    INTERVIEW: "Interview",
    OUTCOME: "Outcome",
    EMAIL: "Email",
    FULL_NAME: "Full Name"
  },

  // Dynamic Role Resolution across Sheet Tabs
  ROLE_MAPPING: {
    "Growth & Strategy": "Growth & Strategy Specialist",
    "M&A": "M&A Analyst",
    "Mergers & Acquisitions": "M&A Analyst",
    "Equity Research": "Equity Research Analyst",
    "ER": "Equity Research Analyst",
    "Equity Capital Market": "Equity Capital Markets (ECM) Analyst",
    "ECM": "Equity Capital Markets (ECM) Analyst",
    "Debt Capital Market": "Debt Capital Markets (DCM) Analyst",
    "DCM": "Debt Capital Markets (DCM) Analyst",
    "Markets": "Markets & Research Analyst",
    "Newsletter": "Markets & Research Analyst"
  },

  // Division-Specific Shared Drive Resource Hubs
  DIVISION_DRIVE_HUBS: {
    "ECM": "https://drive.google.com/drive/folders/DEMO_FOLDER_ECM",
    "Equity Capital Market": "https://drive.google.com/drive/folders/DEMO_FOLDER_ECM",
    "DCM": "https://drive.google.com/drive/folders/DEMO_FOLDER_DCM",
    "Debt Capital Market": "https://drive.google.com/drive/folders/DEMO_FOLDER_DCM",
    "ER": "https://drive.google.com/drive/folders/DEMO_FOLDER_ER",
    "Equity Research": "https://drive.google.com/drive/folders/DEMO_FOLDER_ER",
    "M&A": "https://drive.google.com/drive/folders/DEMO_FOLDER_MA",
    "Mergers & Acquisitions": "https://drive.google.com/drive/folders/DEMO_FOLDER_MA",
    "Growth & Strategy": "https://drive.google.com/drive/folders/DEMO_FOLDER_GROWTH"
  }
};

// =========================================================================
// MASTER ON-EDIT TRIGGER
// =========================================================================
/**
 * Master Installable Trigger: Evaluates cell updates across active sheets
 * and routes execution to the appropriate stage handler.
 *
 * @param {Object} e - Event object passed by Google Sheets onEdit event
 */
function handleHiringPipelineEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.source.getActiveSheet();
  const tabName = sheet.getName().trim();
  const roleTitle = PIPELINE_CONFIG.ROLE_MAPPING[tabName];

  // Guard: Ignore non-pipeline tabs
  if (!roleTitle) return;

  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  // Guard: Ignore header modifications
  if (row <= 1) return;

  // Resolve target column indices dynamically from row 1
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  const interviewColIdx = headers.indexOf(PIPELINE_CONFIG.HEADERS.INTERVIEW) + 1;
  const outcomeColIdx   = headers.indexOf(PIPELINE_CONFIG.HEADERS.OUTCOME) + 1;
  const emailColIdx     = headers.indexOf(PIPELINE_CONFIG.HEADERS.EMAIL) + 1;
  const nameColIdx      = headers.indexOf(PIPELINE_CONFIG.HEADERS.FULL_NAME) + 1;

  // Guard: Exit if edit occurs outside monitored decision columns
  if (col !== interviewColIdx && col !== outcomeColIdx) return;

  const cellValue = String(range.getValue()).trim().toLowerCase();
  const recipientEmail = emailColIdx ? sheet.getRange(row, emailColIdx).getValue() : null;
  const candidateName = nameColIdx ? sheet.getRange(row, nameColIdx).getValue() : "Candidate";

  if (!recipientEmail || !cellValue) return;

  // -----------------------------------------------------------------------
  // ROUTE 1: STAGE 1 INTERVIEW INVITATION (Interview -> "Yes")
  // -----------------------------------------------------------------------
  if (col === interviewColIdx && cellValue === "yes") {
    dispatchInterviewInvite(recipientEmail, candidateName, roleTitle);
  }

  // -----------------------------------------------------------------------
  // ROUTE 2: STAGE 2 FORMAL OFFER (Outcome -> "Offer")
  // -----------------------------------------------------------------------
  else if (col === outcomeColIdx && cellValue === "offer") {
    dispatchFormalOffer(recipientEmail, candidateName, roleTitle);
  }

  // -----------------------------------------------------------------------
  // ROUTE 3: STAGE 2 APPLICATION REJECTION (Outcome -> "Reject")
  // -----------------------------------------------------------------------
  else if (col === outcomeColIdx && cellValue === "reject") {
    dispatchRejectionNotice(recipientEmail, candidateName, roleTitle);
  }

  // -----------------------------------------------------------------------
  // ROUTE 4: STAGE 3 ONBOARDING DISPATCH (Outcome -> "Onboarding")
  // -----------------------------------------------------------------------
  else if (col === outcomeColIdx && cellValue === "onboarding") {
    const driveUrl = PIPELINE_CONFIG.DIVISION_DRIVE_HUBS[tabName] || PIPELINE_CONFIG.PORTAL_URL;
    dispatchOnboardingPack(recipientEmail, candidateName, roleTitle, tabName, driveUrl);
  }
}

// =========================================================================
// EMAIL DISPATCHERS & TEMPLATE GENERATORS
// =========================================================================

/**
 * Stage 1: Sends candidate interview booking dispatch via Calendly.
 */
function dispatchInterviewInvite(recipientEmail, candidateName, roleTitle) {
  const subject = `Interview Invitation: ${roleTitle} @ ${PIPELINE_CONFIG.ORG_NAME}`;
  const htmlBody = `
    <div dir="ltr" style="font-family: Arial, sans-serif; color: #222; font-size: 13.5px; line-height: 1.55;">
      <p>Dear ${candidateName},</p>
      <p>Thank you for submitting your application for the <b>${roleTitle}</b> position at ${PIPELINE_CONFIG.ORG_NAME}.</p>
      <p>Your application stood out for its clarity and quantitative depth. As we finalize candidate screening for our upcoming cohort, we would like to progress your application to an introductory discussion.</p>
      <br>
      <p><b>Next Steps & Scheduling:</b></p>
      <p>&rarr; <a href="${PIPELINE_CONFIG.CALENDLY_URL}" target="_blank" style="font-weight: bold; color: #0b5394;">Schedule Your Introductory Discussion Here</a></p>
      <p><i><font color="#666666">Note: If you need to adjust or reschedule your session, you can do so directly through the booking link above.</font></i></p>
      <br>
      <p><b>What to Expect:</b></p>
      <p>This discussion will cover your analytical experience, problem-solving methodologies, and strategic alignment with our current projects. It is also an opportunity to ask questions regarding team workflows and strategic roadmaps.</p>
      <p>To learn more about our ongoing initiatives, visit our portal at <a href="${PIPELINE_CONFIG.PORTAL_URL}" target="_blank">${PIPELINE_CONFIG.PORTAL_URL}</a>.</p>
      ${renderCorporateSignature()}
    </div>
  `;

  GmailApp.sendEmail(recipientEmail, subject, "", {
    from: PIPELINE_CONFIG.SENDERS.CAREERS.EMAIL,
    replyTo: PIPELINE_CONFIG.SENDERS.CAREERS.EMAIL,
    name: PIPELINE_CONFIG.SENDERS.CAREERS.NAME,
    htmlBody: htmlBody
  });
}

/**
 * Stage 2A: Sends formal offer letter and acceptance guidelines.
 */
function dispatchFormalOffer(recipientEmail, candidateName, roleTitle) {
  const subject = `Application Outcome: Formal Offer — ${roleTitle} @ ${PIPELINE_CONFIG.ORG_NAME}`;
  const htmlBody = `
    <div dir="ltr" style="font-family: Arial, sans-serif; color: #222; font-size: 13.5px; line-height: 1.55;">
      <p>Dear ${candidateName},</p>
      <p>Following our conversations, we are delighted to formally offer you the position of <b>${roleTitle}</b> at ${PIPELINE_CONFIG.ORG_NAME}!</p>
      <p>Out of a competitive applicant pool, your technical profile and strategic mindset stood out. You will work closely alongside leadership to execute core quantitative analysis, streamline operational workflows, and scale our deliverables.</p>
      <hr style="border: 0; border-top: 1px solid #ddd; margin: 18px 0;">
      <p><b><u>Offer Acceptance Procedure</u></b></p>
      <p>To confirm your acceptance, <font color="#cc0000"><b>please reply directly to this email within 48 hours</b></font> with:</p>
      <ul>
        <li>Your primary email address (for official repository and workspace access)</li>
        <li>Your mobile contact number (with country code)</li>
      </ul>
      <p><i><font color="#555555">Please note: To ensure cohort onboarding milestones remain on schedule, unconfirmed offers expire after this 48-hour window.</font></i></p>
      <hr style="border: 0; border-top: 1px solid #ddd; margin: 18px 0;">
      <p><b>Next Steps:</b></p>
      <p>Upon receipt of your confirmation, our team will dispatch your division onboarding documentation, repository access credentials, and kickoff call schedules.</p>
      <p>Welcome aboard!</p>
      ${renderCorporateSignature()}
    </div>
  `;

  GmailApp.sendEmail(recipientEmail, subject, "", {
    from: PIPELINE_CONFIG.SENDERS.OPERATIONS.EMAIL,
    replyTo: PIPELINE_CONFIG.SENDERS.OPERATIONS.EMAIL,
    name: PIPELINE_CONFIG.SENDERS.OPERATIONS.NAME,
    htmlBody: htmlBody
  });
}

/**
 * Stage 2B: Sends professional decline notice.
 */
function dispatchRejectionNotice(recipientEmail, candidateName, roleTitle) {
  const subject = `Update regarding your application with ${PIPELINE_CONFIG.ORG_NAME}`;
  const bodyText = `Dear ${candidateName},\n\n`
    + `Thank you for taking the time to apply for the ${roleTitle} position at ${PIPELINE_CONFIG.ORG_NAME} and for sharing your background with us.\n\n`
    + `Following review across a high volume of competitive profiles, we regret to inform you that we are unable to advance your application for this cohort.\n\n`
    + `This decision was driven by cohort capacity constraints and does not diminish the merit of your qualifications. We appreciate your interest and wish you every success in your academic and professional endeavors.\n\n`
    + `Best regards,\n`
    + `${PIPELINE_CONFIG.SENDERS.CAREERS.NAME}`;

  GmailApp.sendEmail(recipientEmail, subject, bodyText, {
    from: PIPELINE_CONFIG.SENDERS.CAREERS.EMAIL,
    replyTo: PIPELINE_CONFIG.SENDERS.CAREERS.EMAIL,
    name: PIPELINE_CONFIG.SENDERS.CAREERS.NAME
  });
}

/**
 * Stage 3: Dispatches tailored onboarding manuals & cloud repository links.
 */
function dispatchOnboardingPack(recipientEmail, candidateName, roleTitle, tabName, driveUrl) {
  const subject = `Welcome to ${PIPELINE_CONFIG.ORG_NAME} — ${roleTitle} Onboarding Details`;
  const htmlBody = `
    <div dir="ltr" style="font-family: Arial, sans-serif; color: #222; font-size: 13.5px; line-height: 1.55;">
      <p>Dear ${candidateName},</p>
      <p><b>Welcome to ${PIPELINE_CONFIG.ORG_NAME}!</b> We have received your acceptance confirmation and are setting up your workspace credentials.</p>
      <hr style="border: 0; border-top: 1px solid #ddd; margin: 18px 0;">
      <p><b>1. Division Cloud Workspace Access</b></p>
      <p>All handbooks, financial modeling guidelines, and project trackers are located in your division repository:<br>
      <a href="${driveUrl}" target="_blank" style="color: #0b5394; font-weight: bold; text-decoration: underline;">
        Access ${roleTitle} Resource Drive
      </a></p>
      <p><i>Action: Star the shared folder in Google Drive for quick daily navigation.</i></p>
      <br>
      <p><b>2. Pre-Kickoff Preparations</b></p>
      <p>Prior to our team onboarding call, please review the <b>Division Workflow Guide</b> and <b>Operational Handbooks</b> contained within the drive root directory.</p>
      <br>
      <p><b>3. Confidentiality Notice</b></p>
      <p><i><font color="#555555">All data pipelines, modeling workpapers, handbooks, and internal materials shared via Google Drive or official comms remain the proprietary intellectual property of ${PIPELINE_CONFIG.ORG_NAME} and must not be distributed externally.</font></i></p>
      <br>
      <p>For urgent scheduling inquiries, contact ${PIPELINE_CONFIG.CONTACT_NAME} (${PIPELINE_CONFIG.CONTACT_PHONE}).</p>
      ${renderCorporateSignature()}
    </div>
  `;

  GmailApp.sendEmail(recipientEmail, subject, "", {
    from: PIPELINE_CONFIG.SENDERS.OPERATIONS.EMAIL,
    replyTo: PIPELINE_CONFIG.SENDERS.OPERATIONS.EMAIL,
    name: PIPELINE_CONFIG.SENDERS.OPERATIONS.NAME,
    htmlBody: htmlBody
  });
}

/**
 * Helper: Renders standardized corporate HTML disclaimer & signature.
 */
function renderCorporateSignature() {
  return `
    <br>
    <p style="font-family: Georgia, serif; font-weight: bold; color: #222; margin-top: 16px;">
      Excellence in Execution.
    </p>
    <p style="font-family: Tahoma, sans-serif; margin-bottom: 8px;">Best regards,</p>
    <table cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-top: 8px;">
      <tr>
        <td style="padding-top: 12px; padding-bottom: 8px;">
          <div style="width: 500px; border-bottom: 1px solid #ccc;"></div>
        </td>
      </tr>
      <tr>
        <td style="max-width: 500px;">
          <p style="color: #888; font-size: 11.5px; line-height: 1.4; font-family: Georgia, serif; margin: 0;">
            <b>Confidentiality Disclaimer:</b> The contents of this transmission are strictly confidential and intended solely for the designated recipient. If you have received this transmission in error, please notify the sender immediately and purge all copies.
          </p>
        </td>
      </tr>
    </table>
  `;
}
