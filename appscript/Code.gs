/**
 * Satguru SA Outreach Dashboard - Google Apps Script
 *
 * Transforms conversation_dump_latest (step-level) into two clean sheets:
 *   1. "Lead Summary" - one row per lead with aggregated metrics + derived fields
 *   2. "Sequence Summary" - one row per sequence with totals
 *
 * HOW TO USE:
 *   1. Open your Google Sheet containing conversation_dump_latest
 *   2. Go to Extensions > Apps Script
 *   3. Paste this entire script
 *   4. Click Run > transformForDashboard
 *   5. Grant permissions when prompted
 *   6. Two new sheets will be created: "Lead Summary" and "Sequence Summary"
 *   7. Connect Looker Studio to these sheets
 *
 * To auto-refresh: set a time-based trigger on transformForDashboard (e.g., every hour)
 */

// ============================================================
// CONFIGURATION - update these if your sheet name differs
// ============================================================
const SOURCE_SHEET = "conversation_dump_latest";
const LEAD_SHEET = "Lead Summary";
const SEQUENCE_SHEET = "Sequence Summary";
const CAMPAIGN_SHEET = "Campaign Summary";
const CONFIG_SHEET = "Sequence Config";  // optional mapping override
const STEP_DETAIL_SHEET = "Step Detail";
const STEP_PERFORMANCE_SHEET = "Step Performance";

// ============================================================
// MAIN FUNCTION
// ============================================================
function transformForDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(SOURCE_SHEET);

  if (!source) {
    Logger.log("Sheet '" + SOURCE_SHEET + "' not found. Check the name and try again.");
    try { SpreadsheetApp.getUi().alert("Sheet '" + SOURCE_SHEET + "' not found. Check the name and try again."); } catch(e) {}
    return;
  }

  const rawData = source.getDataRange().getValues();
  const headers = rawData[0];
  const data = rawData.slice(1);

  // Build column index map
  const col = {};
  headers.forEach((h, i) => { if (h) col[h.toString().trim()] = i; });

  // ============================================================
  // LOAD OPTIONAL SEQUENCE CONFIG (manual overrides)
  // ============================================================
  const seqConfig = loadSequenceConfig(ss);

  // ============================================================
  // PRE-SCAN: detect channel strategy per sequence from actual step_channel data
  // ============================================================
  const seqChannels = {};  // sequence_name -> Set of step_channel values
  data.forEach(row => {
    const seq = (row[col["sequence_name"]] || "").toString();
    const ch = (row[col["step_channel"]] || "").toString();
    if (seq && ch) {
      if (!seqChannels[seq]) seqChannels[seq] = new Set();
      seqChannels[seq].add(ch);
    }
  });

  // ============================================================
  // 1. AGGREGATE TO LEAD LEVEL
  // ============================================================
  const leads = {};
  const stepDetails = [];  // one entry per step per lead

  data.forEach(row => {
    const convId = row[col["conversation_id"]];
    if (!convId) return;

    if (!leads[convId]) {
      const seqName = (row[col["sequence_name"]] || "").toString();
      const channels = seqChannels[seqName] || new Set();

      // Use config override if available, otherwise derive from data
      const config = seqConfig[seqName] || {};

      leads[convId] = {
        conversation_id: convId,
        person_name: row[col["person_name"]] || "",
        person_email: row[col["person_email"]] || "",
        person_title: row[col["person_title"]] || "",
        person_city: row[col["person_city"]] || "",
        person_country: row[col["person_country"]] || "",
        company_name: row[col["company_name"]] || "",
        company_website: row[col["company_website"]] || "",
        company_city: row[col["company_city"]] || "",
        company_employee_range: row[col["company_employee_range"]] || "",
        primary_industry: row[col["primary_industry"]] || "",
        conversation_stage: row[col["conversation_stage"]] || "",
        automation_status: row[col["automation_status"]] || "",
        temperature: row[col["temperature"]] || "",
        campaign_name_raw: row[col["campaign_name"]] || "",
        campaign_name: normalizeCampaignName(row[col["campaign_name"]] || ""),
        campaign_status: row[col["campaign_status"]] || "",
        campaign_started_at: row[col["campaign_started_at"]] || "",
        sequence_name_raw: seqName,
        sequence_name: normalizeCampaignName(seqName),

        // Derived fields: config override > data-driven > fallback
        sender: config.sender || deriveSender(seqName),
        channel_strategy: config.channel || deriveChannelFromData(channels),
        messaging_version: config.version || deriveVersion(seqName),

        // Send timing (SAST = UTC+2)
        send_day: "",
        send_hour_sast: "",
        send_time_bucket: "",

        // Email metrics (aggregate across steps)
        email_sent: 0,
        email_failed: 0,
        email_delivered: 0,
        email_bounced: 0,
        email_reply: 0,
        email_clicked: 0,

        // SendGrid metrics (more reliable email tracking)
        sg_delivered: 0,
        sg_bounced: 0,
        sg_not_delivered: 0,
        sg_processing: 0,
        sg_opens: 0,
        sg_clicks: 0,
        sg_status: "",
        sg_last_event: "",
        email_activity_matched: false,

        // LinkedIn metrics
        li_conn_actions: 0,
        li_conn_sent: 0,
        li_msg_actions: 0,
        li_msg_sent: 0,
        li_pushed_to_aimfox: 0,
        aimfox_conn_sent: 0,
        aimfox_accepted: 0,
        aimfox_reply: 0,

        // Timestamps (new reliable sources)
        first_email_sent_at: "",
        first_email_response_at: "",
        first_li_conn_sent_at: "",
        first_li_dm_sent_at: "",
        first_li_response_at: "",

        // Timestamps (legacy / cross-channel)
        first_email_delivered_at: "",
        first_email_reply_at: "",
        first_aimfox_reply_at: "",
        last_outbound_message_at: "",
        last_inbound_message_at: "",

        // Step tracking
        total_steps: 0,
        steps_with_activity: 0,
        max_step_order: 0,
        current_step: -1,
        current_step_channel: "",
        reply_step: -1,
        reply_step_channel: "",
      };
    }

    const lead = leads[convId];
    lead.total_steps++;

    const stepOrder = num(row, col["step_order"]);
    if (stepOrder > lead.max_step_order) lead.max_step_order = stepOrder;

    // Email
    lead.email_sent += num(row, col["email_sent_count"]);
    lead.email_failed += num(row, col["email_failed_count"]);
    lead.email_delivered += num(row, col["email_delivered_count"]);
    lead.email_bounced += num(row, col["email_bounced_count"]);
    lead.email_reply += num(row, col["email_reply_count"]);
    lead.email_clicked += num(row, col["email_clicked_count"]);

    // SendGrid metrics
    if (col["sendgrid_email_status"] !== undefined) {
      const sgStatus = (row[col["sendgrid_email_status"]] || "").toString();
      if (sgStatus) {
        lead.sg_status = sgStatus;
        lead.sg_delivered = 0;
        lead.sg_not_delivered = 0;
        lead.sg_processing = 0;
        lead.sg_bounced = 0;
        if (sgStatus === "delivered") lead.sg_delivered = 1;
        else if (sgStatus === "not_delivered") lead.sg_not_delivered = 1;
        else if (sgStatus === "processing") lead.sg_processing = 1;
        else if (sgStatus === "bounced") lead.sg_bounced = 1;
      }
    }
    if (col["sendgrid_opens_count"] !== undefined) {
      const opens = num(row, col["sendgrid_opens_count"]);
      if (opens > lead.sg_opens) lead.sg_opens = opens;
    }
    if (col["sendgrid_clicks_count"] !== undefined) {
      const clicks = num(row, col["sendgrid_clicks_count"]);
      if (clicks > lead.sg_clicks) lead.sg_clicks = clicks;
    }
    if (col["sendgrid_last_event_time"] !== undefined && row[col["sendgrid_last_event_time"]]) {
      lead.sg_last_event = row[col["sendgrid_last_event_time"]];
    }
    if (col["email_activity_matched"] !== undefined) {
      const matched = row[col["email_activity_matched"]];
      if (matched === true || matched === "TRUE" || matched === "true") lead.email_activity_matched = true;
    }

    // LinkedIn step-level
    lead.li_conn_actions += num(row, col["step_linkedin_connection_actions"]);
    lead.li_conn_sent += num(row, col["step_linkedin_connection_action_sent"]);
    lead.li_msg_actions += num(row, col["step_linkedin_message_actions"]);
    lead.li_msg_sent += num(row, col["step_linkedin_message_sent"]);

    // LinkedIn Aimfox
    lead.li_pushed_to_aimfox += num(row, col["linkedin_pushed_to_aimfox_count"]);
    lead.aimfox_conn_sent += num(row, col["aimfox_connection_request_sent_count"]);
    lead.aimfox_accepted += num(row, col["aimfox_connection_accepted_event_count"]);

    // Aimfox reply
    if (col["aimfox_reply_count"] !== undefined) {
      lead.aimfox_reply += num(row, col["aimfox_reply_count"]);
    }

    // Activity check
    const stepChannel = (row[col["step_channel"]] || "").toString();
    const stepEmailSent = num(row, col["email_sent_count"]);
    const stepEmailDelivered = num(row, col["email_delivered_count"]);
    const stepEmailBounced = num(row, col["email_bounced_count"]);
    const stepEmailReply = num(row, col["email_reply_count"]);
    const stepEmailClicked = num(row, col["email_clicked_count"]);
    const stepLiConnSent = num(row, col["step_linkedin_connection_action_sent"]);
    const stepLiMsgSent = num(row, col["step_linkedin_message_sent"]);
    const stepAimfoxConnSent = num(row, col["aimfox_connection_request_sent_count"]);
    const stepAimfoxAccepted = num(row, col["aimfox_connection_accepted_event_count"]);
    const stepAimfoxReply = col["aimfox_reply_count"] !== undefined ? num(row, col["aimfox_reply_count"]) : 0;

    const stepActivity = stepEmailSent + stepLiConnSent + stepLiMsgSent + stepAimfoxConnSent;
    if (stepActivity > 0) lead.steps_with_activity++;

    // Track current step (highest step with activity)
    if (stepActivity > 0 && stepOrder > lead.current_step) {
      lead.current_step = stepOrder;
      lead.current_step_channel = stepChannel;
    }

    // Track which step got a reply
    const stepReplyTotal = stepEmailReply + stepAimfoxReply;
    if (stepReplyTotal > 0) {
      lead.reply_step = stepOrder;
      lead.reply_step_channel = stepChannel;
    }

    // Collect step-level detail row
    stepDetails.push({
      conversation_id: convId,
      person_name: row[col["person_name"]] || "",
      person_email: row[col["person_email"]] || "",
      company_name: row[col["company_name"]] || "",
      sequence_name: normalizeCampaignName((row[col["sequence_name"]] || "").toString()),
      campaign_name: normalizeCampaignName(row[col["campaign_name"]] || ""),
      step_order: stepOrder,
      step_channel: stepChannel,
      step_label: "Step " + stepOrder + ": " + formatChannel(stepChannel),
      email_sent: stepEmailSent,
      email_delivered: stepEmailDelivered,
      email_bounced: stepEmailBounced,
      email_reply: stepEmailReply,
      email_clicked: stepEmailClicked,
      li_conn_sent: stepLiConnSent,
      li_msg_sent: stepLiMsgSent,
      aimfox_conn_sent: stepAimfoxConnSent,
      aimfox_accepted: stepAimfoxAccepted,
      aimfox_reply: stepAimfoxReply,
      has_activity: stepActivity > 0 ? "Yes" : "No",
      has_reply: stepReplyTotal > 0 ? "Yes" : "No",
      conversation_stage: row[col["conversation_stage"]] || "",
      automation_status: row[col["automation_status"]] || "",
      outbound_message: col["step_outbound_message"] !== undefined ? (row[col["step_outbound_message"]] || "") : "",
      reply_message: col["step_reply_message"] !== undefined ? (row[col["step_reply_message"]] || "") : "",
      email_sent_at: col["email_sent_sendgrid_at"] !== undefined ? (row[col["email_sent_sendgrid_at"]] || "") : "",
      li_conn_sent_at: col["linkedin_connection_aimfox_sent_at"] !== undefined ? (row[col["linkedin_connection_aimfox_sent_at"]] || "") : "",
      li_dm_sent_at: col["linkedin_dm_aimfox_at"] !== undefined ? (row[col["linkedin_dm_aimfox_at"]] || "") : "",
    });

    // Timestamps (take first non-empty)
    if (!lead.first_email_delivered_at && row[col["first_email_delivered_at"]]) {
      lead.first_email_delivered_at = row[col["first_email_delivered_at"]];
    }
    if (!lead.first_email_reply_at && row[col["first_email_reply_at"]]) {
      lead.first_email_reply_at = row[col["first_email_reply_at"]];
    }
    if (col["first_aimfox_reply_at"] !== undefined && !lead.first_aimfox_reply_at && row[col["first_aimfox_reply_at"]]) {
      lead.first_aimfox_reply_at = row[col["first_aimfox_reply_at"]];
    }
    if (col["last_outbound_message_at"] !== undefined) {
      const outbound = row[col["last_outbound_message_at"]];
      if (outbound && (!lead.last_outbound_message_at || outbound > lead.last_outbound_message_at)) {
        lead.last_outbound_message_at = outbound;
      }
    }
    if (col["last_inbound_message_at"] !== undefined) {
      const inbound = row[col["last_inbound_message_at"]];
      if (inbound && (!lead.last_inbound_message_at || inbound > lead.last_inbound_message_at)) {
        lead.last_inbound_message_at = inbound;
      }
    }

    // New reliable timestamps (earliest non-empty across steps)
    if (col["email_sent_sendgrid_at"] !== undefined) {
      const ts = row[col["email_sent_sendgrid_at"]];
      if (ts && (!lead.first_email_sent_at || ts < lead.first_email_sent_at)) {
        lead.first_email_sent_at = ts;
      }
    }
    if (col["email_response_sendgrid_at"] !== undefined) {
      const ts = row[col["email_response_sendgrid_at"]];
      if (ts && (!lead.first_email_response_at || ts < lead.first_email_response_at)) {
        lead.first_email_response_at = ts;
      }
    }
    if (col["linkedin_connection_aimfox_sent_at"] !== undefined) {
      const ts = row[col["linkedin_connection_aimfox_sent_at"]];
      if (ts && (!lead.first_li_conn_sent_at || ts < lead.first_li_conn_sent_at)) {
        lead.first_li_conn_sent_at = ts;
      }
    }
    if (col["linkedin_dm_aimfox_at"] !== undefined) {
      const ts = row[col["linkedin_dm_aimfox_at"]];
      if (ts && (!lead.first_li_dm_sent_at || ts < lead.first_li_dm_sent_at)) {
        lead.first_li_dm_sent_at = ts;
      }
    }
    if (col["linkedin_response_aimfox_at"] !== undefined) {
      const ts = row[col["linkedin_response_aimfox_at"]];
      if (ts && (!lead.first_li_response_at || ts < lead.first_li_response_at)) {
        lead.first_li_response_at = ts;
      }
    }
  });

  // ============================================================
  // 2. COMPUTE DERIVED METRICS PER LEAD
  // ============================================================
  const leadArray = Object.values(leads);

  leadArray.forEach(lead => {
    // Prefer new reliable timestamps, fall back to legacy
    var sendTimestamp = lead.first_email_sent_at
      || lead.first_li_conn_sent_at
      || lead.first_li_dm_sent_at
      || lead.last_outbound_message_at
      || lead.campaign_started_at
      || "";
    if (sendTimestamp) {
      try {
        var dt = new Date(sendTimestamp);
        if (!isNaN(dt.getTime())) {
          dt = new Date(dt.getTime() + 2 * 60 * 60 * 1000);
          var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
          lead.send_day = days[dt.getUTCDay()];
          var hour = dt.getUTCHours();
          lead.send_hour_sast = hour;
          if (hour >= 6 && hour < 9) lead.send_time_bucket = "6-9am";
          else if (hour >= 9 && hour < 12) lead.send_time_bucket = "9am-12pm";
          else if (hour >= 12 && hour < 15) lead.send_time_bucket = "12-3pm";
          else if (hour >= 15 && hour < 18) lead.send_time_bucket = "3-6pm";
          else if (hour >= 18 && hour < 21) lead.send_time_bucket = "6-9pm";
          else lead.send_time_bucket = "Off-hours";
        }
      } catch(e) {}
    }

    if (lead.email_activity_matched) {
      lead.email_delivered = lead.sg_delivered;
      if (lead.email_bounced === 0 && lead.sg_bounced > 0) {
        lead.email_bounced = lead.sg_bounced;
      }
      if (lead.email_failed === 0 && lead.sg_not_delivered > 0) {
        lead.email_failed = lead.sg_not_delivered;
      }
    }
    lead.email_opens = lead.sg_opens;
    lead.email_clicks = Math.max(lead.sg_clicks, lead.email_clicked);

    lead.total_touches = lead.email_sent + lead.li_conn_sent + lead.li_msg_sent;
    lead.total_replies = lead.email_reply + lead.aimfox_reply;
    lead.has_replied = lead.total_replies > 0 ? "Yes" : "No";

    lead.sequence_progress = lead.max_step_order > 0
      ? Math.round((lead.steps_with_activity / lead.total_steps) * 100)
      : 0;

    if (lead.email_sent === 0) {
      lead.email_delivery_status = "Not Sent";
    } else if (lead.email_reply > 0) {
      lead.email_delivery_status = "Replied";
    } else if (lead.email_delivered > 0) {
      lead.email_delivery_status = "Delivered";
    } else if (lead.email_bounced > 0) {
      lead.email_delivery_status = "Bounced";
    } else if (lead.email_failed > 0) {
      lead.email_delivery_status = "Failed";
    } else if (lead.sg_status === "processing") {
      lead.email_delivery_status = "Processing";
    } else {
      lead.email_delivery_status = "Sent";
    }

    lead.role_category = categorizeRole(lead.person_title);
  });

  // ============================================================
  // 3. WRITE LEAD SUMMARY SHEET
  // ============================================================
  const leadHeaders = [
    "Conversation ID", "Person Name", "Person Email", "Person Title", "Role Category",
    "Person City", "Person Country",
    "Company Name", "Company Website", "Company City", "Company Employee Range", "Primary Industry",
    "Conversation Stage", "Automation Status", "Temperature",
    "Campaign", "Campaign Status", "Campaign Started",
    "Sequence", "Sender", "Channel Strategy", "Messaging Version",
    "Send Day (SAST)", "Send Hour (SAST)", "Send Time Bucket",
    "Email Sent", "Email Delivered", "Email Bounced", "Email Failed", "Email Replies",
    "Email Opens", "Email Clicks",
    "LI Connection Actions", "LI Connections Sent",
    "LI Connections Accepted", "LI Messages Sent", "LI Replies",
    "Total Touches", "Total Replies", "Has Replied",
    "Total Steps", "Steps with Activity", "Max Step", "Sequence Progress %",
    "First Email Sent", "First Email Delivered", "First Email Response",
    "First LI Connection Sent", "First LI DM Sent", "First LI Response",
    "Last Outbound", "Last Inbound",
    "Current Step", "Current Step Channel", "Reply Step", "Reply Step Channel",
    "SendGrid Status", "Email Delivery Status"
  ];

  const leadRows = leadArray.map(l => [
    l.conversation_id, l.person_name, l.person_email, l.person_title, l.role_category,
    l.person_city, l.person_country,
    l.company_name, l.company_website, l.company_city, l.company_employee_range, l.primary_industry,
    l.conversation_stage, l.automation_status, l.temperature,
    l.campaign_name, l.campaign_status, l.campaign_started_at,
    l.sequence_name, l.sender, l.channel_strategy, l.messaging_version,
    l.send_day, l.send_hour_sast, l.send_time_bucket,
    l.email_sent, l.email_delivered, l.email_bounced, l.email_failed, l.email_reply,
    l.email_opens, l.email_clicks,
    l.li_conn_actions, l.li_conn_sent,
    l.aimfox_accepted, l.li_msg_sent, l.aimfox_reply,
    l.total_touches, l.total_replies, l.has_replied,
    l.total_steps, l.steps_with_activity, l.max_step_order, l.sequence_progress,
    l.first_email_sent_at, l.first_email_delivered_at, l.first_email_response_at,
    l.first_li_conn_sent_at, l.first_li_dm_sent_at, l.first_li_response_at,
    l.last_outbound_message_at, l.last_inbound_message_at,
    l.current_step >= 0 ? l.current_step : "",
    l.current_step_channel ? formatChannel(l.current_step_channel) : "",
    l.reply_step >= 0 ? l.reply_step : "",
    l.reply_step_channel ? formatChannel(l.reply_step_channel) : "",
    l.sg_status || "",
    l.email_delivery_status
  ]);

  writeSheet(ss, LEAD_SHEET, leadHeaders, leadRows);

  // ============================================================
  // 4. BUILD SEQUENCE SUMMARY
  // ============================================================
  const seqMap = {};

  leadArray.forEach(lead => {
    const key = [lead.campaign_name, lead.sender, lead.channel_strategy, lead.messaging_version].join(" | ");
    if (!seqMap[key]) {
      seqMap[key] = {
        sequence_name: lead.campaign_name,
        sender: lead.sender,
        channel_strategy: lead.channel_strategy,
        messaging_version: lead.messaging_version,
        leads: 0,
        companies: new Set(),
        email_sent: 0,
        email_delivered: 0,
        email_bounced: 0,
        email_failed: 0,
        email_reply: 0,
        email_opens: 0,
        email_clicks: 0,
        li_conn_sent: 0,
        aimfox_accepted: 0,
        aimfox_reply: 0,
        li_msg_sent: 0,
        total_touches: 0,
        total_replies: 0,
        leads_contacted: 0,
        leads_scheduled: 0,
        leads_conn_scheduled: 0,
      };
    }

    const s = seqMap[key];
    s.leads++;
    if (lead.company_name) s.companies.add(lead.company_name);
    s.email_sent += lead.email_sent;
    s.email_delivered += lead.email_delivered;
    s.email_bounced += lead.email_bounced;
    s.email_failed += lead.email_failed;
    s.email_reply += lead.email_reply;
    s.email_opens += lead.email_opens;
    s.email_clicks += lead.email_clicks;
    s.li_conn_sent += lead.li_conn_sent;
    s.aimfox_accepted += lead.aimfox_accepted;
    s.aimfox_reply += lead.aimfox_reply;
    s.li_msg_sent += lead.li_msg_sent;
    s.total_touches += lead.total_touches;
    s.total_replies += lead.total_replies;

    if (lead.conversation_stage === "contacted") s.leads_contacted++;
    if (lead.conversation_stage === "contact_scheduled") s.leads_scheduled++;
    if (lead.conversation_stage === "connection_scheduled") s.leads_conn_scheduled++;
  });

  const seqHeaders = [
    "Campaign", "Sender", "Channel Strategy", "Messaging Version",
    "Total Leads", "Unique Companies",
    "Leads Contacted", "Leads Scheduled", "Leads Conn Scheduled",
    "Email Sent", "Email Delivered", "Email Bounced", "Email Failed", "Email Replies",
    "Delivery Rate %", "Email Opens", "Email Clicks", "Open Rate %",
    "LI Connections Sent", "LI Accepted", "LI Acceptance Rate %",
    "LI Messages Sent", "LI Replies",
    "Total Touches", "Total Replies", "Reply Rate (All Touches) %",
    "Email Reply Rate %", "LI DM Reply Rate %", "Overall Reply Rate %"
  ];

  const seqRows = Object.values(seqMap).map(s => {
    const delivRate = s.email_sent > 0 ? Math.round((s.email_delivered / s.email_sent) * 1000) / 10 : 0;
    const openRate = s.email_delivered > 0 ? Math.round((s.email_opens / s.email_delivered) * 1000) / 10 : 0;
    const acceptRate = s.li_conn_sent > 0 ? Math.round((s.aimfox_accepted / s.li_conn_sent) * 1000) / 10 : 0;
    const replyRate = s.total_touches > 0 ? Math.round((s.total_replies / s.total_touches) * 1000) / 10 : 0;
    const emailReplyRate = s.email_delivered > 0 ? Math.round((s.email_reply / s.email_delivered) * 1000) / 10 : 0;
    const liDmReplyRate = s.li_msg_sent > 0 ? Math.round((s.aimfox_reply / s.li_msg_sent) * 1000) / 10 : 0;
    const overallDenom = s.email_delivered + s.li_msg_sent;
    const overallReplyRate = overallDenom > 0 ? Math.round((s.total_replies / overallDenom) * 1000) / 10 : 0;

    return [
      s.sequence_name, s.sender, s.channel_strategy, s.messaging_version,
      s.leads, s.companies.size,
      s.leads_contacted, s.leads_scheduled, s.leads_conn_scheduled,
      s.email_sent, s.email_delivered, s.email_bounced, s.email_failed, s.email_reply,
      delivRate, s.email_opens, s.email_clicks, openRate,
      s.li_conn_sent, s.aimfox_accepted, acceptRate,
      s.li_msg_sent, s.aimfox_reply,
      s.total_touches, s.total_replies, replyRate,
      emailReplyRate, liDmReplyRate, overallReplyRate
    ];
  });

  seqRows.sort((a, b) => b[4] - a[4]);

  writeSheet(ss, SEQUENCE_SHEET, seqHeaders, seqRows);

  // ============================================================
  // 5. BUILD CAMPAIGN SUMMARY
  // ============================================================
  const campMap = {};

  leadArray.forEach(lead => {
    const key = lead.campaign_name;
    if (!key) return;

    if (!campMap[key]) {
      campMap[key] = {
        campaign_name: lead.campaign_name,
        campaign_status: lead.campaign_status,
        campaign_started_at: lead.campaign_started_at,
        leads: 0,
        companies: new Set(),
        sequences: new Set(),
        senders: new Set(),
        email_sent: 0,
        email_delivered: 0,
        email_bounced: 0,
        email_failed: 0,
        email_reply: 0,
        li_conn_sent: 0,
        aimfox_accepted: 0,
        aimfox_reply: 0,
        li_msg_sent: 0,
        total_touches: 0,
        total_replies: 0,
        leads_contacted: 0,
        leads_scheduled: 0,
        leads_conn_scheduled: 0,
      };
    }

    const c = campMap[key];
    c.leads++;
    if (lead.company_name) c.companies.add(lead.company_name);
    c.sequences.add(lead.sequence_name);
    c.senders.add(lead.sender);
    c.email_sent += lead.email_sent;
    c.email_delivered += lead.email_delivered;
    c.email_bounced += lead.email_bounced;
    c.email_failed += lead.email_failed;
    c.email_reply += lead.email_reply;
    c.li_conn_sent += lead.li_conn_sent;
    c.aimfox_accepted += lead.aimfox_accepted;
    c.aimfox_reply += lead.aimfox_reply;
    c.li_msg_sent += lead.li_msg_sent;
    c.total_touches += lead.total_touches;
    c.total_replies += lead.total_replies;

    if (lead.conversation_stage === "contacted") c.leads_contacted++;
    if (lead.conversation_stage === "contact_scheduled") c.leads_scheduled++;
    if (lead.conversation_stage === "connection_scheduled") c.leads_conn_scheduled++;
  });

  const campHeaders = [
    "Campaign Name", "Status", "Started At",
    "Total Leads", "Unique Companies", "Sequences", "Senders",
    "Leads Contacted", "Leads Scheduled", "Leads Conn Scheduled",
    "Email Sent", "Email Delivered", "Email Bounced", "Email Failed", "Email Replies",
    "Delivery Rate %",
    "LI Connections Sent", "LI Accepted", "LI Acceptance Rate %",
    "LI Messages Sent", "LI Replies",
    "Total Touches", "Total Replies", "Reply Rate (All Touches) %",
    "Email Reply Rate %", "LI DM Reply Rate %", "Overall Reply Rate %"
  ];

  const campRows = Object.values(campMap).map(c => {
    const delivRate = c.email_sent > 0 ? Math.round((c.email_delivered / c.email_sent) * 1000) / 10 : 0;
    const acceptRate = c.li_conn_sent > 0 ? Math.round((c.aimfox_accepted / c.li_conn_sent) * 1000) / 10 : 0;
    const replyRate = c.total_touches > 0 ? Math.round((c.total_replies / c.total_touches) * 1000) / 10 : 0;
    const emailReplyRate = c.email_delivered > 0 ? Math.round((c.email_reply / c.email_delivered) * 1000) / 10 : 0;
    const liDmReplyRate = c.li_msg_sent > 0 ? Math.round((c.aimfox_reply / c.li_msg_sent) * 1000) / 10 : 0;
    const overallDenom = c.email_delivered + c.li_msg_sent;
    const overallReplyRate = overallDenom > 0 ? Math.round((c.total_replies / overallDenom) * 1000) / 10 : 0;

    return [
      c.campaign_name, c.campaign_status, c.campaign_started_at,
      c.leads, c.companies.size,
      Array.from(c.sequences).join(", "),
      Array.from(c.senders).join(", "),
      c.leads_contacted, c.leads_scheduled, c.leads_conn_scheduled,
      c.email_sent, c.email_delivered, c.email_bounced, c.email_failed, c.email_reply,
      delivRate,
      c.li_conn_sent, c.aimfox_accepted, acceptRate,
      c.li_msg_sent, c.aimfox_reply,
      c.total_touches, c.total_replies, replyRate,
      emailReplyRate, liDmReplyRate, overallReplyRate
    ];
  });

  campRows.sort((a, b) => b[3] - a[3]);

  writeSheet(ss, CAMPAIGN_SHEET, campHeaders, campRows);

  // ============================================================
  // 6. WRITE STEP DETAIL SHEET (one row per step per lead)
  // ============================================================
  const stepDetailHeaders = [
    "Conversation ID", "Person Name", "Person Email", "Company Name",
    "Sequence", "Campaign",
    "Step Order", "Step Channel", "Step Label",
    "Email Sent", "Email Delivered", "Email Bounced", "Email Reply", "Email Clicked",
    "LI Connection Sent", "LI Message Sent",
    "Aimfox Conn Sent", "Aimfox Accepted", "Aimfox Reply",
    "Has Activity", "Has Reply",
    "Conversation Stage", "Automation Status",
    "Outbound Message", "Reply Message",
    "Email Sent At", "LI Connection Sent At", "LI DM Sent At"
  ];

  stepDetails.sort((a, b) => {
    const nameCompare = a.person_name.toString().localeCompare(b.person_name.toString());
    if (nameCompare !== 0) return nameCompare;
    return a.step_order - b.step_order;
  });

  const stepDetailRows = stepDetails.map(sd => [
    sd.conversation_id, sd.person_name, sd.person_email, sd.company_name,
    sd.sequence_name, sd.campaign_name,
    sd.step_order, formatChannel(sd.step_channel), sd.step_label,
    sd.email_sent, sd.email_delivered, sd.email_bounced, sd.email_reply, sd.email_clicked,
    sd.li_conn_sent, sd.li_msg_sent,
    sd.aimfox_conn_sent, sd.aimfox_accepted, sd.aimfox_reply,
    sd.has_activity, sd.has_reply,
    sd.conversation_stage, sd.automation_status,
    sd.outbound_message, sd.reply_message,
    sd.email_sent_at, sd.li_conn_sent_at, sd.li_dm_sent_at
  ]);

  writeSheet(ss, STEP_DETAIL_SHEET, stepDetailHeaders, stepDetailRows);

  // ============================================================
  // 7. BUILD STEP PERFORMANCE SHEET (aggregate per sequence per step)
  // ============================================================
  const stepPerfMap = {};

  stepDetails.forEach(sd => {
    const key = sd.sequence_name + "|" + sd.step_order + "|" + sd.step_channel;
    if (!stepPerfMap[key]) {
      stepPerfMap[key] = {
        sequence_name: sd.sequence_name,
        step_order: sd.step_order,
        step_channel: sd.step_channel,
        leads_in_step: 0,
        leads_with_activity: 0,
        leads_with_reply: 0,
        email_sent: 0,
        email_delivered: 0,
        email_bounced: 0,
        email_reply: 0,
        li_conn_sent: 0,
        li_msg_sent: 0,
        aimfox_accepted: 0,
        aimfox_reply: 0,
      };
    }

    const sp = stepPerfMap[key];
    sp.leads_in_step++;
    if (sd.has_activity === "Yes") sp.leads_with_activity++;
    if (sd.has_reply === "Yes") sp.leads_with_reply++;
    sp.email_sent += sd.email_sent;
    sp.email_delivered += sd.email_delivered;
    sp.email_bounced += sd.email_bounced;
    sp.email_reply += sd.email_reply;
    sp.li_conn_sent += sd.li_conn_sent;
    sp.li_msg_sent += sd.li_msg_sent;
    sp.aimfox_accepted += sd.aimfox_accepted;
    sp.aimfox_reply += sd.aimfox_reply;
  });

  const stepPerfHeaders = [
    "Sequence", "Step Order", "Step Channel",
    "Leads in Step", "Leads with Activity", "Leads with Reply",
    "Email Sent", "Email Delivered", "Email Bounced", "Email Replies",
    "Email Reply Rate %",
    "LI Connections Sent", "LI Messages Sent",
    "Aimfox Accepted", "Aimfox Replies",
    "LI DM Reply Rate %",
    "Total Touches", "Total Replies", "Step Reply Rate %", "Overall Reply Rate %"
  ];

  const stepPerfRows = Object.values(stepPerfMap).map(sp => {
    const emailReplyRate = sp.email_delivered > 0
      ? Math.round((sp.email_reply / sp.email_delivered) * 1000) / 10 : 0;
    const liDmReplyRate = sp.li_msg_sent > 0
      ? Math.round((sp.aimfox_reply / sp.li_msg_sent) * 1000) / 10 : 0;
    const totalTouches = sp.email_sent + sp.li_conn_sent + sp.li_msg_sent;
    const totalReplies = sp.email_reply + sp.aimfox_reply;
    const stepReplyRate = totalTouches > 0
      ? Math.round((totalReplies / totalTouches) * 1000) / 10 : 0;
    const overallDenom = sp.email_delivered + sp.li_msg_sent;
    const overallReplyRate = overallDenom > 0
      ? Math.round((totalReplies / overallDenom) * 1000) / 10 : 0;

    return [
      sp.sequence_name, sp.step_order, formatChannel(sp.step_channel),
      sp.leads_in_step, sp.leads_with_activity, sp.leads_with_reply,
      sp.email_sent, sp.email_delivered, sp.email_bounced, sp.email_reply,
      emailReplyRate,
      sp.li_conn_sent, sp.li_msg_sent,
      sp.aimfox_accepted, sp.aimfox_reply,
      liDmReplyRate,
      totalTouches, totalReplies, stepReplyRate, overallReplyRate
    ];
  });

  stepPerfRows.sort((a, b) => {
    const seqCompare = a[0].toString().localeCompare(b[0].toString());
    if (seqCompare !== 0) return seqCompare;
    return a[1] - b[1];
  });

  writeSheet(ss, STEP_PERFORMANCE_SHEET, stepPerfHeaders, stepPerfRows);

  // ============================================================
  // 8. SUMMARY LOG
  // ============================================================
  const msg = "Dashboard data refreshed!\n\n" +
    "Campaign Summary: " + Object.keys(campMap).length + " campaigns\n" +
    "Sequence Summary: " + Object.keys(seqMap).length + " sequences\n" +
    "Lead Summary: " + leadArray.length + " leads\n" +
    "Step Detail: " + stepDetails.length + " step rows\n" +
    "Step Performance: " + Object.keys(stepPerfMap).length + " sequence-step combos\n" +
    "Companies: " + new Set(leadArray.map(l => l.company_name).filter(Boolean)).size + "\n\n" +
    "Email sent: " + leadArray.reduce((s, l) => s + l.email_sent, 0) + "\n" +
    "Email delivered: " + leadArray.reduce((s, l) => s + l.email_delivered, 0) + "\n" +
    "LI conn sent: " + leadArray.reduce((s, l) => s + l.li_conn_sent, 0) + "\n" +
    "Total replies: " + leadArray.reduce((s, l) => s + l.total_replies, 0) + "\n\n" +
    "5 sheets created: '" + CAMPAIGN_SHEET + "', '" + SEQUENCE_SHEET + "', '" + LEAD_SHEET +
    "', '" + STEP_DETAIL_SHEET + "', '" + STEP_PERFORMANCE_SHEET + "'.\n" +
    "Connect Looker Studio to any or all of them.";

  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================

function num(row, colIndex) {
  if (colIndex === undefined || colIndex === null) return 0;
  const val = row[colIndex];
  if (val === null || val === undefined || val === "") return 0;
  const n = Number(val);
  return isNaN(n) ? 0 : n;
}

function deriveSender(seqName) {
  const senderNames = ["Renier", "Samihah", "Samhiah", "Shaunelle"];
  const seqLower = seqName.toLowerCase();
  for (const name of senderNames) {
    if (seqLower.includes(name.toLowerCase())) {
      if (name.toLowerCase() === "samhiah") return "Samihah";
      return name;
    }
  }
  return "Team";
}

function deriveChannelFromData(channels) {
  const hasEmail = channels.has("email");
  const hasLinkedIn = channels.has("linkedin_conn") || channels.has("linkedin_msg");
  if (hasEmail && hasLinkedIn) return "Hybrid";
  if (hasEmail && !hasLinkedIn) return "Email Only";
  if (!hasEmail && hasLinkedIn) return "LinkedIn Only";
  return "Unknown";
}

function normalizeCampaignName(name) {
  if (!name) return "Unknown";
  var s = name.toLowerCase();
  if (s.includes("pharma")) return "Pharmaceuticals";
  if (s.includes("embassy") || s.includes("consulate")) return "Embassies & Consulates";
  if (s.includes("travel") || s.includes("global travel")) return "Travel & Corporate";
  if (s.includes("lead") && (s.includes("targeting") || s.includes("outreach"))) return "Lead Targeting";
  return name;
}

function deriveVersion(seqName) {
  var vMap = {"1": "A", "2": "B", "A": "A", "B": "B"};
  var match = seqName.match(/version[_\s]?([a-z0-9]+)/i);
  if (match) {
    var key = match[1].toUpperCase();
    return vMap[key] || key;
  }
  var vMatch = seqName.match(/[_\-]V(\d+)/i);
  if (vMatch) return vMap[vMatch[1]] || "V" + vMatch[1];
  var vMatch2 = seqName.match(/\bV(\d+)\b/i);
  if (vMatch2) return vMap[vMatch2[1]] || "V" + vMatch2[1];
  return "N/A";
}

function loadSequenceConfig(ss) {
  const config = {};
  const sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet) return config;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return config;
  const headers = data[0].map(h => h.toString().trim().toLowerCase());
  const nameIdx = headers.indexOf("sequence name");
  const senderIdx = headers.indexOf("sender");
  const channelIdx = headers.indexOf("channel");
  const versionIdx = headers.indexOf("version");
  if (nameIdx === -1) return config;
  for (let i = 1; i < data.length; i++) {
    const name = (data[i][nameIdx] || "").toString().trim();
    if (!name) continue;
    config[name] = {};
    if (senderIdx !== -1 && data[i][senderIdx]) config[name].sender = data[i][senderIdx].toString().trim();
    if (channelIdx !== -1 && data[i][channelIdx]) config[name].channel = data[i][channelIdx].toString().trim();
    if (versionIdx !== -1 && data[i][versionIdx]) config[name].version = data[i][versionIdx].toString().trim();
  }
  return config;
}

function formatChannel(channel) {
  if (!channel) return "Unknown";
  var map = {
    "email": "Email",
    "linkedin_conn": "LinkedIn Connection",
    "linkedin_msg": "LinkedIn Message"
  };
  return map[channel.toLowerCase()] || channel;
}

function categorizeRole(title) {
  if (!title) return "Unknown";
  const t = title.toLowerCase();
  if (/financ/i.test(t)) return "Finance";
  if (/procurement|sourcing|supply chain/i.test(t)) return "Procurement";
  if (/personal assistant|^pa$|^pa /i.test(t) || /executive assistant/i.test(t) || /exec.*pa/i.test(t)) return "PA / Executive Assistant";
  if (/office manager|admin/i.test(t)) return "Office Manager / Admin";
  if (/general manager|gm|director|head of|chief|ceo|cfo|coo/i.test(t)) return "Senior Leadership";
  if (/travel/i.test(t)) return "Travel";
  if (/business development|sales/i.test(t)) return "Business Development";
  return "Other";
}

function writeSheet(ss, sheetName, headers, rows) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(10);
  headerRange.setBackground("#F3F4F6");
  headerRange.setWrap(true);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // Auto-resize first 15 columns in one batch call
  const colsToResize = Math.min(headers.length, 15);
  if (colsToResize > 0) {
    sheet.autoResizeColumns(1, colsToResize);
  }

  sheet.setFrozenRows(1);
}

// ============================================================
// OPTIONAL: SET UP AUTO-REFRESH TRIGGER
// ============================================================
function createHourlyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "transformForDashboard") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("transformForDashboard")
    .timeBased()
    .everyHours(1)
    .create();

  SpreadsheetApp.getUi().alert("Auto-refresh trigger set: runs every hour.");
}

function removeAutoRefreshTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "transformForDashboard") {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  SpreadsheetApp.getUi().alert("Removed " + removed + " trigger(s).");
}
