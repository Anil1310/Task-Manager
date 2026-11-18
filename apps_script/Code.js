/**
 * Aarcon Backend — Connected to Google Sheet
 * Sheet ID: 1sXbqbx2CmwvxXvdkdaB55HeOuhsSPy1RwwEXSKHI8Xc
 * Handles all CRUD operations for Users and Tasks
 */

const SHEET_ID = '1sXbqbx2CmwvxXvdkdaB55HeOuhsSPy1RwwEXSKHI8Xc';

// ============================================================
// 1. MAIN ENTRY POINT — Serves the web app
// ============================================================
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Aarcon — Construction ReortTask Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// 2. GET SPREADSHEET — Opens by ID (works in deployed web app)
// ============================================================
function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

// ============================================================
// 3. INITIAL DATA — Fetches Users + Tasks for the frontend
// ============================================================
function getInitialData() {
  const ss = getSpreadsheet();
  
  const userSheet = ss.getSheetByName('Users');
  const taskSheet = ss.getSheetByName('Tasks');
  const projectSheet = ss.getSheetByName('Projects');
  const logSheet = ss.getSheetByName('Logs');
  const settingsSheet = ss.getSheetByName('Settings');
  const sessionSheet = ss.getSheetByName('Sessions');
  
  if (userSheet && taskSheet && projectSheet && logSheet && settingsSheet && sessionSheet) {
    const settingsData = getSheetDataAsJson(settingsSheet);
    const settings = {};
    settingsData.forEach(function(row) {
      if (row.setting) settings[row.setting] = row.value;
    });
    
    // Fill default values in memory
    const defaults = {
      showLogsToUsers: 'false',
      showAttachmentsToUsers: 'false',
      showLogsToAdmins: 'true',
      showAttachmentsToAdmins: 'true',
      showTeamToAdmins: 'true',
      showProjectsToAdmins: 'true',
      showSettingsToAdmins: 'true',
      showTeamToManagers: 'false',
      showTeamToUsers: 'false',
      showProjectsToManagers: 'false',
      showProjectsToUsers: 'false',
      showSettingsToManagers: 'false',
      showSettingsToUsers: 'false',
      allowNewTaskToManagers: 'true',
      allowNewTaskToUsers: 'true',
      disableDemoSeed: 'false'
    };
    
    const settingsToCreate = [];
    for (let key in defaults) {
      if (settings[key] === undefined) {
        settings[key] = defaults[key];
        settingsToCreate.push([key, defaults[key]]);
      }
    }
    
    // Batch write any missing settings in a single sheet call instead of one by one
    if (settingsToCreate.length > 0) {
      try {
        const lastRow = settingsSheet.getLastRow();
        settingsSheet.getRange(lastRow + 1, 1, settingsToCreate.length, 2).setValues(settingsToCreate);
      } catch (e) {
        console.log("Failed to batch write settings: " + e.message);
      }
    }
    
    return {
      users: getSheetDataAsJson(userSheet),
      tasks: getSheetDataAsJson(taskSheet),
      projects: getSheetDataAsJson(projectSheet),
      settings: settings
    };
  }
  
  return doSlowSetupAndGetInitialData(ss);
}

function doSlowSetupAndGetInitialData(ss) {
  // Ensure sheets exist with correct headers
  const userSheet = getOrCreateSheet(ss, 'Users', [
    'Username', 'Password', 'Name', 'Role', 'Initials', 'Color', 'Department', 'Manager', 'Project', 'Email', 'Mobile', 'Photo'
  ]);
  ensureHeadersExist(userSheet, ['Role', 'Initials', 'Color', 'Department', 'Manager', 'Project', 'Email', 'Mobile', 'Photo']);
  const taskSheet = getOrCreateSheet(ss, 'Tasks', [
    'Id', 'Title', 'Desc', 'Project', 'Assignee', 'Email', 'Assigned By', 'Priority',
    'Status', 'Progress', 'Deadline', 'Created', 'Comments'
  ]);
  ensureHeadersExist(taskSheet, ['Assigned By']);
  const projectSheet = getOrCreateSheet(ss, 'Projects', [
    'Name', 'Location', 'Progress', 'Budget'
  ]);
  const logSheet = getOrCreateSheet(ss, 'Logs', [
    'Timestamp', 'TaskId', 'FileName', 'FileUrl', 'UploadedBy'
  ]);
  const sessionSheet = getOrCreateSheet(ss, 'Sessions', [
    'Timestamp', 'Username', 'Action', 'Details'
  ]);
  const settingsSheet = getOrCreateSheet(ss, 'Settings', [
    'Setting', 'Value', 'Description', 'LastModified'
  ]);

  const settingsData = getSheetDataAsJson(settingsSheet);
  const settings = {};
  settingsData.forEach(function(row) {
    if (row.setting) settings[row.setting] = row.value;
  });

  // Fill default values in memory
  const defaults = {
    showLogsToUsers: 'false',
    showAttachmentsToUsers: 'false',
    showLogsToAdmins: 'true',
    showAttachmentsToAdmins: 'true',
    showTeamToAdmins: 'true',
    showProjectsToAdmins: 'true',
    showSettingsToAdmins: 'true',
    showTeamToManagers: 'false',
    showTeamToUsers: 'false',
    showProjectsToManagers: 'false',
    showProjectsToUsers: 'false',
    showSettingsToManagers: 'false',
    showSettingsToUsers: 'false',
    allowNewTaskToManagers: 'true',
    allowNewTaskToUsers: 'true',
    disableDemoSeed: 'false'
  };

  const settingsToCreate = [];
  for (let key in defaults) {
    if (settings[key] === undefined) {
      settings[key] = defaults[key];
      settingsToCreate.push([key, defaults[key]]);
    }
  }

  // Batch write any missing settings in a single sheet call instead of one by one
  if (settingsToCreate.length > 0) {
    try {
      const lastRow = settingsSheet.getLastRow();
      settingsSheet.getRange(lastRow + 1, 1, settingsToCreate.length, 2).setValues(settingsToCreate);
    } catch (e) {
      console.log("Failed to batch write settings: " + e.message);
    }
  }

  // Dynamic user list update
  var usersList = getSheetDataAsJson(userSheet);

  return {
    users: usersList,
    tasks: getSheetDataAsJson(taskSheet),
    projects: getSheetDataAsJson(projectSheet),
    settings: settings
  };
}


// ============================================================
// 4. SAVE TASK — Create or Update a task
// ============================================================
function saveTaskToSheet(taskObj) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Tasks');
  if (!sheet) throw new Error('Tasks sheet not found');

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Convert comments array to JSON string for storage
  const taskData = {
    ...taskObj,
    comments: JSON.stringify(taskObj.comments || [])
  };

  // Map object keys to correct column positions
  const normalizedHeaders = normalizeHeaders(headers);
  const rowValues = normalizedHeaders.map(function(key) {
    return taskData[key] !== undefined ? taskData[key] : '';
  });

  const idIndex = normalizedHeaders.indexOf('id');
  const rowIndex = data.findIndex(function(row) {
    return String(row[idIndex]) === String(taskObj.id);
  });

  var isNewTask = rowIndex === -1;
  var oldAssignee = '';
  
  if (rowIndex > -1) {
    // Update existing row - check if assignee changed
    const assigneeIndex = normalizedHeaders.indexOf('assignee');
    if (assigneeIndex > -1) {
      oldAssignee = String(data[rowIndex][assigneeIndex] || '').trim();
    }
    sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([rowValues]);
  } else {
    // Append new row
    sheet.appendRow(rowValues);
  }

  // Send email notification on assignment or reassignment
  if (taskObj.assignee && taskObj.assignee.trim()) {
    var assigneeChanged = !isNewTask && oldAssignee.trim().toLowerCase() !== taskObj.assignee.trim().toLowerCase();
    if (isNewTask || assigneeChanged) {
      try {
        // Notify new assignee
        sendTaskAssignmentEmail(taskObj, isNewTask ? 'new' : 'reassigned', oldAssignee);
      } catch(e) {
        console.log('Email send to new assignee failed: ' + e.message);
      }
      // If reassigned, also notify the previous assignee that task was taken away
      if (assigneeChanged && oldAssignee.trim()) {
        try {
          sendTaskUnassignedEmail(taskObj, oldAssignee.trim());
        } catch(e) {
          console.log('Email send to old assignee failed: ' + e.message);
        }
      }
    }
  }

  return { success: true };
}

// ============================================================
// SEND EMAIL NOTIFICATION FOR TASK ASSIGNMENT / REASSIGNMENT
// ============================================================
function sendTaskAssignmentEmail(taskObj, type, oldAssignee) {
  try {
    if (!taskObj || !taskObj.assignee || !taskObj.assignee.trim()) return;

    const userSheet = getSpreadsheet().getSheetByName('Users');
    if (!userSheet) return;

    const userData = getSheetDataAsJson(userSheet);

    // Look up new assignee's email
    const assigneeUser = userData.find(function(u) {
      return String(u.username || '').toLowerCase().trim() === String(taskObj.assignee || '').toLowerCase().trim();
    });
    if (!assigneeUser || !assigneeUser.email) return;
    const assigneeEmail = String(assigneeUser.email).trim();
    if (!assigneeEmail) return;

    // Look up assigner's display name (assigned by field)
    var assignedByName = taskObj.assignedby || taskObj['assigned by'] || taskObj.assignedBy || '';
    if (assignedByName) {
      var assignerUser = userData.find(function(u) {
        return String(u.username || '').toLowerCase().trim() === String(assignedByName).toLowerCase().trim();
      });
      if (assignerUser && assignerUser.name) assignedByName = assignerUser.name;
    }

    var isSelfAssign = !isReassigned && assignedByName && String(assignedByName).toLowerCase().trim() === String(taskObj.assignee || '').toLowerCase().trim();
    var isReassigned = (type === 'reassigned');
    var assigneeDisplayName = assigneeUser.name || taskObj.assignee;
    var appUrl = ScriptApp.getService().getUrl();
    var taskTitle = taskObj.title || 'Untitled Task';
    var priorityColor = getPriorityColor(taskObj.priority);
    var priorityBg  = getPriorityBg(taskObj.priority);
    var statusColor = getStatusColor(taskObj.status);

    var subjectPrefix = isReassigned ? '🔄 Task Reassigned to You' : (isSelfAssign ? '📝 New Task Created by You' : '📌 New Task Assigned to You');
    var subject = subjectPrefix + ': ' + taskTitle;

    var reassignedBanner = '';
    if (isReassigned && oldAssignee) {
      var oldUser = userData.find(function(u) {
        return String(u.username || '').toLowerCase().trim() === String(oldAssignee).toLowerCase().trim();
      });
      var oldName = (oldUser && oldUser.name) ? oldUser.name : oldAssignee;
      reassignedBanner = `
        <div style="background:#fff3cd;border:1px solid #ffc107;border-radius:8px;padding:12px 16px;margin:0 0 20px 0;display:flex;align-items:center;gap:10px;">
          <span style="font-size:20px;">🔄</span>
          <span style="color:#856404;font-size:13px;">
            This task was previously assigned to <strong>${oldName}</strong> and has been reassigned to you.
          </span>
        </div>`;
    }

    var htmlBody = `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f0f4f8;font-family:Arial,Helvetica,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0f4f8;padding:30px 0;">
  <tr><td align="center">
    <table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.10);">

      <!-- HEADER -->
      <tr>
        <td style="background:linear-gradient(135deg,#1a2540 0%,#0f1829 100%);padding:32px 36px 28px;">
          <table width="100%" cellpadding="0" cellspacing="0">
            <tr>
              <td>
                <div style="display:inline-block;background:#f59e0b;color:#1a1a1a;font-weight:700;font-size:13px;letter-spacing:1px;padding:5px 12px;border-radius:20px;text-transform:uppercase;">⚙️ AARCON</div>
                <h1 style="color:#ffffff;margin:14px 0 4px;font-size:22px;font-weight:700;">${isReassigned ? '🔄 Task Reassigned to You' : (isSelfAssign ? '📝 New Task Created by You' : '📌 New Task Assigned to You')}</h1>
                <p style="color:#8fa3c0;margin:0;font-size:13px;">Hi <strong style="color:#e2eaf5;">${assigneeDisplayName}</strong>, ${isSelfAssign ? 'you created a new task for yourself' : 'you have a task waiting for your attention'}.</p>
              </td>
            </tr>
          </table>
        </td>
      </tr>

      <!-- BODY -->
      <tr>
        <td style="background:#ffffff;padding:32px 36px;">

          ${reassignedBanner}

          <!-- TASK TITLE CARD -->
          <div style="background:#f7f9fc;border:1px solid #e4e9f0;border-radius:10px;padding:20px 22px;margin-bottom:22px;">
            <p style="margin:0 0 6px;font-size:11px;font-weight:700;letter-spacing:1px;color:#8fa3c0;text-transform:uppercase;">Task</p>
            <h2 style="margin:0;font-size:18px;font-weight:700;color:#1a2540;">${taskTitle}</h2>
          </div>

          <!-- DETAILS GRID -->
          <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:22px;">
            <tr>
              <!-- Project -->
              <td width="50%" style="padding-right:8px;vertical-align:top;">
                <div style="background:#f7f9fc;border:1px solid #e4e9f0;border-radius:10px;padding:16px 18px;">
                  <p style="margin:0 0 4px;font-size:11px;font-weight:700;letter-spacing:1px;color:#8fa3c0;text-transform:uppercase;">Project</p>
                  <p style="margin:0;font-size:14px;font-weight:600;color:#1a2540;">${taskObj.project || 'N/A'}</p>
                </div>
              </td>
              <!-- Priority -->
              <td width="50%" style="padding-left:8px;vertical-align:top;">
                <div style="background:#f7f9fc;border:1px solid #e4e9f0;border-radius:10px;padding:16px 18px;">
                  <p style="margin:0 0 4px;font-size:11px;font-weight:700;letter-spacing:1px;color:#8fa3c0;text-transform:uppercase;">Priority</p>
                  <div style="display:inline-block;background:${priorityBg};color:${priorityColor};font-size:13px;font-weight:700;padding:3px 10px;border-radius:12px;border:1px solid ${priorityColor}33;">${taskObj.priority || 'N/A'}</div>
                </div>
              </td>
            </tr>
            <tr><td colspan="2" style="padding:8px 0 0;"></td></tr>
            <tr>
              <!-- Deadline -->
              <td width="50%" style="padding-right:8px;vertical-align:top;">
                <div style="background:#f7f9fc;border:1px solid #e4e9f0;border-radius:10px;padding:16px 18px;">
                  <p style="margin:0 0 4px;font-size:11px;font-weight:700;letter-spacing:1px;color:#8fa3c0;text-transform:uppercase;">⏰ Deadline</p>
                  <p style="margin:0;font-size:14px;font-weight:600;color:#1a2540;">${taskObj.deadline || 'Not Set'}</p>
                </div>
              </td>
              <!-- Status -->
              <td width="50%" style="padding-left:8px;vertical-align:top;">
                <div style="background:#f7f9fc;border:1px solid #e4e9f0;border-radius:10px;padding:16px 18px;">
                  <p style="margin:0 0 4px;font-size:11px;font-weight:700;letter-spacing:1px;color:#8fa3c0;text-transform:uppercase;">Status</p>
                  <div style="display:inline-block;background:${statusColor}1a;color:${statusColor};font-size:13px;font-weight:700;padding:3px 10px;border-radius:12px;border:1px solid ${statusColor}33;">${taskObj.status || 'Pending'}</div>
                </div>
              </td>
            </tr>
          </table>

          ${isSelfAssign ? `
          <!-- SELF ASSIGNED -->
          <div style="background:#e8f5e9;border-radius:10px;padding:14px 18px;margin-bottom:22px;display:flex;align-items:center;gap:10px;">
            <span style="font-size:18px;">✍️</span>
            <div>
              <p style="margin:0;font-size:11px;font-weight:700;letter-spacing:1px;color:#2e7d32;text-transform:uppercase;">Created By</p>
              <p style="margin:2px 0 0;font-size:14px;font-weight:600;color:#1b5e20;">You (Self-Assigned)</p>
            </div>
          </div>` : (assignedByName ? `
          <!-- ASSIGNED BY -->
          <div style="background:#f0f4f8;border-radius:10px;padding:14px 18px;margin-bottom:22px;display:flex;align-items:center;gap:10px;">
            <span style="font-size:18px;">👤</span>
            <div>
              <p style="margin:0;font-size:11px;font-weight:700;letter-spacing:1px;color:#8fa3c0;text-transform:uppercase;">Assigned By</p>
              <p style="margin:2px 0 0;font-size:14px;font-weight:600;color:#1a2540;">${assignedByName}</p>
            </div>
          </div>` : '')}

          ${taskObj.desc ? `
          <!-- DESCRIPTION -->
          <div style="border-left:3px solid #f59e0b;padding:12px 16px;background:#fffbf0;border-radius:0 8px 8px 0;margin-bottom:22px;">
            <p style="margin:0 0 4px;font-size:11px;font-weight:700;letter-spacing:1px;color:#8fa3c0;text-transform:uppercase;">Description</p>
            <p style="margin:0;font-size:14px;color:#374151;line-height:1.6;">${taskObj.desc}</p>
          </div>` : ''}

          <!-- CTA BUTTON -->
          <div style="text-align:center;margin:28px 0 8px;">
            <a href="${appUrl}" style="display:inline-block;background:linear-gradient(135deg,#f59e0b,#e58b00);color:#1a1a1a;font-size:15px;font-weight:700;padding:14px 36px;text-decoration:none;border-radius:8px;letter-spacing:0.5px;">
              🚀 Open Aarcon &amp; View Task
            </a>
          </div>

        </td>
      </tr>

      <!-- FOOTER -->
      <tr>
        <td style="background:#f7f9fc;border-top:1px solid #e4e9f0;padding:20px 36px;text-align:center;">
          <p style="margin:0;font-size:12px;color:#8fa3c0;">© 2026 <strong>Aarcon</strong> — Construction Task Manager</p>
          <p style="margin:4px 0 0;font-size:11px;color:#b0bec5;">This is an automated notification. Please do not reply to this email.</p>
        </td>
      </tr>

    </table>
  </td></tr>
</table>
</body>
</html>`;

    var textBody = (isReassigned ? 'Task Reassigned to You: ' : (isSelfAssign ? 'New Task Created by You: ' : 'New Task Assigned: ')) + taskTitle + '\n\n' +
      'Project: ' + (taskObj.project || 'N/A') + '\n' +
      'Priority: ' + (taskObj.priority || 'N/A') + '\n' +
      'Status: ' + (taskObj.status || 'Pending') + '\n' +
      'Deadline: ' + (taskObj.deadline || 'Not Set') + '\n' +
      (isSelfAssign ? 'Created By: You (Self-Assigned)' + '\n' : (assignedByName ? 'Assigned By: ' + assignedByName + '\n' : '')) +
      (taskObj.desc ? 'Description: ' + taskObj.desc + '\n' : '') +
      '\nPlease log in to Aarcon to view more details: ' + appUrl;

    MailApp.sendEmail(assigneeEmail, subject, textBody, {
      htmlBody: htmlBody,
      name: 'Aarcon Task Manager',
      replyTo: 'admin@aarcon@gmail.com'
    });

  } catch(e) {
    console.log('Error sending assignment email: ' + e.message);
  }
}

// ============================================================
// SEND EMAIL TO OLD ASSIGNEE WHEN TASK IS REASSIGNED AWAY
// ============================================================
function sendTaskUnassignedEmail(taskObj, oldAssigneeUsername) {
  try {
    if (!oldAssigneeUsername) return;

    const userSheet = getSpreadsheet().getSheetByName('Users');
    if (!userSheet) return;

    const userData = getSheetDataAsJson(userSheet);
    const oldUser = userData.find(function(u) {
      return String(u.username || '').toLowerCase().trim() === oldAssigneeUsername.toLowerCase().trim();
    });
    if (!oldUser || !oldUser.email) return;
    const oldEmail = String(oldUser.email).trim();
    if (!oldEmail) return;

    // Look up new assignee display name
    var newAssigneeUser = userData.find(function(u) {
      return String(u.username || '').toLowerCase().trim() === String(taskObj.assignee || '').toLowerCase().trim();
    });
    var newAssigneeName = (newAssigneeUser && newAssigneeUser.name) ? newAssigneeUser.name : taskObj.assignee;
    var oldDisplayName = oldUser.name || oldAssigneeUsername;
    var taskTitle = taskObj.title || 'Untitled Task';
    var appUrl = ScriptApp.getService().getUrl();
    var subject = '🔁 Task Reassigned: ' + taskTitle;

    var htmlBody = `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f0f4f8;font-family:Arial,Helvetica,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0f4f8;padding:30px 0;">
  <tr><td align="center">
    <table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.10);">
      <tr>
        <td style="background:linear-gradient(135deg,#1a2540 0%,#0f1829 100%);padding:32px 36px 28px;">
          <div style="display:inline-block;background:#6b7280;color:#fff;font-weight:700;font-size:13px;letter-spacing:1px;padding:5px 12px;border-radius:20px;text-transform:uppercase;">⚙️ AARCON</div>
          <h1 style="color:#ffffff;margin:14px 0 4px;font-size:22px;font-weight:700;">🔁 Task Reassigned</h1>
          <p style="color:#8fa3c0;margin:0;font-size:13px;">Hi <strong style="color:#e2eaf5;">${oldDisplayName}</strong>, a task you were working on has been reassigned.</p>
        </td>
      </tr>
      <tr>
        <td style="background:#ffffff;padding:32px 36px;">
          <div style="background:#fff3cd;border:1px solid #ffc107;border-radius:10px;padding:16px 18px;margin-bottom:22px;">
            <p style="margin:0;font-size:14px;color:#856404;">The task <strong>${taskTitle}</strong> has been reassigned to <strong>${newAssigneeName}</strong>. It has been removed from your task list.</p>
          </div>
          <div style="background:#f7f9fc;border:1px solid #e4e9f0;border-radius:10px;padding:16px 18px;margin-bottom:22px;">
            <p style="margin:0 0 4px;font-size:11px;font-weight:700;letter-spacing:1px;color:#8fa3c0;text-transform:uppercase;">Task</p>
            <h2 style="margin:0;font-size:18px;font-weight:700;color:#1a2540;">${taskTitle}</h2>
            <p style="margin:8px 0 0;font-size:13px;color:#6b7280;">Project: ${taskObj.project || 'N/A'}</p>
          </div>
          <div style="text-align:center;margin:24px 0 8px;">
            <a href="${appUrl}" style="display:inline-block;background:linear-gradient(135deg,#6b7280,#4b5563);color:#ffffff;font-size:14px;font-weight:700;padding:12px 30px;text-decoration:none;border-radius:8px;">Open Aarcon</a>
          </div>
        </td>
      </tr>
      <tr>
        <td style="background:#f7f9fc;border-top:1px solid #e4e9f0;padding:18px 36px;text-align:center;">
          <p style="margin:0;font-size:12px;color:#8fa3c0;">© 2026 <strong>Aarcon</strong> — Construction Task Manager</p>
          <p style="margin:4px 0 0;font-size:11px;color:#b0bec5;">This is an automated notification. Please do not reply.</p>
        </td>
      </tr>
    </table>
  </td></tr>
</table>
</body>
</html>`;

    var textBody = 'Task Reassigned: ' + taskTitle + '\n\n' +
      'This task has been reassigned to ' + newAssigneeName + ' and removed from your task list.\n' +
      'Project: ' + (taskObj.project || 'N/A') + '\n\n' +
      'Open Aarcon: ' + appUrl;

    MailApp.sendEmail(oldEmail, subject, textBody, {
      htmlBody: htmlBody,
      name: 'Aarcon Task Manager',
      replyTo: 'admin@aarcon@gmail.com'
    });

  } catch(e) {
    console.log('Error sending unassignment email: ' + e.message);
  }
}

function getPriorityColor(priority) {
  const colors = {
    'High': '#ef4444',
    'Medium': '#d97706',
    'Low': '#16a34a',
    'Critical': '#dc2626'
  };
  return colors[priority] || '#6b7280';
}

function getPriorityBg(priority) {
  const colors = {
    'High': '#fef2f2',
    'Medium': '#fffbeb',
    'Low': '#f0fdf4',
    'Critical': '#fff1f2'
  };
  return colors[priority] || '#f9fafb';
}

function getStatusColor(status) {
  const colors = {
    'Pending': '#f59e0b',
    'In Progress': '#3b82f6',
    'Completed': '#22c55e',
    'On Hold': '#8b5cf6',
    'Cancelled': '#ef4444'
  };
  return colors[status] || '#6b7280';
}

// ============================================================
// 5. DELETE TASK — Remove by ID
// ============================================================
function deleteTaskFromSheet(id) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Tasks');
  if (!sheet) return { success: false, error: 'Tasks sheet not found' };

  const data = sheet.getDataRange().getValues();
  const normalizedHeaders = normalizeHeaders(data[0]);
  const idIndex = normalizedHeaders.indexOf('id');
  const rowIndex = data.findIndex(function(row) {
    return String(row[idIndex]) === String(id);
  });

  if (rowIndex > -1) {
    sheet.deleteRow(rowIndex + 1);
    return { success: true };
  }
  return { success: false, error: 'Task not found' };
}

// ============================================================
// 6. GET NEXT TASK ID — Finds max ID in sheet and returns +1
// ============================================================
function getNextTaskId() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Tasks');
  if (!sheet) return 1;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return 1;

  const normalizedHeaders = normalizeHeaders(data[0]);
  const idIndex = normalizedHeaders.indexOf('id');
  let maxId = 0;
  for (var i = 1; i < data.length; i++) {
    var num = parseInt(data[i][idIndex]);
    if (!isNaN(num) && num > maxId) maxId = num;
  }
  return maxId + 1;
}

// ============================================================
// 7. FILE UPLOAD — Upload proof files to Google Drive
// ============================================================

/**
 * Gets or creates the main uploads folder in Google Drive.
 */
function getUploadsFolder() {
  var folderName = 'Aarcon Uploads';
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

/**
 * Gets or creates a sub-folder for a specific task.
 */
function getTaskFolder(taskId) {
  var parentFolder = getUploadsFolder();
  var taskFolderName = 'Task_' + taskId;
  var folders = parentFolder.getFoldersByName(taskFolderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(taskFolderName);
}

/**
 * Uploads a file (sent as base64) to Google Drive, prepending today's timestamp to its name,
 * and logs it in both the Attachments and Logs sheet.
 * @param {Object} fileData - { taskId, fileName, mimeType, base64, uploadedBy }
 * @returns {Object} - { success, attachment }
 */
function uploadFileToTask(fileData) {
  try {
    var taskFolder = getTaskFolder(fileData.taskId);
    
    // Create today's timestamp (e.g. 2026-05-21_11-46-02)
    var timestampStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    var driveFileName = timestampStr + '_' + fileData.fileName;
    
    // Decode base64 and create blob
    var decoded = Utilities.base64Decode(fileData.base64);
    var blob = Utilities.newBlob(decoded, fileData.mimeType, driveFileName);
    
    // Save to Drive
    var file = taskFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var fileUrl = file.getUrl();
    var fileId = file.getId();
    var thumbnailUrl = '';
    
    // For images, create a direct view link
    if (fileData.mimeType.indexOf('image') > -1) {
      thumbnailUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w200';
    }
    
    var ss = getSpreadsheet();
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    
    // 1. Log attachment in the Attachments sheet (with the timestamped filename)
    var attachSheet = getOrCreateSheet(ss, 'Attachments', [
      'AttachId', 'TaskId', 'FileName', 'FileType', 'FileUrl', 'FileId', 'ThumbnailUrl', 'UploadedBy', 'UploadedAt'
    ]);
    
    var attachId = 'ATT_' + new Date().getTime();
    
    attachSheet.appendRow([
      attachId,
      String(fileData.taskId),
      driveFileName,
      fileData.mimeType,
      fileUrl,
      fileId,
      thumbnailUrl,
      fileData.uploadedBy,
      now
    ]);
    
    // 2. Log detail in the new sheet tab "Logs"
    var logSheet = getOrCreateSheet(ss, 'Logs', [
      'Timestamp', 'TaskId', 'FileName', 'FileUrl', 'UploadedBy'
    ]);
    
    logSheet.appendRow([
      now,
      String(fileData.taskId),
      driveFileName,
      fileUrl,
      fileData.uploadedBy
    ]);
    
    return {
      success: true,
      attachment: {
        attachId: attachId,
        taskId: String(fileData.taskId),
        fileName: driveFileName,
        fileType: fileData.mimeType,
        fileUrl: fileUrl,
        fileId: fileId,
        thumbnailUrl: thumbnailUrl,
        uploadedBy: fileData.uploadedBy,
        uploadedAt: now
      }
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/**
 * Gets all attachments for a specific task.
 */
function getTaskAttachments(taskId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Attachments');
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var normalizedHeaders = normalizeHeaders(headers);
  var taskIdIndex = normalizedHeaders.indexOf('taskId');
  
  var attachments = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][taskIdIndex]) === String(taskId)) {
      var obj = {};
      headers.forEach(function(h, idx) {
        obj[normalizedHeaders[idx]] = (data[i][idx] !== null && data[i][idx] !== undefined) ? String(data[i][idx]).trim() : '';
      });
      attachments.push(obj);
    }
  }
  return attachments;
}

/**
 * Deletes an attachment from Drive and the Attachments sheet.
 */
function deleteAttachment(attachId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Attachments');
    if (!sheet) return { success: false, error: 'Attachments sheet not found' };
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var normalizedHeaders = normalizeHeaders(headers);
    var attachIdIndex = normalizedHeaders.indexOf('attachId');
    var fileIdIndex = normalizedHeaders.indexOf('fileId');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][attachIdIndex]).trim() === String(attachId).trim()) {
        var driveFileId = data[i][fileIdIndex];
        if (driveFileId) {
          try {
            DriveApp.getFileById(driveFileId).setTrashed(true);
          } catch(e) {
            // File may already be deleted
          }
        }
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Attachment not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================

/**
 * Converts sheet rows into an array of JSON objects.
 * Parses 'comments' column from JSON string to array.
 * Converts date values to ISO date strings.
 */
function normalizeHeader(header) {
  if (!header) return '';
  var withSpaces = String(header).trim().replace(/([a-z])([A-Z])/g, '$1 $2');
  var parts = withSpaces.replace(/_/g, ' ').split(/\s+/);
  return parts.map(function(part, idx) {
    var lower = part.toLowerCase();
    return idx === 0 ? lower : lower.charAt(0).toUpperCase() + lower.slice(1);
  }).join('');
}

function normalizeHeaders(headers) {
  return headers.map(normalizeHeader);
}

function getSheetDataAsJson(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const normalizedHeaders = normalizeHeaders(headers);
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      var val = row[i];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      val = (val !== null && val !== undefined) ? String(val).trim() : '';
      obj[normalizedHeaders[i]] = val;
    });

    // Parse comments from JSON string
    if (obj.comments && typeof obj.comments === 'string') {
      try {
        obj.comments = JSON.parse(obj.comments);
      } catch(e) {
        obj.comments = [];
      }
    } else if (!obj.comments) {
      obj.comments = [];
    }

    // Ensure progress is a number
    if (obj.progress !== undefined) {
      obj.progress = parseInt(obj.progress) || 0;
    }

    // Ensure id is a number
    if (obj.id !== undefined) {
      obj.id = parseInt(obj.id) || 0;
    }

    return obj;
  });
}

/**
 * Gets a sheet by name or creates it with headers.
 * Also inserts headers if the sheet exists but is missing them.
 */
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    // Sheet doesn't exist — create it with headers
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#000000')
      .setFontColor('#ffffff');
  } else {
    // Remove any leading completely empty rows
    var lastRow = sheet.getLastRow();
    var lastCol = Math.max(sheet.getLastColumn(), headers.length);
    while (lastRow > 0) {
      var checkRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var allEmpty = checkRow.every(function(cell) {
        return cell === '' || cell === null || cell === undefined;
      });
      if (allEmpty) {
        sheet.deleteRow(1);
        lastRow--;
      } else {
        break;
      }
    }
    if (lastRow === 0) {
      // All rows were empty — add headers
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#000000')
      .setFontColor('#ffffff');
    } else {
      var expectedFirst = String(headers[0]).trim().toLowerCase();
      var headerRowFound = -1;
      var allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      for (var r = 0; r < allData.length; r++) {
        if (String(allData[r][0]).trim().toLowerCase() === expectedFirst) {
          headerRowFound = r + 1; // 1-indexed
          break;
        }
      }
      if (headerRowFound === 1) {
        // Already correct — ensure bold formatting
        sheet.getRange(1, 1, 1, headers.length)
          .setFontWeight('bold')
          .setBackground('#000000')
      .setFontColor('#ffffff');
      } else if (headerRowFound > 1) {
        // Header exists at lower row — move it to row 1
        var headerValues = sheet.getRange(headerRowFound, 1, 1, headers.length).getValues()[0];
        sheet.deleteRow(headerRowFound);
        sheet.insertRowBefore(1);
        sheet.getRange(1, 1, 1, headers.length).setValues([headerValues]);
        sheet.getRange(1, 1, 1, headers.length)
          .setFontWeight('bold')
          .setBackground('#000000')
      .setFontColor('#ffffff');
      } else {
        // No header found — insert at row 1
        sheet.insertRowBefore(1);
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length)
          .setFontWeight('bold')
          .setBackground('#000000')
      .setFontColor('#ffffff');
      }
    }
  }
  return sheet;
}

// ============================================================
// 9. SETTINGS & LOGS & ALL ATTACHMENTS VIEW
// ============================================================

function getAllLogs() {
  const ss = getSpreadsheet();
  const logSheet = ss.getSheetByName('Logs');
  if (!logSheet) return [];
  return getSheetDataAsJson(logSheet);
}

function getAllAttachments() {
  const ss = getSpreadsheet();
  const attachSheet = ss.getSheetByName('Attachments');
  if (!attachSheet) return [];
  return getSheetDataAsJson(attachSheet);
}

function saveSettings(settingsObj) {
  try {
    const ss = getSpreadsheet();
    const settingsSheet = getOrCreateSheet(ss, 'Settings', ['Setting', 'Value', 'Description', 'LastModified']);
    const dataRange = settingsSheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let key in settingsObj) {
      const val = String(settingsObj[key]);
      let found = false;
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]).trim() === key) {
          settingsSheet.getRange(i + 1, 2).setValue(val);
          found = true;
          break;
        }
      }
      if (!found) {
        settingsSheet.appendRow([key, val]);
      }
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function clearDummyData() {
  try {
    const ss = getSpreadsheet();
    const userSheet = ss.getSheetByName('Users');
    const taskSheet = ss.getSheetByName('Tasks');
    const projectSheet = ss.getSheetByName('Projects');
    const logSheet = ss.getSheetByName('Logs');
    const settingsSheet = getOrCreateSheet(ss, 'Settings', ['Setting', 'Value', 'Description', 'LastModified']);

    const demoUsernames = ['rajesh','anil','priya','vikram','deepa','suresh','kavita','admin'];
    const demoTaskIds = ['1','2','3','4','5','6'];
    const demoTaskTitles = [
      'Foundation waterproofing — Block A',
      'Steel reinforcement inspection — Level 3',
      'Concrete pour scheduling — Basement 2',
      'Facade panel delivery coordination',
      'Quality audit — concrete curing',
      'Electrical panel wiring — Basement'
    ];
    const demoProjectNames = ['Skyline Towers','Green Valley Residency','Metro Bridge Overpass','Sunrise Mall'];

    function removeRows(sheet, predicate) {
      if (!sheet || sheet.getLastRow() < 2) return;
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
      const values = range.getValues();
      const kept = values.filter(function(row, index) {
        return !predicate(row, index + 2);
      });
      if (kept.length === values.length) return;
      if (kept.length > 0) {
        sheet.getRange(2, 1, kept.length, lastCol).setValues(kept);
        if (values.length > kept.length) {
          sheet.deleteRows(2 + kept.length, values.length - kept.length);
        }
      } else {
        sheet.deleteRows(2, values.length);
      }
    }

    removeRows(userSheet, function(row) {
      var username = String(row[0] || '').trim().toLowerCase();
      return demoUsernames.indexOf(username) !== -1;
    });

    removeRows(taskSheet, function(row) {
      var taskId = String(row[0] || '').trim();
      var title = String(row[1] || '').trim();
      return demoTaskIds.indexOf(taskId) !== -1 && demoTaskTitles.indexOf(title) !== -1;
    });

    if (logSheet && logSheet.getLastRow() >= 2) {
      removeRows(logSheet, function(row) {
        var taskId = String(row[1] || '').trim();
        return demoTaskIds.indexOf(taskId) !== -1;
      });
    }

    if (projectSheet && projectSheet.getLastRow() >= 2) {
      const remainingTasks = taskSheet ? getSheetDataAsJson(taskSheet) : [];
      const activeProjectNames = remainingTasks.map(function(task) { return String(task.project || '').trim(); });
      removeRows(projectSheet, function(row) {
        var projectName = String(row[0] || '').trim();
        return demoProjectNames.indexOf(projectName) !== -1 && activeProjectNames.indexOf(projectName) === -1;
      });
    }

    saveSettings({ disableDemoSeed: 'true' });
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/**
 * Checks the header row of a sheet and appends missing column headers if necessary.
 */
function ensureHeadersExist(sheet, expectedHeaders) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headersLower = headers.map(function(h) { return String(h).trim().toLowerCase(); });
  
  var added = false;
  expectedHeaders.forEach(function(h) {
    if (headersLower.indexOf(h.toLowerCase()) === -1) {
      sheet.getRange(1, lastCol + 1).setValue(h)
        .setFontWeight('bold')
        .setBackground('#000000')
      .setFontColor('#ffffff');
      lastCol++;
      added = true;
    }
  });
  if (added) {
    SpreadsheetApp.flush();
  }
}

function addUser(userObj) {
  try {
    const ss = getSpreadsheet();
    const userSheet = ss.getSheetByName('Users');
    if (!userSheet) throw new Error('Users sheet not found');
    
    const data = userSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userObj.username) {
        return { success: false, error: 'Username already exists' };
      }
    }
    
    var photoUrl = userObj.photo || '';
    if (userObj.photoData) {
      try {
        var folder = getUploadsFolder();
        var decoded = Utilities.base64Decode(userObj.photoData.base64);
        var blob = Utilities.newBlob(decoded, userObj.photoData.mimeType, userObj.photoData.fileName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=s200';
      } catch(e) {
        photoUrl = '';
      }
    }
    
    userSheet.appendRow([
      userObj.username,
      userObj.password,
      userObj.name,
      userObj.role,
      userObj.initials,
      userObj.color || '#3b82f6',
      userObj.department || '',
      userObj.manager || '',
      userObj.project || '',
      userObj.email || '',
      userObj.mobile || '',
      photoUrl
    ]);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function editUser(userObj) {
  try {
    const ss = getSpreadsheet();
    const userSheet = ss.getSheetByName('Users');
    if (!userSheet) throw new Error('Users sheet not found');

    const data = userSheet.getDataRange().getValues();
    var headers = data[0];
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(userObj.username).trim()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: 'User not found' };
    }

    var photoUrl = userObj.photo || '';
    if (userObj.photoData) {
      try {
        var folder = getUploadsFolder();
        var decoded = Utilities.base64Decode(userObj.photoData.base64);
        var blob = Utilities.newBlob(decoded, userObj.photoData.mimeType, userObj.photoData.fileName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=s200';
      } catch(e) {
        photoUrl = userObj.photo || '';
      }
    }

    userSheet.getRange(rowIndex, 3, 1, 10).setValues([[
      userObj.name || '',
      userObj.role || '',
      userObj.initials || '',
      userObj.color || '#d97706',
      userObj.department || '',
      userObj.manager || '',
      userObj.project || '',
      userObj.email || '',
      userObj.mobile || '',
      photoUrl
    ]]);

    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function changeUserPassword(username, newPassword) {
  try {
    const ss = getSpreadsheet();
    const userSheet = ss.getSheetByName('Users');
    if (!userSheet) throw new Error('Users sheet not found');
    
    const data = userSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        userSheet.getRange(i + 1, 2).setValue(newPassword);
        found = true;
        break;
      }
    }
    if (!found) return { success: false, error: 'User not found' };
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function deleteUser(username) {
  try {
    const ss = getSpreadsheet();
    const userSheet = ss.getSheetByName('Users');
    if (!userSheet) throw new Error('Users sheet not found');
    
    const data = userSheet.getDataRange().getValues();
    let foundIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        foundIndex = i + 1;
        break;
      }
    }
    if (foundIndex === -1) return { success: false, error: 'User not found' };
    userSheet.deleteRow(foundIndex);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}
// ============================================================
// LOG SESSION ACTION
// ============================================================
function logSessionAction(username, action, details) {
  try {
    const ss = getSpreadsheet();
    let sessionSheet = ss.getSheetByName('Sessions');
    if (!sessionSheet) {
      sessionSheet = ss.insertSheet('Sessions');
      sessionSheet.appendRow(['Timestamp', 'Username', 'Action', 'Details']);
    }
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    sessionSheet.appendRow([timestamp, username, action, details || '']);
  } catch (e) {
    console.error("Failed to log session action:", e);
  }
}
