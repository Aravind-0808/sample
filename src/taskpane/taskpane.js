Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("email-details").style.display = "none";
    getEmailDetails();
  }
});

// Fetch Email Details
async function getEmailDetails() {
  try {
    const item = Office.context.mailbox.item;
    const subject = item.subject;

    item.body.getAsync("html", function (result) {
      if (result.status === "succeeded") {
        const body = result.value;
        document.getElementById("email-subject").textContent = subject;
        document.getElementById("email-body").innerHTML = body;
        document.getElementById("email-details").style.display = "block";
      } else {
        console.error("Error fetching email body:", result.error);
      }
    });
  } catch (error) {
    console.error("Error fetching email details:", error);
  }
}

// Convert email content to EML format and save (WITH ATTACHMENTS)
async function saveEmail() {
    const subject = document.getElementById("email-subject").textContent.trim();
    const body = document.getElementById("email-body").innerHTML;
    const projectName = document.getElementById("project-name").value.trim();
    const cleanedSubject = subject.replace(/[^a-zA-Z0-9\s]/g, ""); // Sanitize subject for filename
    const statusMessage = document.getElementById("status-message"); // Get status message element
  
    statusMessage.style.display = "none"; // Hide previous messages
  
    if (!projectName) {
      statusMessage.textContent = "Please enter a project name.";
      statusMessage.style.color = "red";
      statusMessage.style.display = "block";
      return;
    }
  
    try {
      const directoryHandle = await window.showDirectoryPicker();
      let projectFolder;
  
      try {
        projectFolder = await directoryHandle.getDirectoryHandle(projectName);
      } catch {
        projectFolder = await directoryHandle.getDirectoryHandle(projectName, { create: true });
      }
  
      // **Check if file already exists**
      for await (const file of projectFolder.values()) {
        if (file.kind === "file" && file.name === `${cleanedSubject}.eml`) {
          statusMessage.textContent = "The email is already saved in this Project !";
          statusMessage.style.color = "red";
          statusMessage.style.display = "block"; // Show warning message
          return;
        }
      }
  
      // Fetch Email Details
      const item = Office.context.mailbox.item;
      const emailDate = new Date().toUTCString();
      const from = item.sender ? item.sender.emailAddress : "unknown@domain.com";
      const to = item.to ? item.to.map(t => t.emailAddress).join(", ") : "unknown@domain.com";
  
      let attachmentParts = "";
      let boundary = "----=_NextPart_000_001"; // Unique boundary for MIME parts
  
      // Fetch and convert attachments
      if (item.attachments && item.attachments.length > 0) {
        await Promise.all(
          item.attachments.map(async (attachment) => {
            if (attachment.isInline) return; // Ignore inline images
  
            return new Promise((resolve, reject) => {
              item.getAttachmentContentAsync(attachment.id, async function (result) {
                if (result.status === "succeeded") {
                  try {
                    const base64Data = result.value.content;
  
                    attachmentParts += `
  --${boundary}
  Content-Type: ${attachment.contentType}; name="${attachment.name}"
  Content-Transfer-Encoding: base64
  Content-Disposition: attachment; filename="${attachment.name}"
  
  ${base64Data}
  `;
  
                    resolve();
                  } catch (err) {
                    console.error(`Error processing attachment ${attachment.name}:`, err);
                    reject(err);
                  }
                } else {
                  console.error("Error fetching attachment:", result.error);
                  reject(result.error);
                }
              });
            });
          })
        );
      }
  
      // Construct EML format with attachments
      const emlContent = `From: ${from}
  To: ${to}
  Subject: ${subject}
  Date: ${emailDate}
  MIME-Version: 1.0
  Content-Type: multipart/mixed; boundary="${boundary}"
  
  --${boundary}
  Content-Type: text/html; charset=UTF-8
  Content-Transfer-Encoding: quoted-printable
  
  ${body}
  
  ${attachmentParts}
  --${boundary}--`;
  
      // **Save EML File**
      const fileHandle = await projectFolder.getFileHandle(`${cleanedSubject}.eml`, { create: true });
      const writableStream = await fileHandle.createWritable();
      await writableStream.write(emlContent);
      await writableStream.close();
  
      // **Show success message in the p tag**
      statusMessage.textContent = "Email saved successfully";
      statusMessage.style.color = "green";
      statusMessage.style.display = "block";
  
    } catch (error) {
      console.error("Error saving email:", error);
      statusMessage.textContent = "An error occurred while saving the email.";
      statusMessage.style.color = "red";
      statusMessage.style.display = "block";
    }
  }
  
  


// List Saved Emails
async function listSavedEmails() {
  const historyTableBody = document.getElementById("history-table-body");
  document.getElementById("email-history").style.display = "block";
  historyTableBody.innerHTML = "";

  try {
    const directoryHandle = await window.showDirectoryPicker();

    for await (const entry of directoryHandle.values()) {
      if (entry.kind === "directory") {
        const projectFolder = entry;

        for await (const file of projectFolder.values()) {
          if (file.kind === "file" && file.name.endsWith(".eml")) {
            const row = document.createElement("tr");

            const projectCell = document.createElement("td");
            projectCell.textContent = projectFolder.name;

            const fileCell = document.createElement("td");
            fileCell.textContent = file.name;

            row.appendChild(projectCell);
            row.appendChild(fileCell);
            historyTableBody.appendChild(row);
          }
        }
      }
    }
  } catch (error) {
    console.error("Error listing saved emails:", error);
  }
}
