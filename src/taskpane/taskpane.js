/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = run;
    document.getElementById("saveNote").onclick = saveNote;
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;
  let insertAt = document.getElementById("item-subject");
  let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject));
  insertAt.appendChild(document.createElement("br"));
}



async function saveNote() {
  const text = document.getElementById("noteText").value;
  const statusDiv = document.getElementById("status");

  if (!text.trim()) {
    statusDiv.textContent = "Please enter some text!";
    statusDiv.style.color = "red";
    return;
  }

  try {
    const API_BASE = window.location.origin + '/api';

    const item = Office.context.mailbox.item;
    const userEmail = Office.context.mailbox?.userProfile?.emailAddress || 'anonymous';
    const subject = item.subject; // simpler: item.subject is available
    const itemId = item.itemId;

    const response = await fetch(`${API_BASE}/save-note`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        note: text,
        userEmail,
        emailSubject: subject,
        emailId: itemId
      })
    });

    const result = await response.json();

    if (result.success) {
      statusDiv.textContent = "✓ Saved successfully!";
      statusDiv.style.color = "green";
      document.getElementById("noteText").value = '';
      loadNotes();
    } else {
      statusDiv.textContent = "Save failed: " + (result.error || 'Unknown');
      statusDiv.style.color = "red";
    }
  } catch (error) {
    statusDiv.textContent = "Error: " + error.message;
    statusDiv.style.color = "red";
    console.error(error);
  }
}

async function loadNotes() {
  try {
    const API_BASE = window.location.origin + '/api';
    const response = await fetch(`${API_BASE}/notes`);
    const notes = await response.json();
    console.log('Your notes:', notes);
  } catch (error) {
    console.error('Load failed:', error);
  }
}
