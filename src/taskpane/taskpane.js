/* global Office, document, Blob, URL, window, console */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Attach the export function to the button
    document.getElementById("exportButton").onclick = exportEmailAsJS;
  }
});

function exportEmailAsJS() {
  const item = Office.context.mailbox.item;

  if (!item) {
    showMessage("No email item selected.");
    return;
  }

  // Get the email body as plain text
  item.body.getAsync("text", (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailContent = result.value;
      const fileName = `${item.subject || "email"}.js`;

      // Create a blob and download it as a .js file
      const blob = new Blob([`const emailContent = \`${emailContent}\`;`], { type: "application/javascript" });
      const url = URL.createObjectURL(blob);

      const a = document.createElement("a");
      a.style.display = "none";
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();

      window.URL.revokeObjectURL(url);
      showMessage(`Successfully exported as ${fileName}`);
    } else {
      showMessage("Failed to get email content.");
      console.error("Error:", result.error);
    }
  });
}

// Display a status message to the user
function showMessage(message) {
  const statusElement = document.getElementById("statusMessage");
  statusElement.textContent = message;
}
