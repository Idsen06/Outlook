Office.onReady(() => {
  console.log("✅ Office Add-in is ready!"); // Debugging log

  let button = document.getElementById("exportButton");

  if (button) {
      console.log("✅ Button found in DOM.");
      button.addEventListener("click", () => {
          console.log("🟢 Button clicked!"); // Debugging log
          exportEmail();
      });
  } else {
      console.error("❌ Button not found in the DOM.");
  }
});

function exportEmail() {
  console.log("🟡 Starting email export...");

  if (!Office.context.mailbox || !Office.context.mailbox.item) {
      console.error("❌ No email selected or Office.js not available.");
      document.getElementById("statusMessage").textContent = "No email selected!";
      return;
  }

  let emailItem = Office.context.mailbox.item;
  console.log("📩 Fetching email details for:", emailItem.subject);

  emailItem.body.getAsync(Office.CoercionType.Text, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("✅ Email body retrieved successfully.");

          let emailData = {
              subject: emailItem.subject || "No Subject",
              sender: emailItem.from ? `${emailItem.from.displayName} <${emailItem.from.emailAddress}>` : "Unknown Sender",
              body: result.value || "No Body Content",
              receivedTime: emailItem.dateTimeCreated ? emailItem.dateTimeCreated.toISOString() : "Unknown Time"
          };

          console.log("📜 Email Data:", emailData);
          saveAsJson(emailData);
      } else {
          console.error("❌ Failed to retrieve email body:", result.error);
          document.getElementById("statusMessage").textContent = "Failed to retrieve email.";
      }
  });
}

function saveAsJson(data) {
  console.log("💾 Saving email as JSON...");

  try {
      let jsonString = JSON.stringify(data, null, 2);
      let blob = new Blob([jsonString], { type: "application/json" });

      let a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "email.json";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);

      document.getElementById("statusMessage").textContent = "✅ Email exported successfully!";
      console.log("🎉 Email exported successfully!");
  } catch (error) {
      console.error("❌ Error saving JSON file:", error);
      document.getElementById("statusMessage").textContent = "Failed to save email as JSON.";
  }
}
