const { sendPrompt } = require('../phishingDetectionUsingGPT');

Office.onReady(() => {
  console.log("Office Add-in is ready");

  let button = document.getElementById("exportButton");

  if (button) {
      console.log("Button found in DOM");
      button.addEventListener("click", () => {
          console.log("Button clicked");
          exportEmail();
      });
  } else {
      console.error("Button not found in the DOM");
  }

  // Load and display custom properties when the task pane is opened
  //loadCustomProperties();
});

function exportEmail() {
  console.log("Starting email export...");

  if (!Office.context.mailbox || !Office.context.mailbox.item) {
      console.error("No email selected or Office.js not available");
      document.getElementById("statusMessage").textContent = "No email selected";
      return;
  }

  let emailItem = Office.context.mailbox.item;
  console.log("Fetching email details for:", emailItem.subject);

  emailItem.body.getAsync(Office.CoercionType.Html, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Email body retrieved successfully");

          let emailData = {
              subject: emailItem.subject || "No Subject",
              sender: emailItem.from ? `${emailItem.from.displayName} <${emailItem.from.emailAddress}>` : "Unknown Sender",
              body: result.value || "No Body Content",
              receivedTime: emailItem.dateTimeCreated ? emailItem.dateTimeCreated.toISOString() : "Unknown Time"
          };

          console.log("email Data:", emailData);
          saveAsJson(emailData);
      } else {
          console.error("Failed to retrieve email body:", result.error);
          document.getElementById("statusMessage").textContent = "Failed to retrieve email.";
      }
  });
}

let savedJsonData;

function saveAsJson(data) {
  console.log("Saving email as JSON...");

  try {
      let jsonString = JSON.stringify(data, null, 2);
      savedJsonData = jsonString;

      document.getElementById("statusMessage").textContent = "Detecting phishing...";
      console.log("email exported successfully!");
      console.log("saved JSON Data:", savedJsonData);

      let emailBody = cleanseEmailBody(data.body);

      console.log("cleaned Email Body:", emailBody); //log the cleaned email body for debugging (can be removed later)

      //contact GPT
      sendPrompt(emailBody).then(response => {
          console.log("Phishing Detection Response:", response);
          //document.getElementById("statusMessage").textContent = "Detecting phishing...";


          //phishing detection response in task pane TODO: if possible, modify email to display it?
          //document.getElementById("statusMessage").textContent = `Phishing Detection Response: ${response}`;
          if (response.is_phishing) {
              document.getElementById("statusMessage").textContent = "Phishing Detected.";
          }
          else { document.getElementById("statusMessage").textContent = "No Phishing Detected"; }

          document.getElementById("percentage").textContent += `Likelihood for this to be phishing: ${response.likelihood_percentage}%`;
          document.getElementById("explanation").textContent += `Reason: `;
          document.getElementById("reason").textContent += `${response.explanation}`;

          //add a custom property to the email item (IN THE WORKS FOR SORTING IN THE APPLICATION)
          //addCustomProperty(response);

      }).catch(error => {
          console.error("Error in phishing detection:", error);
          document.getElementById("statusMessage").textContent = `Phishing Detection Failed.`;
      });

  } catch (error) {
      console.error("Error saving JSON file:", error);
      document.getElementById("statusMessage").textContent = "Failed to save email as JSON.";
  }
}

function cleanseEmailBody(htmlBody) {
    //temporary DOM element to parse HTML
    const tempDiv = document.createElement("div");
    tempDiv.innerHTML = htmlBody;

    //remove <style> <script> tags
    const styleAndScriptTags = tempDiv.querySelectorAll("style, script");
    styleAndScriptTags.forEach(tag => tag.remove());

    //remove anchor tags (links) but keep their text content
    const anchorTags = tempDiv.querySelectorAll("a");
    anchorTags.forEach(anchor => {
        const textNode = document.createTextNode(anchor.textContent);
        anchor.replaceWith(textNode);
    });

    const plainText = tempDiv.textContent || tempDiv.innerText || "";

    //remove any remaining HTML entities - &nbsp; and &amp;
    const decodedText = plainText.replace(/&nbsp;/g, " ").replace(/&amp;/g, "&");

    return decodedText.trim(); //also trims all extra spaces
}

//functions i will get back to in order to sort emails based on the phishing detection response
/*
function addCustomProperty(phishingResponse) {
  Office.context.mailbox.item.loadCustomPropertiesAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          let customProps = result.value;
          console.log("Custom properties: ", customProps);
          
          customProps.set("PhishingDetectionResponse", phishingResponse);
          customProps.saveAsync(function (asyncResult) {
              if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log("Custom property saved successfully.");
                  // Move the email based on the custom property
                  moveEmailBasedOnCustomProperty(phishingResponse);
              } else {
                  console.error("Failed to save custom property:", asyncResult.error);
              }
          });
      } else {
          console.error("Failed to load custom properties:", result.error);
      }
  });
}
  */
/*
function loadCustomProperties() {
  Office.context.mailbox.item.loadCustomPropertiesAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          let customProps = result.value;
          let phishingResponse = customProps.get("PhishingDetectionResponse");
          if (phishingResponse) {
              document.getElementById("statusMessage").textContent += `\nLoaded Phishing Detection Response: ${phishingResponse}`;
              console.log("Loaded Phishing Detection Response:", phishingResponse);
          } else {
              console.log("No Phishing Detection Response found.");
          }
      } else {
          console.error("Failed to load custom properties:", result.error);
      }
  });
}

function moveEmailBasedOnCustomProperty(phishingResponse) {
  // Define the folder ID where you want to move the email
  const folderId = phishingResponse.includes("phishing") ? "phishing-folder-id" : "safe-folder-id";

  Office.context.mailbox.item.moveAsync(folderId, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Email moved successfully.");
          document.getElementById("statusMessage").textContent += "\nEmail moved successfully.";
      } else {
          console.error("Failed to move email:", result.error);
          document.getElementById("statusMessage").textContent += "\nFailed to move email.";
      }
  });
  
}*/
