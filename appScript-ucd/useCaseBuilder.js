// Global Scope Variables
var ucdFileId = ''
const allSections = [
  {
    SectionHeader: "Configure Role Based Incident Assignment",
    Type: "Playbook Input",
    Purpose: "The default role to assign the incident too. This role must exist in XSOARs RBAC setting. To configure this feature review the following section from out admin guide https://docs-cortex.paloaltonetworks.com/r/Cortex-XSOAR/8/Cortex-XSOAR-Cloud-Documentation/Manage-roles-in-the-Cortex-XSOAR-tenant",
    Input: "inputs.Role: Role Name",
    Command: "N/A",
  },
  {
    SectionHeader: "Configure On-Call Assignment",
    Type: "Playbook Input",
    Purpose: "When set to True this configures the playbook to only assign the incident to a user that is currently on shift. Read about setting up Shift Management in XSOAR from this link in the admin guide, https://docs-cortex.paloaltonetworks.com/r/Cortex-XSOAR/8/Cortex-XSOAR-Cloud-Documentation/Manage-roles-in-the-Cortex-XSOAR-tenant",
    Input: "inputs.OnCall: True/False",
    Command: "N/A",
  },
  {
    SectionHeader: "Email Hunting Create New Incidents",
    Type: "Playbook Input",
    Purpose: "When set to True the sub-playbook, Phishing - Handle Microsoft 365 Defender Results will open new incidents if indicators from this Phishing incident are discovered on additional endpoints.",
    Input: "inputs.EmailHuntingCreateNewIncidents: True/False",
    Command: "N/A",
  },
  {
    SectionHeader: "Set Listener MailBox",
    Type: "Playbook Input",
    Purpose: "The mailbox which is being used to fetch phishing incidents. This mailbox will be excluded in the Phishing - Indicators Hunting playbook. In case the value of this input is empty, the value of the Email To incident field will be automatically used as the listener mailbox.",
    Input: "inputs.ListenerMailbox: Phishing Mailbox Name",
    Command: "N/A",
  },
  {
    SectionHeader: "Set Send Mail Instance",
    Type: "Playbook Input", 
    Purpose: "The name of the instance to be used when executing the send-mail command in the playbook. In case it will be empty, all available instances will be used (default). This can lead to multiple emails being sent if multiple mail instances are configured.",
    Input: "inputs.SendMailInstance: Instance Name",
    Command: "N/A",
  },
  {
    SectionHeader: "Set End User Engagement",
    Type: "Playbook Input",
    Purpose: "Specify whether to engage with the user via email for investigation updates.Set the value to 'True' to allow user engagement, or 'False' to avoid user engagement.",
    Input: "inputs.EndUserEngagement: True/False",
    Command: "N/A"
  },
  {
    SectionHeader: "Set Take Manual Actions",
    Type: "Playbook Input",
    Purpose: "Specify whether to stop the playbook to take additional action before closing the incident. Set the value to 'True' to stop the playbook before closing the incidents, or 'False' to close the incident once the playbook flow is done.",
    Input: "inputs.TakeManualActions: True/False",
    Command: "N/A",
  },
  {
    SectionHeader: "Non Malicious Email Review",
    Type: "Playbook Input",
    Purpose: "In case of an incident with a non-malicious email, it is possible either to close the incident or to review and approve it by an analyst.Set the value 'True' for a review of the incident by an analyst. Set the value 'False' to close the incident.",
    Input: "inputs.NonMaliciousEmailReview",
    Command: "N/A",
  },
  {
    SectionHeader: "Detonate URL",
    Type: "Playbook Input",
    Purpose: "Determines whether to use the ‘URL Detonation’ playbook. Detonating a URL may take a few minutes. This playbook will leverage third party integrations to perform the sandboxing. Make sure these integrations are configured in the customer environment if they use URL detonation services.",
    Input: "inputs.DetonateUrl: True/False",
    Command: "N/A",
  },
  {
    SectionHeader: "Internal Range",
    Type: "Playbook Input",
    Purpose: "This input is used in the ‘Entity Enrichment - Phishing v2’ playbook.A list of internal IP ranges to check IP addresses against. The comma-separated list should be provided in CIDR notation. For example, a list of ranges is: '172.16.0.0/12,10.0.0.0/8,192.168.0.0/16' (without quotes). The default list name is called PrivateIPs. This list would need to exist in XSOAR before it is able to be leveraged in the playbook. Here is a link that provides instructions on how to create a list https://docs-cortex.paloaltonetworks.com/r/Cortex-XSOAR/8/Cortex-XSOAR-Cloud-Documentation/Create-a-list",
    Input: "inputs.InternalRange: lists.PrivateIPs",
    Command: "N/A",
  },
  {
    SectionHeader: "Internal Domains",
    Type: "Playbook Input",
    Purpose: "A CSV list of internal domains. The list is used to determine whether an email address is internal or external. This list would need to exist in XSOAR before it is able to be leveraged in the playbook. Here is a link that provides instructions on how to create a list https://docs-cortex.paloaltonetworks.com/r/Cortex-XSOAR/8/Cortex-XSOAR-Cloud-Documentation/Create-a-list",
    Input: "inputs.InternalDomains: lists.InternalDomains",
    Command: "N/A",
  },
  {
    SectionHeader: "Authenticate Email",
    Type: "Playbook Input",
    Purpose: " Determines whether the authenticity of the email should be verified using SPF, DKIM, and DMARC. SPF ensures that the email is being sent from an authorized mail server for the given domain. DKIM uses cryptographic signatures to ensure that the email contents were not altered in transit. DMARC allows domain owners to define how to handle emails that fail SPF or DKIM (reject, quarantine, or allow).",
    Input: "inputs.AuthenticateEmail: True/False",
    Command: "N/A",
  },
  {
    SectionHeader: "Check Microsoft Headers",
    Type: "Playbook Input",
    Purpose: "Whether to check Microsoft headers for BCL/PCL/SCL scores and set the 'Severity' and 'Email Classification' accordingly.BCL (Bulk Containment Level) measures how likely an email is a bulk (mass) email on a scale of 0 - 9. 0-3 likely legitimate bulk email, 4-7, questionable bulk email (possible spam), and 8-9 High-Risk bulk email (more likely to be junk or spam). PCL (Phishing Confidence Level) determines the likelihood of an email being phishing on a scale of 1 - 8. 1-3 is likely safe (low risk), 4-5 is suspicious (requires manual review), and 6-8 is a high confidence score that the email is phishing and should be blocked or quarantined. SCL (Spam Confidence Level) assigns a spam likelihood score to an email on a scale of -1 - 9. -1 is a safe email (whitelisted), 0-1 is a normal email (not spam), 5-6 likely spam (goes to junk), and 7-9 is High confidence spam (blocked or quarantined).",
    Input: "inputs.CheckMicrosoftHeaders: True/False",
    Command: "N/A",
  },
  {
    SectionHeader: "Get Original Email",
    Type: "Playbook Input",
    Purpose: "For forwarded emails. When 'True', retrieves the original email in the thread.You must have the necessary permissions in your email service to execute global search.- For EWS: eDiscovery - For Gmail: Google Apps Domain-Wide Delegation of Authority - For MSGraph: As described in these links https://docs.microsoft.com/en-us/graph/api/message-get, https://docs.microsoft.com/en-us/graph/api/user-list-messages",
    Input: "inputs.GetORiginalEmail: True/False",
    Command: "N/A",
  },
  {
    SectionHeader: "Original Authentication Header",
    Type: "Playbook Input",
    Purpose: "This input will be used as the 'original_authentication_header' argument in the 'CheckEmailAuthenticity' script under the 'Authenticate email' task.The header that holds the original Authentication-Results header value. This can be used when an intermediate server changes the original email and holds the original header value in a different header. Note - Use this only if you trust the server creating this header.",
    Input: "inputs.OriginalAuthenticationHeader: True/False",
    Command: "N/A",
  },
  {
    SectionHeader: "Phishing Model Name",
    Type: "Playbook Input",
    Purpose: "Optional - the name of a pre-trained phishing model to predict phishing type using machine learning.",
    Input: "inputs.PhishingModelName: Phishing Model",
    Command: "N/A",
  }
]

const integrationInstructions = [
  {
    Name: "Cisco Secure Malware Analytics (Threat Grid) v2",
    ContentPack: "Cisco Secure Malware Analytics",
    Actions: "Detonate URL - ThreatGrid v2 Playbook",
    Purpose: "Detonate URLs as part of the Detonate URL - Generic v1.5 sub-playbook in the Phishing Generic v3 Playbook.",
    Url: "https://xsoar.pan.dev/docs/reference/integrations/threat-gridv2",
    Implement: [
      "1. Install the Cisco Secure Malware Analytics content pack from the Marketplace",
      "2. Navigate to the Cisco Secure Malware Analytics (Threat Grid) v2 integration and click 'Add Instance'",
      "3. Follow the instructions in the integration guide to configure the instance.",
      "4. Test the instance and save it."
    ],
    Remove: [
      "1. Make a copy of the Detonate URL - Generic v1.5 playbook. If you already have one for the phishing playbook modify that copy.",
      "2. Remove task # 24 from the playbook copy.",
      "3. Make sure the copy of the phishing playbook is using the newly modified copy of the Detonate URL - Generic v1.5 playbook."
    ],
  },
  {
    Name: "Joe Security v2",
    ContentPack: "Joe Security",
    Actions: "joe-submit-url script",
    Purpose: "Detonate URLs as part of the Detonate URL - Generic v1.5 sub-playbook in the Phishing Generic v3 Playbook.",
    Url: "https://xsoar.pan.dev/docs/reference/integrations/joe-security-v2",
    Implement: [
      "1. Install the Joe Security content pack from the Marketplace",
      "2. Navigate to the Joe Security v2 integration and click 'Add Instance'",
      "3. Follow the instructions in the integration guide to configure the instance.",
      "4. Test the instance and save it."
    ],
    Remove: [
      "1. Make a copy of the Detonate URL - Generic v1.5 playbook. If you already have one for the phishing playbook modify that copy.",
      "2. Remove task # 23 from the playbook copy.",
      "3. Make sure the copy of the phishing playbook is using the newly modified copy of the Detonate URL - Generic v1.5 playbook."
    ],
  },
]

function onOpen() {
  // Creates the AutoFill UCD menu when the spreadsheet
  // is opened

  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Use Case Tools');
  menu.addItem('Create Use Case', 'createUcd');
  menu.addToUi();
}

function createUcd() {
  // Creates the copy of the ucd template and adds customer contact info.

  const useCaseTemplate = DriveApp.getFileById('1ELp3JTOjJzs8zoxpWf-FEmkOh0qwDQvKtOCFcg6uN04');
  const destinationFolder = DriveApp.getFolderById('1VLnx-aeGKTHU6fg_L1RDW5gogXWolenS');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customer_Info');
  const rows = sheet.getDataRange().getValues();
  const row = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  const customerName = row[0];
  const copy = useCaseTemplate.makeCopy(`${customerName} Phishing v3 Use Case`, destinationFolder);
  const doc = DocumentApp.openById(copy.getId());
  
  const body = doc.getBody();
  ucdFileId = copy.getId();

  rows.forEach(function(row, index){
    if (index === 0) return;
    body.replaceText('{{FIRST}}', row[1]);
    body.replaceText('{{LAST}}', row[2]);
    body.replaceText('{{EMAIL}}', row[3]);
  })

  useCaseInstructions();
}

function useCaseInstructions() {
  const playbookSections = getSectionsToImplement();
  const integrationSections = getEnabledIntegrationSections();
  const sections = [...integrationSections, ...playbookSections];
  const doc = DocumentApp.openById(ucdFileId);
  const body = doc.getBody();

  // Handle {{SECTION_START}} placeholder
  const placeholder = "{{SECTION_START}}";
  const firstMatch = body.findText(placeholder);
  if (!firstMatch) throw new Error("Placeholder not found.");
  const placeholderElement = firstMatch.getElement().getParent();
  const insertIndex = body.getChildIndex(placeholderElement);
  body.removeChild(placeholderElement);

  // Insert integration table at placeholder {{INTEGRATIONS_TABLE}}
  insertIntegrationsAtPlaceholder(body, "{{INTEGRATIONS_TABLE}}");

  // Begin inserting sections
  let currentIndex = insertIndex;
  let sectionCounter = 1;

  // 🔹 Insert Integration Sections First
  if (integrationSections.length > 0) {
    const integrationHeader = body.insertParagraph(currentIndex++, "Integrations");
    integrationHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);

    integrationSections.forEach(section => {
      currentIndex = insertFormattedSection(body, currentIndex, section, sectionCounter++);
    });
  }

  // 🔹 Then Insert Playbook Input Sections
  if (playbookSections.length > 0) {
    const playbookHeader = body.insertParagraph(currentIndex++, "Playbook Configuration");
    playbookHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);

    playbookSections.forEach(section => {
      currentIndex = insertFormattedSection(body, currentIndex, section, sectionCounter++);
    });
  }
  // 🔹 Then Create Project Tracker
  appendToProjectTracker(sections);
  doc.saveAndClose();
}

function insertFormattedSection(body, index, section, sectionNumber) {
  //Heading 3
  const numberedTitle = `${sectionNumber}. ${section.SectionHeader}`;
  const header = body.insertParagraph(index++, numberedTitle);
  header.setHeading(DocumentApp.ParagraphHeading.HEADING4)

  // Each KVP as bold key + normal value
  if (section.Type) index = InsertBoldLabelParagraph(body, index, "Type", section.Type);
  if (section.Purpose) index = InsertBoldLabelParagraph(body, index, "Purpose", section.Purpose);
  if (section.Input) index = InsertBoldLabelParagraph(body, index, "Input", section.Input);
  if (section.Command) index = InsertBoldLabelParagraph(body, index, "Command", section.Command);
  if (section.Steps && Array.isArray(section.Steps)) {
    section.Steps.forEach(step => {
      body.insertParagraph(index++, "\t" + step);
    });
  }

  
  body.insertParagraph(index++, "");

  return index;
}

function InsertBoldLabelParagraph(body, index, label, value) {
  const fullText = `\t${label}: ${value}`;
  const paragraph = body.insertParagraph(index, fullText);
  paragraph.setIndentStart(36);
  const text = paragraph.editAsText();
  text.setBold(0, label.length + 0, true);
  return index + 1;
}

function getEnabledFeaturesFromSheet() {
   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('UCD_Features');
   const data = sheet.getDataRange().getValues();
   const headers = data[0];
   const rows = data.slice(1);
   const enabledfeatures = [];

   for (const row of rows) {
    const rowData = Object.fromEntries(headers.map((key, i) => [key, row[i]]));

    if (String(rowData["Implement"]).toLowerCase() === "yes") {
      const name = rowData["key"].trim().toLowerCase();
      Logger.log("Enabled feature from sheet:" + name);
      enabledfeatures.push(name);
    }
   }
   return enabledfeatures;
}

function getSectionsToImplement() {
  const enabledfeaturesNames = getEnabledFeaturesFromSheet();
  Logger.log("Enabled features: " + JSON.stringify(enabledfeaturesNames));

  const matched = allSections.filter(section =>{
    const name = section.SectionHeader.trim().toLowerCase();
    const match = enabledfeaturesNames.includes(name);
    Logger.log(`Checking "${name}" -> ${match}`);
    return match; 
  });

  Logger.log("Matched sections count" + matched.length);
  return matched;
}

function getEnabledIntegrationsFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Integrations");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const enabled = [];

  for (const row of rows) {
    const rowData = {};
    for (let i = 0; i < headers.length; i ++) {
      rowData[headers[i]] = row[i];
    }

    const integrationName = rowData["Integration Name"];
    const implementFlag = String(rowData["Implement"] || "").toLowerCase();

    if (implementFlag === "yes") {
      enabled.push(integrationName.trim().toLowerCase());
    }
  }

  return enabled;
}

function getIntegrationsToDocument() {
  const enabledNames = getEnabledIntegrationsFromSheet();
  Logger.log("Enabled (normalized): " + JSON.stringify(enabledNames));

  return integrationInstructions.filter(instr =>
    enabledNames.includes(instr.Name.trim().toLowerCase())
  );
}

function insertIntegrationsAtPlaceholder(body, placeholder) {
  const integrations = getIntegrationsToDocument();
  if (integrations.length === 0) return index;

  const match = body.findText(placeholder);
  if (!match) {
    Logger.log("Placeholder not found:" + placeholder);
    return;
  }

  const element = match.getElement();
  const parent = element.getParent();
  const index = body.getChildIndex(parent);

  body.removeChild(parent);

  const heading = body.insertParagraph(index, "Integrations Required");
  heading.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  const table = body.insertTable(index + 1);
  const headerRow = table.appendTableRow();
  headerRow.appendTableCell("Product Name").setBold(true);
  headerRow.appendTableCell("Actions Needed").setBold(true);

  integrations.forEach(integration => {
    const row = table.appendTableRow();
    row.appendTableCell(integration.Name || "Unknown Product");
    row.setBold(false);
    const cell = row.appendTableCell();
    const implTitle = cell.appendParagraph(integration.Actions);
    implTitle.setBold(false);
  });
}

function getEnabledIntegrationSections() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Integrations");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const output = [];

  for (const row of rows) {
    const rowData = {};
    for (let i = 0; i < headers.length; i++) {
      rowData[headers[i]] = row[i];
    }

    const name = (rowData["Integration Name"] || "").trim().toLowerCase();
    const implementFlag = (rowData["Implement"] || "").trim().toLowerCase();

    if (!name || (implementFlag!== "yes" && implementFlag !== "no")) continue;

    const integration = integrationInstructions.find(i => i.Name.toLowerCase() == name);
    if (!integration) continue;

    const actionType = implementFlag === "yes" ? "Implementation Steps" : "Removal Steps"
    const steps = implementFlag === "yes" ? integration.Implement : integration.Remove;

    output.push({
      SectionHeader: `${integration.Name}: ${actionType}`,
      Type: "Integration",
      Purpose: `${integration.Purpose}`,
      Input: "N/A",
      Command: "N/A",
      Steps: steps
    });
  }

  return output;
}

function appendToProjectTracker(sections) {
  const trackerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Project_Tracker");
  const featuresSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UCD_Features");
  if (!trackerSheet || !featuresSheet) throw new Error("Project_Tracker or UCD_Features sheet not found.");

  const headers = [
    "Feature",
    "Use Case Document Step",
    "OOTB Feature",
    "Custom Feature",
    "Implemented",
    "Description"
  ];

  if (trackerSheet.getLastRow() === 0) {
    trackerSheet.appendRow(headers);
  }

  if (trackerSheet.getLastRow() > 1) {
    trackerSheet.getRange(2, 1, trackerSheet.getLastRow() - 1, trackerSheet.getLastColumn()).clearContent();
  }

  const featureData = featuresSheet.getDataRange().getValue();
  const featureHeaders = featureData[0];
  const featureRows = featureData.slice(1);
  const featureDescMap = {};
  const featureIndex = featureHeaders.indexOf("Feature");
  const descIndex = featureHeaders.indexOf("Description");

  featureRows.forEach(rows => {
    const key = String(row[featureIndex]).trim().toLowerCase();
    const Description = row[descIndex] || "";
    featureDescMap[key] = Description;
  });

  let stepCounter = 1;
  sections.forEach(section => {
    const row = [
      section.SectionHeader || "Untitled Feature",
      `${stepCounter}. ${section.SectionHeader || "Untitled"}`,
      section.Type === "Integration" ? "No":"Yes",
      section.Type === "Integration" ? "Yes":"No",
      "No",
      section.Purpose || ""
    ];
    trackerSheet.appendRow(row);
    stepCounter++;
  });
  const lastRow = trackerSheet.getLastRow();
  const implementedRange = trackerSheet.getRange(2, 5, lastRow -1);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Yes", "No"], true)
    .setAllowInvalid(false)
    .build();

  implementedRange.setDataValidation(rule);

  const rules = trackerSheet.getConditionalFormatRules();

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Yes")
      .setBackground("#c6efce")
      .setFontColor("#006100")
      .setRanges([implementedRange])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("No")
      .setBackground("#ffc7ce")
      .setFontColor("#9c0006")
      .setRanges([implementedRange])
      .build()
  );
  trackerSheet.setConditionalFormatRules(rules);
}