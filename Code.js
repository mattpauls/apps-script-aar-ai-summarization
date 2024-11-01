function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  console.log("onOpen fired");
  try {
    // Add-on menu creation
    const addonMenu = SpreadsheetApp.getUi().createAddonMenu();
    addonMenu.addItem("Generate AAR Summary", "getSettings");
    addonMenu.addToUi();
  } catch (error) {
    console.error("Add-on menu creation failed: " + error);
  }
}

function getSettings() {
  // https://developers.google.com/apps-script/guides/dialogs#custom_dialogs
  // Display a modal dialog box with custom HtmlService content.
  const htmlOutput = HtmlService.createHtmlOutputFromFile("class_select")
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Generate Summary");
}

// Function to get unique class numbers from the sheet
function getUniqueClassNumbers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getRange("C2:C").getValues();
  const classes = data.flat().filter(Boolean); // Flatten array and remove empty values
  const uniqueClasses = [...new Set(classes)]; // Get unique values
  return uniqueClasses;
}

function summarizeAAR(class_number, activity_title) {
  class_number = class_number || "[no class]";
  activity_title = activity_title
    ? `AAR ${activity_title}`
    : "AAR AI-Generated Summary";
  const doc_title = `${class_number} - ${activity_title}`;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const [headers, ...rows] = data;

  // Get column indices for relevant columns
  const classIndex = headers.indexOf("What class?");
  const wellIndex = headers.findIndex((v) => v.includes("What went well?")); // use .includes because there was a trailing space on the document
  const improveIndex = headers.indexOf("What are some ideas to improve it?");

  // Filter rows by class_number
  const filteredRows = rows.filter((row) => row[classIndex] === class_number);

  // Extract "what went well" and "what can be improved upon" data from filtered rows
  const wellData = filteredRows
    .map((row) => row[wellIndex]?.replace(/\n/g, " "))
    .join(" ");
  const improveData = filteredRows
    .map((row) => row[improveIndex]?.replace(/\n/g, " "))
    .join(" ");

  // Call OpenAI API for summarization
  const summaryWell = callOpenAI("Summarize in bullet form: " + wellData);
  const summaryImprove = callOpenAI("Summarize in bullet form: " + improveData);

  // Output summaries to Google Doc with formatting
  const doc = DocumentApp.create(doc_title);
  const body = doc.getBody();

  // Format the "What Went Well" section with heading and bullets
  body
    .appendParagraph(doc_title)
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  // Format the "What Went Well" section with heading and bullets
  body
    .appendParagraph("What went well")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  appendFormattedContent(body, summaryWell);

  // Format the "Improvements" section with heading and bullets
  body
    .appendParagraph("Ideas for improvement")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  appendFormattedContent(body, summaryImprove);

  // Remove first "paragraph" if it's blank
  const paras = body.getParagraphs();
  const firstPara = paras[0];
  if (paras.length > 1 && !firstPara.getText()) firstPara.removeFromParent();
}

function appendFormattedContent(body, content) {
  // Split content into lines while preserving empty lines
  const lines = content.split(/\r?\n/);
  let currentList = null;
  let currentListLevel = 0;
  let isFirstParagraph = true; // Track if we're at the start of the document

  lines.forEach((line, index) => {
    const trimmedLine = line.trim();

    // Skip empty lines, but only if they're not being used to separate sections
    if (!trimmedLine) {
      if (currentList) {
        // Don't add empty lines between list items
        return;
      }
      if (!isFirstParagraph) {
        body.appendParagraph("");
      }
      currentList = null; // End current list
      return;
    }

    // Calculate indentation level based on leading spaces
    const indentMatch = line.match(/^(\s*)/);
    const indentLevel = Math.floor((indentMatch[1].length || 0) / 2);

    // Extract the actual content without list markers and initial spacing
    let content = trimmedLine.replace(/^[-*]\s*/, "");

    // Handle different types of formatting
    if (trimmedLine.match(/^#+\s/)) {
      // Handle headers
      const headerLevel = trimmedLine.match(/^#+/)[0].length;
      content = content.replace(/^#+\s/, "");
      if (!isFirstParagraph) {
        body.appendParagraph(""); // Add space before header if not first paragraph
      }
      const paragraph = body.appendParagraph(content);
      switch (headerLevel) {
        case 1:
          paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
          break;
        case 2:
          paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
          break;
        case 3:
          paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING3);
          break;
        default:
          paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING4);
      }
      currentList = null; // End current list
    } else if (trimmedLine.match(/^[-*]\s/)) {
      // Handle list items
      content = processInlineFormatting(content);

      if (!currentList || currentListLevel !== indentLevel) {
        currentList = body.appendListItem(content);
        currentListLevel = indentLevel;
      } else {
        currentList = body.appendListItem(content);
      }

      // Apply indentation based on nesting level
      currentList
        .setIndentStart(indentLevel * 18)
        .setGlyphType(DocumentApp.GlyphType.BULLET)
        .setLineSpacing(1); // Reduce space between list items

      // Apply nested list formatting if needed
      if (indentLevel > 0) {
        currentList.setNestingLevel(indentLevel);
      }
    } else {
      // Handle regular paragraphs
      if (!isFirstParagraph) {
        content = processInlineFormatting(content);
        body.appendParagraph(content);
      }
      currentList = null; // End current list
    }

    isFirstParagraph = false; // Mark that we've processed the first paragraph
  });

  // Clean up any extra spacing at the end of lists
  const numChildren = body.getNumChildren();
  for (let i = 0; i < numChildren - 1; i++) {
    const element = body.getChild(i);
    const nextElement = body.getChild(i + 1);

    if (
      element.getType() === DocumentApp.ElementType.LIST_ITEM &&
      nextElement.getType() === DocumentApp.ElementType.LIST_ITEM
    ) {
      element.setSpacingAfter(0);
    }
  }
}

function processInlineFormatting(text) {
  let element = null;

  // Process bold text with both ** and __ markers
  text = text.replace(/(\*\*|__)(.*?)\1/g, (match, marker, content) => {
    return content; // Return content to be made bold later
  });

  // Process italic text with both * and _ markers
  text = text.replace(/(\*|_)(.*?)\1/g, (match, marker, content) => {
    return content; // Return content to be made italic later
  });

  // Process code blocks with backticks
  text = text.replace(/`(.*?)`/g, (match, content) => {
    return content; // Return content to be made monospace later
  });

  return text;
}

function callOpenAI(prompt) {
  // moved here instead of at top of script, I don't think AuthMode.NONE has access to script properties
  // so was erroring out, but I could be wrong about that.
  const OPENAI_API_KEY =
    PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  const formattedPrompt = `${prompt}

Please format the response using the following guidelines:
- Use "-" for bullet points
- Use proper indentation for nested points (2 spaces)
- Use ** for bold text
- Keep formatting compact with minimal blank lines
- Only use blank lines to separate major sections`;

  const url = "https://api.openai.com/v1/chat/completions";
  const options = {
    method: "post",
    headers: {
      Authorization: `Bearer ${OPENAI_API_KEY}`,
      "Content-Type": "application/json",
    },
    payload: JSON.stringify({
      model: "gpt-4o-mini",
      messages: [
        {
          role: "system",
          content:
            "You are a helpful assistant that creates well-structured, compact summaries. Use minimal spacing between bullet points while maintaining readability.",
        },
        {
          role: "user",
          content: formattedPrompt,
        },
      ],
      max_tokens: 1000,
    }),
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  return data.choices[0].message.content;
}

// Add a utility function to apply text formatting
function applyTextFormatting(textElement) {
  let text = textElement.getText();
  let indices = [];

  // Find bold text
  let boldPattern = /\*\*(.*?)\*\*/g;
  let match;
  while ((match = boldPattern.exec(text)) !== null) {
    indices.push({
      start: match.index,
      end: match.index + match[1].length,
      type: "bold",
    });
  }

  // Apply formatting from end to start to maintain indices
  indices
    .sort((a, b) => b.start - a.start)
    .forEach((index) => {
      switch (index.type) {
        case "bold":
          textElement.setBold(index.start, index.end, true);
          break;
      }
    });

  // Clean up markdown symbols
  text = text.replace(/\*\*(.*?)\*\*/g, "$1");
  textElement.setText(text);
}

function runFlow(settings) {
  // called by pop up dialog entered in settings
  Logger.log(settings);

  const { class_number, activity_title } = settings;

  summarizeAAR(class_number, activity_title);
}
