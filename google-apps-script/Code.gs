function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
      .setTitle('קולם - תיעוד ביוגרפיה קולית')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// CORS POST Handler to support standalone static hosting (GitHub Pages)
function doPost(e) {
  var response = {};
  try {
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    
    if (action === "getSettings") {
      response = getSettings();
    } else if (action === "getStories") {
      response = getStories();
    } else if (action === "selectStory") {
      response = selectStory(requestData.storyId);
    } else if (action === "createNewStory") {
      response = createNewStory(requestData.storyName);
    } else if (action === "saveSettings") {
      response = saveSettings(requestData.userName, requestData.apiKey);
    } else if (action === "getLastParagraphs") {
      response = getLastParagraphs();
    } else if (action === "transcribeAudio") {
      response = transcribeAudio(requestData.base64Audio);
    } else {
      response = { status: "error", message: "פעולה לא מוכרת: " + action };
    }
  } catch (err) {
    response = { status: "error", message: "שגיאת שרת בעיבוד הבקשה: " + err.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(response))
                       .setMimeType(ContentService.MimeType.JSON);
}

// Get folder named "קולם" or create it if not exists
function getStoriesFolder() {
  var folderName = "קולם";
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

// List all stories (Google Docs) inside "קולם" folder
function getStories() {
  try {
    var folder = getStoriesFolder();
    var files = folder.getFilesByType(MimeType.GOOGLE_DOCS);
    var stories = [];
    
    while (files.hasNext()) {
      var file = files.next();
      stories.push({
        id: file.getId(),
        name: file.getName()
      });
    }
    
    var activeStoryId = PropertiesService.getScriptProperties().getProperty('ACTIVE_DOC_ID') || "";
    var activeStoryName = "";
    
    if (activeStoryId) {
      try {
        var activeFile = DriveApp.getFileById(activeStoryId);
        activeStoryName = activeFile.getName();
      } catch (e) {
        // Active doc was deleted or inaccessible, reset
        activeStoryId = "";
        PropertiesService.getScriptProperties().deleteProperty('ACTIVE_DOC_ID');
      }
    }
    
    // If no active story but stories exist, make the first one active
    if (!activeStoryId && stories.length > 0) {
      activeStoryId = stories[0].id;
      activeStoryName = stories[0].name;
      PropertiesService.getScriptProperties().setProperty('ACTIVE_DOC_ID', activeStoryId);
    }
    
    return {
      status: "success",
      stories: stories,
      active_story_id: activeStoryId,
      active_story_name: activeStoryName
    };
  } catch (err) {
    return { status: "error", message: err.toString(), stories: [] };
  }
}

// Select active story
function selectStory(storyId) {
  try {
    var file = DriveApp.getFileById(storyId);
    PropertiesService.getScriptProperties().setProperty('ACTIVE_DOC_ID', storyId);
    return {
      status: "success",
      active_story_id: storyId,
      active_story_name: file.getName()
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

// Create a new story (Google Doc)
function createNewStory(storyName) {
  try {
    var folder = getStoriesFolder();
    var doc = DocumentApp.create(storyName);
    var docFile = DriveApp.getFileById(doc.getId());
    
    // Add default title paragraph and format RTL
    var body = doc.getBody();
    var p = body.appendParagraph(storyName);
    p.setLeftToRight(false);
    p.setHeading(DocumentApp.ParagraphHeading.TITLE);
    
    doc.saveAndClose();
    
    // Move to "קולם" folder
    folder.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile); // Remove from root
    
    PropertiesService.getScriptProperties().setProperty('ACTIVE_DOC_ID', doc.getId());
    
    return {
      status: "success",
      active_story_id: doc.getId(),
      active_story_name: storyName
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

// Read settings
function getSettings() {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('GEMINI_API_KEY') || "";
  var userName = props.getProperty('USER_NAME') || "אבא";
  var activeStoryId = props.getProperty('ACTIVE_DOC_ID') || "";
  
  var activeStoryName = "";
  var activeStoryUrl = "";
  if (activeStoryId) {
    try {
      var file = DriveApp.getFileById(activeStoryId);
      activeStoryName = file.getName();
      activeStoryUrl = file.getUrl();
    } catch (e) {}
  }
  
  return {
    user_name: userName,
    has_api_key: apiKey.length > 0,
    masked_key: apiKey ? (apiKey.substring(0, 6) + "..." + apiKey.substring(apiKey.length - 4)) : "",
    active_story_name: activeStoryName,
    active_story_url: activeStoryUrl,
    version: "0.1.20"
  };
}

// Save settings
function saveSettings(userName, apiKey) {
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('USER_NAME', userName);
    if (apiKey && !apiKey.includes("...")) {
      props.setProperty('GEMINI_API_KEY', apiKey);
    }
    return { status: "success" };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

// Read last paragraphs from active document
function getLastParagraphs() {
  try {
    var activeStoryId = PropertiesService.getScriptProperties().getProperty('ACTIVE_DOC_ID');
    if (!activeStoryId) {
      return { status: "empty", paragraphs: [] };
    }
    
    var doc = DocumentApp.openById(activeStoryId);
    var body = doc.getBody();
    var paragraphs = body.getParagraphs();
    
    var activeStoryName = "";
    if (activeStoryId) {
      try {
        var activeFile = DriveApp.getFileById(activeStoryId);
        activeStoryName = activeFile.getName();
      } catch (e) {}
    }

    var validParagraphs = [];
    var isFirstNonEmpty = true;
    for (var i = 0; i < paragraphs.length; i++) {
      var text = paragraphs[i].getText().trim();
      // Skip title or empty paragraphs
      if (text) {
        if (isFirstNonEmpty) {
          isFirstNonEmpty = false;
          continue;
        }
        if (text !== activeStoryName && !text.startsWith("הביוגרפיה של")) {
          validParagraphs.push(text);
        }
      }
    }
    
    var last3 = validParagraphs.slice(-3);
    if (last3.length === 0) {
      return { status: "empty", paragraphs: [] };
    }
    
    return { status: "success", paragraphs: last3 };
  } catch (err) {
    return { status: "error", message: err.toString(), paragraphs: [] };
  }
}

// Transcribe base64 WAV audio via Gemini 2.5 Flash and append to active Doc
function transcribeAudio(base64Audio) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return { status: "error", message: "מפתח ה-API של Gemini לא הוגדר במערכת. אנא הגדר אותו במסך ההגדרות." };
  }
  
  var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;
  
  var promptText = "תמלל את השמע בעברית בצורה מדויקת ככל הניתן ומילה במילה (Verbatim). שמור על הניסוח, הסגנון המקורי וקולו של המספר. אל תשנה את המילים, אל תשכתב ספרותית ואל תוסיף דבר. הסר רק גמגומים או מילות קישור חוזרות כמו 'אממ' או 'אה'. פלוט אך ורק את הטקסט המתומלל ללא שום תוספת.";
  
  var payload = {
    "contents": [
      {
        "parts": [
          {
            "inlineData": {
              "mimeType": "audio/wav",
              "data": base64Audio
            }
          },
          {
            "text": promptText
          }
        ]
      }
    ]
  };
  
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    var resCode = response.getResponseCode();
    var resText = response.getContentText();
    
    if (resCode !== 200) {
      var errMsg = "שגיאה בפנייה לשרת Gemini";
      try {
        var parsedErr = JSON.parse(resText);
        if (parsedErr.error && parsedErr.error.message) {
          errMsg += ": " + parsedErr.error.message;
        }
      } catch (e) {
        errMsg += " (קוד " + resCode + "): " + resText;
      }
      return { status: "error", message: errMsg };
    }
    
    var json = JSON.parse(resText);
    if (!json.candidates || json.candidates.length === 0 || !json.candidates[0].content) {
      return { status: "error", message: "תשובת שרת ה-API של Gemini הייתה ריקה." };
    }
    
    var transcribedText = json.candidates[0].content.parts[0].text.trim();
    
    var activeDocId = PropertiesService.getScriptProperties().getProperty('ACTIVE_DOC_ID');
    if (!activeDocId) {
      var defaultDoc = createNewStory("הביוגרפיה שלי");
      activeDocId = defaultDoc.active_story_id;
    }
    
    var doc = DocumentApp.openById(activeDocId);
    var body = doc.getBody();
    
    var paragraphs = body.getParagraphs();
    var firstParagraph = paragraphs[0];
    
    // Clean initial empty document placeholders
    if (paragraphs.length === 1 && firstParagraph.getText().trim() === "") {
      firstParagraph.setText(transcribedText);
      firstParagraph.setLeftToRight(false);
    } else {
      var newP = body.appendParagraph(transcribedText);
      newP.setLeftToRight(false);
    }
    
    doc.saveAndClose();
    return { status: "success", text: transcribedText };
  } catch (err) {
    return { status: "error", message: "שגיאת שרת: " + err.toString() };
  }
}

// Temporary function to force authorization of UrlFetchApp scope
function forceAuth() {
  UrlFetchApp.fetch("https://generativelanguage.googleapis.com/");
}
