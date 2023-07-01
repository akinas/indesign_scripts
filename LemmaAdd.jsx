// Prompt the user to input the range of endnote IDs
var startID = parseInt(prompt("Enter the start ID:", ""));
var endID = parseInt(prompt("Enter the end ID:", ""));

// Get the active document
var doc = app.activeDocument;

// Get the story containing the endnotes
//var endnoteStory = doc.stories.itemByName("Endnotes");
var endnoteStory = doc.stories[0];

// Loop through the endnotes in the specified range
for (var i = startID-1; i <= endID-1; i++) {
  var endnote = endnoteStory.endnotes[i];

  // Find the endnote reference insertion point
  var refInsertionPoint = endnote.storyOffset;


  // Get the previous word before the reference insertion point
  var previousWord = getPrecedingText(refInsertionPoint.index);
  alert("Endnote insertion point and previous word: " + refInsertionPoint.index + " - " + previousWord);
  // Check if the previous word exists and is not punctuation
  if (previousWord) {
    // Copy the previous word
    var wordText = previousWord;

    // Insert the word at the beginning of the endnote
    endnote.insertionPoints[2].contents = wordText + " ] ";

    // Prompt the user for the next action
    var userChoice = prompt(
      "Endnote " + i + " processed. What would you like to do?\n1. Proceed to the next endnote\n2. Skip the next endnote\n3. Stop running the script\nEnter the number corresponding to your choice:"
    );

    // Process the user's choice
    if (userChoice == "1") {
      // Proceed to the next endnote
      continue;
    } else if (userChoice == "2") {
      // Skip the next endnote
      i++;
    } else if (userChoice == "3") {
      // Stop running the script
      break;
    } else {
      // Invalid choice, stop running the script
      break;
    }
  }
}

// Function to retrieve preceding text based on insertionIndex
function getPrecedingText(insertionIndex) {
  var precedingText = "";
  var currentIndex = insertionIndex;

  // Iterate backwards through the story until a space or paragraph marker is encountered
  while (currentIndex > 0) {
    var character = endnoteStory.characters[currentIndex - 1];
    var characterContents = character.contents;

    if (characterContents === " " || characterContents === "\r") {
      // Stop iterating if a space or paragraph marker is encountered
      break;
    }

    precedingText = characterContents + precedingText;
    currentIndex--;
  }

  return precedingText;
}

// Alert the user when the script has finished processing
alert("Script has finished processing endnotes.");
