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
  var rawPreviousWord = getPrecedingText(refInsertionPoint.index);

    // Define a regular expression pattern to match digits (\d)
    var pattern = /\d/g;

    var previousWord = rawPreviousWord.replace(pattern, '');

  // Prompt the user to edit the previous word
  var userInput = prompt("Please edit the previous word:", previousWord);

  // Check if the user clicked OK or Cancel
  if (userInput !== null) {
    // User clicked OK, update the previousWord with the user input
    previousWord = userInput;
  }  

  // Check if the previous word exists
  if (previousWord) {
    // Copy the previous word
    var wordText = previousWord;

    // Insert the word at the beginning of the endnote
    endnote.insertionPoints[2].contents = wordText + " ] ";

    // Prompt the user for the next action
    // Create a dialog
    var dialog = new Window("dialog", "Endnote Processing");

    // Add a static text
    var staticText = dialog.add("statictext", undefined, "Endnote " + i + " processed. What would you like to do?");

    // Function to handle processing the endnote
    function processEndnote(choice) {
      switch (choice) {
        case 1: // Proceed to the next endnote
          break;
        case 2: // Skip the next endnote
          i++;
          break;
        case 3: // Stop running the script
          break;
        default: // Invalid choice, stop running the script
          break;
      }
    }

    // Add buttons
    var nextButton = dialog.add("button", undefined, "Proceed to the next endnote");
    var skipButton = dialog.add("button", undefined, "Skip the next endnote");
    var stopButton = dialog.add("button", undefined, "Stop running the script");

    // Define button click handlers
    nextButton.onClick = function() {
      processEndnote(1);
      dialog.close();
    };

    skipButton.onClick = function() {
      processEndnote(2);
      dialog.close();
    };

    stopButton.onClick = function() {
      processEndnote(3);
      dialog.close();
    };

    // Show the dialog
    dialog.show();


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
