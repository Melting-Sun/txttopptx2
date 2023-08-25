const fs = require("fs");
const pptxgen = require("pptxgenjs");

function modifyExistingPresentation() {
  // Load the existing PowerPoint file
  const existingPptx = fs.readFileSync("temp.pptx");

  // Initialize pptxgen and load the existing presentation
  const pptx = new pptxgen();
  pptx.load(existingPptx);

  // Create a new slide and add text to it
  const newSlide = pptx.addSlide();
  newSlide.addText("New Slide Text", {
    x: 1,
    y: 1,
    w: "100%",
    h: 2,
    fontFace: "Arial",
    fontSize: 18
  });

  // Save the modified presentation to a file
  pptx.writeFile("modified.pptx", () => {
    console.log("Modified presentation saved!");
  });
}

modifyExistingPresentation();

