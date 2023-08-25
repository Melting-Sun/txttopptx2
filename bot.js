// const pptxgen = require("pptxgenjs");

// function createPresentation(inputText) {
//   const pptx = new pptxgen();

//   // Create a slide with a title and content layout
//   const slide = pptx.addSlide();

//   // Add text to the slide
//   slide.addText(inputText, { x: 1, y: 1, w: "80%", h: "50%" });

//   // Convert the presentation to a buffer
//   pptx.writeFile("output.pptx");
// }

// const userInput = "This is the text I want to add to the slide.";
// createPresentation(userInput);


const pptxgen = require("pptxgenjs");

function createPresentation(inputText) {
  const pptx = new pptxgen();

  // Create a slide with a title and content layout
  const slide = pptx.addSlide();

  // Add text to the slide
  slide.addText(inputText, { x: 1, y: 1, w: "80%", h: "50%" });

  // Convert the presentation to a buffer and save to a file
  pptx.writeFile("output.pptx", () => {
    console.log("Presentation saved!");
  });
}

// Get input text from the command line arguments
const userInput = process.argv[2] || "Default text if no input provided";
createPresentation(userInput);
