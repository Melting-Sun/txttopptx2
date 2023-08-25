const fs = require('fs');
const pptxgen = require('pptxgenjs');

// Extracted text from the original PPTX (replace with your extracted text)
const extractedText = "This is the extracted text from the PPTX file.";

// Create a new PowerPoint presentation
const pptx = new pptxgen();

// Add a slide with your extracted text
const slide = pptx.addSlide();
slide.addText(extractedText, {
    x: '50%',
    y: '50%',
    fontSize: 20,
    color: '000000',  // Black color
});

// Generate the PowerPoint presentation
const pptxBuffer = pptx.writeFile();

// Save the new PPTX to a file using the fs module
fs.writeFileSync('new_presentation.pptx', pptxBuffer, 'binary');

console.log('Presentation saved.');
