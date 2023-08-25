// const fs = require('fs');
// const pptxgen = require('pptxgenjs');

// async function generateNewPresentation(pptxFilePath) {
//     const pptx = new pptxgen();

//     // Load the uploaded PowerPoint file
//     const originalPresentation = new pptxgen();
//     originalPresentation.load(pptxFilePath);

//     // Modify the styles of the original presentation (example: change font color)
//     originalPresentation.slides.forEach(slide => {
//         slide.elements.forEach(element => {
//             if (element.options && element.options.color) {
//                 element.options.color = 'FF0000'; // Change font color to red
//             }
//         });
//     });

//     // Generate the new presentation
//     const newPptxData = originalPresentation.save();

//     // Save the new PPTX to a file
//     fs.writeFileSync('new_presentation.pptx', newPptxData, 'binary');
// }

// // Call the function with the path to the uploaded PowerPoint file
// const uploadedFilePath = './'; // Replace with the actual file path
// generateNewPresentation(uploadedFilePath);


const fs = require('fs');
const officegen = require('officegen');

// Replace 'original.pptx' with the path to your existing PowerPoint file
const originalPptxFilePath = 'modified.pptx';

// Create a new PowerPoint presentation
const pptx = officegen('pptx');
const slide = pptx.makeSlide();

// Extracted text from the original PPTX (replace with your extracted text)
const extractedText = "This is the extracted text from the original PPTX file.";

// Add text to the slide with your extracted text and styles
const textElement = slide.addText(extractedText, {
    x: 'c',   // Center the text horizontally
    y: 'c',   // Center the text vertically
    font_size: 20,
    color: 'FF0000',  // Red color
});

// Save the new PPTX to a file
const outputStream = fs.createWriteStream('new_presentation.pptx');
pptx.generate(outputStream);

console.log('New presentation generated.');
