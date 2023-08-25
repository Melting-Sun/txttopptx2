const fs = require('fs');
const PDFParser = require('pdf-parse');

const pdfPath = 'Salam.pdf';

// Read the PDF file
const pdfBuffer = fs.readFileSync(pdfPath);

// Parse the PDF and extract text
PDFParser(pdfBuffer).then(data => {
  const extractedText = data.text;
  console.log(extractedText);
}).catch(error => {
  console.error('Error parsing PDF:', error);
});
