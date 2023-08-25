


// const mammoth = require("mammoth");
// const fs = require("fs");



// mammoth.extractRawText({ path: "./myWord.docx" }).then(result => {
//   const content = result.value;

//   // Split content into paragraphs
//   const paragraphs = content.split("\n").filter(para => para.trim() !== "");

//   // Initialize arrays to store headings and bullet points
//   const headings = [];
//   const bulletPoints = [];

//   // Iterate through paragraphs to identify headings and bullet points
//   let isHeading = false;
//   let currentHeading = "";
//   let currentBulletPoints = [];

//   for (const paragraph of paragraphs) {
//     // Identify headings based on text size, bold, or other formatting
//     if (isHeading) {
//       if (paragraph.trim() === "") {
//         // End of heading
//         headings.push(currentHeading);
//         isHeading = false;
//         currentHeading = "";
//       } else {
//         currentHeading += paragraph + "\n";
//       }
//     } else if (paragraph.startsWith("*") || paragraph.startsWith("-") || paragraph.startsWith("•")) {
//       // Identify bullet points based on common bullet symbols
//       currentBulletPoints.push(paragraph);
//     } else if (paragraph.trim() !== "") {
//       // Start of heading
//       isHeading = true;
//       currentHeading = paragraph + "\n";
//     }
//   }

//   // Log the extracted headings and bullet points
//   console.log("Headings:", headings);
//   console.log("Bullet Points:", bulletPoints);

//   // Call your function to modify the presentation using the extracted data
//   modifyExistingPresentation(headings, bulletPoints);
// }).catch(error => {
//   console.error("An error occurred:", error);
// });









// const PPTX = require('nodejs-pptx');

// async function modifyExistingPresentation(title, bulletPoints) {
//   let pptx = new PPTX.Composer();

//   // Load the existing presentation
//   await pptx.load('./trmp1.pptx');

//   // Compose and modify the presentation
//   await pptx.compose(async pres => {
//     await pres.getSlide('slide1').addText(text => {
//         text
//           .value(title)
//           .x(400)
//           .y(50)
//           .fontFace('Alien Encounters')
//           .fontSize(100)
//           .textColor('CC0000')
//           .textWrap('none')
//           .textAlign('center')
//           .textVerticalAlign('center')
//           .line({ color: '0000FF', dashType: 'dash', width: 1.0 })
//           .margin(0);
//     });
//     await pres.getSlide('slide1').addText(text => {
//         text
//           .value(bulletPoints)
//           .x(20)
//           .y(400)
//           .fontFace('Alien Encounters')
//           .fontSize(30)
//           .textColor('CC0000')
//           .textWrap('none')
//           .textAlign('left')
//           .textVerticalAlign('left')
//           .line({ color: '0000FF', dashType: 'dash', width: 1.0 })
//           .margin(0);
//     });

//   });

//   // Save the modified presentation
//   await pptx.save('./newtep44.pptx');

//   console.log('Modified presentation saved!');
// }




const mammoth = require("mammoth");
const PPTX = require('nodejs-pptx');

mammoth.extractRawText({ path: "./myWord.docx" }).then(result => {
  const content = result.value;

  // Split content into paragraphs
  const paragraphs = content.split("\n").filter(para => para.trim() !== "");

  // Initialize arrays to store headings and bullet points
  const headings = [];
  const bulletPoints = [];

  let currentHeading = null;
  let isBulletPoint = false;
  let currentBulletPoints = [];

  for (const paragraph of paragraphs) {
    if (paragraph.trim() !== "") {
      if (paragraph.trim().length > 0 && paragraph.trim().length < 50 && !isBulletPoint) {
        // Assuming smaller length indicates a bullet point
        isBulletPoint = true;
        currentBulletPoints.push(paragraph.trim());
      } else if (isBulletPoint) {
        // This is a bullet point
        currentBulletPoints.push(paragraph.trim());
      } else {
        // This is a heading
        if (currentHeading) {
          headings.push(currentHeading);
        }
        currentHeading = paragraph.trim();
      }
    }
  }

  // Log the extracted headings and bullet points
  console.log("Headings:", headings);
  console.log("Bullet Points:", bulletPoints);

  // Call the function to modify the presentation using the extracted data
  // modifyExistingPresentation(headings, bulletPoints);
}).catch(error => {
  console.error("An error occurred:", error);
});

