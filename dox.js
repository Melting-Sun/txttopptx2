


const mammoth = require("mammoth");
const fs = require("fs");

// Load the DOCX file and extract content
mammoth.extractRawText({ path: "./myWord.docx" }).then(result => {
  const content = result.value;

  // Split content into paragraphs
  const paragraphs = content.split("\n").filter(para => para.trim() !== "");

  // Extract the title (first paragraph)
  const title = paragraphs[0];

  // Extract bullet points (remaining paragraphs)
  const bulletPoints = paragraphs.slice(1);

  console.log("Title:", title);
  console.log("Bullet Points:", bulletPoints);
  modifyExistingPresentation(title,bulletPoints)
}).catch(error => {
  console.error("An error occurred:", error);
});








const PPTX = require('nodejs-pptx');

async function modifyExistingPresentation(title, bulletPoints) {
  let pptx = new PPTX.Composer();

  // Load the existing presentation
  await pptx.load('./trmp1.pptx');

  // Compose and modify the presentation
  await pptx.compose(async pres => {
    await pres.getSlide('slide1').addText(text => {
        text
          .value(title)
          .x(400)
          .y(50)
          .fontFace('Alien Encounters')
          .fontSize(100)
          .textColor('CC0000')
          .textWrap('none')
          .textAlign('center')
          .textVerticalAlign('center')
          .line({ color: '0000FF', dashType: 'dash', width: 1.0 })
          .margin(0);
    });
    await pres.getSlide('slide1').addText(text => {
        text
          .value(bulletPoints)
          .x(20)
          .y(400)
          .fontFace('Alien Encounters')
          .fontSize(30)
          .textColor('CC0000')
          .textWrap('none')
          .textAlign('left')
          .textVerticalAlign('left')
          .line({ color: '0000FF', dashType: 'dash', width: 1.0 })
          .margin(0);
    });

  });

  // Save the modified presentation
  await pptx.save('./newtep44.pptx');

  console.log('Modified presentation saved!');
}


