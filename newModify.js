
const PPTX = require('nodejs-pptx');
const fs = require('fs/promises'); // Import fs.promises for async file operations

async function modifyExistingPresentation() {
  let pptx = new PPTX.Composer();

  // Load the existing presentation
  await pptx.load('./trmp1.pptx');

  // Compose and modify the presentation
  await pptx.compose(async pres => {
    await pres.getSlide('slide1').addText(text => {
        text
          .value('Hello World!')
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
  });

  // Save the modified presentation
  await pptx.save('./newtep2.pptx');

  console.log('Modified presentation saved!');
}

modifyExistingPresentation();



