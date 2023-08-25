


const PPTX = require('nodejs-pptx');
const fs = require('fs/promises'); // Import fs.promises for async file operations


async function modifyExistingPresentation() {
  try {
    // Read configuration from config.json
    const configData = await fs.readFile('config.json', 'utf-8');
    const config = JSON.parse(configData);

    let pptx = new PPTX.Composer();

    // Load the existing presentation
    await pptx.load(config.inputPath);

    // Compose and modify the presentation
    await pptx.compose(async pres => {
      await pres.getSlide('slide1').addText(text => {
        text
          .value(config.textProperties.value)
          .x(config.textProperties.x)
          .y(config.textProperties.y)
          .fontFace(config.textProperties.fontFace)
          .fontSize(config.textProperties.fontSize)
          .textColor(config.textProperties.textColor)
          .textWrap(config.textProperties.textWrap)
          .textAlign(config.textProperties.textAlign)
          .textVerticalAlign(config.textProperties.textVerticalAlign)
          .line(config.textProperties.line)
          .margin(config.textProperties.margin);
      });
    });

    // Save the modified presentation
    await pptx.save(config.outputPath);

    console.log('Modified presentation saved!');
  } catch (error) {
    console.error('An error occurred:', error);
  }
}

modifyExistingPresentation();



