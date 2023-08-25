const mammoth = require("mammoth");
const PPTX = require('nodejs-pptx');

async function createSlidesFromDocx(docxPath) {
  const result = await mammoth.extractRawText({ path: docxPath });
  const content = result.value;
  const paragraphs = content.split("\n").filter(para => para.trim() !== "");
  const slidesData = [];
  let currentSlide = null;

  for (const paragraph of paragraphs) {
    if (paragraph.includes("\t")) {
      if (currentSlide) {
        currentSlide.bulletPoints.push(paragraph);
      }
    } else {
      if (currentSlide) {
        slidesData.push(currentSlide);
      }
      currentSlide = { title: paragraph, bulletPoints: [] };
    }
  }

  if (currentSlide) {
    slidesData.push(currentSlide);
  }

  return slidesData;
}

async function createPptxFromSlides(slidesData, templatePptxPath, outputPptxPath) {
  let pptx = new PPTX.Composer();
  await pptx.load(templatePptxPath);

  await pptx.compose(async pres => {
    for (const slideData of slidesData) {
      const { title, bulletPoints } = slideData;

      await pres.addSlide(async slide => {
        await slide.addText(text => {
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

        for (const bulletPoint of bulletPoints) {
          await slide.addText(text => {
            text
              .value(bulletPoint)
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
        }
      });
    }
  });

  await pptx.save(outputPptxPath);
  console.log('New presentation saved!');
}

const docxPath = "./myWord.docx";
const templatePptxPath = './trmp1.pptx';
const outputPptxPath = './newtep44.pptx';

(async () => {
  const slidesData = await createSlidesFromDocx(docxPath);
  await createPptxFromSlides(slidesData, templatePptxPath, outputPptxPath);
})();
