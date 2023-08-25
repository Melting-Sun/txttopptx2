const WordExtractor = require("word-extractor");
const PPTX = require('nodejs-pptx');

async function extractHeadersFromDocx(docxPath) {
  const extractor = new WordExtractor();
  const extracted = await extractor.extract(docxPath);
  const headerString = extracted.getHeaders();
  const headers = headerString.split('\n').filter(header => header.trim() !== '');
  return headers;
}

async function addHeadersToPptx(headers, pptxPath, outputPptxPath) {
  let pptx = new PPTX.Composer();
  await pptx.load('./trmp1.pptx');

  for (const [i, header] of headers.entries()) {
    await pptx.compose(async pres => {
      await pres.getSlide(i + 1).addText(text => {
        text
          .value(header)
          .x(400)
          .y(50)
          .fontFace('Alien Encounters')
          .fontSize(10)
          .textColor('CC0000')
          .textWrap('none')
          .textAlign('center')
          .textVerticalAlign('center')
          .line({ color: '0000FF', dashType: 'dash', width: 1.0 })
          .margin(0);
      });
    });
  }

  await pptx.save('./newwwwwwwwwww.pptx');
  console.log('New presentation with headers saved!');
}

const headersDocxPath = "./word1.docx";
const templatePptxPath = './template.pptx';
const outputPptxPath = './new_presentation.pptx';

(async () => {
  const headers = await extractHeadersFromDocx(headersDocxPath);
  await addHeadersToPptx(headers, templatePptxPath, outputPptxPath);
})();














// const WordExtractor = require("word-extractor"); 
// const extractor = new WordExtractor();
// const extracted = extractor.extract("./word.docx");


// extracted.then(function(doc) { console.log(doc.getHeaders())});
// extracted.then(function(doc) { const headers = doc.getHeaders()
    
//     // modifyExistingPresentation(headers)
// } );



// const WordExtractor = require("word-extractor");
// const mammoth = require("mammoth");
// const PPTX = require('nodejs-pptx');

// async function extractHeadersFromDocx(docxPath) {
//   const extractor = new WordExtractor();
//   const extracted = await extractor.extract(docxPath);
//   const headerString = extracted.getHeaders();
//   const headers = headerString.split('\n').filter(header => header.trim() !== '');
//   return headers;
// }

// async function createSlidesFromHeaders(headers) {
//   const slidesData = headers.map(header => ({ title: header, bulletPoints: [] }));
//   return slidesData;
// }

// async function createPptxFromSlides(slidesData, templatePptxPath, outputPptxPath) {
//   let pptx = new PPTX.Composer();
//   await pptx.load('./trmp1.pptx');

//   await pptx.compose(async pres => {
//     for (const slideData of slidesData) {
//       const { title, bulletPoints } = slideData;

//       await pres.getSlide('slide1').addText(text => {
//         text
//           .value(title)
//           .x(400)
//           .y(50)
//           .fontFace('Alien Encounters')
//           .fontSize(10)
//           .textColor('CC0000')
//           .textWrap('none')
//           .textAlign('center')
//           .textVerticalAlign('center')
//           .line({ color: '0000FF', dashType: 'dash', width: 1.0 })
//           .margin(0);
//     });


//     // await pres.getSlide('slide1').addText(text => {
//     //     text
//     //       .value(bulletPoints)
//     //       .x(20)
//     //       .y(400)
//     //       .fontFace('Alien Encounters')
//     //       .fontSize(30)
//     //       .textColor('CC0000')
//     //       .textWrap('none')
//     //       .textAlign('left')
//     //       .textVerticalAlign('left')
//     //       .line({ color: '0000FF', dashType: 'dash', width: 1.0 })
//     //       .margin(0);
//     // });

//     //   await pres.addSlide(async slide => {
//     //     await slide.addText(text => {
//     //       text
//     //         .value(title)
//     //         .x(400)
//     //         .y(50)
//     //         .fontFace('Alien Encounters')
//     //         .fontSize(100)
//     //         .textColor('CC0000')
//     //         .textWrap('none')
//     //         .textAlign('center')
//     //         .textVerticalAlign('center')
//     //         .line({ color: '0000FF', dashType: 'dash', width: 1.0 })
//     //         .margin(0);
//     //     });

//     //     for (const bulletPoint of bulletPoints) {
//     //       await slide.addText(text => {
//     //         text
//     //           .value(bulletPoint)
//     //           .x(20)
//     //           .y(400)
//     //           .fontFace('Alien Encounters')
//     //           .fontSize(30)
//     //           .textColor('CC0000')
//     //           .textWrap('none')
//     //           .textAlign('left')
//     //           .textVerticalAlign('left')
//     //           .line({ color: '0000FF', dashType: 'dash', width: 1.0 })
//     //           .margin(0);
//     //       });
//     //     }
//     //   });
//     }
//   });

//   await pptx.save(outputPptxPath);
//   console.log('New presentation saved!');
// }

// const docxPath = "./word.docx";
// const templatePptxPath = './trmp1.pptx';
// const outputPptxPath = './newtep44.pptx';

// (async () => {
//   const headers = await extractHeadersFromDocx(docxPath);
//   const slidesData = await createSlidesFromHeaders(headers);
//   await createPptxFromSlides(slidesData, templatePptxPath, outputPptxPath);
// })();
































































































// const WordExtractor = require("word-extractor");
// const PPTX = require('nodejs-pptx');
// const fs = require('fs/promises');

// async function extractHeadersFromDocx(docxPath) {
//   const extractor = new WordExtractor();
//   const extracted = await extractor.extract(docxPath);
//   const headerString = extracted.getHeaders();
//   const headers = headerString.split('\n').filter(header => header.trim() !== '');
//   return headers;
// }

// async function modifyExistingPresentationWithHeaders(headers) {
//   try {
//     const pptx = new PPTX.Composer();

//     // Load the existing presentation
//     await pptx.load('./allah2.pptx');

//     // Get the slides from the loaded presentation
//     const slides = pptx.slides;

//     if (!slides || slides.length === 0) {
//       throw new Error('No slides found in the existing presentation.');
//     }

//     // Iterate through the slides and update their content with headers
//     for (let i = 0; i < slides.length && i < headers.length; i++) {
//       const slide = slides[i];
//       const header = headers[i];

//       if (!slide) {
//         throw new Error(`Slide ${i + 1} not found in the existing presentation.`);
//       }

//       // Clear existing content
//       slide.clear();

//       // Add header as text to the slide
//       await slide.addText(text => {
//         text
//           .value(header)
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
//       });
//     }

//     // Save the modified presentation
//     await pptx.save('./trmp1.pptx');

//     console.log('Modified presentation saved!');
//   } catch (error) {
//     console.error('An error occurred:', error);
//   }
// }

// (async () => {
//   const docxPath = "./word.docx";
//   const headers = await extractHeadersFromDocx(docxPath);
  
//   await modifyExistingPresentationWithHeaders(headers);
// })();









































// const PPTX = require('nodejs-pptx');

// async function modifyExistingPresentation(headers) {
//   let pptx = new PPTX.Composer();

//   // Load the existing presentation
//   await pptx.load('./trmp1.pptx');

//   // Compose and modify the presentation
//   await pptx.compose(async pres => {
//     await pres.getSlide('slide1').addText(text => {
//         text
//           .value(headers)
//           .x(400)
//           .y(400)
//           .fontFace('Alien Encounters')
//           .fontSize(100)
//           .textColor('CC0000')
//           .textWrap('none')
//           .textAlign('center')
//           .textVerticalAlign('center')
//           .line({ color: '0000FF', dashType: 'dash', width: 1.0 })
//           .margin(0);
//     });
    

//   });

//   // Save the modified presentation
//   await pptx.save('./newtep99.pptx');

//   console.log('Modified presentation saved!');
// }









// const fs = require('fs');
// const WordExtractor = require('word-extractor');
// const officegen = require('officegen');

// async function generatePptxWithHeaders() {
//   try {
//     // Extract headers from the Word document
//     const extractor = new WordExtractor();
//     const extracted = await extractor.extract('./word.docx');
//     const headers = extracted.getHeaders();

//     const pptx = officegen('pptx');

//     for (const header of headers) {
//       const slide = pptx.makeNewSlide();
//       const title = slide.addText(header, {
//         x: 'c',
//         y: 'c',
//         font_face: 'Arial',
//         font_size: 24,
//         color: '000000',
//         bold: true,
//         align: 'center',
//         valign: 'middle',
//       });
//     }

//     const outputPath = './output.pptx';
//     const stream = fs.createWriteStream(outputPath);
//     pptx.generate(stream);

//     stream.on('finish', () => {
//       console.log('New presentation with headers generated!');
//     });
//   } catch (error) {
//     console.error('An error occurred:', error);
//   }
// }

// generatePptxWithHeaders();














// const fs = require('fs');
// const WordExtractor = require('word-extractor');
// const officegen = require('officegen');

// async function generatePptxWithHeaders() {
//   try {
//     // Extract headers from the Word document
//     const extractor = new WordExtractor();
//     const extracted = await extractor.extract('./myWord.docx');
//     const headers = extracted.getHeaders();

//     const pptx = officegen('pptx');

//     for (const header of headers) {
//       const slide = pptx.makeNewSlide();

//       // Split header into words and create a slide for each word
//       const words = header.split(' ');
//       for (const word of words) {
//         const title = slide.addText(word, {
//           x: 'c',
//           y: 'c',
//           font_face: 'Arial',
//           font_size: 24,
//           color: '000000',
//           bold: true,
//           align: 'center',
//           valign: 'middle',
//         });

//         // Add a new slide after each word, except the last one
//         if (word !== words[words.length - 1]) {
//           slide.addNewSlide();
//         }
//       }
//     }

//     const outputPath = './output.pptx';
//     const stream = fs.createWriteStream(outputPath);
//     pptx.generate(stream);

//     stream.on('finish', () => {
//       console.log('New presentation with headers generated!');
//     });
//   } catch (error) {
//     console.error('An error occurred:', error);
//   }
// }

// generatePptxWithHeaders();






