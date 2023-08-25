// const fs = require('fs');
// const unzipper = require('unzipper');
// const archiver = require('archiver');

// function modifyExistingPresentation() {
//   const existingPptxPath = 'temp.pptx'; // Path to the existing presentation
//   const outputPath = 'allah.pptx'; // Path for the modified presentation

//   // Read the content of the existing PowerPoint file
//   const pptxBuffer = fs.readFileSync(existingPptxPath);

//   // Unzip the PowerPoint file
//   fs.mkdirSync('temp');
//   fs.writeFileSync('temp/presentation.zip', pptxBuffer);
//   fs.createReadStream('temp/presentation.zip')
//     .pipe(unzipper.Extract({ path: 'temp' }))
//     .on('close', () => {
//       // Modify the XML content of the slides

//       // Re-zip the modified content
//       const output = fs.createWriteStream(outputPath);
//       const archive = archiver('zip');
//       archive.pipe(output);
//       archive.directory('temp/', false);
//       archive.finalize();

//       output.on('close', () => {
//         console.log('Modified presentation saved!');
//         fs.rmdirSync('temp', { recursive: true });
//       });
//     });
// }

// modifyExistingPresentation();


// const fs = require('fs');
// const unzipper = require('unzipper');
// const archiver = require('archiver');
// const xml2js = require('xml2js');

// function modifyExistingPresentation() {
//   const existingPptxPath = 'temp.pptx'; // Path to the existing presentation
//   const outputPath = 'yaallah2.pptx'; // Path for the modified presentation

//   // Read the content of the existing PowerPoint file
//   const pptxBuffer = fs.readFileSync(existingPptxPath);

//   // Unzip the PowerPoint file
//   fs.mkdirSync('temp');
//   fs.writeFileSync('temp/presentation.zip', pptxBuffer);
//   fs.createReadStream('temp/presentation.zip')
//     .pipe(unzipper.Extract({ path: 'temp' }))
//     .on('close', () => {
//       // Modify the XML content of the slides
//       const slidesDir = 'temp/ppt/slides';
//       const slideFiles = fs.readdirSync(slidesDir);

//       for (const slideFile of slideFiles) {
//         const slidePath = `${slidesDir}/${slideFile}`;
//         const slideXml = fs.readFileSync(slidePath, 'utf-8');
        
//         // Parse the slide XML
//         xml2js.parseString(slideXml, (err, result) => {
//           if (err) {
//             console.error('Error parsing XML:', err);
//             return;
//           }

//           // Add the desired text to the XML
//           const newText = {
//             t: 'p',
//             r: [
//               {
//                 t: 't',
//                 _: 'New Text'
//               }
//             ]
//           };
//           result['p:sld']['p:cSld'][0]['p:spTree'][0]['p:sp'][1]['p:txBody'][0]['a:p'].push(newText);

//           // Convert the modified XML back to string
//           const builder = new xml2js.Builder();
//           const modifiedXml = builder.buildObject(result);

//           // Write the modified XML back to the slide file
//           fs.writeFileSync(slidePath, modifiedXml);
//         });
//       }

//       // Re-zip the modified content
//       const output = fs.createWriteStream(outputPath);
//       const archive = archiver('zip');
//       archive.pipe(output);
//       archive.directory('temp/', false);
//       archive.finalize();

//       output.on('close', () => {
//         console.log('Modified presentation saved!');
//         fs.rmdirSync('temp', { recursive: true });
//       });
//     });
// }

// modifyExistingPresentation();



const fs = require('fs');
const unzipper = require('unzipper');
const archiver = require('archiver');
const xml2js = require('xml2js');

function modifyExistingPresentation() {
  const existingPptxPath = 'temp.pptx'; // Path to the existing presentation
  const outputPath = 'modifieddd.pptx'; // Path for the modified presentation

  // Read the content of the existing PowerPoint file
  const pptxBuffer = fs.readFileSync(existingPptxPath);

  // Ensure the temp directory doesn't exist before creating it
  if (!fs.existsSync('temp')) {
    fs.mkdirSync('temp');
  }

  fs.writeFileSync('temp/presentation.zip', pptxBuffer);
  fs.createReadStream('temp/presentation.zip')
    .pipe(unzipper.Extract({ path: 'temp' }))
    .on('close', () => {
      // The rest of your code remains unchanged
      // ...

      // Re-zip the modified content
      const output = fs.createWriteStream(outputPath);
      const archive = archiver('zip');
      archive.pipe(output);
      archive.directory('temp/', false);
      archive.finalize();

      output.on('close', () => {
        console.log('Modified presentation saved!');
        fs.rmdirSync('temp', { recursive: true });
      });
    });
}

modifyExistingPresentation();



