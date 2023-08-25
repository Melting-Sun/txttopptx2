// import Automizer from 'pptx-automizer';
// const fs = require('fs');
const Automizer = require('pptx-automizer')

// First, let's set some preferences!
const automizer = new Automizer({
  // this is where your template pptx files are coming from:
  templateDir: `./temp.pptx`,

  // use a fallback directory for e.g. generic templates:
  templateFallbackDir: `my/pptx/fallback-templates`,

  // specify the directory to write your final pptx output files:
  outputDir: `my/pptx/output`,

  // turn this to true if you want to generally use
  // Powerpoint's creationIds instead of slide-numbers
  // or shape names:
  useCreationIds: false,

  // Always use the original slideMaster and slideLayout of any
  // imported slide:
  autoImportSlideMasters: true,

  // truncate root presentation and start with zero slides
  removeExistingSlides: true,

  // activate `cleanup` to eventually remove unused files:
  cleanup: false,

  // Set a value from 0-9 to specify the zip-compression level.
  // The lower the number, the faster your output file will be ready.
  // Higher compression levels produce smaller files.
  compression: 0,

  // You can enable 'archiveType' and set mode: 'fs'.
  // This will extract all templates and output to disk.
  // It will not improve performance, but it can help debugging:
  // You don't have to manually extract pptx contents, which can
  // be annoying if you need to look inside your files.
  // archiveType: {
  //   mode: 'fs',
  //   baseDir: `${__dirname}/../__tests__/pptx-cache`,
  //   workDir: 'tmpWorkDir',
  //   cleanupWorkDir: true,
  // },

  // use a callback function to track pptx generation process.
  // statusTracker: myStatusTracker,
});

// Now we can start and load a pptx template.
// With removeExistingSlides set to 'false', each addSlide will append to
// any existing slide in RootTemplate.pptx. Otherwise, we are going to start
// with a truncated root template.
let pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  // We want to make some more files available and give them a handy label.
  .load(`SlideWithShapes.pptx`, 'shapes')
  .load(`SlideWithGraph.pptx`, 'graph')
  // Skipping the second argument will not set a label.
  .load(`SlideWithImages.pptx`);

// addSlide takes two arguments: The first will specify the source
// presentation's label to get the template from, the second will set the
// slide number to require.
pres
  .addSlide('graph', 1)
  .addSlide('shapes', 1)
  .addSlide(`SlideWithImages.pptx`, 2);

// Finally, we want to write the output file.
pres.write(`myPresentation.pptx`).then((summary) => {
  console.log(summary);
});

// It is also possible to get a ReadableStream.
// stream() accepts JSZip.JSZipGeneratorOptions for 'nodebuffer' type.
const stream = await pres.stream({
  compressionOptions: {
    level: 9,
  },
});
// You can e.g. output the pptx archive to stdout instead of writing a file:
stream.pipe(process.stdout);

// If you need any other output format, you can eventually access
// the underlying JSZip instance:
const finalJSZip = await pres.getJSZip();
// Convert the output to whatever needed:
const base64 = await finalJSZip.generateAsync({ type: 'base64' });