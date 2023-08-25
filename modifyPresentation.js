// const fs = require("fs")
// const pptxgen = require("pptxgenjs")

// const currentPptx = fs.readFileSync("output.pptx")

// const pptx = new pptxgen()

// pptx.load(currentPptx)


// pptx.defineSlideMaster({
//     background: {color: "F2F2F2"}
// })

// pptx.addSlide({ masterName: "MASTER_SLIDE" }).addTitle("modified",{x: 1, y: 1, w: "100%", h: 1, align: "center", fontFace: "Arial", fontSize: 44})

// const contentSlide = pptx.addSlide({masterName: 'master slide'})
// contentSlide.addText('intro', {x: 1, y: 1, w: "100%", h: 1, align: "center", fontFace: "Arial", fontSize: 28})
// contentSlide.addText("Lorem ipsum dolor sit amet, consectetur adipiscing elit.", { x: 1, y: 2, w: "100%", h: 5, fontFace: "Arial", fontSize: 18 });

// pptx.writeFile("modified.pptx", () => {
//     console.log("Modified presentation saved!");
//   });



// const pptxgen = require("pptxgenjs");

// function createModifiedPresentation() {
//   const pptx = new pptxgen();

//   // Set a custom background color for all slides
//   pptx.defineSlideMaster({
//     background: { color: "F2F2F2" } // Light gray background
//   });

//   // Add a title slide with custom formatting
//   const titleSlide = pptx.addSlide("MASTER_SLIDE");
//   titleSlide.addText("Modified Presentation", {
//     x: 1,
//     y: 1,
//     w: "100%",
//     h: 1,
//     align: "center",
//     fontFace: "Arial",
//     fontSize: 44
//   });

//   // Add a content slide with custom formatting
//   const contentSlide = pptx.addSlide("MASTER_SLIDE");
//   contentSlide.addText("Introduction", {
//     x: 1,
//     y: 1,
//     w: "100%",
//     h: 1,
//     align: "center",
//     fontFace: "Arial",
//     fontSize: 28
//   });
//   contentSlide.addText("Lorem ipsum dolor sit amet, consectetur adipiscing elit.", {
//     x: 1,
//     y: 2,
//     w: "100%",
//     h: 5,
//     fontFace: "Arial",
//     fontSize: 18
//   });

//   // Save the modified presentation to a file
//   pptx.writeFile("modified.pptx", () => {
//     console.log("Modified presentation saved!");
//   });
// }

// createModifiedPresentation();


const pptxgen = require("pptxgenjs");

function createModifiedPresentation() {
  const pptx = new pptxgen();

  // Add title slide with custom formatting
  const titleSlide = pptx.addSlide();
  titleSlide.addText("Modified Presentation", {
    x: 2.5,
    y: 2.5,
    w: "50%",
    h: 1,
    align: "center",
    fontFace: "Arial",
    fontSize: 44
  });

  // Add content slide with custom formatting
  const contentSlide = pptx.addSlide();
  contentSlide.addText("Introduction", {
    x: 2.5,
    y: 2.5,
    w: "50%",
    h: 1,
    align: "center",
    fontFace: "Arial",
    fontSize: 28
  });
  contentSlide.addText("Lorem ipsum dolor sit amet, consectetur adipiscing elit.", {
    x: 1,
    y: 2,
    w: "100%",
    h: 5,
    fontFace: "Arial",
    fontSize: 18
  });

  // Save the modified presentation to a file
  pptx.writeFile("modified.pptx", () => {
    console.log("Modified presentation saved!");
  });
}

createModifiedPresentation();
