const officegen = require('officegen')
const fs = require('fs')

// Create an empty PowerPoint object:
let pptx = officegen('pptx')

// Let's add a title slide:

let slide = pptx.makeTitleSlide('salam mashti')

// Pie chart slide example:

slide = pptx.makeNewSlide()
// slide.name = 'Pie Chart slide'
// slide.back = 'ffff00'


return new Promise((resolve, reject) => {
  let out = fs.createWriteStream('bb.pptx')

  // This one catch only the officegen errors:
  pptx.on('error', function(err) {
    reject(err)
  })

  // Catch fs errors:
  out.on('error', function(err) {
    reject(err)
  })

  // End event after creating the PowerPoint file:
  out.on('close', function() {
    resolve()
  })

  // This async method is working like a pipe - it'll generate the pptx data and put it into the output stream:
  pptx.generate(out)
})