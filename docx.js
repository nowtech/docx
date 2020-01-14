const officegen = require('officegen')
const fs = require('fs')  // so you can write a local file

// Create an empty Word object
let docx = officegen('docx')

// Officegen calling this function after finishing to generate the docx document:
docx.on('finalize', function(written) {
    console.log(
        'Finish to create a Microsoft Word document.'
    )
})

// Officegen calling this function to report errors:
docx.on('error', function(err) {
    console.log(err)
})

// First page
let pObj = docx.createP()



pObj.addText('Lorem ipsum dolor sit amet, consectetur adipiscing elit. Aenean molestie vitae dolor in semper. Aenean vulputate eu arcu id tincidunt. Quisque efficitur imperdiet consectetur. Nulla viverra dolor nulla, sit amet sodales mauris ultrices et. Nunc efficitur at augue lacinia vulputate. Cras laoreet magna leo, sed tincidunt risus consequat a. Sed sit amet nunc augue. Proin porta neque purus, at pretium est sollicitudin cursus.')


pObj = docx.createP()

pObj.addText('This is a paragraph with a background color and a text color added. Maecenas eget nibh felis. Phasellus ullamcorper nec tortor sit amet bibendum. Pellentesque fringilla tellus id pulvinar pretium. Donec nisl mi, accumsan vitae gravida quis, ultrices ut erat. Aenean mollis, libero a euismod sollicitudin, lacus sem malesuada tellus, ac consectetur nibh felis non nulla. Ut sit amet posuere augue, eu viverra nibh. Ut aliquet risus a iaculis placerat. Phasellus vel semper mauris. After this paragraph there will be a page break.', { color: '00ffff', back: '000088' })

//Pagebreak,jumps to next page
docx.putPageBreak()

// table
const table = [
    [{
        val: "No.",
        opts: {
            cellColWidth: 4261,
            b:true,
            sz: '48',
            spacingBefor: 120,
            spacingAfter: 120,
            spacingLine: 240,
            spacingLineRule: 'atLeast',
            shd: {
                fill: "7F7F7F",
                themeFill: "text1",
                themeFillTint: "80"
            },
            fontFamily: "Avenir Book"
        }
    },{
        val: "Title1",
        opts: {
          //  cellColWidth not adjusting here
          // Text in this column still adjusting vertical....
            cellColWidth: 4261,
            b:true,
            sz: "48",
            color: "A00000",
            // align: "right",
            shd: {
                fill: "92CDDC",
                themeFill: "text1",
                "themeFillTint": "80"
            }
        }
    }],
    [1,'All grown-ups were once children'],
    [2,'there is no harm in putting off a piece of work until another day.']
]

const tableStyle = {
    tableColWidth: 8261,
    tableSize: 24,
    tableColor: "ada",
    tableAlign: "left",
    tableFontFamily: "Comic Sans MS",
    spacingBefor: 120, // default is 100
    spacingAfter: 120, // default is 100
    spacingLine: 240, // default is 240
    spacingLineRule: 'atLeast', // default is atLeast
    indent: 100, // table indent, default is 0
    fixedLayout: true, // default is false
    borders: true, // default is false. if true, default border size is 4
    borderSize: 2, // To use this option, the 'borders' must set as true, default is 4
}
docx.createTable (table, tableStyle);

docx.putPageBreak()

// Create a new paragraph:
pObj = docx.createP()

pObj.addText('Simple')
pObj.addText(' with color', { color: '000088' })
pObj.addText(' and back color.', { color: '00ffff', back: '000088' })

pObj = docx.createP()

pObj.addText('Since ')
pObj.addText('officegen 0.2.12', {
    back: '00ffff',
    shdType: 'pct12',
    shdColor: 'ff0000'
}) // Use pattern in the background.
pObj.addText(' you can do ')
pObj.addText('more cool ', { highlight: true }) // Highlight!
pObj.addText('stuff!', { highlight: 'darkGreen' }) // Different highlight color.

pObj = docx.createP()

pObj.addText('Even add ')
pObj.addText('external link', { link: 'https://github.com' })
pObj.addText('!')

pObj = docx.createP()

pObj.addText('Bold + underline', { bold: true, underline: true })

pObj = docx.createP({ align: 'center' })

pObj.addText('Center this text', {
    border: 'dotted',
    borderSize: 12,
    borderColor: '88CCFF'
})

pObj = docx.createP()
pObj.options.align = 'right'

pObj.addText('Align this text to the right.')

pObj = docx.createP()

pObj.addText('Those two lines are in the same paragraph,')
pObj.addLineBreak()
pObj.addText('but they are separated by a line break.')

docx.putPageBreak()

pObj = docx.createP()

pObj.addText('Fonts face only.', { font_face: 'Arial' })
pObj.addText(' Fonts face and size.', { font_face: 'Arial', font_size: 40 })

docx.putPageBreak()

pObj = docx.createP()

// We can even add images:
pObj.addImage('knifeblock.jpg')



// Let's generate the Word document into a file:
let out = fs.createWriteStream('example.docx')

out.on('error', function(err) {
    console.log(err)
})

// Async call to generate the output file:
docx.generate(out)
