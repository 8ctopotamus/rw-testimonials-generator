const fs = require('fs')
const XLSX = require('xlsx')
const workbook = XLSX.readFile('./test.xlsx')
const cells = workbook.Sheets.Sheet0

let combined = []
let firstNameCol
let lastNameCol

// combine the rows values
Object.keys(cells).forEach(cell => {
  let row = cell.slice(1)
  let col = cell.charAt(0)
  let val = cells[cell].v

  if (val === undefined) return

  // we need to determine which cols are for names.
  if (val === 'First Name') {
    firstNameCol = col
  }
  else if(val === 'Last Name') {
    lastNameCol = col
  }

  val = val.trim()

  if (col === firstNameCol) {
    combined[row] += '-' + val + ' '
  } else {
    combined[row] += val + ' '
  }
})

// turn into array
const combinedArr = Object.keys(combined).map(key => { return combined[key] })

// format testimonials
const formattedTestimonials = combinedArr.map((testimonial, i) => {
  // the first row is the list of questions.
  // we don't need this.
  if (i === 0) return

  // for some reason, each line started with 'undefined'
  // so we'll just slice that out real quick.
  testimonial = testimonial.slice(9)

  // remove blank testimonials
  if (testimonial.charAt(0) === '-') return

  return `<blockquote>${testimonial}</blockquote>\n`
})
.join('\n')

// generate the file
fs.writeFile('./testimonials.txt', formattedTestimonials, 'utf8', (err) => {
  if (err) throw err
  console.log('Your file is ready :)')
})
