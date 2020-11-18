var officegen = require('officegen')
// var OfficeChart = require('../lib/officechart.js')
var async = require('async')

var fs = require('fs')
var path = require('path')

var pptx = officegen('pptx')

var outDir = path.join(__dirname, './')

var slide

pptx.on('finalize', function (written) {
  console.log(
    'Finish to create a PowerPoint file.\nTotal bytes created: ' +
      written +
      '\n'
  )

  // clear the temporatory files
})

pptx.on('error', function (err) {
  console.log(err)
})

pptx.setDocTitle('Sample PPTX Document')

var chartsData = [
 
  {
    title: 'My production',
    renderType: 'pie',
    data: [
      {
        name: 'Oil',
        labels: [
          'Czech Republic',
          'Ireland',
          'Germany',
          'Australia',
          'Austria',
          'UK',
          'Belgium'
        ],
        values: [301, 201, 165, 139, 128, 99, 60],
        colors: [
          'ff0000',
          '00ff00',
          '0000ff',
          'ffff00',
          'ff00ff',
          '00ffff',
          '000000'
        ]
      }
    ]
  }
]

function generateOneChart(chartInfo, callback) {
  slide = pptx.makeNewSlide()
  slide.name = 'OfficeChart slide'
  slide.back = 'ffffff'
  slide.addChart(chartInfo, callback, callback)
}

function generateCharts(callback) {
  async.each(chartsData, generateOneChart, callback)
}

function finalize() {
  var out = fs.createWriteStream('Chart2.pptx')

  out.on('error', function (err) {
    console.log(err)
  })

  pptx.generate(out)
}

async.series(
  [
    generateCharts // new
  ],
  finalize
)
