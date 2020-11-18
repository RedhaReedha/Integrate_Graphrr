/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest

//import * as officegen from 'officegen'

////var officegen = require('officegen')
var officegen = require('officegen')
var pptx = officegen('pptx')
// var pptx = officegen('pptx')
//var async1 = require('async')

Office.onReady(info => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});



export async function run() {
  console.log("rida", pptx);
 
  if(pptx){ 
    //appedOK 
    Office.context.document.setSelectedDataAsync(

      "OK !",
      {
        coercionType: Office.CoercionType.Text
      },
      result => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(result.error.message);
        }
      }
    );
  } else { 
    //apendNope 
    Office.context.document.setSelectedDataAsync(

      " Nope !",
      {
        coercionType: Office.CoercionType.Text
      },
      result => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(result.error.message);
        }
      }
    );
  }
  
  const slide = pptx.makeNewSlide();

  /**
   * Insert your PowerPoint code here
   */
/*  var chartsData = [
 
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
  slide = pptx.makeNewSlide()
  pptx.addChart(chartsData)*/
 /* function generateCharts(callback) {
    async1.each(chartsData, generateOneChart, callback)
  }
  function generateOneChart(chartInfo, callback) {
    //slide = pptx.makeNewSlide()
    //slide.name = 'OfficeChart slide'
    //slide.back = 'ffffff'
    pptx.addChart(chartInfo, callback, callback)
  }
  async1.series(

    [
      generateCharts // new
    ]
  )  */

  /*Office.context.document.setSelectedDataAsync(

    "Hello testing !",
    {
      coercionType: Office.CoercionType.Text
    },
    result => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error(result.error.message);
      }
    }
  );*/
}
