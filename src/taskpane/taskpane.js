/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    testCORS();
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("btnGetEntities").onclick = extractEntities;
  }
});

function testCORS() {
  let url = 'https://hubbleentity.azurewebsites.net/';
  // let url = 'https://hubbleentity.azurewebsites.net/extract?text=What is the price of Pixel 3&name=product annotaions';
  // let url = 'https://hubbleentity.azurewebsites.net/extract?text=What is the price of Pixel 3';
  // let url = "https://hubbleentity.azurewebsites.net/extract?text=Per our conversation, HP is aligned to help Walmart promote the What’s Your Color units (DeskJet 3722 – Blue, Purple and Pink) for another $10 off. To help fund this program, HP will fund $5.00 per unit for all on hand units to date (excluding units already sold) on top of the current front-end buy price received by Walmart ($43.66).";
  // let url = "https://hubbleinferenceapi.azurewebsites.net/extract/?text=Per our conversation, HP is aligned to help Walmart promote the What’s Your Color units (DeskJet 3722 – Blue, Purple and Pink) for another $10 off. To help fund this program, HP will fund $5.00 per unit for all on hand units to date (excluding units already sold) on top of the current front-end buy price received by Walmart ($43.66)."
  $.ajax({
    type: "GET",
    url: url,
    async: false,
    success: function(response) {
      console.log("Success - Test CORS");
      console.log(response);
    },
    error: function(request, status, error) {
      console.log("Fail - Test CORS");
      console.log(error);
    }
  });
}

function extractEntities() {
  Office.context.mailbox.item.body.getAsync("text", function callback(result) {
    let emailText = result.value;
    emailText = emailText.replace(/(\r\n|\n|\r)/gm, " "); // Removing line breaks in the text
    let sentences = emailText.split(". ");
    let allEntities = [];
    sentences.forEach(eachSentense => {
      eachSentense = eachSentense.trim();
      if (!eachSentense.endsWith(".")){
        eachSentense = eachSentense + "."
      }
      if (!isNullOrEmpty(eachSentense)) {
        let entities;
        let count = 0
        while (count <= 10 && (entities == undefined || entities == null)) {
          entities = getEntities(eachSentense);
          count++
        }
        let thisEntity = {
          text: eachSentense,
          entities: entities
        };
        if (thisEntity.entities.length >= 0) {
          thisEntity.entities.forEach(eachEnitity => {
            eachEnitity["color"] = getRandomColor();
          });
          allEntities.push(thisEntity);
        }
        console.log(thisEntity);
      }
    });
    buildHtmlForEntities(allEntities);
  });
}

function preprocess(emailText){
  let paras = emailText.split("\r");
  let filteredParas = []
  paras.forEach(eachPara => {
    eachPara = eachPara.trim();
    if (!isNullOrEmpty(eachPara)){
      eachPara = eachPara.replace(/(\r\n|\n|\r)/gm, ""); // Removing line breaks in the text
      filteredParas.push(eachPara)
    }
  });
  console.log(filteredParas);
  let sentences = [];

  filteredParas.forEach(eachPara => {
    let thisSentences = eachPara.split(". ");
    sentences = sentences.concat(thisSentences)
  });
  console.log(sentences);
  return sentences;
}

function buildHtmlForEntities(allEntities) {
  let html = "";
  allEntities.forEach(eachEntity => {
    html += buildHtmlForEntity(eachEntity);
  });
  $("#entityResults").empty();
  $("#entityResults").append(html);
}

function getEntities(text) {
  console.log("Getting entities for: " + text);
  let entities;
  let url = "https://hubbleentity.azurewebsites.net/extract?text=" + text;
  $.ajax({
    type: "GET",
    url: url,
    success: function(response) {
      if (response.status) {
        entities = response.Entities;
      }
    },
    error: function(request, status, error) {
      // console.log("Error in getting entities for this email text: " + text);
    },
    async: false
  });
  return entities;
}

function getRandomColor() {
  // storing all letter and digit combinations
  // for html color code
  let letters = "0123456789ABCDEF";

  // html color code starts with #
  let color = "#";

  // generating 6 times as HTML color code consist
  // of 6 letter or digits
  for (var i = 0; i < 6; i++) {
    color += letters[Math.floor(Math.random() * 16)];
  }

  let rgb = ColorLuminance(color)
  // return rgb;
  return color;
}


function ColorLuminance(hex, lum = -0.5) {
  // validate hex string
  hex = String(hex).replace(/[^0-9a-f]/gi, '');
  if (hex.length < 6) {
    hex = hex[0]+hex[0]+hex[1]+hex[1]+hex[2]+hex[2];
  }
  lum = lum || 0;

  // convert to decimal and change luminosity
  var rgb = "#", c, i;
  for (i = 0; i < 3; i++) {
    c = parseInt(hex.substr(i*2,2), 16);
    c = Math.round(Math.min(Math.max(0, c + (c * lum)), 255)).toString(16);
    rgb += ("00"+c).substr(c.length);
  }

  return rgb;
}

String.prototype.insert = function(index, string) {
  if (index > 0) {
    return this.substring(0, index) + string + this.substring(index, this.length);
  }

  return string + this;
};

function buildHtmlForEntity(entity) {
  let text = entity.text;
  let entities = entity.entities;
  let l = entities.length;
  if (l > 0) {
    for (i = l - 1; i >= 0; i--) {
      let eachEntity = entities[i];
      text = text.insert(
        eachEntity.end_char,
        '</span><span>]<sub style="font-weight: bold; font-size:0.75em">' + eachEntity.label + "</sub></span>"
      );
      text = text.insert(eachEntity.start_char, '[<span style="color:' + eachEntity.color + '; font-weight:bold">');
    }
  }
  text = "<div class='entity-sentence'>" + text + "</div>";
  return text;
}

function isNullOrEmpty(text) {
  if (text.trim() === "" || text == undefined || text === null) return true;
  return false;
}
