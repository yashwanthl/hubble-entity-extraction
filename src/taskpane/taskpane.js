/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
var baseURL = 'https://hubbleentity.azurewebsites.net/';
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {    
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("btnGetEntities").onclick = extractEntities;
    if (testCORS()){
      extractEntities();
    }
  }
});

function testCORS(){
  let isSuccess = false
  let i = 1;
  while(!isSuccess && i <= 3) {
    console.log("Try: " + i);
    isSuccess = getBaseURLResponse();
    if (isSuccess){
      break;
    }
    i++;
  }
  return isSuccess
}

function getBaseURLResponse() {
  let url = baseURL
  let result = false;
  $.ajax({
    type: "GET",
    url: url,
    async: false,
    success: function(response) {
      console.log("Success - Test CORS");
      console.log(response);
      result = true;
    },
    error: function(request, status, error) {
      console.log("Fail - Test CORS");
      console.log(error);
      return false;
    }
  });

  return result;
}

function setupEntities(allEntities){
  console.log(allEntities);
  let filteredEntities = []
  allEntities.forEach(eachObj => {
    if (eachObj["entities"].length > 0){
      filteredEntities = filteredEntities.concat(eachObj["entities"])
    }
  });
  console.log(filteredEntities);
  filteredEntities.forEach(eachObj => {
    delete eachObj["color"];
    delete eachObj["end_char"];
    delete eachObj["start_char"]
  });
  console.log(filteredEntities);
  
  var p = {
    "ORG": ["HP", "Walmart"],
    "Date": ["Oct", "1/1/21"]
};

  let objLabel = {};
  filteredEntities.forEach(eachObj => {
    let key = eachObj["label"]
    if (objLabel.hasOwnProperty(key)) {
      let words = objLabel[key];
      if (!words.includes(eachObj["text"])){
        words.push(eachObj["text"])
      }
    }
    else {
      objLabel[key] = [eachObj["text"]]
    }
  });
  console.log(objLabel);
  htmlForSetUpEntities(objLabel);
}

function htmlForSetUpEntities(objLabel){
  let wholeHtml = ""
  for (let key in objLabel) {
    if (objLabel.hasOwnProperty(key)) {
      let words = objLabel[key];
      if (words.length > 0){
        let html = '<div style="padding: 0.75em;">';
        html += ('<div style="font-weight:bold">' + key + '</div>')
        html += '<div style="padding: 0.5em;">';
        html += '<ul style="list-style: square inside;">';
        words.forEach(eachWord => {
          html += ('<li>' + eachWord + '</li>');
        });
        html += '</ul>';
        html += '</div>';
        html += '</div>';
        wholeHtml += html;
      }
    }
  }
  $('#display-entities-result').empty();
  $('#display-entities-result').append(wholeHtml);
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
          entities = filterUnwantedLabels(entities);
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
    setupEntities(allEntities);
  });
}

function filterUnwantedLabels(entities){
  let newArray = entities.filter(function (el) {
    return el.label != "WORK_OF_ART" &&
          el.label != "CARDINAL"
  });
  newArray.forEach(element => {
    if (element["label"] == "MONTH_FIRST_DATE"){
      element["label"] = "DATE"
    }
  });
  return newArray;
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
  let url = baseURL + "extract?text=" + text;
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
