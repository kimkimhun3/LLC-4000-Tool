// const { ipcRenderer } = require("electron");

var inputText = "";
var simpleB;
let newArr = [];
let decodedStringFde9 = [];
let decodedStringFde8 = []; 
function hexToAscii(hex) {
  let asciiString = "";
  for (let i = 0; i < hex.length; i += 2) {
    const hexCode = parseInt(hex.substr(i, 2), 16);
    asciiString += String.fromCharCode(hexCode);
  }
  return asciiString;
}

let selectedDirectory;
let fileContent = "";
function submitForm() {}

function openFile() {
  ipcRenderer.send("open-file-dialog");
}
function openSaveDialog() {
  if (document.getElementById("portNumber1").value === "") {
    showAlert("ポート番号を入力してください！");
  } else {
    ipcRenderer.send("open-save-dialog");
  }
}
function openSaveDialog2() {
  if (document.getElementById("portNumber2").value === "") {
    showAlert("ポート番号を入力してください！");
  } else {
    ipcRenderer.send("open-save-dialog2");
  }
}
function openSaveDialog3() {
  if (
    document.getElementById("portNumber1").value === "" &&
    document.getElementById("portNumber2").value === ""
  ) {
    showAlert("ポート番号を入力してください！");
  } else {
    ipcRenderer.send("open-save-dialog3");
  }
}
function openSaveDialog4() {
  if (
    document.getElementById("portNumber1").value === "" &&
    document.getElementById("portNumber2").value === ""
  ) {
    showAlert("ポート番号を入力してください！");
  } else {
    ipcRenderer.send("open-save-dialog4");
  }
}
function saveFile() {
  // const downloadBtn = document.getElementById("downloadBtn");
  const loadingOverlay = document.getElementById("loadingOverlay");
  handleConvert();
  const content = document.getElementById("fileContent").value;
  const content2 = document.getElementById("fileContent2").value;
  const content3 = document.getElementById("fileContent3").value;
  // const content4 = document.getElementById("fileContent4").value;
  if (
    !window.currentFilePath &&
    !window.currentFilePath2 &&
    !window.currentFilePath3 &&
    !window.currentFilePath4
  ) {
    showAlert("ファイルパスを入力してください！");
  } else {
    loadingOverlay.style.display = "flex";
    ipcRenderer.send("save-filesssss", {
      filePath: window.currentFilePath,
      content,
      filePath2: window.currentFilePath2,
      content2,
      filePath3: window.currentFilePath3,
      content3,
    });
    if (window.currentFilePath4) {
      createExcelFile(window.currentFilePath4);
    }
    else {
      setTimeout(() => {
        // Hide loading overlay when the download is complete
        loadingOverlay.style.display = "none";
        showAlert("ファイルダウンロードを保存しました");
        ipcRenderer.send("download-complete");
      }, 1500); // Adjust the duration based on your actual download time
    }
   
  }
}
ipcRenderer.on("file-content", (event, content) => {
  inputText = content;
});
ipcRenderer.on("file-created", (event, filePath) => {
  window.currentFilePath = filePath;
  const filePathElement = document.getElementById("created-file1");
  const filePathLink = document.createElement("a");
  filePathLink.href = "#"; // Set the href attribute to "#" or a valid URL
  filePathLink.textContent = filePath;
  filePathLink.classList.add("no-underline");
  // Add click event listener to the link
  filePathLink.addEventListener("click", (e) => {
    e.preventDefault();
    shell.showItemInFolder(filePath);
  });
  filePathElement.innerHTML = "";
  filePathElement.appendChild(filePathLink);
});

ipcRenderer.on("file-created-2", (event, filePath) => {
  window.currentFilePath2 = filePath;
  const filePathElement = document.getElementById("created-file2");
  const filePathLink = document.createElement("a");
  filePathLink.href = "#"; // Set the href attribute to "#" or a valid URL
  filePathLink.textContent = filePath;
  filePathLink.classList.add("no-underline");
  // Add click event listener to the link
  filePathLink.addEventListener("click", (e) => {
    e.preventDefault();
    shell.showItemInFolder(filePath);
  });
  filePathElement.innerHTML = "";
  filePathElement.appendChild(filePathLink);
});
ipcRenderer.on("file-created-3", (event, filePath) => {
  window.currentFilePath3 = filePath;
  const filePathElement = document.getElementById("created-file3");
  const filePathLink = document.createElement("a");
  filePathLink.href = "#"; // Set the href attribute to "#" or a valid URL
  filePathLink.textContent = filePath;
  filePathLink.classList.add("no-underline");
  // Add click event listener to the link
  filePathLink.addEventListener("click", (e) => {
    e.preventDefault();
    shell.showItemInFolder(filePath);
  });
  filePathElement.innerHTML = "";
  filePathElement.appendChild(filePathLink);
});

ipcRenderer.on("file-created-4", (event, filePath) => {
  window.currentFilePath4 = filePath;
  const filePathElement = document.getElementById("created-file4");
  const filePathLink = document.createElement("a");
  filePathLink.href = "#"; // Set the href attribute to "#" or a valid URL
  filePathLink.textContent = filePath;
  filePathLink.classList.add("no-underline");
  // Add click event listener to the link
  filePathLink.addEventListener("click", (e) => {
    e.preventDefault();
    shell.showItemInFolder(filePath);
  });
  filePathElement.innerHTML = "";
  filePathElement.appendChild(filePathLink);
});

ipcRenderer.on("file-path", (event, filePath) => {
  const filePathElement = document.getElementById("uploaded-file-path");
  const filePathLink = document.createElement("a");
  filePathLink.href = "#"; // Set the href attribute to "#" or a valid URL
  filePathLink.textContent = filePath;
  filePathLink.classList.add("no-underline");
  // Add click event listener to the link
  filePathLink.addEventListener("click", (e) => {
    e.preventDefault();
    shell.showItemInFolder(filePath);
  });
  filePathElement.innerHTML = "";
  filePathElement.appendChild(filePathLink);
});
function handleConvert() {
  const fileContent = document.getElementById("fileContent");
  const fileContent2 = document.getElementById("fileContent2");
  const fileContent3 = document.getElementById("fileContent3");
  const portNumber1 = document.getElementById("portNumber1").value;
  const portNumber2 = document.getElementById("portNumber2").value;
  const hexPort1 = parseInt(portNumber1).toString(16); //port1 convert to hexadecial
  const hexPort2 = parseInt(portNumber2).toString(16); //port2 convert to hexadecial
  const hexValues = inputText
    .split("\n")
    .map((line) => {
      const hexString = line.substring(6, 56).replace(/ /g, ""); // Extracting the hex values
      return hexString;
    })
    .flat();
  const resultArray = [];
  let currentSum = [];

  for (const hexString of hexValues) {
    let sum = "";
    if (hexString === "") {
      // Empty string encountered, add the current sum to the result array and reset for the next index
      resultArray.push(currentSum);
      currentSum = [];
    } else {
      // Convert each pair of hex characters to decimal and sum them
      sum += hexString;
      currentSum.push(sum);
    }
  }
  const groupFde8 = [];
  const groupFde9 = [];
  const allPort = [];
  const target = resultArray
    .filter((i) => i.length > 0)
    .map((item) => item.reduce((a, b) => a + b));
  for (const item of target) {
    const lastIndex = item.lastIndexOf('0a');
    const cleanedItem = lastIndex !== -1
      ? item.substring(0, lastIndex + 2).replace(/0+$/g, '')
      : item;
  
    if (item.includes(hexPort1)) {
      groupFde8.push(cleanedItem);
    } else if (item.includes(hexPort2)) {
      groupFde9.push(cleanedItem);
    }
  }

  // console.log("fde8", groupFde8);
  // console.log("fde9", groupFde9);
  for (const item of target) {
    if (item.includes(hexPort1) || item.includes(hexPort2)) {
      allPort.push(item);
    }
  }
  const decodeFde8 = groupFde8.map((item) => {
    const findFde8 = item.indexOf(hexPort1);
    return item.substr(findFde8 + 12);
  });
  const decodeFde9 = groupFde9.map((item) => {
    const findFde8 = item.indexOf(hexPort2);
    return item.substr(findFde8 + 12);
  });
  const allDecode = allPort.map(item => {
    const findPort1 = item.indexOf(hexPort1);
    const findPort2 = item.indexOf(hexPort2);
    const indexToUse = (findPort1 !== -1) ? findPort1 : (findPort2 !== -1) ? findPort2 : -1;
    if (indexToUse !== -1) {
      const lastIndex0a = item.lastIndexOf('0a');
      return (lastIndex0a !== -1)
        ? item.substring(indexToUse + 12, lastIndex0a)
        : item;
    } else {
      return null;
    }
  });

  // console.log("allDecoed: ",allDecode);

  decodedStringFde8 = decodeFde8.map(hexToAscii);
  decodedStringFde9 = decodeFde9.map(hexToAscii);
  // console.log("Full Data: ", decodedStringFde9);
  // for (let i = 0; i < decodedStringFde9.length; i++) {
  //   newArr.filter(element => !elementsToRemove.includes(element)).map(element => element.replace(/[^a-zA-Z0-9=,:-]/g, ''));
  // }
  const elementsToRemove = [
    'pLuWgCCcC2BB4D3ECrEB2EB4DcBEB5EC2EBEC5C',
    'ouWgCCCcBcC2B4DScECEBEB4DSEB5ECsEBECC5C'
  ];
  // console.log("All Fde9, ",decodedStringFde9);
  
  newArr = decodedStringFde9.map((element, index) => {
    for (const pattern of elementsToRemove) {
      element = element.replace(pattern, '');
    }
  

    // Convert sequence numbers to integers
    element = element.replace(/(\d+)\s/g, (match, group1) => `${parseInt(group1, 10)} `);
  
    return element.replace(/[^a-zA-Z0-9=,:-]/g, '');
  }).filter(element => element !== undefined);
  // console.log("New Array: ",newArr);
// New Array include all of fde9, also evt

  const decodedAll = allDecode.map(hexToAscii);
  // Assign decoded strings to global variables
  //console.log(typeof decodedStringFde8);
  decodedStringFde8 = decodedStringFde8
  .map(line => line.replace(/\x00/g, '').trim()).filter(line => line !== '').join('\n');
  //console.log(typeof decodedStringFde8);
  window.decodedStringFde8 = decodedStringFde8;
  window.decodedStringFde9 = decodedStringFde9
  .map((line) => line.trim())
  .filter((line) => line !== '')
  .join("\n");
  window.decodedAll = decodedAll
  .map(line => line.replace(/\x00/g, '').trim()).filter(line => line !== '').join('\n');
  // window.excelData = excelData;

  fileContent.value = window.decodedStringFde8;
  fileContent2.value = window.decodedStringFde9;
  fileContent3.value = window.decodedAll;
  // fileContent4.value = window.excelData;
}
function showAlert(message) {
  const alertOverlay = document.getElementById("alertOverlay");
  const alertMessage = document.getElementById("alertMessage");
  const alertCloseBtn = document.getElementById("alertCloseBtn");

  alertMessage.innerHTML = message.replace("\n", "<br>");
  alertOverlay.style.display = "flex";
  alertCloseBtn.addEventListener("click", () => {
    alertOverlay.style.display = 'none';
    closeAlert();
  });
  // Close the alert when the user presses Enter, Space, or Esc
  document.addEventListener("keydown", handleKeyDown);
}
function closeAlert() {
  const alertOverlay = document.getElementById("alertOverlay");
  alertOverlay.style.display = "none";
  // Remove the event listener to avoid potential memory leaks
  document.removeEventListener("keydown", handleKeyDown);
}
function handleKeyDown(event) {
  // Check if the pressed key is Enter (key code 13), Space (key code 32), or Esc (key code 27)
  if (event.keyCode === 13 || event.keyCode === 32 || event.keyCode === 27) {
    closeAlert();
  }
}

// function parseEventData(dataString) {
//   const eventData = {};
//   const keyValuePairs = dataString.split(':').pop().split(',');

//   keyValuePairs.forEach((pair) => {
//     const [key, value] = pair.split('=');
//     eventData[key] = parseInt(value, 10);
//   });

//   return eventData;
// }

function parseEventData(dataString) {
  const eventData = {};
  const keyValuePairs = dataString.split(':').pop().split(',');

  keyValuePairs.forEach((pair) => {
    const [key, value] = pair.split('=');
    eventData[key] = parseInt(value, 10);
  });

  // Extract and add "evt" value
  const evtValue = dataString.split(':')[0].split('=')[1];
  eventData['evt'] = parseInt(evtValue, 10);

  return eventData;
}



function getSelectedColumns() {
  const selectedColumns = [];
  if (document.getElementById('checkboxJT').checked) selectedColumns.push('JT');
  if (document.getElementById('checkboxJTA').checked) selectedColumns.push('JT-A');
  if (document.getElementById('checkboxJTSDA').checked) selectedColumns.push('JT-SDA');
  if (document.getElementById('checkboxRTT').checked) selectedColumns.push('RTT');
  if (document.getElementById('checkboxRTTA').checked) selectedColumns.push('RTT-A');
  if (document.getElementById('checkboxRTTSDA').checked) selectedColumns.push('RTT-SDA');
  if (document.getElementById('checkboxRTTSTRD').checked) selectedColumns.push('RTT-STRD');
  if (document.getElementById('checkboxRTTLTRD').checked) selectedColumns.push('RTT-LTRD');
  if (document.getElementById('checkboxPLOST').checked) selectedColumns.push('PLOST');

  return selectedColumns;
}
function getUncheckedColumns() {
  const uncheckedColumns = [];
  if (!document.getElementById('checkboxJT').checked) uncheckedColumns.push('JT');
  if (!document.getElementById('checkboxJTA').checked) uncheckedColumns.push('JT-A');
  if (!document.getElementById('checkboxJTSDA').checked) uncheckedColumns.push('JT-SDA');
  if (!document.getElementById('checkboxRTT').checked) uncheckedColumns.push('RTT');
  if (!document.getElementById('checkboxRTTA').checked) uncheckedColumns.push('RTT-A');
  if (!document.getElementById('checkboxRTTSDA').checked) uncheckedColumns.push('RTT-SDA');
  if (!document.getElementById('checkboxRTTSTRD').checked) uncheckedColumns.push('RTT-STRD');
  if (!document.getElementById('checkboxRTTLTRD').checked) uncheckedColumns.push('RTT-LTRD');
  if (!document.getElementById('checkboxPLOST').checked) uncheckedColumns.push('PLOST');

  return uncheckedColumns;
}

function createExcelFile(filePath) {
  // Assuming newArr is your array of data
  const dataObjects = newArr
    .filter((dataString) => !dataString.includes('pLuWgCCcC2BB4D3ECrEB2EB4DcBEB5EC2EBEC5C') && !dataString.includes('ouWgCCCcBcC2B4DScECEBEB4DSEB5ECsEBECC5C'))
    .map(parseEventData);
  if (!dataObjects || dataObjects.length === 0) {
    console.error('Data objects are undefined or have length 0.');
    return;
  }
  // console.log("Data Object: ",dataObjects);
  //const evtValues = dataObjects.map((data) => data.evt);
  // const evtValues = dataObjects.map((data) => (data.evt === 0 ? "" : data.evt));
  // console.log("ONLY evt: ",evtValues);


  // Extract all unique property names from all data points
  //const allProperties = Array.from(new Set(dataObjects.flatMap((item) => Object.keys(item))));

  // Prepare data for xlsx-chart for all data
  // const chartData = {};
  // allProperties.forEach((columnName) => {
  //   chartData[columnName] = dataObjects.map((item) => (item && item[columnName] !== undefined ? item[columnName] : null));
  // });
  // console.log("Chart Data: ",chartData);


  /* Each Graph one by one */
  const jtChartData = {
    JT: dataObjects.map((item) => (item && item['JT'] !== undefined ? item['JT'] : null)),
  }
  const jtAChartData = {
    'JT-A': dataObjects.map((item) => (item && item['JT-A'] !== undefined ? item['JT-A'] : null)),
  };
  const jtSDAChartData = {
    'JT-SDA': dataObjects.map((item) => (item && item['JT-SDA'] !== undefined ? item['JT-SDA'] : null)),
  };
  const rttChartData = {
    RTT: dataObjects.map((item) => (item && item['RTT'] !== undefined ? item['RTT'] : null)),
  };
  const rttAChartData = {
    'RTT-A': dataObjects.map((item) => (item && item['RTT-A'] !== undefined ? item['RTT-A'] : null)),
  };
  const rttSDAChartData = {
    'RTT-SDA': dataObjects.map((item) => (item && item['RTT-SDA'] !== undefined ? item['RTT-SDA'] : null)),
  };
  const rttSTRDChartData = {
    'RTT-STRD': dataObjects.map((item) => (item && item['RTT-STRD'] !== undefined ? item['RTT-STRD'] : null)),
  };
  const rttLTRDChartData = {
    'RTT-LTRD': dataObjects.map((item) => (item && item['RTT-LTRD'] !== undefined ? item['RTT-LTRD'] : null)),
  };
  const plostChartData = {
    PLOST: dataObjects.map((item) => (item && item['PLOST'] !== undefined ? item['PLOST'] : null)),
  }
  // const evtChartData = {
  //   evt: dataObjects.map((item) => (item && item['evt'] !== undefined ? item['evt'] : null)),
  // }
  // const evtChartData = {
  //   evt: dataObjects.map((item) => (item && item['evt'] !== undefined ? (item['evt'] === 0 ? " " : item['evt']) : null)),
  // };

  const filteredEvtData = {
    "PLOST-F": dataObjects.map((item) => (item && item['evt'] === 1 ? item['evt'] : " ")),
    "JT-F": dataObjects.map((item) => (item && item['evt'] === 2 ? item['evt'] : " ")),
    "RTT-SDA-F": dataObjects.map((item) => (item && item['evt'] === 4 ? item['evt'] : " ")),
    "RTT-STRD-F": dataObjects.map((item) => (item && item['evt'] === 8 ? item['evt'] : " ")),
    "RTT-LTRD-F": dataObjects.map((item) => (item && item['evt'] === 10 ? item['evt'] : " ")),
    "PLOST-M": dataObjects.map((item) => (item && item['evt'] === 20 ? item['evt'] : " ")),
    "JT-M": dataObjects.map((item) => (item && item['evt'] === 40 ? item['evt'] : " ")),
    "RTT-SDA-M": dataObjects.map((item) => (item && item['evt'] === 80 ? item['evt'] : " ")),
    "RTT-STRD-M": dataObjects.map((item) => (item && item['evt'] === 100 ? item['evt'] : " ")),
    "RTT-LTRD-M": dataObjects.map((item) => (item && item['evt'] === 200 ? item['evt'] : " ")),
  };
  console.log("FilteredEvt Data: ",filteredEvtData);  
  const nonEmptyArrays = [];
  // Helper function to check if an array has values other than " ", zero, or empty strings
  const hasNonEmptyValues = (array) => array.some((value) => value !== " " && value !== 0 && value !== "");

  // Check each property in filteredEvtData
  Object.entries(filteredEvtData).forEach(([propertyName, array]) => {
    if (hasNonEmptyValues(array)) {
      nonEmptyArrays.push(propertyName);
    }
  });
  const xlsxChart = new XLSXChart();
  const selectedColumns = getSelectedColumns();
  const newArrayData = {};
  selectedColumns.forEach((columnName) => {
    newArrayData[columnName] = dataObjects.map((item) =>
      item && item[columnName] !== undefined ? item[columnName] : null
    );
  });
  // Object.assign(newArrayData,filteredArr);
  // console.log("Type of NEW DATA: ", typeof newArrayData);
  // console.log("NEW DATA: ",newArrayData);
  const selectedAndEvt = selectedColumns.concat(nonEmptyArrays);
  // console.log("Selected + Evt data: ",selectedAndEvt);
  // console.log("Selected Column: ",selectedColumns);
  //love
  // Add column data to the data object
  var datas = {};
  let maxValue = 0
  var dataNamesArray = [];
  selectedColumns.forEach(title => {
    datas[title] = {
      "chart": "column",
    };
    newArrayData[title].forEach((value, index) => {
      var dataKey = index + 1;
      if (value !== 0 && value !== "") {
        if (maxValue < value) {
          maxValue = value + maxValue/20;
        }
        datas[title][dataKey] = value;
      }
    });
    // Check if there are properties other than 'chart'
    if (Object.keys(datas[title]).length > 1) {
      dataNamesArray.push(title);
    } else {
      // If no properties, remove the empty object
      delete datas[title];
    }
  });

  nonEmptyArrays.forEach(title => {
    datas[title] = {
      "chart": "line",
    };
    filteredEvtData[title].forEach((value, index) => {
      var dataKey = index + 1;
      // console.log(value !== 0 && value !== "")
      if (value !== 0 && value !== "") {
        datas[title][dataKey] = value === " " ? value : maxValue + maxValue/20
      }
    });
    // Check if there are properties other than 'chart'
    if (Object.keys(datas[title]).length > 1) {
      dataNamesArray.push(title);
    } else {
      // If no properties, remove the empty object
      delete datas[title];
    }
  });
  console.log("dataNames: ",dataNamesArray);
  console.log("Data Data: ",datas);
  const fieldDataset = dataObjects.map((_, i) => i+1) //start from 1 rather than 0

  // console.log("Field Dataset: ,",fieldDataset);
  const ourOpts = {
    charts: [
      {
        position: {
          fromColumn: 1,
          toColumn: 28,
          fromRow: 22,
          toRow: 42,
        },
        customColors: {
          points: {
              "PLOST": {
                  "PLOST": 'ff0000',
              },
          },
          series: {
              "PLOST": {
                  fill: 'ff0000',
                  line: 'ff0000',
              }
          }
        },
        chart: 'column',
        titles: ['PLOST'],
        fields: fieldDataset,
        data: plostChartData,
        chartTitle: 'PLOST Chart',
        lineWidth: 0.2,
      },
      {
        position: {
          fromColumn: 1,
          toColumn: 28,
          fromRow: 1,
          toRow: 21,
        },
        customColors: {
          points: {
              "PLOST": {
                  "PLOST": 'ff0000',
              },
              "JT-A": {
                "JT-A": '808080'
              },
              "RTT-A": {
                "RTT-A": '097969'
              },
              "PLOST-F": {
                "PLOST-F": 'ffffff'
              },
              "JT-F": {
                "JT-F": 'ffffff'
              },
              "RTT-SDA-F": {
                "RTT-SDA-F": 'ffffff'
              },
              "RTT-STRD-F": {
                "RTT-STRD-F": 'ffffff'
              },
              "RTT-LTRD-F": {
                "RTT-LTRD-F": 'ffffff'
              },
              "PLOST-M": {
                "PLOST-M": 'ffffff'
              },
              "JT-M": {
                "JT-M": 'ffffff'
              },
              "RTT-SDA-M": {
                "RTT-SDA-M": 'ffffff'
              },
              "RTT-STRD-M": {
                "RTT-STRD-M": 'ffffff'
              },
              "RTT-LTRD-M": {
                "RTT-LTRD-M": 'ffffff'
              }
          },
          series: {
              "PLOST": {
                  fill: 'ff0000',
                  line: 'ff0000',
              },
              "JT-A": {
                fill: '808080',
                line: '808080'
              },
              "RTT-A": {
                fill: '097969',
                line: '097969'
              },
              "PLOST-F": {
                fill: 'ffffff',
                line: 'ffffff'
              },
              "JT-F": {
                fill: 'ffffff',
                line: 'ffffff'
              },
              "RTT-SDA-F": {
                fill: 'ffffff',
                line: 'ffffff'
              },
              "RTT-STRD-F": {
                fill: 'ffffff',
                line: 'ffffff'
              },
              "RTT-LTRD-F": {
                fill: 'ffffff',
                line: 'ffffff'
              },
              "PLOST-M": {
                fill: 'ffffff',
                line: 'ffffff'
              },
              "JT-M": {
                fill: 'ffffff',
                line: 'ffffff'
              },
              "RTT-SDA-M": {
                fill: 'ffffff',
                line: 'ffffff'
              },
              "RTT-STRD-M": {
                fill: 'ffffff',
                line: 'ffffff'
              },
              "RTT-LTRD-M": {
                fill: 'ffffff',
                line: 'ffffff'
              }
            }
        },
        titles: selectedAndEvt,
        fields: fieldDataset,
        data: datas,
        chartTitle: 'All Data Chart'
      },
      {
        position: {
          fromColumn: 1,
          toColumn: 28,
          fromRow: 43,
          toRow: 63,
        },
        chart: 'column',
        titles: ['JT'],
        fields: fieldDataset,
        data: jtChartData,
        chartTitle: 'JT Chart',
        lineWidth: 0.2,
      },
      {
        position: {
          fromColumn: 1,
          toColumn: 28,
          fromRow: 64,
          toRow: 84,
        },
        chart: 'column',
        titles: ['JT-A'],
        fields: fieldDataset,
        data: jtAChartData,
        chartTitle: 'JT-A Chart',
        lineWidth: 0.2,
      },
      {
        position: {
          fromColumn: 1,
          toColumn: 28,
          fromRow: 85,
          toRow: 105,
        },
        chart: 'column',
        titles: ['JT-SDA'],
        fields: fieldDataset,
        data: jtSDAChartData,
        chartTitle: 'JT-SDA Chart',
        lineWidth: 0.2,
      },
      {
        position: {
          fromColumn: 1,
          toColumn: 28,
          fromRow: 106,
          toRow: 126,
        },
        chart: 'column',
        titles: ['RTT'],
        fields: fieldDataset, // Use allProperties for fields
        data: rttChartData,
        chartTitle: 'RTT Chart',
        lineWidth: 0.2,
      },
      {
        position: {
          fromColumn: 1,
          toColumn: 28,
          fromRow: 127,
          toRow: 147,
        },
        chart: 'column',
        titles: ['RTT-A'],
        fields: fieldDataset,
        data: rttAChartData,
        chartTitle: 'RTT-A Chart',
        lineWidth: 0.2,
      },
      {
        position: {
          fromColumn: 1,
          toColumn: 28,
          fromRow: 148,
          toRow: 168,
        },
        chart: 'column',
        titles: ['RTT-SDA'],
        fields: fieldDataset,
        data: rttSDAChartData,
        chartTitle: 'RTT-SDA Chart',
        lineWidth: 0.2,
      },
      {
        position: {
          fromColumn: 1,
          toColumn: 28,
          fromRow: 169,
          toRow: 189,
        },
        chart: 'column',
        titles: ['RTT-STRD'],
        fields: fieldDataset,
        data: rttSTRDChartData,
        chartTitle: 'RTT-STRD Chart',
        lineWidth: 0.2,
      },
      {        
        position: {
        fromColumn: 1,
        toColumn: 28,
        fromRow: 190,
        toRow: 210,
      },
        chart: 'column',
        titles: ['RTT-LTRD'],
        fields: fieldDataset,
        data: rttLTRDChartData,
        chartTitle: 'RTT-LTRD Chart',
        lineWidth: 0.2,
      },
      // {
      //   position: {
      //     fromColumn: 1,
      //     toColumn: 28,
      //     fromRow: 211,
      //     toRow: 231
      //   },
      //   chart: "line",
      //   titles: [
      //     "Price",
      //     "Number"
      //   ],
      //   fields: [
      //     "Apple",
      //     "Blackberry",
      //     "Strawberry",
      //     "Cowberry"
      //   ],
      //   data: {
      //     "Price": {
      //       "Apple": 10,
      //       "Blackberry": 5,
      //       "Strawberry": 15,
      //       "Cowberry": 20
      //     },
      //     "Number": {
      //       "Apple": 5,
      //       "Blackberry": 2,
      //       "Strawberry": 9,
      //       "Cowberry": 3
      //     }
      //   },
      //   chartTitle: "Scatter chart"
      // }
    
      
    ],
  };
  xlsxChart.generate(ourOpts, function (err, data) {
  const loadingOverlay = document.getElementById("loadingOverlay");
    if (err) {
      console.error(err);
    } else {
      loadingOverlay.style.display = "flex";
      fs.writeFileSync(filePath, data);
      loadingOverlay.style.display = "none";
      showAlert("ファイルダウンロードを保存しました");
      console.log('Excel file with line chart created successfully at:', filePath);
    }
  });
}