var inputText = "";
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
  handleConvert(true);
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
// ipcRenderer.on("file-created", (event, filePath) => {
//   window.currentFilePath = filePath;
//   const filePathElement = document.getElementById("created-file1");
//   const filePathLink = document.createElement("a");
//   filePathLink.href = "#"; // Set the href attribute to "#" or a valid URL
//   filePathLink.textContent = filePath;
//   filePathLink.classList.add("no-underline");
//   // Add click event listener to the link
//   filePathLink.addEventListener("click", (e) => {
//     e.preventDefault();
//     shell.showItemInFolder(filePath);
//   });
//   filePathElement.innerHTML = "";
//   filePathElement.appendChild(filePathLink);
// });

// ipcRenderer.on("file-created-2", (event, filePath) => {
//   window.currentFilePath2 = filePath;
//   const filePathElement = document.getElementById("created-file2");
//   const filePathLink = document.createElement("a");
//   filePathLink.href = "#"; // Set the href attribute to "#" or a valid URL
//   filePathLink.textContent = filePath;
//   filePathLink.classList.add("no-underline");
//   // Add click event listener to the link
//   filePathLink.addEventListener("click", (e) => {
//     e.preventDefault();
//     shell.showItemInFolder(filePath);
//   });
//   filePathElement.innerHTML = "";
//   filePathElement.appendChild(filePathLink);
// });
// ipcRenderer.on("file-created-3", (event, filePath) => {
//   window.currentFilePath3 = filePath;
//   const filePathElement = document.getElementById("created-file3");
//   const filePathLink = document.createElement("a");
//   filePathLink.href = "#"; // Set the href attribute to "#" or a valid URL
//   filePathLink.textContent = filePath;
//   filePathLink.classList.add("no-underline");
//   // Add click event listener to the link
//   filePathLink.addEventListener("click", (e) => {
//     e.preventDefault();
//     shell.showItemInFolder(filePath);
//   });
//   filePathElement.innerHTML = "";
//   filePathElement.appendChild(filePathLink);
// });

// ipcRenderer.on("file-created-4", (event, filePath) => {
//   window.currentFilePath4 = filePath;
//   const filePathElement = document.getElementById("created-file4");
//   const filePathLink = document.createElement("a");
//   filePathLink.href = "#"; // Set the href attribute to "#" or a valid URL
//   filePathLink.textContent = filePath;
//   filePathLink.classList.add("no-underline");
//   // Add click event listener to the link
//   filePathLink.addEventListener("click", (e) => {
//     e.preventDefault();
//     shell.showItemInFolder(filePath);
//   });
//   filePathElement.innerHTML = "";
//   filePathElement.appendChild(filePathLink);
// });

ipcRenderer.on("file-created", (event, filePath) => {
  window.currentFilePath = filePath;
  displayFilePath(filePath, "created-file1");
});
ipcRenderer.on("file-created-2", (event, filePath) => {
  window.currentFilePath2 = filePath;
  displayFilePath(filePath, "created-file2");
});
ipcRenderer.on("file-created-3", (event, filePath) => {
  window.currentFilePath3 = filePath;
  displayFilePath(filePath, "created-file3");
});
ipcRenderer.on("file-created-4", (event, filePath) => {
  window.currentFilePath4 = filePath;
  displayFilePath(filePath, "created-file4");
});

function displayFilePath(filePath, elementId) {
  const filePathElement = document.getElementById(elementId);
  const filePathLink = document.createElement("a");
  filePathLink.href = "#";
  filePathLink.textContent = filePath;
  filePathLink.classList.add("no-underline");

  filePathLink.addEventListener("click", (e) => {
    e.preventDefault();
    shell.showItemInFolder(filePath);
  });
  filePathElement.innerHTML = "";
  filePathElement.appendChild(filePathLink);
}




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
function handleConvert(isFromSaveFile = false) {
  const fileContent = document.getElementById("fileContent");
  const fileContent2 = document.getElementById("fileContent2");
  const fileContent3 = document.getElementById("fileContent3");
  const portNumber1 = document.getElementById("portNumber1").value;
  const portNumber2 = document.getElementById("portNumber2").value;
  const fde9InputStart = document.getElementById('fde9InputStart').value;
  const fde9InputEnd = document.getElementById('fde9InputEnd').value;
  // console.log("Start: ", fde9InputStart);
  // console.log("End: ", fde9InputStart);
  const hexPort1 = parseInt(portNumber1).toString(16); //port1 convert to hexadecial
  const hexPort2 = parseInt(portNumber2).toString(16); //port2 convert to hexadecial
  console.log("Port number1: ",portNumber1)
  console.log("Port number2", portNumber2)
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

  // console.log("hexValues :",hexValues );
  // console.log("Result array: ",resultArray);




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

  decodedStringFde8 = decodeFde8.map(hexToAscii);
  decodedStringFde9 = decodeFde9.map(hexToAscii);
  const maxStringFde9 = decodedStringFde9.length
  console.log("All Fde9, ", maxStringFde9);
  document.getElementById('maxLengthDisplay').innerText = `Max Length: ${decodedStringFde9.length}`;
  if (!isFromSaveFile) {
    document.getElementById('fde9InputEnd').value = decodedStringFde9.length;
  }
  // document.getElementById('fde9InputEnd').value = decodedStringFde9.length;


  // const finaldecodedStringFde9 = decodedStringFde9.slice(fde9InputStart, fde9InputEnd);
  const start = parseInt(fde9InputStart, 10);
  const end = parseInt(fde9InputEnd, 10);
  let finaldecodedStringFde9; // = (fde9InputEnd > decodedStringFde9.length || fde9InputStart >= fde9InputEnd || fde9InputStart <= 0 || !fde9InputStart || !fde9InputEnd ) ? decodedStringFde9 : decodedStringFde9.slice(fde9InputStart - 1, fde9InputEnd);
  if (
    end > decodedStringFde9.length ||    // End exceeds length
    start >= end ||                      // Start is greater than or equal to End
    start <= 0 ||                        // Start is zero or negative
    isNaN(start) ||                      // Start is not a valid number
    isNaN(end)                           // End is not a valid number
  ) {
    finaldecodedStringFde9 = decodedStringFde9;
  } else {
    finaldecodedStringFde9 = decodedStringFde9.slice(start - 1, end);
  }

  console.log("Final length: ", finaldecodedStringFde9.length)
  const elementsToRemove = [
    'pLuWgCCcC2BB4D3ECrEB2EB4DcBEB5EC2EBEC5C',
    'ouWgCCCcBcC2B4DScECEBEB4DSEB5ECsEBECC5C'
  ];

  // const maxBrValue = decodedStringFde9["MAX-BR"];

  // const firstMaxBrValue = +decodedStringFde9.map(item => (item.match(/MAX-BR=(\d+)/) || [])[1]).find(Boolean);
  // console.log("MAX-BR:   ", firstMaxBrValue);
  
  newArr = finaldecodedStringFde9.map((element, index) => {
    for (const pattern of elementsToRemove) {
      element = element.replace(pattern, '');
    }
    // Convert sequence numbers to integers
    element = element.replace(/(\d+)\s/g, (match, group1) => `${parseInt(group1, 10)} `);
    return element.replace(/[^a-zA-Z0-9=,:-]/g, '');
  }).filter(element => element !== undefined);
  const decodedAll = allDecode.map(hexToAscii);
  decodedStringFde8 = decodedStringFde8
  .map(line => line.replace(/\x00/g, '').trim()).filter(line => line !== '').join('\n');
  window.decodedStringFde8 = decodedStringFde8;
  window.finaldecodedStringFde9 = finaldecodedStringFde9
  .map((line) => line.trim())
  .filter((line) => line !== '')
  .join("\n");
  window.decodedAll = decodedAll
  .map(line => line.replace(/\x00/g, '').trim()).filter(line => line !== '').join('\n');
  // window.excelData = excelData;

  fileContent.value = window.decodedStringFde8;
  fileContent2.value = window.finaldecodedStringFde9;
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
}
function closeAlert() {
  const alertOverlay = document.getElementById("alertOverlay");
  alertOverlay.style.display = "none";
}

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
  if (document.getElementById('checkboxRTT').checked) selectedColumns.push('RTT');
  if (document.getElementById('checkboxRTTA').checked) selectedColumns.push('RTT-A');
  if (document.getElementById('checkboxRTTSDA').checked) selectedColumns.push('RTT-SDA');
  if (document.getElementById('checkboxRTTSTRD').checked) selectedColumns.push('RTT-STRD');
  if (document.getElementById('checkboxRTTLTRD').checked) selectedColumns.push('RTT-LTRD');
  if (document.getElementById('checkboxJT').checked) selectedColumns.push('JT');
  if (document.getElementById('checkboxJTA').checked) selectedColumns.push('JT-A');
  if (document.getElementById('checkboxJTSDA').checked) selectedColumns.push('JT-SDA');
  if (document.getElementById('checkboxPLOST').checked) selectedColumns.push('PLOST');
  return selectedColumns;
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

  /* Each Graph one by one */

  const jtAChartData = {
    'JT-A': dataObjects
      .map((item) => (item && item['JT-A'] !== undefined ? item['JT-A'] : null))
      .reduce((acc, value, index) => {
        acc[index + 1] = value;  // Start the index from 1
        return acc;
      }, {}),
  };
  // console.log("JT-A: ", jtAChartData)
  const jtSDAChartData = {
    'JT-SDA': dataObjects
      .map((item) => (item && item['JT-SDA'] !== undefined ? item['JT-SDA'] : null))
      .reduce((acc, value, index) => {
        acc[index + 1] = value; // Start index from 1
        return acc;
      }, {}),
  };
  const rttChartData = {
    RTT: dataObjects
      .map((item) => (item && item['RTT'] !== undefined ? item['RTT'] : null))
      .reduce((acc, value, index) => {
        acc[index + 1] = value; // Start index from 1
        return acc;
      }, {}),
  };
  const rttAChartData = {
    'RTT-A': dataObjects
      .map((item) => (item && item['RTT-A'] !== undefined ? item['RTT-A'] : null))
      .reduce((acc, value, index) => {
        acc[index + 1] = value; // Start index from 1
        return acc;
      }, {}),
  };
  const rttSDAChartData = {
    'RTT-SDA': dataObjects
      .map((item) => (item && item['RTT-SDA'] !== undefined ? item['RTT-SDA'] : null))
      .reduce((acc, value, index) => {
        acc[index + 1] = value; // Start index from 1
        return acc;
      }, {}),
  };
  const rttSTRDChartData = {
    'RTT-STRD': dataObjects
      .map((item) => (item && item['RTT-STRD'] !== undefined ? item['RTT-STRD'] : null))
      .reduce((acc, value, index) => {
        acc[index + 1] = value; // Start index from 1
        return acc;
      }, {}),
  };
  const rttLTRDChartData = {
    'RTT-LTRD': dataObjects
      .map((item) => (item && item['RTT-LTRD'] !== undefined ? item['RTT-LTRD'] : null))
      .reduce((acc, value, index) => {
        acc[index + 1] = value; // Start index from 1
        return acc;
      }, {}),
  };
const plostChartData = {
  PLOST: dataObjects
    .map((item) => (item && item['PLOST'] !== undefined ? item['PLOST'] : null))
    .reduce((acc, value, index) => {
      acc[index + 1] = value; // Start index from 1
      return acc;
    }, {}),
};
  const jtChartData = {
    JT: dataObjects
      .map((item) => (item && item['JT'] !== undefined ? item['JT'] : null))
      .reduce((acc, value, index) => {
        acc[index + 1] = value; // Start index from 1
        return acc;
      }, {}),
  };
  // const isAllPLOSTZero = plostChartData.PLOST.every(value => value === 0);
  const isAllPLOSTZero = Object.values(plostChartData.PLOST).every(value => value === 0);
  // console.log("PLOST chart data: ", plostChartData);


  //New function
  const filteredEvtData = {
    "PLOST-M": dataObjects.map((item) => (item && item['evt'] === 20 ? item['evt'] : 0)),
    "PLOST-F": dataObjects.map((item) => (item && item['evt'] === 1 ? item['evt'] : 0)),
    "JT-M": dataObjects.map((item) => (item && item['evt'] === 40 ? item['evt'] : 0)),
    "JT-F": dataObjects.map((item) => (item && item['evt'] === 2 ? item['evt'] : 0)),
    "RTT-M": dataObjects.map((item) => (item && item['evt'] === 80 ? item['evt'] : 0)),
    "RTT-F": dataObjects.map((item) => (item && item['evt'] === 4 ? item['evt'] : 0)),
    "RTT-STRD-M": dataObjects.map((item) => (item && item['evt'] === 100 ? item['evt'] : 0)),
    "RTT-STRD-F": dataObjects.map((item) => (item && item['evt'] === 8 ? item['evt'] : 0)),
    "RTT-LTRD-M": dataObjects.map((item) => (item && item['evt'] === 200 ? item['evt'] : 0)),
    "RTT-LTRD-F": dataObjects.map((item) => (item && item['evt'] === 10 ? item['evt'] : 0)),
  };
  // console.log("FilteredEvt Data: ",filteredEvtData);  
  const nonEmptyArrays = [];
  // Helper function to check if an array has values other than " ", zero, or empty strings
  const hasNonEmptyValues = (array) => array.some((value) => value !== " " && value !== 0 && value !== "");
  // Check each property in filteredEvtData
  Object.entries(filteredEvtData).forEach(([propertyName, array]) => {
    if (hasNonEmptyValues(array)) {
      nonEmptyArrays.push(propertyName);
    }
  });
  // console.log("All Event: ",nonEmptyArrays);

  const xlsxChart = new XLSXChart();
  const selectedColumns = getSelectedColumns();
  // Filter elements based on prefixes
  const hasColumn = (prefix) => selectedColumns.includes(prefix);

  // Your existing array filtering based on conditions
  const filteredNonEmptyArrays = nonEmptyArrays.filter((elem) => {
    const prefix = elem.replace(/-(F|M)$/, ''); // Remove the suffix to match with selectedColumns
    return hasColumn(prefix) && (elem.endsWith("-F") || elem.endsWith("-M"));
  });

  const newArrayData = {};
  selectedColumns.forEach((columnName) => {
    const arrayValues = dataObjects.map((item) =>
      item && item[columnName] !== undefined ? item[columnName] : null
    );
    // Check if the array name starts with "JT"
    if (columnName.startsWith("JT")) {
      newArrayData[columnName] = arrayValues.map((value) => value / 1000);
    } else {
      newArrayData[columnName] = arrayValues;
    }
  });

  //maybe check here.
  const selectedAndEvt = selectedColumns.concat(filteredNonEmptyArrays, "NOW-BR");
  if (isAllPLOSTZero) {
    const indexToRemove = selectedAndEvt.indexOf("PLOST");
    if (indexToRemove !== -1) {
      selectedAndEvt.splice(indexToRemove, 1);
    }
  }
  

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
    const hasNonChartProperties = Object.values(datas[title]).some(val => val !== "chart");
    if (hasNonChartProperties) {
      dataNamesArray.push(title);
    } else {
      // If no properties other than 'chart', remove the object if all values are 0
      const allValuesAreZero = Object.values(datas[title]).every(val => val === 0);
      if (!allValuesAreZero) {
        delete datas[title];
      }
    }
  });
  filteredNonEmptyArrays.forEach(title => {
    datas[title] = {
      "chart": "scatter",
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


  const nowBitrateChartData = {
    'NOW-BR': {
      "chart": "scatter",
      ...dataObjects.reduce((acc, item, index) => {
        acc[index + 1] = item && item['NOW-BR'] !== undefined ? item['NOW-BR'] : null;
        return acc;
      }, {}),
    },
  };
  Object.assign(datas, nowBitrateChartData);
  // if (Object.keys(datas['NOW-BR']).length > 1) {
  //   dataNamesArray.push('NOW-BR');
  // } else {
  //   delete datas['NOW-BR'];
  // }
  // Check for MAX-BR properties
  // if (Object.keys(datas['MAX-BR']).length > 1) {
  //   dataNamesArray.push('MAX-BR');
  // } else {
  //   // If no properties, remove the empty object
  //   delete datas['MAX-BR'];
  // }
  // console.log("Updated datas object: ", datas);

  const fieldDataset = dataObjects.map((_, i) => i+1) //start from 1
  const ourOpts = {
    charts: [
      {
        position: {
          fromColumn: 0,
          toColumn: 27,
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
          fromColumn: 0,
          toColumn: 29,
          fromRow: 1,
          toRow: 21,
        },
        customColors: {
          points: {
              "PLOST": {
                  "PLOST": 'ff0000',
              },
              "NOW-BR":{
                "NOW-BR": 'ff0000'
              },
              "JT": {
                "JT": '008000'
              },
              "RTT": {
                "RTT": '6699CC'
              }
          },
          series: {
              "PLOST": {
                  fill: 'ff0000',
                  line: 'ff0000',
              },
              "NOW-BR": {
                fill: 'ff0000',
                line: 'ff0000'
              },
              "JT": {
                fill: '008000',
                line: '008000'
              },
              "RTT": {
                fill: '6699CC',
                line: '6699CC'
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
          fromColumn: 0,
          toColumn: 27,
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
          fromColumn: 0,
          toColumn: 27,
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
          fromColumn: 0,
          toColumn: 27,
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
          fromColumn: 0,
          toColumn: 27,
          fromRow: 106,
          toRow: 126,
        },
        chart: 'column',
        titles: ['RTT'],
        fields: fieldDataset,
        data: rttChartData,
        chartTitle: 'RTT Chart',
        lineWidth: 0.2,
      },
      {
        position: {
          fromColumn: 0,
          toColumn: 27,
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
          fromColumn: 0,
          toColumn: 27,
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
          fromColumn: 0,
          toColumn: 27,
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
        fromColumn: 0,
        toColumn: 27,
        fromRow: 190,
        toRow: 210,
      },
        chart: 'column',
        titles: ['RTT-LTRD'],
        fields: fieldDataset,
        data: rttLTRDChartData,
        chartTitle: 'RTT-LTRD Chart',
        lineWidth: 0.2,
      }
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
      executePythonScript(filePath);
      showAlert("ファイルダウンロードを保存しました");
      console.log('Excel file with line chart created successfully at:', filePath);
    }
  });
  // Execute Python script after Excel file generation
  const { exec } = require('child_process');
  function executePythonScript(filePath) {
    // Execute the Python script
    const pythonScriptPath = 'python.exe'; // Update with the actual name of your Python script
    exec(`${pythonScriptPath} "${filePath}"`, (error, stdout, stderr) => {
        if (error) {
            console.error(`Error executing Python script: ${error}`);
            return;
        }
        console.log(`Python script output: ${stdout}`);
    });
  }


}
