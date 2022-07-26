/* set up XMLHttpRequest */
let url = "./doc.xlsx";
let oReq = new XMLHttpRequest();
oReq.open("GET", url, true);
oReq.responseType = "arraybuffer";
let HVAC_test = true;


oReq.onload = function (e) {
  let arraybuffer = oReq.response;

  /* convert data to binary string */
  let data = new Uint8Array(arraybuffer);
  let arr = new Array();
  for (let i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
  let bstr = arr.join("");

  /* Call XLSX */
  let workbook = XLSX.read(bstr, { type: "binary" });

  /* DO SOMETHING WITH workbook HERE */
  let first_sheet_name = workbook.SheetNames[0];
  /* Get worksheet */
  let worksheet = workbook.Sheets[first_sheet_name];

  const arrTable = XLSX.utils.sheet_to_json(worksheet, { raw: true });
  console.log(arrTable);

  const renderSwitch = (state, info,) => {
    const span = document.querySelector('.title__span');
    const log = document.querySelector('.log');
    span.textContent = state;
    log.textContent = info;
  }

  const renderChannel = (Channel = 0) => {
    const channelSpan = document.querySelector('.channel__span');
    channelSpan.textContent = Channel;
  }

  const switchLogic = (EXT1, A1_IN, A2_IN, temp, Channel) => {
    console.log(Channel)
    if (A1_IN === true) {
      renderSwitch('включен', '');
      EXT1 = true;
      if (temp > 25) {
        renderChannel(800);
      } else {
        renderChannel(2000);
      }
      setTimeout((A2_IN) => {
        if (!A2_IN) {
          renderSwitch('выключен', 'ALARM A2_IN');
          renderChannel();
          EXT1 = false;
        }
      }, 20000)
    } else {
      renderSwitch('выключен', 'ALARM A1_IN');
      EXT1 = false;
    }

  }

  let EXT1, A1_IN, A2_IN, temp, Channel;

  arrTable.forEach((line) => {
    switch (line.__EMPTY) {
      case 'EXT1_R3A1':
        EXT1 = line.__EMPTY_3;
        if (EXT1 === 'true') {
          EXT1 = true
        } else {
          EXT1 = false
        }
        break;
      case 'A1_IN':
        A1_IN = line.__EMPTY_3;
        if (A1_IN === 'true') {
          A1_IN = true
        } else {
          EXT1 = false
        }
        break;
      case 'A2_IN':
        A2_IN = line.__EMPTY_3;
        if (A2_IN === 'true') {
          A2_IN = true
        } else {
          A2_IN = false
        }
        break;
      case '28-00000d6b460c':
        temp = Number(line.__EMPTY_3);
        break;
      case 'Channel 1':
        Channel = Number(line.__EMPTY_3);
        break;
    }
  });

  switchLogic(EXT1, A1_IN, A2_IN, temp, Channel);
}

oReq.send();