const SHEET_SOAL = 'SOAL';
const SHEET_NILAI = 'NILAI';
const SHEET_USER = 'USER';

function doGet(){

  return HtmlService
  .createHtmlOutputFromFile('index');

}

// LOGIN
function login(username,password){

  if(username == '' || password == ''){

    return {
      status:false
    };

  }

  let role = 'siswa';

  // USER KHUSUS ADMIN
  if(
    username.toLowerCase() == 'admin'
  ){

    role = 'admin';

  }

  return {

    status:true,

    nama:username,

    role:role

  };

}

// AMBIL SOAL
function getQuestions(){

  const ss =
  SpreadsheetApp.getActiveSpreadsheet();

  const sh =
  ss.getSheetByName(SHEET_SOAL);

  const data =
  sh.getDataRange().getValues();

  let hasil = [];

  for(let i=1;i<data.length;i++){

    hasil.push({

      mapel:data[i][0],

      soal:data[i][1],

      gambar:data[i][2],

      pilihan:[

        data[i][3],
        data[i][4],
        data[i][5],
        data[i][6]

      ],

      benar:data[i][7]

    });

  }

  return hasil;

}

// SIMPAN NILAI
function saveScore(nama,mapel,skor){

  const ss =
  SpreadsheetApp.getActiveSpreadsheet();

  let sh =
  ss.getSheetByName(SHEET_NILAI);

  if(!sh){

    sh =
    ss.insertSheet(SHEET_NILAI);

    sh.appendRow([

      'NAMA',
      'MAPEL',
      'SKOR',
      'TANGGAL'

    ]);

  }

  sh.appendRow([

    nama,
    mapel,
    skor,
    new Date()

  ]);

  return true;

}

function getLeaderboard(){

  const ss =
  SpreadsheetApp.getActiveSpreadsheet();

  const sh =
  ss.getSheetByName(SHEET_NILAI);

  if(!sh){

    return [];

  }

  const data =
  sh.getDataRange().getValues();

  if(data.length < 2){

    return [];

  }

  let hasil = [];

  for(let i=1;i<data.length;i++){

    hasil.push([

      data[i][0], // nama
      data[i][1], // mapel
      data[i][2]  // skor

    ]);

  }

  hasil.sort(function(a,b){

    return b[2] - a[2];

  });

  return hasil;

}

// SIMPAN SOAL
function saveQuestion(data){

  const ss =
  SpreadsheetApp.getActiveSpreadsheet();

  const sh =
  ss.getSheetByName(SHEET_SOAL);

  sh.appendRow([

    data.mapel,
    data.soal,
    data.gambar,
    data.a,
    data.b,
    data.c,
    data.d,
    data.benar

  ]);

  return true;

}

// AMBIL MAPEL
function getMapel(){

  const ss =
  SpreadsheetApp.getActiveSpreadsheet();

  const sh =
  ss.getSheetByName(SHEET_SOAL);

  const data =
  sh.getDataRange().getValues();

  let mapel = [];

  for(let i=1;i<data.length;i++){

    let nama =
    data[i][0];

    if(

      nama &&
      !mapel.includes(nama)

    ){

      mapel.push(nama);

    }

  }

  return mapel;

}

function saveUser(username,password){

  const ss =
  SpreadsheetApp.getActiveSpreadsheet();

  let sh =
  ss.getSheetByName('USER');

  // BUAT SHEET JIKA BELUM ADA
  if(!sh){

    sh = ss.insertSheet('USER');

    sh.appendRow([

      'USERNAME',
      'PASSWORD'

    ]);

  }

  sh.appendRow([

    username,
    password

  ]);

  return true;

}

function resetLeaderboard(){

  const ss =
  SpreadsheetApp.getActiveSpreadsheet();

  const sh =
  ss.getSheetByName('NILAI');

  const lastRow =
  sh.getLastRow();

  if(lastRow > 1){

    sh.deleteRows(

      2,
      lastRow - 1

    );

  }

  return true;

}

function hapusSemuaSoal(){

  const ss =
  SpreadsheetApp.getActiveSpreadsheet();

  const sh =
  ss.getSheetByName('SOAL');

  const lastRow =
  sh.getLastRow();

  // SISAKAN HEADER
  if(lastRow > 1){

    sh.deleteRows(

      2,
      lastRow - 1

    );

  }

  return true;

}