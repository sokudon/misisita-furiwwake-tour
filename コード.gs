
//共有が必要

//  1. Enter sheet name where data is to be written below           
        
//  2. Run > setup
//
//  3. Publish > Deploy as web app 
//    - enter Project Version name and click 'Save New Version' 
//    - set security level and enable service (most likely execute as 'me' and access 'anyone, even anonymously) 
//
//  4. Copy the 'Current web app URL' and post this in your form/script action 
//
//  5. Insert column names on your destination sheet matching the parameter names of the data you are passing in (exactly matching case)

var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service

// If you don't want to expose either GET or POST methods you can comment out the appropriate function
function doPost(e) {
  // shortly after my original solution Google announced the LockService[1]
  // this prevents concurrent access overwritting data
  // [1] http://googleappsdeveloper.blogspot.co.uk/2011/10/concurrency-and-google-apps-script.html
  // we want a public lock, one that locks for all invocations
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.

    var sname="100位";//シート名
    
    
  try {
    // next set where we write the data - you could write to multiple/alternate destinations
    var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    
     //JSONでーたを受け取り
     var JSON_DATA= JSON.parse(e.parameter["BD_JSON"]);
     var log="(*^_^*)でーた入力にご協力ありがとうございます(*^_^*)～\r\nごっごるしーと受け皿でのろぐをお伝えします～\r\n\r\n";        
     log= log+"送信したJSONでーた:\r\n" + JSON.stringify(JSON_DATA)+"\r\n\r\n";
    　var temp=JSON_DATA.data;  
    
    
  var sheet =doc.getSheetByName(sname);
          sheet.getRange(2, 2).setValue(temp);//2行目に貼り付ける

    doGet();
    log= log + "でーたふえいわけがおわりました";
    
    
    return ContentService
          .createTextOutput(log)
          .setMimeType(ContentService.MimeType.JSON)
          
    
  } catch(e){
    // 何か例外発生時のろぐ
    return ContentService
          .createTextOutput(log + JSON.stringify({"結果":"エラー", "エラー": e}))
          .setMimeType(ContentService.MimeType.JSON);
  } finally { //release lock
    lock.releaseLock();
  }
}

function status(d){
var st;
if(d>0){
st=",成功,"
}
else if(d==0){
st=",同じすこあが存在,"
}
else if(d==-1){
st=",検索対象なし,"
}
else if(d==-2){
st=",送信すこあが前回より小さい,"
}

return st;
}

function setup() {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    SCRIPT_PROP.setProperty("key", doc.getId());
}

/* ちゃんげろぐ
2015/07/29
*/



/**
 * Return a list of sheet names in the Spreadsheet with the given ID.
 * @param {String} a Spreadsheet ID.
 * @return {Array} A list of sheet names.
 */

//"100位: 805,183 (+304) 100いの最終２１時のとこにいれてボタンを押す
//2,500位: 27,021 (+88)
//5,000位: 17,907 (+85)
//10,000位: 12,805 (+67)
//25,000位: 7,024 (+35)
//50,000位: 3,030 (+16)
//2020/02/09 00:00 #ミリシタボーダー
//https://si.ster.li/events/121"

//var sid="1-6Uz0g3AFTHGv8a4bZYgIttn60bQw0Vl2UqSb1OKVx8";
var data=["100位","2500位","5000位","10000位"];
var mkdt=["いべ時点","最終21時","8日17時","8日0時","7日17時","6日17時","5日17時","4日17時","3日17時","2日17時","1日17時","0日17時"];

function doGet() {
  var ss = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key")); //SpreadsheetApp.openById(sid);
　var last_row =2;
　var last_col = 2;
  var sname=data[0];
  var sheets = ss.getSheetByName(sname);
  var s = ss.getSheetByName('しーと版いべたいま');
   
  var values= sheets.getRange(2,2,last_row ,last_col).getValues();
  var str=JSON.stringify(values);  
  
  str=str.replace(/,/g,'');
  var reg = /\d+位: \d+/gi;
  var m=str.match(reg);
  var regt = /\d\d\d\d.\d\d.\d\d \d\d:\d\d/g;
  
  var values= s.getRange(22,3,1 ,2).getValues();
  var ibe=JSON.parse(JSON.stringify(values));   
  var st =str.match(regt).toString();
  st =st.replace(/\//gi,'-');
  st =st.replace(/ /gi,'T');
  st +="+09:00"; 
  
  var moment = Moment.load();
  var t=moment(st);
  
  var ibelen=(moment(ibe[0][0])-moment(ibe[0][1]));
  var ibeh=ibelen.valueOf()/3600/1000;
  var ibed=ibeh/24;
  var ibeoff=-(ibed-7.25);
  var ibeday=[];
  for(var i=ibeoff;i<ibed;i++){
    ibeday[i]=[];
    if(ibed-i<=0.25){
  ibeday[i][0]=moment(ibe[0][1]).add("hours",-17+2+24*i);
    ibeday[i+1]=[];
    ibeday[i+2]=[];
  ibeday[i+1][0]=moment(ibe[0][1]).add("hours",+2+24*i);
  ibeday[i+2][0]=moment(ibe[0][1]).add("hours",4+2+24*i);
  ibeday[i][1]=(i+1)+moment(ibe[0][1]).add("hours",-17+2+24*i).format("日H時");
  ibeday[i+1][1]=(i+1)+moment(ibe[0][1]).add("hours",+2+24*i).format("日H時");
  ibeday[i+2][1]="最終"+moment(ibe[0][1]).add("hours",4+2+24*i).format("H時");
    }
    else{
  ibeday[i][0]=moment(ibe[0][1]).add("hours",2+24*i);
  ibeday[i][1]=(i+1)+moment(ibe[0][1]).add("hours",2+24*i).format("日H時");
    }}
  
  var tmp="";
  for(var i=0;i<ibeday.length;i++){
    if(ibeday[i][0].valueOf()==t.valueOf()){
      tmp=ibeday[i][1]
    }
  }
  
  
  for(var i=0;i<data.length;i++){
  var sheets = ss.getSheetByName(data[i]);
  for(var k=0;k<mkdt.length;k++){
    if(tmp==mkdt[k]){
      var mm=m[i].toString();
  var reg = /\d+/gi;
  var mm=mm.match(reg);
      sheets.getRange(k+1,2).setValue(mm[1]);}
  }
  }
  return
  //return ContentService.createTextOutput(JSON.stringify(m)).setMimeType(ContentService.MimeType.TEXT);
  //JSON.stringify(sheet.getName());
}

function wmap_getSheetsName(sheets){
  //var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheet_names = new Array();
  
  if (sheets.length >= 1) {  
    for(var i = 0;i < sheets.length; i++)
    {
      sheet_names.push(sheets[i].getName());
    }
  }
  return sheet_names;
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();           // Uiクラスを取得する
  var menu = ui.createMenu('ぼだJS振り分け');  // Uiクラスからメニューを作成する
  menu.addItem('ぼだ', 'doGet');   // メニューにアイテムを追加する
  menu.addToUi();                            // メニューをUiクラスに追加する
}
