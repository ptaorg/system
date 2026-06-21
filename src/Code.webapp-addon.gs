function doGet(e){
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('PTA管理アプリ');
}

function showAdminApp(){
  var html=HtmlService.createHtmlOutputFromFile('Index').setTitle('PTA管理アプリ').setWidth(1100).setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html,'PTA管理アプリ');
}

function admin_getInitialData(){
  var apps=admin_listApplications();
  var members=admin_getMembers();
  var confirmed=apps.filter(function(r){return r.status===STATUS_MEMBER;}).length;
  var withdrawn=apps.filter(function(r){return r.status===STATUS_WITHDRAWN;}).length;
  return {dashboard:{totalApplications:apps.length,pendingApplications:apps.length-confirmed-withdrawn,confirmedApplications:confirmed,memberCount:members.length},applications:apps,members:members};
}

function admin_listApplications(){
  var sh=getOrCreateRoster_();
  var last=sh.getLastRow();
  if(last<2)return [];
  var v=sh.getRange(2,1,last-1,ROSTER_COLS).getValues();
  return v.map(function(r,i){return {rowNumber:i+2,id:r[COL.APPLICATION_ID-1]||'',timestamp:admin_fmt_(r[COL.TIMESTAMP-1]),guardian:r[COL.GUARDIAN_NAME-1]||'',email:r[COL.EMAIL-1]||'',status:r[COL.STATUS-1]||''};}).filter(function(r){return r.id||r.guardian||r.email;});
}

function admin_getMembers(){
  var sh=getOrCreateMemberSheet_();
  var last=sh.getLastRow();
  if(last<2)return [];
  var v=sh.getRange(2,1,last-1,5).getValues();
  return v.map(function(r){return {id:r[0]||'',guardian:r[1]||'',child:r[2]||'',grade:r[3]||'',email:r[4]||''};}).filter(function(r){return r.id||r.guardian||r.email;});
}

function admin_confirmMember(rowNumber){
  var sh=getOrCreateRoster_();
  rowNumber=Number(rowNumber);
  sh.getRange(rowNumber,COL.STATUS).setValue(STATUS_MEMBER);
  if(!sh.getRange(rowNumber,COL.JOIN_DATE).getValue())sh.getRange(rowNumber,COL.JOIN_DATE).setValue(new Date());
  exportMembers();
  return admin_getInitialData();
}

function admin_withdrawApplication(rowNumber,reason){
  var sh=getOrCreateRoster_();
  rowNumber=Number(rowNumber);
  sh.getRange(rowNumber,COL.STATUS).setValue(STATUS_WITHDRAWN);
  sh.getRange(rowNumber,COL.NOTE).setValue(reason||'管理アプリから取下げ');
  exportMembers();
  return admin_getInitialData();
}

function admin_fmt_(v){
  if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Tokyo','yyyy/MM/dd HH:mm');
  return v==null?'':String(v);
}
