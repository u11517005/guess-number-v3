/**
 * 猜數字遊戲版本三：Google Apps Script 後端
 * 第一次執行 setup()，會自動建立 Games 與 Guesses 工作表。
 * 部署為網頁應用程式後，將 /exec 網址貼到 index.html 的 API_URL。
 */
const SPREADSHEET_ID = ""; // 留空：使用本 Apps Script 綁定的試算表；獨立專案請填入試算表 ID
const GAMES_SHEET = "Games";
const GUESSES_SHEET = "Guesses";

function setup() {
  const ss = getSpreadsheet_();
  let games = ss.getSheetByName(GAMES_SHEET);
  if (!games) games = ss.insertSheet(GAMES_SHEET);
  if (games.getLastRow() === 0) games.appendRow(["userId","gameId","answer","status","createdAt","updatedAt"]);
  let guesses = ss.getSheetByName(GUESSES_SHEET);
  if (!guesses) guesses = ss.insertSheet(GUESSES_SHEET);
  if (guesses.getLastRow() === 0) guesses.appendRow(["userId","gameId","guess","a","b","time"]);
  return "設定完成。請確認已建立 Games 與 Guesses 工作表。";
}
function doGet(e) {
  const p = e.parameter || {};
  const callback = /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(p.callback || "") ? p.callback : "callback";
  let result;
  try {
    if (!p.userId || !/^[A-Za-z0-9_-]{3,40}$/.test(p.userId)) throw new Error("識別碼格式錯誤");
    result = handle_(p);
    result.ok = true;
  } catch (err) {
    result = {ok:false,error:String(err.message || err)};
  }
  const output = callback + "(" + JSON.stringify(result) + ");";
  return ContentService.createTextOutput(output).setMimeType(ContentService.MimeType.JAVASCRIPT);
}
function handle_(p) {
  const userId = p.userId;
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet_();
    const gs = ss.getSheetByName(GAMES_SHEET);
    const qs = ss.getSheetByName(GUESSES_SHEET);
    if (!gs || !qs) throw new Error("試算表尚未設定，請先在 Apps Script 執行 setup()");
    if (p.action === "load") return loadUser_(gs, qs, userId);
    if (p.action === "newGame") {
      const now = new Date().toISOString();
      const gameId = Utilities.getUuid();
      const answer = makeAnswer_();
      gs.appendRow([userId,gameId,answer,"unfinished",now,now]);
      return {game: {gameId,answer,status:"unfinished",createdAt:now,updatedAt:now,guesses:[]}};
    }
    if (p.action === "guess") return addGuess_(gs,qs,userId,p.gameId,p.guess);
    if (p.action === "debug") {
      const row = findGameRow_(gs,userId,p.gameId);
      if (!row) throw new Error("找不到這局遊戲");
      return {answer:String(gs.getRange(row,3).getValue())};
    }
    throw new Error("未知的操作");
  } finally {
    lock.releaseLock();
  }
}
function getSpreadsheet_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("請將 Apps Script 綁定到 Google 試算表，或在程式中填入 SPREADSHEET_ID");
  return ss;
}
function makeAnswer_() {
  const digits = "0123456789".split("");
  for (let i=digits.length-1;i>0;i--) {
    const j=Math.floor(Math.random()*(i+1));[digits[i],digits[j]]=[digits[j],digits[i]];
  }
  return digits.slice(0,4).join("");
}
function findGameRow_(sheet,userId,gameId) {
  if (!gameId) return 0;
  const values=sheet.getDataRange().getValues();
  for(let i=1;i<values.length;i++) if(String(values[i][0])===userId && String(values[i][1])===gameId) return i+1;
  return 0;
}
function loadUser_(gs,qs,userId) {
  const gvals=gs.getDataRange().getValues(), qvals=qs.getDataRange().getValues();
  const all=[];
  for(let i=1;i<gvals.length;i++){
    const r=gvals[i];if(String(r[0])!==userId)continue;
    const guesses=[];
    for(let j=1;j<qvals.length;j++){
      const q=qvals[j];if(String(q[0])===userId && String(q[1])===String(r[1]))
        guesses.push({guess:String(q[2]),a:Number(q[3]),b:Number(q[4]),time:String(q[5])});
    }
    all.push({gameId:String(r[1]),answer:String(r[2]),status:String(r[3]),createdAt:String(r[4]),updatedAt:String(r[5]),guesses});
  }
  all.sort((a,b)=>new Date(a.createdAt)-new Date(b.createdAt));
  let current=null;
  for(let i=all.length-1;i>=0;i--) if(all[i].status==="unfinished"){current=all[i];break;}
  if(!current && all.length) current=all[all.length-1];
  return {games:all,currentGame:current};
}
function addGuess_(gs,qs,userId,gameId,guess) {
  if (!/^\d{4}$/.test(guess || "") || new Set(guess).size!==4) throw new Error("請輸入四位且不重複的數字");
  const row=findGameRow_(gs,userId,gameId);
  if(!row)throw new Error("找不到這局遊戲，請重新連線");
  const r=gs.getRange(row,1,1,6).getValues()[0];
  if(String(r[3])==="completed")throw new Error("這局已完成，請開始新遊戲");
  const answer=String(r[2]);
  const a=guess.split("").filter((d,i)=>d===answer[i]).length;
  const b=guess.split("").filter(d=>answer.includes(d)).length-a;
  const now=new Date().toISOString();
  qs.appendRow([userId,gameId,guess,a,b,now]);
  if(a===4)gs.getRange(row,4).setValue("completed");
  gs.getRange(row,6).setValue(now);
  const loaded=loadUser_(gs,qs,userId);
  const game=loaded.games.find(x=>x.gameId===gameId);
  return {game,result:{a,b}};
}
