/**
 * Book Bingo Fair — Web API + email OTP login (@company domain).
 *
 * Script Properties (optional): ALLOWED_EMAIL_DOMAIN = e.g. amanotes.com (no @)
 * Sheet "Users" is auto-created with headers on first web request.
 *
 * Deploy: Apps Script editor → Deploy → Manage deployments → Edit (pencil) →
 *   Version: New version → Deploy. Ensure execute-as has permission to send mail.
 *
 * OAuth: First web or mail send may prompt for Gmail + Spreadsheet access.
 */
function getSpreadsheet(){return SpreadsheetApp.getActiveSpreadsheet();}

function getAllowedDomain(){
  var d=PropertiesService.getScriptProperties().getProperty('ALLOWED_EMAIL_DOMAIN');
  d=(d||'amanotes.com').trim().toLowerCase().replace(/^@/,'');
  return d;
}

function isAllowedEmail(email){
  var dom=getAllowedDomain();
  var e=(email||'').trim().toLowerCase();
  var suf='@'+dom;
  return e.length>suf.length&&e.substring(e.length-suf.length)===suf;
}

function sha256Hex(s){
  var bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s),Utilities.Charset.UTF_8);
  var hex='';
  for(var i=0;i<bytes.length;i++){
    var b=bytes[i]&0xff;
    hex+=(b<16?'0':'')+b.toString(16);
  }
  return hex;
}

function ensureUsersSheet(){
  var ss=getSpreadsheet();
  var sh=ss.getSheetByName('Users');
  if(sh)return sh;
  sh=ss.insertSheet('Users');
  sh.appendRow(['email','token_hash','expires_at','display_name','created_at','updated_at']);
  return sh;
}

function upsertUserSession(normEmail,tokenHash,expiresIso,displayName){
  var sh=getSpreadsheet().getSheetByName('Users');
  var data=sh.getDataRange().getValues();
  var nowIso=new Date().toISOString();
  var hdr=data[0];
  var ei=hdr.indexOf('email');
  if(data.length<2){
    sh.appendRow([normEmail,tokenHash,expiresIso,displayName,nowIso,nowIso]);
    return;
  }
  for(var i=1;i<data.length;i++){
    if(String(data[i][ei]).toLowerCase()===normEmail){
      sh.getRange(i+1,hdr.indexOf('token_hash')+1).setValue(tokenHash);
      sh.getRange(i+1,hdr.indexOf('expires_at')+1).setValue(expiresIso);
      sh.getRange(i+1,hdr.indexOf('display_name')+1).setValue(displayName);
      sh.getRange(i+1,hdr.indexOf('updated_at')+1).setValue(nowIso);
      return;
    }
  }
  sh.appendRow([normEmail,tokenHash,expiresIso,displayName,nowIso,nowIso]);
}

function validateSession(sessionToken){
  if(!sessionToken)return {ok:false};
  var h=sha256Hex(sessionToken);
  var sh=getSpreadsheet().getSheetByName('Users');
  if(!sh)return {ok:false};
  var data=sh.getDataRange().getValues();
  if(data.length<2)return {ok:false};
  var hdr=data[0];
  var hi=hdr.indexOf('token_hash'),ei=hdr.indexOf('email'),di=hdr.indexOf('display_name'),xi=hdr.indexOf('expires_at');
  for(var i=1;i<data.length;i++){
    if(String(data[i][hi])!==h)continue;
    var expStr=String(data[i][xi]||'');
    if(!expStr)return {ok:false};
    try{
      if(new Date(expStr).getTime()<=Date.now())return {ok:false};
    }catch(e){return {ok:false};}
    return {ok:true,email:String(data[i][ei]||''),displayName:String(data[i][di]||'')};
  }
  return {ok:false};
}

function handleRequestLoginCode(p){
  var raw=(p.email||'').trim().toLowerCase();
  if(!raw)return out({error:'Missing email'});
  if(!isAllowedEmail(raw))return out({error:'Only company email addresses are allowed.'});
  var cache=CacheService.getDocumentCache();
  var cdKey='otp_cd_'+raw;
  if(cache.get(cdKey))return out({error:'Please wait a minute before requesting another code.'});
  var hourKey='otp_hr_'+raw;
  var cnt=parseInt(cache.get(hourKey)||'0',10);
  if(cnt>=10)return out({error:'Too many code requests this hour. Try again later.'});
  cache.put(hourKey,String(cnt+1),3600);
  var code=(''+Math.floor(100000+Math.random()*900000)).slice(0,6);
  cache.put('otp_'+raw,code,600);
  cache.put(cdKey,'1',60);
  var subject='Your Amanotes Book Bingo Fair login code';
  var body='Your verification code is: '+code+'\n\nIt expires in 10 minutes. If you did not request this, ignore this email.';
  try{
    GmailApp.sendEmail(raw,subject,body);
  }catch(err){
    try{
      MailApp.sendEmail(raw,subject,body);
    }catch(err2){
      return out({error:'Could not send email: '+String(err2.message||err2)});
    }
  }
  return out({success:true});
}

function handleVerifyLoginCode(p){
  var raw=(p.email||'').trim().toLowerCase();
  var code=(p.code||'').trim().replace(/\s/g,'');
  if(!raw||!code)return out({error:'Missing email or code'});
  if(!isAllowedEmail(raw))return out({error:'Invalid email'});
  var cache=CacheService.getDocumentCache();
  var stored=cache.get('otp_'+raw);
  if(!stored||String(stored)!==code)return out({error:'Invalid or expired code'});
  cache.remove('otp_'+raw);
  var token=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
  var tokenHash=sha256Hex(token);
  var displayName=(p.display_name||'').trim();
  if(!displayName){
    var at=raw.indexOf('@');
    displayName=at>0?raw.substring(0,at):raw;
  }
  var expires=new Date(Date.now()+30*24*60*60*1000).toISOString();
  upsertUserSession(raw,tokenHash,expires,displayName);
  return out({success:true,sessionToken:token,email:raw,displayName:displayName});
}

function handleInvalidateSession(p){
  var tok=p.sessionToken||'';
  if(!tok)return out({error:'Missing sessionToken'});
  var h=sha256Hex(tok);
  var sh=getSpreadsheet().getSheetByName('Users');
  if(!sh)return out({success:true});
  var data=sh.getDataRange().getValues();
  if(data.length<2)return out({success:true});
  var hdr=data[0];
  var hi=hdr.indexOf('token_hash'),xi=hdr.indexOf('expires_at');
  for(var i=1;i<data.length;i++){
    if(String(data[i][hi])===h){
      sh.getRange(i+1,hi+1).setValue('');
      sh.getRange(i+1,xi+1).setValue('');
      return out({success:true});
    }
  }
  return out({success:true});
}

function handleSessionPing(p){
  var s=validateSession(p.sessionToken||'');
  if(!s.ok)return out({ok:false});
  return out({ok:true,email:s.email,displayName:s.displayName});
}

function requireSession(p){
  var s=validateSession(p.sessionToken||'');
  if(!s.ok)return null;
  return s;
}

function doPost(e){
  var p={};
  if(e.parameter){
    Object.keys(e.parameter).forEach(function(k){p[k]=e.parameter[k];});
  }
  if(e.postData&&e.postData.contents){
    try{
      var body=JSON.parse(e.postData.contents);
      Object.keys(body).forEach(function(k){p[k]=body[k];});
    }catch(err){}
  }
  return doGet({parameter:p});
}

function doGet(e){
  var p=e.parameter;
  try{
    ensureUsersSheet();
    var action=String(p.action||'');

    if(action==='requestLoginCode')return handleRequestLoginCode(p);
    if(action==='verifyLoginCode')return handleVerifyLoginCode(p);
    if(action==='invalidateSession')return handleInvalidateSession(p);
    if(action==='sessionPing')return handleSessionPing(p);

    if(action==='getBooks'){
      var books=getData('Books');
      var rets=getData('Returns');
      books.forEach(function(b){
        b.returnReviews=rets.filter(function(r){return String(r.bookId)===String(b.id);});
      });
      return out(books);
    }
    if(action==='getBingoBonus'){
      return out(getData('BingoBonus'));
    }
    if(action==='getLibrary'){
      return out(getData('Library'));
    }
    if(action==='getBingoLinks'){
      var sheet=getSpreadsheet().getSheetByName('BingoLinks');
      if(!sheet)return out([]);
      var data=sheet.getDataRange().getValues();
      if(data.length<2)return out([]);
      var hdr=data[0];
      var result=data.slice(1).filter(function(row){
        return String(row[hdr.indexOf('name')]).toLowerCase()===String(p.name||'').toLowerCase();
      }).map(function(row){
        var obj={};
        hdr.forEach(function(h,i){obj[h]=row[i];});
        return obj;
      });
      return out(result);
    }

    var sess=requireSession(p);
    if(!sess)return out({error:'Unauthorized'});
    var who=sess.displayName;

    if(action==='addBook'){
      var idA=Date.now();
      getSpreadsheet().getSheetByName('Books').appendRow([idA,p.title,p.author,p.category,who,p.microReview,p.coverUrl||'',0,0,p.date,p.fullReview||'']);
      return out({success:true,id:idA});
    }
    if(action==='like'){bump('Books',p.bookId,'likes',Number(p.delta));return out({success:true});}
    if(action==='borrow'){bump('Books',p.bookId,'wantToBorrow',Number(p.delta));return out({success:true});}
    if(action==='addReturn'){
      getSpreadsheet().getSheetByName('Returns').appendRow([p.bookId,who,p.text,p.date]);
      return out({success:true});
    }
    if(action==='editBook'){
      var sheetE=getSpreadsheet().getSheetByName('Books');
      var dataE=sheetE.getDataRange().getValues();
      var hdrE=dataE[0];
      var iiE=hdrE.indexOf('id'),recI=hdrE.indexOf('recommender');
      for(var j=1;j<dataE.length;j++){
        if(String(dataE[j][iiE])===String(p.bookId)){
          if(String(dataE[j][recI]||'').toLowerCase()!==String(who).toLowerCase()){
            return out({error:'You can only edit your own posts.'});
          }
          sheetE.getRange(j+1,hdrE.indexOf('title')+1).setValue(p.title);
          sheetE.getRange(j+1,hdrE.indexOf('author')+1).setValue(p.author);
          sheetE.getRange(j+1,hdrE.indexOf('category')+1).setValue(p.category);
          sheetE.getRange(j+1,hdrE.indexOf('recommender')+1).setValue(who);
          sheetE.getRange(j+1,hdrE.indexOf('microReview')+1).setValue(p.microReview);
          sheetE.getRange(j+1,hdrE.indexOf('fullReview')+1).setValue(p.fullReview||'');
          sheetE.getRange(j+1,hdrE.indexOf('coverUrl')+1).setValue(p.coverUrl||'');
          return out({success:true});
        }
      }
      return out({error:'Book not found'});
    }
    if(action==='claimBingo'){
      getSpreadsheet().getSheetByName('BingoBonus').appendRow([who,p.type,p.lines,p.pts,p.date]);
      return out({success:true});
    }
    if(action==='donateBook'){
      var idD=Date.now();
      getSpreadsheet().getSheetByName('Library').appendRow([
        idD,p.title,p.author,p.description||'',who,
        p.coverUrl||'','available','',p.date,p.categories||''
      ]);
      return out({success:true,id:idD});
    }
    if(action==='borrowLibraryBook'){
      var sheetB=getSpreadsheet().getSheetByName('Library');
      var dataB=sheetB.getDataRange().getValues();
      var hdrB=dataB[0],iiB=hdrB.indexOf('id');
      for(var k=1;k<dataB.length;k++){
        if(String(dataB[k][iiB])===String(p.bookId)){
          sheetB.getRange(k+1,hdrB.indexOf('status')+1).setValue('borrowed');
          sheetB.getRange(k+1,hdrB.indexOf('borrowedBy')+1).setValue(who);
          return out({success:true});
        }
      }
      return out({error:'Book not found'});
    }
    if(action==='returnLibraryBook'){
      var sheetR=getSpreadsheet().getSheetByName('Library');
      var dataR=sheetR.getDataRange().getValues();
      var hdrR=dataR[0],iiR=hdrR.indexOf('id'),bbI=hdrR.indexOf('borrowedBy');
      for(var r=1;r<dataR.length;r++){
        if(String(dataR[r][iiR])===String(p.bookId)){
          if(String(dataR[r][bbI]||'').toLowerCase()!==String(who).toLowerCase()){
            return out({error:'Only the borrower can return this book.'});
          }
          sheetR.getRange(r+1,hdrR.indexOf('status')+1).setValue('available');
          sheetR.getRange(r+1,bbI+1).setValue('');
          return out({success:true});
        }
      }
      return out({error:'Book not found'});
    }

    if(action==='addBingoLink'){
      var sheetL=getSpreadsheet().getSheetByName('BingoLinks');
      if(!sheetL)return out({error:'BingoLinks sheet not found'});
      var dataL=sheetL.getDataRange().getValues();
      var hdrL=dataL[0];
      var ni=hdrL.indexOf('name'),bi=hdrL.indexOf('bookId'),si=hdrL.indexOf('squareIndex');
      var bingoName=who;
      for(var m=1;m<dataL.length;m++){
        if(String(dataL[m][ni]).toLowerCase()===bingoName.toLowerCase()&&String(dataL[m][bi])===String(p.bookId)){
          return out({error:'This book is already linked to another square'});
        }
        if(String(dataL[m][ni]).toLowerCase()===bingoName.toLowerCase()&&String(dataL[m][si])===String(p.squareIndex)){
          return out({error:'This square is already linked'});
        }
      }
      sheetL.appendRow([bingoName,p.squareIndex,p.bookId,p.bookTitle,p.date]);
      return out({success:true});
    }

    return out({error:'Unknown action'});
  }catch(err){return out({error:err.message});}
}

function getData(sheetName){
  var sheet=getSpreadsheet().getSheetByName(sheetName);
  if(!sheet)return[];
  var rows=sheet.getDataRange().getValues();
  if(rows.length<2)return[];
  var headers=rows[0];
  return rows.slice(1).map(function(row){
    var obj={};
    headers.forEach(function(h,i){obj[h]=row[i];});
    return obj;
  });
}

function bump(sheetName,id,col,delta){
  var sheet=getSpreadsheet().getSheetByName(sheetName);
  var data=sheet.getDataRange().getValues();
  var ci=data[0].indexOf(col);
  var ii=data[0].indexOf('id');
  for(var i=1;i<data.length;i++){
    if(String(data[i][ii])===String(id)){
      var cell=sheet.getRange(i+1,ci+1);
      cell.setValue(Math.max(0,(cell.getValue()||0)+delta));
      return;
    }
  }
}
function out(data){
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
function fetchMissingCovers() {
  var sheet = getSpreadsheet().getSheetByName('Library');
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0];
  var ti    = hdr.indexOf('title');
  var ai    = hdr.indexOf('author');
  var ci    = hdr.indexOf('coverUrl');
  var count = 0;

  for (var i = 1; i < data.length; i++) {
    // Chỉ fetch những dòng chưa có cover
    if (data[i][ci] && data[i][ci] !== '') continue;

    var title  = data[i][ti];
    var author = data[i][ai];
    if (!title) continue;

    try {
      var query = 'intitle:' + encodeURIComponent(title);
      if (author) query += '+inauthor:' + encodeURIComponent(author);
      var url = 'https://www.googleapis.com/books/v1/volumes?q=' + query + '&maxResults=1&fields=items/volumeInfo/imageLinks';
      var res  = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      var json = JSON.parse(res.getContentText());

      if (json.items && json.items[0] &&
          json.items[0].volumeInfo &&
          json.items[0].volumeInfo.imageLinks) {
        var cover = json.items[0].volumeInfo.imageLinks.thumbnail || '';
        // Nâng chất lượng ảnh + đổi sang HTTPS
        cover = cover.replace('zoom=1','zoom=2').replace('http://','https://');
        sheet.getRange(i + 1, ci + 1).setValue(cover);
        count++;
      }
    } catch(e) {
      // Bỏ qua nếu lỗi, tiếp tục cuốn tiếp theo
    }

    // Tránh rate limit — nghỉ 300ms mỗi cuốn
    Utilities.sleep(300);
  }

  Logger.log('Done! Fetched covers for ' + count + ' books.');
}
function populateLibraryAndFetchCovers() {
  var sheet = getSpreadsheet().getSheetByName('Library');
  
  // Step 1: Clear existing data (keep header)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 10).clearContent();
  
  // Step 2: Insert all 233 books
  var books = [
    [1001,'Sapiens','Yuval Noah Harari','Lịch sử loài người','Ly Dang','','available','','2026-05-15','A book that changed how you think|A book related to your work'],
    [1002,'Triệu phú khu ổ chuột','Vikas Swarup','','Ly Dang','','available','','2026-05-15',''],
    [1003,'Người đua diều','Khaled Hosseini','','Ly Dang','','available','','2026-05-15',''],
    [1004,'Ngàn mặt trời rực rỡ','Khaled Hosseini','','Ly Dang','','available','','2026-05-15',''],
    [1005,'Hai số phận','Jeffrey Archer','','Ly Dang','','available','','2026-05-15',''],
    [1006,'Ăn gì không chết','Michael Greger','','Ly Dang','','available','','2026-05-15',''],
    [1007,'No More Plastic','Martin Dorey','','Ly Dang','','available','','2026-05-15',''],
    [1008,'Yêu những điều không hoàn hảo','Hae Min','','Ly Dang','','available','','2026-05-15',''],
    [1009,'Bước chậm lại giữa thế gian vội vã','Hae Min','','Ly Dang','','available','','2026-05-15',''],
    [1010,'Tìm mình trong thế giới hậu tuổi thơ','Đặng Hoàng Giang','','Ly Dang','','available','','2026-05-15',''],
    [1011,'Siddhartha','Hermann Hesse','','Ly Dang','','available','','2026-05-15',''],
    [1012,'Việt Nam sử lược','Trần Trọng Kim','','Ly Dang','','available','','2026-05-15',''],
    [1013,'Lịch sử (Historiai)','Herodotos','','Ly Dang','','available','','2026-05-15',''],
    [1014,'Tam quốc diễn nghĩa (3 tập)','La Quán Trung','','Ly Dang','','available','','2026-05-15',''],
    [1015,'Lịch sử Do Thái','Paul Johnson','','Ly Dang','','available','','2026-05-15',''],
    [1016,'Chim cổ đỏ','Jo Nesbø','','Ly Dang','','available','','2026-05-15',''],
    [1017,'Sự im lặng của bầy cừu','Thomas Harris','','Ly Dang','','available','','2026-05-15',''],
    [1018,'Hannibal','Thomas Harris','','Ly Dang','','available','','2026-05-15',''],
    [1019,'Battle Royale','Koushun Takami','','Ly Dang','','available','','2026-05-15',''],
    [1020,'Điều Kỳ Diệu Của Tiệm Tạp Hóa Namiya','Higashino Keigo','','Ly Dang','','available','','2026-05-15',''],
    [1021,'Bạch dạ hành','Higashino Keigo','','Ly Dang','','available','','2026-05-15',''],
    [1022,'Phía sau nghi can X','Higashino Keigo','','Ly Dang','','available','','2026-05-15',''],
    [1023,'Biên niên ký chim vặn dây cót','Haruki Murakami','','Ly Dang','','available','','2026-05-15',''],
    [1024,'Coraline','Neil Gaiman','','Ly Dang','','available','','2026-05-15',''],
    [1025,'Charlotte','David Foenkinos','','Ly Dang','','available','','2026-05-15',''],
    [1026,'Chết chịu','Louis-Ferdinand Céline','','Ly Dang','','available','','2026-05-15',''],
    [1027,'Chuyện người tùy nữ','Margaret Atwood','','Ly Dang','','available','','2026-05-15',''],
    [1028,'Sapiens: Lược Sử Loài Người','Yuval Noah Harari','','Ly Dang','','available','','2026-05-15',''],
    [1029,'Homo Deus: Lược Sử Tương Lai','Yuval Noah Harari','','Ly Dang','','available','','2026-05-15',''],
    [1030,'Học cách học','Barbara Oakley','','Ly Dang','','available','','2026-05-15',''],
    [1031,'21 Bài Học Cho Thế Kỷ 21','Yuval Noah Harari','','Ly Dang','','available','','2026-05-15',''],
    [1032,'Tôi tự học','Nguyễn Duy Cần','','Ly Dang','','available','','2026-05-15',''],
    [1033,'Phương Pháp Điều Trị Trầm Cảm','Stephen S. Ilardi','','Ly Dang','','available','','2026-05-15',''],
    [1034,'Chủ nghĩa khắc kỷ','William B. Irvine','','Ly Dang','','available','','2026-05-15',''],
    [1035,'Factfulness','Hans Rosling','','Ly Dang','','available','','2026-05-15',''],
    [1036,'Nơi khu rừng chạm tới những vì sao','Glendy Vanderah','','Ly Dang','','available','','2026-05-15',''],
    [1037,'Danh sách của Schindler','Thomas Keneally','','Ly Dang','','available','','2026-05-15',''],
    [1038,'Shoe Dog','Phil Knight','','Ly Dang','','available','','2026-05-15',''],
    [1039,'Phi lý trí','Dan Ariely','','Ly Dang','','available','','2026-05-15',''],
    [1040,'7 thói quen để thành đạt','Stephen Covey','','Ly Dang','','available','','2026-05-15',''],
    [1041,'Thằng Cười','Victor Hugo','','Ly Dang','','available','','2026-05-15',''],
    [1042,'Bố Già','Mario Puzo','','Ly Dang','','available','','2026-05-15',''],
    [1043,'Tư duy nhanh và chậm','Daniel Kahneman','','Ly Dang','','available','','2026-05-15',''],
    [1044,'Ác quỷ Nam Kinh','Mo Hayder','','Ly Dang','','available','','2026-05-15',''],
    [1045,'Đi tìm lẽ sống','Viktor Frankl','','Ly Dang','','available','','2026-05-15',''],
    [1046,'Thú tội','Minato Kanae','','Ly Dang','','available','','2026-05-15',''],
    [1047,'Another (2 tập)','Yukito Ayatsuji','','Ly Dang','','available','','2026-05-15',''],
    [1048,'Totto-chan bên cửa sổ','Tetsuko Kuroyanagi','','Ly Dang','','available','','2026-05-15',''],
    [1049,'Mùi hương','Patrick Süskind','','Ly Dang','','available','','2026-05-15',''],
    [1050,'Sử Việt 12 khúc tráng ca','Dũng Phan','','Ly Dang','','available','','2026-05-15',''],
    [1051,'Thảm kịch vĩ nhân','Hoàng Minh Tường','','Ly Dang','','available','','2026-05-15',''],
    [1052,'Nghiệt duyên','Thommayanti','','Ly Dang','','available','','2026-05-15',''],
    [1053,'Đội gạo lên chùa','Nguyễn Xuân Khánh','','Ly Dang','','available','','2026-05-15',''],
    [1054,'Thiên thần và ác quỷ','Dan Brown','','Ly Dang','','available','','2026-05-15',''],
    [1055,'Nguồn cội','Dan Brown','','Ly Dang','','available','','2026-05-15',''],
    [1056,'Biểu tượng thất truyền','Dan Brown','','Ly Dang','','available','','2026-05-15',''],
    [1057,'Trường ca Achilles','Madeline Miller','','Ly Dang','','available','','2026-05-15',''],
    [1058,'Thiếu nữ đánh cờ vây','Sơn Táp','','Ly Dang','','available','','2026-05-15',''],
    [1059,'Xa ngoài kia nơi loài tôm hát','Delia Owens','','Ly Dang','','available','','2026-05-15',''],
    [1060,'The midnight library','Matt Haig','','Ly Dang','','available','','2026-05-15',''],
    [1061,'Sa môn Không Hải thết yến bầy quỷ Đại Đường - tập 1','Yumemakura Baku','','Ly Dang','','available','','2026-05-15',''],
    [1062,'Chiến binh cầu vồng','Andrea Hirata','','Ly Dang','','available','','2026-05-15',''],
    [1063,'Giải mã siêu trí nhớ','Mai Tường Vân','','Ly Dang','','available','','2026-05-15',''],
    [1064,'Hoả ngục','Dan Brown','','Ly Dang','','available','','2026-05-15',''],
    [1065,'Hoa vẫn nở mỗi ngày','Valerie Perrin','','Ly Dang','','available','','2026-05-15',''],
    [1066,'Nuôi dạy con bằng trái tim của một vị phật','Dr. C. L. Claridge','','Ly Dang','','available','','2026-05-15',''],
    [1067,'Bão táp triều Trần','Hoàng Quốc Hải','','Ly Dang','','available','','2026-05-15',''],
    [1068,'Một chiến dịch ở Bắc Kỳ','Bác sĩ Hocquard','','Ly Dang','','available','','2026-05-15',''],
    [1069,'Thế giới phẳng','Thomas L. Friedman','','Ly Dang','','available','','2026-05-15',''],
    [1070,'Cội nguồn','David Christian','','Ly Dang','','available','','2026-05-15',''],
    [1071,'Vòm rừng','Richard Powers','','Ly Dang','','available','','2026-05-15',''],
    [1072,'Gen: Lịch sử và tương lai của nhân loại','Siddhartha Mukherjee','','Ly Dang','','available','','2026-05-15',''],
    [1073,'Nguồn gốc dịch bệnh','David Quammen','','Ly Dang','','available','','2026-05-15',''],
    [1074,'Utopia - Địa đàng trần gian','Thomas More','','Ly Dang','','available','','2026-05-15',''],
    [1075,'Sapiens: Lược Sử Loài Người Bằng Tranh - Tập 1','Yuval Noah Harari','','Ly Dang','','available','','2026-05-15',''],
    [1076,'Cha mẹ độc hại','Susan Forward','','Ly Dang','','available','','2026-05-15',''],
    [1077,'Luật trí não','John Medina','','Ly Dang','','available','','2026-05-15',''],
    [1078,'Giải Mã Siêu Trí Tuệ','Vishen Lakhiani','','Ly Dang','','available','','2026-05-15',''],
    [1079,'Bí Mật Của Một Trí Nhớ Siêu Phàm','Eran Katz','','Ly Dang','','available','','2026-05-15',''],
    [1080,'Mật mã tài năng','Daniel Coyle','','Ly Dang','','available','','2026-05-15',''],
    [1081,'Bản Chất Của Dối Trá','Dan Ariely','','Ly Dang','','available','','2026-05-15',''],
    [1082,'Luật trí não dành cho trẻ','John Medina','','Ly Dang','','available','','2026-05-15',''],
    [1083,'Thiên tài bên trái kẻ điên bên phải','Cao Minh','','Ly Dang','','available','','2026-05-15',''],
    [1084,'Đừng Bao Giờ Chia Đôi Lợi Ích Trong Mọi Cuộc Đàm Phán','Chris Voss, Tahl Raz','','Ly Dang','','available','','2026-05-15',''],
    [1085,'Câu chuyện nghệ thuật','E. H. Gombrich','','Ly Dang','','available','','2026-05-15',''],
    [1086,'Đừng bao giờ đi ăn một mình','Keith Ferrazzi, Tahl Raz','','Ly Dang','','available','','2026-05-15',''],
    [1087,'Thế giới của Sophie','Jostein Gaarder','','Ly Dang','','available','','2026-05-15',''],
    [1088,'Range - Hiểu Sâu Biết Rộng Kiểu Gì Cũng Thắng','David Epstein','','Ly Dang','','available','','2026-05-15',''],
    [1089,'Michelangelo- Sáu Kiệt Tác Cuộc Đời','Miles J. Unger','','Ly Dang','','available','','2026-05-15',''],
    [1090,'Leonardo Da Vinci','Walter Isaacson','','Ly Dang','','available','','2026-05-15',''],
    [1091,'Điểm đến của cuộc đời','Đặng Hoàng Giang','','Ly Dang','','available','','2026-05-15',''],
    [1092,'Bức xúc không làm ta vô can','Đặng Hoàng Giang','','Ly Dang','','available','','2026-05-15',''],
    [1093,'Những Người Khốn Khổ (Boxet 2 Tập)','Victor Hugo','','Ly Dang','','available','','2026-05-15',''],
    [1094,'Thiền Sư Và Em Bé 5 Tuổi','Thích Nhất Hạnh','','Ly Dang','','available','','2026-05-15',''],
    [1095,'Diana','Isabelle Rivère','','Ly Dang','','available','','2026-05-15',''],
    [1096,'Dịch Bệnh- Kẻ Thù Nguy Hiểm Nhất','Michael T.Osterholm','','Ly Dang','','available','','2026-05-15',''],
    [1097,'Học Viện','Stephen King','','Ly Dang','','available','','2026-05-15',''],
    [1098,'Jeff Bezos Và Kỷ Nguyên Amazon','Brad Stone','','Ly Dang','','available','','2026-05-15',''],
    [1099,'Pricing Done Right - Định Giá Dựa Trên Giá Trị','nhiều tác giả','','Ly Dang','','available','','2026-05-15',''],
    [1100,'Bill Gates: Tham Vọng Lớn Lao Và Quá Trình Hình Thành Đế Chế Microsoft','James Wallace, Jim Erickson','','Ly Dang','','available','','2026-05-15',''],
    [1101,'Hồi ký Alex Ferguson','Alex Ferguson','','Ly Dang','','available','','2026-05-15',''],
    [1102,'Mã Vân Giày Vải','Lý Tường, Vương Lợi Phân','','Ly Dang','','available','','2026-05-15',''],
    [1103,'Những Người Khổng Lồ Trong Giới Kinh Doanh','Richard S.Tedlow','','Ly Dang','','available','','2026-05-15',''],
    [1104,'Quy tắc làm việc của Google','Laszlo Bock','','Ly Dang','','available','','2026-05-15',''],
    [1105,'Made In Korea','Richard M. Steers','','Ly Dang','','available','','2026-05-15',''],
    [1106,'Elon Musk','Ashlee Vance','','Ly Dang','','available','','2026-05-15',''],
    [1107,'Được học','Tara Westover','','Ly Dang','','available','','2026-05-15',''],
    [1108,'Titan – Gia Tộc Rockefeller','Ron Chernow','','Ly Dang','','available','','2026-05-15',''],
    [1109,'Gia tộc Morgan','Ron Chernow','','Ly Dang','','available','','2026-05-15',''],
    [1110,'Becoming','Michelle Obama','','Ly Dang','','available','','2026-05-15',''],
    [1111,'Bay trên tổ chim cúc cu','Ken Kesey','','Ly Dang','','available','','2026-05-15',''],
    [1112,'An lạc từng bước chân','Thích Nhất Hạnh','','Ly Dang','','available','','2026-05-15',''],
    [1113,'Cây cam ngọt của tôi','JOSÉ MAURO DE VASCONCELOS','','Ly Dang','','available','','2026-05-15',''],
    [1114,'Thợ Xăm Ở Auschwitz','Heather Morris','','Ly Dang','','available','','2026-05-15',''],
    [1115,'Thiên Táng','Hân Nhiên','','Ly Dang','','available','','2026-05-15',''],
    [1116,'Chân Dung Của Dorian Gray','Oscar Wilde','','Ly Dang','','available','','2026-05-15',''],
    [1117,'Hảo Nữ Trung Hoa','Hân Nhiên','','Ly Dang','','available','','2026-05-15',''],
    [1118,'Bài Học Diệu Kỳ Từ Chiếc Xe Rác','David J. Pollay','','Ly Dang','','available','','2026-05-15',''],
    [1119,'Hiểu về trái tim','Minh Niệm','','Ly Dang','','available','','2026-05-15',''],
    [1120,'Vị Tu Sĩ Bán Chiếc Ferrari','Robin Sharma','','Ly Dang','','available','','2026-05-15',''],
    [1121,'Sinh Ra Để Chạy','Christopher McDougall','','Ly Dang','','available','','2026-05-15',''],
    [1122,'Đồi Gió Hú','Emily Bronte','','Ly Dang','','available','','2026-05-15',''],
    [1123,'Lịch sử ung thư - Hoàng đế của bách bệnh','Siddhartha Mukherjee','','Ly Dang','','available','','2026-05-15',''],
    [1124,'Của chuột và người','John Steinbeck','','Ly Dang','','available','','2026-05-15',''],
    [1125,'Đại dương đen','Đặng Hoàng Giang','','Ly Dang','','available','','2026-05-15',''],
    [1126,'Thảm họa khí hậu','Bill Gates','','Ly Dang','','available','','2026-05-15',''],
    [1127,'Tôi biết tại sao chim trong lồng vẫn hót','Maya Angelou','','Ly Dang','','available','','2026-05-15',''],
    [1128,'Kẻ trộm sách','Markus Zusak','','Ly Dang','','available','','2026-05-15',''],
    [1129,'Hội hè miên man','Earnest Hemingway','','Ly Dang','','available','','2026-05-15',''],
    [1130,'Kafka bên bờ biển','Haruki Murakami','','Ly Dang','','available','','2026-05-15',''],
    [1131,'Ngài cóc đi gặp bác sĩ tâm lý','Robert de Board','','Ly Dang','','available','','2026-05-15',''],
    [1132,'Bệnh nhân câm lặng','Alex Michaelides','','Ly Dang','','available','','2026-05-15',''],
    [1133,'Chú bé mang pijama sọc','John Boyne','','Ly Dang','','available','','2026-05-15',''],
    [1134,'Netflix: Phá Bỏ Nguyên Tắc Để Bứt Phá','REED HASTINGS,  Erin Meyer','','Ly Dang','','available','','2026-05-15',''],
    [1135,'Người đàn ông mang tên Ove','Fredrick Backman','','Ly Dang','','available','','2026-05-15',''],
    [1136,'Những linh hồn chết','Nikolai Gogol','','Ly Dang','','available','','2026-05-15',''],
    [1137,'Ô nhục','John Maxwell Coetzee','','Ly Dang','','available','','2026-05-15',''],
    [1138,'Moby Dick Cá voi trắng','Herman Melville','','Ly Dang','','available','','2026-05-15',''],
    [1139,'Jane Eyre','Charlotte Bronte','','Ly Dang','','available','','2026-05-15',''],
    [1140,'Hoa súng đen','Michel Bussi','','Ly Dang','','available','','2026-05-15',''],
    [1141,'Lẽ thường','Thomas Paine','','Ly Dang','','available','','2026-05-15',''],
    [1142,'Lược sử tôn giáo','Richard Holloway','','Ly Dang','','available','','2026-05-15',''],
    [1143,'Bà Bovary','Gustave Flaubert','','Ly Dang','','available','','2026-05-15',''],
    [1144,'Tâm lý học về tiền','Morgan Housel','','Ly Dang','','available','','2026-05-15',''],
    [1145,'Suy tưởng','Marcus Aurelius','','Ly Dang','','available','','2026-05-15',''],
    [1146,'Van Gogh: The Life','Gregory White Smith, Steven Naifeh','','Ly Dang','','available','','2026-05-15',''],
    [1147,'Làn sóng thứ ba','Alvin Toffler','','Ly Dang','','available','','2026-05-15',''],
    [1148,'Sử ký Tư Mã Thiên','Tư Mã Thiên','','Ly Dang','','available','','2026-05-15',''],
    [1149,'Xây dựng xã hội học tập','Joseph E Stiglitz, Bruce C Greenwald','','Ly Dang','','available','','2026-05-15',''],
    [1150,'Lược sử ngôn ngữ - Chuyện kể về phát minh vĩ đại nhất của loài người','Daniel L. Everett','','Ly Dang','','available','','2026-05-15',''],
    [1151,'Red Nile - Tiểu sử của dòng sông vĩ đại nhất thế giới (BC)','Robert Twigger','','Ly Dang','','available','','2026-05-15',''],
    [1152,'Bản Sắc','Francis Fukuyama','','Ly Dang','','available','','2026-05-15',''],
    [1153,'Marco Polo - Từ Venice tới Thượng Đô (BM)','Laurence Bergreen','','Ly Dang','','available','','2026-05-15',''],
    [1154,'Bảy người chồng của Evelyn Hugo','Taylor Jenkins Reid','','Ly Dang','','available','','2026-05-15',''],
    [1155,'Những điều giữ tôi còn sống','Matt Haig','','Ly Dang','','available','','2026-05-15',''],
    [1156,'Người thu gió','William Kamkwamba','','Ly Dang','','available','','2026-05-15',''],
    [1157,'Phương Pháp Đọc Sách Hiệu Quả','Mortimer J.Adler, Charles Van Doren','','Ly Dang','','available','','2026-05-15',''],
    [1158,'Cuốn Sách Bạn Ước Cha Mẹ Mình Từng Đọc','Philippa Perry','','Ly Dang','','available','','2026-05-15',''],
    [1159,'The Builder\'s Guide to the Tech Galaxy','Martin Schilling, Thomas Klugkist','','Ly Dang','','available','','2026-05-15',''],
    [1160,'Người hướng nội trong thế giới hướng ngoại','Insook Nam','','Ly Dang','','available','','2026-05-15',''],
    [1161,'Tên của đoá hồng','Umberto Eco','','Ly Dang','','available','','2026-05-15',''],
    [1162,'Tái tạo tổ chức','Frederick Laloux','','Ly Dang','','available','','2026-05-15',''],
    [1163,'Tái Tạo Doanh Nghiệp - Một Hành Trình Quả Cảm','Aaron Dignan','','Ly Dang','','available','','2026-05-15',''],
    [1164,'Mẫu thượng ngàn','Nguyễn Xuân Khánh','','Ly Dang','','available','','2026-05-15',''],
    [1165,'Lịch sử chữ quốc ngữ (1615-1919)','Phạm Thị Kiều Ly','','Ly Dang','','available','','2026-05-15',''],
    [1166,'Bút ký người đi săn','i. s. turgenev','','Ly Dang','','available','','2026-05-15',''],
    [1167,'Đêm trường tăm tối','Tử Kim Trần','','Ly Dang','','available','','2026-05-15',''],
    [1168,'Người Châu Á có biết tư duy?','Kishore Mahbubani','','Ly Dang','','available','','2026-05-15',''],
    [1169,'Storytelling with data','knaflic','','Ly Dang','','available','','2026-05-15',''],
    [1170,'Hệ sinh thái Toyota','Taiichi Ohno','','Ly Dang','','available','','2026-05-15',''],
    [1171,'Thay đổi cuộc sống với thần số học','Lê Đỗ Quỳnh Hương','','Ly Dang','','available','','2026-05-15',''],
    [1172,'Tư duy đột phá','Shozo Hibino','','Ly Dang','','available','','2026-05-15',''],
    [1173,'Triết học','DK','','Ly Dang','','available','','2026-05-15',''],
    [1174,'12 quy luật cuộc đời','Jordan Peterson','','Ly Dang','','available','','2026-05-15',''],
    [1175,'Đến Sahara mở quán trà đá','Minh Phan','','Ly Dang','','available','','2026-05-15',''],
    [1176,'Nexus','Yuval Noah Harari','','Ly Dang','','available','','2026-05-15',''],
    [1177,'Dám hạnh phúc','Kishimi Ichiro','','Ly Dang','','available','','2026-05-15',''],
    [1178,'Hoá thân','Franz Kafka','','Ly Dang','','available','','2026-05-15',''],
    [1179,'Trắng','Han Kang','','Ly Dang','','available','','2026-05-15',''],
    [1180,'Dám bị ghét','Kishimi Ichiro','','Ly Dang','','available','','2026-05-15',''],
    [1181,'Xa lạ trong tôi','Orhan Pamuk','','Ly Dang','','available','','2026-05-15',''],
    [1182,'Nghe thổ dân kể chuyện dạy con','Michaeleen Doucleff','','Ly Dang','','available','','2026-05-15',''],
    [1183,'Sự Thật Trần Trụi Về Tiền','Charles Wheelan','','Ly Dang','','available','','2026-05-15',''],
    [1184,'Sự Thật Trần Trụi Về Thống Kê','Charles Wheelan','','Ly Dang','','available','','2026-05-15',''],
    [1185,'Tội Ác Và Hình Phạt','Fyodor Dostoevsky','','Ly Dang','','available','','2026-05-15',''],
    [1186,'Đời Sống Bí Ẩn Của Cây','Peter Wohlleben','','Ly Dang','','available','','2026-05-15',''],
    [1187,'Thế hệ lo âu','Jonathan Haidt','','Ly Dang','','available','','2026-05-15',''],
    [1188,'Con đường chẳng mấy ai đi','M. Scott Peck','','Ly Dang','','available','','2026-05-15',''],
    [1189,'Chùm nho phẫn nộ','John Steinbeck','','Ly Dang','','available','','2026-05-15',''],
    [1190,'Bố già (ấn bản kỷ niệm 55 năm xuất bản)','Mario Puzo','','Ly Dang','','available','','2026-05-15',''],
    [1191,'Khi mọi điều không như ý','Hae Min','','Ly Dang','','available','','2026-05-15',''],
    [1192,'Khúc bi ca của gã dân quê','J.D. Vance','','Ly Dang','','available','','2026-05-15',''],
    [1193,'Sao chúng ta lại ngủ','Matthew Walker','','Ly Dang','','available','','2026-05-15',''],
    [1194,'Bốn Thỏa Ước','Janet Mills, don Miguel Ruiz','','Ly Dang','','available','','2026-05-15',''],
    [1195,'Đúng việc','Giản Tư Trung','','Ly Dang','','available','','2026-05-15',''],
    [1196,'Hồi ký Nguyễn Thị Bình: Gia đình, bạn bè và đất nước','Nguyễn Thị Bình','','Ly Dang','','available','','2026-05-15',''],
    [1197,'Thành phố Hồ Chí Minh – Giờ Khắc Số 0','Borries Gallasch','','Ly Dang','','available','','2026-05-15',''],
    [1198,'Khúc ca của Tế bào','Siddhartha Mukherjee','','Ly Dang','','available','','2026-05-15',''],
    [1199,'Số 0 - Tiểu Sử Một Phát Kiến Nguy Hiểm','Charles Seife','','Ly Dang','','available','','2026-05-15',''],
    [1200,'Giải Trí Đến Chết','Neil Postman','','Ly Dang','','available','','2026-05-15',''],
    [1201,'Người Ăn Chay','Han Kang','','Ly Dang','','available','','2026-05-15',''],
    [1202,'Con người và biểu tượng','Carl Gustav Jung','','Ly Dang','','available','','2026-05-15',''],
    [1203,'Lịch sử Phật giáo','Andrew Skilton','','Ly Dang','','available','','2026-05-15',''],
    [1204,'Triết học về sự ham muốn','Frédéric Lenoir','','Ly Dang','','available','','2026-05-15',''],
    [1205,'Phong cách Nvidia','Tae Kim','','Ly Dang','','available','','2026-05-15',''],
    [1206,'Genesis Khởi nguyên','Henry A. Kissinger','','Ly Dang','','available','','2026-05-15',''],
    [1207,'Nghề viết tiểu thuyết','Haruki Murakami','','Ly Dang','','available','','2026-05-15',''],
    [1208,'Đừng trách tôi mê game - Là do người làm game quá hiểu tâm lý học','Puella','','Ly Dang','','available','','2026-05-15',''],
    [1209,'Hồi ức, giấc mơ, suy ngẫm','Carl Gustav Jung','','Ly Dang','','available','','2026-05-15',''],
    [1210,'Ván cờ bất tử','David Shenk','','Ly Dang','','available','','2026-05-15',''],
    [1211,'Sống đời giản đơn','Charles Wagner','','Ly Dang','','available','','2026-05-15',''],
    [1212,'Sanshiro','Natsume Soseki','','Ly Dang','','available','','2026-05-15',''],
    [1213,'Hồng lâu mộng','Tào Tuyết Cần','','Ly Dang','','available','','2026-05-15',''],
    [1214,'Phono Sapiens','Jae Boong Choi','','Ly Dang','','available','','2026-05-15',''],
    [1215,'Từ tốt đến vĩ đại','Jim Collins','','Ly Dang','','available','','2026-05-15',''],
    [1216,'Spotify và những chuyện chưa kể','Sven Carlsson','','Ly Dang','','available','','2026-05-15',''],
    [1217,'Tự bạch','Thánh Augustine','','Ly Dang','','available','','2026-05-15',''],
    [1218,'Lược sử vạn vật','Bill Bryson','','Ly Dang','','available','','2026-05-15',''],
    [1219,'Cộng hòa','Plato','','Ly Dang','','available','','2026-05-15',''],
    [1220,'Sơn ca vẫn hót','Kristin Hannah','','Ly Dang','','available','','2026-05-15',''],
    [1221,'Dogra Magra Tiếng vọng từ tiền kiếp','Yumeno Kyusaku','','Ly Dang','','available','','2026-05-15',''],
    [1222,'Quả chuông ác mộng','Sylvia Plath','','Ly Dang','','available','','2026-05-15',''],
    [1223,'Khi loài vật lên ngôi','Karel Capek','','Ly Dang','','available','','2026-05-15',''],
    [1224,'Tuyển tập Dino Buzzati','Dino Buzzati','','Ly Dang','','available','','2026-05-15',''],
    [1225,'Bốn mùa, trời và đất','Sándor Márai','','Ly Dang','','available','','2026-05-15',''],
    [1226,'Những ngọn nến cháy tàn','Sándor Márai','','Ly Dang','','available','','2026-05-15',''],
    [1227,'Hoành Sơn một dải','Tô Như','','Ly Dang','','available','','2026-05-15',''],
    [1228,'Tương lai nhân loại','Michio Kaku','','Ly Dang','','available','','2026-05-15',''],
    [1229,'Biên sử nước','Nguyễn Ngọc Tư','','Ly Dang','','available','','2026-05-15',''],
    [1230,'Pedro Paramo','Juan Rulfo','','Ly Dang','','available','','2026-05-15',''],
    [1231,'Shosha','Isaac Bashevis Singer','','Ly Dang','','available','','2026-05-15',''],
    [1232,'Giáo dục tình cảm','Gustave Flaubert','','Ly Dang','','available','','2026-05-15',''],
    [1233,'Vành đai Sao Thổ','W. G. Sebald','','Ly Dang','','available','','2026-05-15','']
  ];
  
  sheet.getRange(2, 1, books.length, 10).setValues(books);
  Logger.log('Inserted ' + books.length + ' books into Library.');
  
  // Step 3: Fetch missing covers
  var data = sheet.getDataRange().getValues();
  var hdr  = data[0];
  var ti   = hdr.indexOf('title');
  var ai   = hdr.indexOf('author');
  var ci   = hdr.indexOf('coverUrl');
  var count = 0;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][ci] && data[i][ci] !== '') continue;
    var title  = data[i][ti];
    var author = data[i][ai];
    if (!title) continue;
    try {
      var query = 'intitle:' + encodeURIComponent(title);
      if (author) query += '+inauthor:' + encodeURIComponent(author);
      var url = 'https://www.googleapis.com/books/v1/volumes?q=' + query + '&maxResults=1&fields=items/volumeInfo/imageLinks';
      var res  = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      var json = JSON.parse(res.getContentText());
      if (json.items && json.items[0] && json.items[0].volumeInfo && json.items[0].volumeInfo.imageLinks) {
        var cover = json.items[0].volumeInfo.imageLinks.thumbnail || '';
        cover = cover.replace('zoom=1','zoom=2').replace('http://','https://');
        sheet.getRange(i + 1, ci + 1).setValue(cover);
        count++;
      }
    } catch(e) { /* skip */ }
    Utilities.sleep(300);
  }
  Logger.log('Done! Fetched covers for ' + count + ' / ' + books.length + ' books.');
}
function fetchCoversOnly() {
  var sheet = getSpreadsheet().getSheetByName('Library');
  var data  = sheet.getDataRange().getValues();
  var hdr   = data[0];
  var ti    = hdr.indexOf('title');
  var ai    = hdr.indexOf('author');
  var ci    = hdr.indexOf('coverUrl');
  var fetched = 0, skipped = 0, failed = 0;

  for (var i = 1; i < data.length; i++) {
    if (data[i][ci] && String(data[i][ci]).trim() !== '') { skipped++; continue; }
    var title  = String(data[i][ti] || '').trim();
    var author = String(data[i][ai] || '').trim();
    if (!title) continue;
    var cover = '';

    // Nguồn 1: Google Books
    try {
      var q = encodeURIComponent(title);
      if (author) q += '+' + encodeURIComponent(author);
      var url = 'https://www.googleapis.com/books/v1/volumes?q=' + q + '&maxResults=1&fields=items/volumeInfo/imageLinks';
      var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      var json = JSON.parse(res.getContentText());
      if (json.items && json.items[0] && json.items[0].volumeInfo && json.items[0].volumeInfo.imageLinks) {
        cover = json.items[0].volumeInfo.imageLinks.thumbnail || '';
        cover = cover.replace('zoom=1','zoom=2').replace('http://','https://');
      }
    } catch(e) { Logger.log('GB err '+i+': '+e.message); }

    // Nguồn 2: Open Library (fallback)
    if (!cover) {
      try {
        var olUrl = 'https://openlibrary.org/search.json?title=' + encodeURIComponent(title) + '&limit=1&fields=cover_i';
        var olRes = UrlFetchApp.fetch(olUrl, { muteHttpExceptions: true });
        var olJson = JSON.parse(olRes.getContentText());
        if (olJson.docs && olJson.docs[0] && olJson.docs[0].cover_i) {
          cover = 'https://covers.openlibrary.org/b/id/' + olJson.docs[0].cover_i + '-M.jpg';
        }
      } catch(e) { Logger.log('OL err '+i+': '+e.message); }
    }

    if (cover) {
      sheet.getRange(i + 1, ci + 1).setValue(cover);
      fetched++;
      Logger.log('✅ '+title);
    } else {
      failed++;
      Logger.log('❌ '+title);
    }
    Utilities.sleep(400);
  }

  Logger.log('Fetched: '+fetched+' | Skipped: '+skipped+' | No cover: '+failed);
}
function testAuth() {
  var res = UrlFetchApp.fetch('https://www.googleapis.com/books/v1/volumes?q=test&maxResults=1');
  Logger.log(res.getContentText().substring(0, 100));
}
function fetchCoversOpenLibrary() {
  var sheet   = getSpreadsheet().getSheetByName('Library');
  var data    = sheet.getDataRange().getValues();
  var hdr     = data[0];
  var ti      = hdr.indexOf('title');
  var ai      = hdr.indexOf('author');
  var ci      = hdr.indexOf('coverUrl');
  var fetched = 0, failed = 0;

  for (var i = 1; i < data.length; i++) {
    // Skip nếu đã có cover
    if (data[i][ci] && String(data[i][ci]).trim() !== '') continue;

    var title  = String(data[i][ti] || '').trim();
    var author = String(data[i][ai] || '').trim();
    if (!title) continue;

    var cover = '';

    // Nguồn 1: Open Library search by title + author
    try {
      var q = 'title=' + encodeURIComponent(title);
      if (author) q += '&author=' + encodeURIComponent(author);
      var url = 'https://openlibrary.org/search.json?' + q + '&limit=1&fields=cover_i,isbn';
      var res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      var json = JSON.parse(res.getContentText());

      if (json.docs && json.docs.length > 0) {
        var doc = json.docs[0];
        // Ưu tiên cover_i
        if (doc.cover_i) {
          cover = 'https://covers.openlibrary.org/b/id/' + doc.cover_i + '-M.jpg';
        }
        // Fallback: dùng ISBN nếu có
        else if (doc.isbn && doc.isbn.length > 0) {
          cover = 'https://covers.openlibrary.org/b/isbn/' + doc.isbn[0] + '-M.jpg';
        }
      }
    } catch(e) {
      Logger.log('OL err row ' + (i+1) + ': ' + e.message);
    }

    // Nguồn 2: Tìm chỉ bằng title nếu không có kết quả
    if (!cover) {
      try {
        var url2 = 'https://openlibrary.org/search.json?title=' + encodeURIComponent(title) + '&limit=1&fields=cover_i';
        var res2  = UrlFetchApp.fetch(url2, { muteHttpExceptions: true });
        var json2 = JSON.parse(res2.getContentText());
        if (json2.docs && json2.docs[0] && json2.docs[0].cover_i) {
          cover = 'https://covers.openlibrary.org/b/id/' + json2.docs[0].cover_i + '-M.jpg';
        }
      } catch(e) {}
    }

    if (cover) {
      sheet.getRange(i + 1, ci + 1).setValue(cover);
      fetched++;
      Logger.log('✅ [' + (i) + '] ' + title);
    } else {
      failed++;
      Logger.log('❌ [' + (i) + '] ' + title);
    }

    // Nghỉ 200ms — Open Library không giới hạn quota nhưng vẫn nên lịch sự
    Utilities.sleep(200);
  }

  Logger.log('✅ Fetched: ' + fetched + ' | ❌ No cover: ' + failed);
}
function fetchCoversV3() {
  var sheet   = getSpreadsheet().getSheetByName('Library');
  var data    = sheet.getDataRange().getValues();
  var hdr     = data[0];
  var ti      = hdr.indexOf('title');
  var ai      = hdr.indexOf('author');
  var ci      = hdr.indexOf('coverUrl');
  var fetched = 0, failed = 0;

  for (var i = 1; i < data.length; i++) {
    if (data[i][ci] && String(data[i][ci]).trim() !== '') continue;
    var title  = String(data[i][ti] || '').trim();
    var author = String(data[i][ai] || '').trim();
    if (!title) continue;

    var cover = '';

    // ── Nguồn 1: Tiki Books (tốt cho sách tiếng Việt) ──────────────────
    if (!cover) {
      try {
        var tikiQ   = encodeURIComponent(title + (author ? ' ' + author : ''));
        var tikiUrl = 'https://tiki.vn/api/v2/products?q=' + tikiQ + '&category=8322&limit=3';
        var tikiRes = UrlFetchApp.fetch(tikiUrl, {
          muteHttpExceptions: true,
          headers: {
            'User-Agent': 'Mozilla/5.0',
            'Accept': 'application/json'
          }
        });
        var tikiJson = JSON.parse(tikiRes.getContentText());
        if (tikiJson.data && tikiJson.data.length > 0) {
          var item = tikiJson.data[0];
          cover = item.thumbnail_url || item.image_url || '';
          // Nâng quality ảnh Tiki
          if (cover) cover = cover.replace('/280x280/', '/500x500/');
        }
      } catch(e) {
        Logger.log('Tiki err row ' + (i+1) + ': ' + e.message);
      }
    }

    // ── Nguồn 2: Open Library (title + author) ──────────────────────────
    if (!cover) {
      try {
        var olQ   = 'title=' + encodeURIComponent(title) + (author ? '&author=' + encodeURIComponent(author) : '');
        var olUrl = 'https://openlibrary.org/search.json?' + olQ + '&limit=1&fields=cover_i,isbn';
        var olRes = UrlFetchApp.fetch(olUrl, { muteHttpExceptions: true });
        var olJson = JSON.parse(olRes.getContentText());
        if (olJson.docs && olJson.docs.length > 0) {
          var doc = olJson.docs[0];
          if (doc.cover_i) {
            cover = 'https://covers.openlibrary.org/b/id/' + doc.cover_i + '-M.jpg';
          } else if (doc.isbn && doc.isbn.length > 0) {
            cover = 'https://covers.openlibrary.org/b/isbn/' + doc.isbn[0] + '-M.jpg';
          }
        }
      } catch(e) {
        Logger.log('OL err row ' + (i+1) + ': ' + e.message);
      }
    }

    // ── Nguồn 3: Open Library (title only fallback) ─────────────────────
    if (!cover) {
      try {
        var olUrl2 = 'https://openlibrary.org/search.json?title=' + encodeURIComponent(title) + '&limit=1&fields=cover_i';
        var olRes2 = UrlFetchApp.fetch(olUrl2, { muteHttpExceptions: true });
        var olJson2 = JSON.parse(olRes2.getContentText());
        if (olJson2.docs && olJson2.docs[0] && olJson2.docs[0].cover_i) {
          cover = 'https://covers.openlibrary.org/b/id/' + olJson2.docs[0].cover_i + '-M.jpg';
        }
      } catch(e) {}
    }

    if (cover) {
      sheet.getRange(i + 1, ci + 1).setValue(cover);
      fetched++;
      Logger.log('✅ [' + i + '] ' + title);
    } else {
      failed++;
      Logger.log('❌ [' + i + '] ' + title);
    }

    Utilities.sleep(300);
  }

  Logger.log('=== DONE: ✅ ' + fetched + ' fetched | ❌ ' + failed + ' no cover ===');
}
function fetchCoversV3() {
  var sheet   = getSpreadsheet().getSheetByName('Library');
  var data    = sheet.getDataRange().getValues();
  var hdr     = data[0];
  var ti      = hdr.indexOf('title');
  var ai      = hdr.indexOf('author');
  var ci      = hdr.indexOf('coverUrl');
  var fetched = 0, failed = 0;

  for (var i = 1; i < data.length; i++) {
    if (data[i][ci] && String(data[i][ci]).trim() !== '') continue;
    var title  = String(data[i][ti] || '').trim();
    var author = String(data[i][ai] || '').trim();
    if (!title) continue;

    var cover = '';

    // ── Nguồn 1: Tiki Books (tốt cho sách tiếng Việt) ──────────────────
    if (!cover) {
      try {
        var tikiQ   = encodeURIComponent(title + (author ? ' ' + author : ''));
        var tikiUrl = 'https://tiki.vn/api/v2/products?q=' + tikiQ + '&category=8322&limit=3';
        var tikiRes = UrlFetchApp.fetch(tikiUrl, {
          muteHttpExceptions: true,
          headers: {
            'User-Agent': 'Mozilla/5.0',
            'Accept': 'application/json'
          }
        });
        var tikiJson = JSON.parse(tikiRes.getContentText());
        if (tikiJson.data && tikiJson.data.length > 0) {
          var item = tikiJson.data[0];
          cover = item.thumbnail_url || item.image_url || '';
          // Nâng quality ảnh Tiki
          if (cover) cover = cover.replace('/280x280/', '/500x500/');
        }
      } catch(e) {
        Logger.log('Tiki err row ' + (i+1) + ': ' + e.message);
      }
    }

    // ── Nguồn 2: Open Library (title + author) ──────────────────────────
    if (!cover) {
      try {
        var olQ   = 'title=' + encodeURIComponent(title) + (author ? '&author=' + encodeURIComponent(author) : '');
        var olUrl = 'https://openlibrary.org/search.json?' + olQ + '&limit=1&fields=cover_i,isbn';
        var olRes = UrlFetchApp.fetch(olUrl, { muteHttpExceptions: true });
        var olJson = JSON.parse(olRes.getContentText());
        if (olJson.docs && olJson.docs.length > 0) {
          var doc = olJson.docs[0];
          if (doc.cover_i) {
            cover = 'https://covers.openlibrary.org/b/id/' + doc.cover_i + '-M.jpg';
          } else if (doc.isbn && doc.isbn.length > 0) {
            cover = 'https://covers.openlibrary.org/b/isbn/' + doc.isbn[0] + '-M.jpg';
          }
        }
      } catch(e) {
        Logger.log('OL err row ' + (i+1) + ': ' + e.message);
      }
    }

    // ── Nguồn 3: Open Library (title only fallback) ─────────────────────
    if (!cover) {
      try {
        var olUrl2 = 'https://openlibrary.org/search.json?title=' + encodeURIComponent(title) + '&limit=1&fields=cover_i';
        var olRes2 = UrlFetchApp.fetch(olUrl2, { muteHttpExceptions: true });
        var olJson2 = JSON.parse(olRes2.getContentText());
        if (olJson2.docs && olJson2.docs[0] && olJson2.docs[0].cover_i) {
          cover = 'https://covers.openlibrary.org/b/id/' + olJson2.docs[0].cover_i + '-M.jpg';
        }
      } catch(e) {}
    }

    if (cover) {
      sheet.getRange(i + 1, ci + 1).setValue(cover);
      fetched++;
      Logger.log('✅ [' + i + '] ' + title);
    } else {
      failed++;
      Logger.log('❌ [' + i + '] ' + title);
    }

    Utilities.sleep(300);
  }

  Logger.log('=== DONE: ✅ ' + fetched + ' fetched | ❌ ' + failed + ' no cover ===');
}