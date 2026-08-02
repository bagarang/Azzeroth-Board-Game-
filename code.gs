const SHEET_NAME = 'Users';
const ONLINE_SHEET = 'Online';
const INVITE_SHEET = 'Invites';

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Username', 'Password', 'SaveData']);
    }
    
    let invSheet = ss.getSheetByName(INVITE_SHEET);
    if (!invSheet) {
      invSheet = ss.insertSheet(INVITE_SHEET);
      invSheet.appendRow(['Id', 'From', 'To', 'Mode', 'Status', 'Timestamp', 'Payload']);
    }

    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const username = data.username;
    const password = data.password;
    const userData = sheet.getDataRange().getValues();
    
    if (action === 'ping') {
      let onlineSheet = ss.getSheetByName(ONLINE_SHEET);
      if (!onlineSheet) { onlineSheet = ss.insertSheet(ONLINE_SHEET); onlineSheet.appendRow(['Username', 'LastSeen']); }
      let rowFound = -1;
      let oData = onlineSheet.getDataRange().getValues();
      for (let i = 1; i < oData.length; i++) {
        if (oData[i][0] === username) { rowFound = i + 1; break; }
      }
      let now = new Date();
      if (rowFound !== -1) onlineSheet.getRange(rowFound, 2).setValue(now);
      else onlineSheet.appendRow([username, now]);
      
      let invData = invSheet.getDataRange().getValues();
      let pendingInvites = [];
      let inviteReplies = [];
      
      for (let i = 1; i < invData.length; i++) {
        if (invData[i][2] === username && invData[i][4] === 'pending') {
          pendingInvites.push({ id: invData[i][0], from: invData[i][1], mode: invData[i][3], payload: JSON.parse(invData[i][6]) });
        }
        if (invData[i][1] === username && (invData[i][4] === 'accepted' || invData[i][4] === 'rejected')) {
          inviteReplies.push({ id: invData[i][0], to: invData[i][2], mode: invData[i][3], status: invData[i][4], payload: JSON.parse(invData[i][6]) });
          invSheet.getRange(i + 1, 5).setValue('notified'); 
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, pendingInvites: pendingInvites, inviteReplies: inviteReplies })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // API PVP: Kirim undangan dan snapshot status P1
    if (action === 'sendInvite') {
      const invId = Date.now().toString();
      const initialPayload = { p1: data.payload, p2: null, combatState: null };
      invSheet.appendRow([invId, username, data.target, data.mode, 'pending', new Date(), JSON.stringify(initialPayload)]);
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    }

    // API PVP: Respon undangan dan buat Combat Shared State
    if (action === 'respondInvite') {
      let invData = invSheet.getDataRange().getValues();
      for(let i = 1; i < invData.length; i++) {
        if(invData[i][0] === data.inviteId) {
          invSheet.getRange(i + 1, 5).setValue(data.response);
          if (data.response === 'accepted') {
             let payload = JSON.parse(invData[i][6]);
             payload.p2 = data.payload; 
             payload.combatState = {
                 turn: 'dice',
                 rolls: {},
                 logs: ["The Duel Begins! Roll the dice."],
                 p1_hp: payload.p1.maxHp, p2_hp: payload.p2.maxHp,
                 p1_mana: payload.p1.maxMana, p2_mana: payload.p2.maxMana
             };
             invSheet.getRange(i + 1, 7).setValue(JSON.stringify(payload));
          }
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    }

    // API PVP: Sinkronisasi Aksi Pertempuran & Dadu
    if (action === 'duelAction') {
       let invData = invSheet.getDataRange().getValues();
       for(let i = 1; i < invData.length; i++) {
         if(invData[i][0] === data.inviteId) {
           let payload = JSON.parse(invData[i][6]);
           let newCs = data.combatState;

           // Server-side Dice Resolution (Mencegah Race Condition)
           if (payload.combatState.rolls && newCs.rolls) {
               newCs.rolls = { ...payload.combatState.rolls, ...newCs.rolls };
           }
           payload.combatState = newCs;

           if (payload.combatState.turn === 'dice' && payload.combatState.rolls[payload.p1.username] && payload.combatState.rolls[payload.p2.username]) {
               let r1 = payload.combatState.rolls[payload.p1.username];
               let r2 = payload.combatState.rolls[payload.p2.username];
               // P1 menang jika seri
               payload.combatState.turn = (r1 >= r2) ? payload.p1.username : payload.p2.username;
           }

           invSheet.getRange(i + 1, 7).setValue(JSON.stringify(payload));
           return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
         }
       }
    }

    // API PVP: Fetch Realtime Combat State
    if (action === 'getDuelState') {
        let invData = invSheet.getDataRange().getValues();
        for(let i = 1; i < invData.length; i++) {
           if(invData[i][0] === data.inviteId) {
              return ContentService.createTextOutput(JSON.stringify({ success: true, payload: JSON.parse(invData[i][6]) })).setMimeType(ContentService.MimeType.JSON);
           }
        }
    }

    if (action === 'getOnline') {
      let onlineSheet = ss.getSheetByName(ONLINE_SHEET);
      if (!onlineSheet) return ContentService.createTextOutput(JSON.stringify({ success: true, players: [] })).setMimeType(ContentService.MimeType.JSON);
      let oData = onlineSheet.getDataRange().getValues();
      let onlinePlayers = [];
      let now = new Date();
      for (let i = 1; i < oData.length; i++) {
        let lastSeen = new Date(oData[i][1]);
        let diff = (now - lastSeen) / 1000 / 60;
        if (diff < 2 && oData[i][0] !== username) onlinePlayers.push(oData[i][0]);
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, players: onlinePlayers })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'register') {
      for (let i = 1; i < userData.length; i++) {
        if (userData[i][0] === username) return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Username sudah digunakan' })).setMimeType(ContentService.MimeType.JSON);
      }
      sheet.appendRow([username, password, '']);
      return ContentService.createTextOutput(JSON.stringify({ success: true, message: 'Registrasi berhasil' })).setMimeType(ContentService.MimeType.JSON);
    } 
    
    if (action === 'login') {
      for (let i = 1; i < userData.length; i++) {
        if (userData[i][0] === username && userData[i][1] === password) return ContentService.createTextOutput(JSON.stringify({ success: true, saveData: userData[i][2] })).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Username atau password salah' })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'changePassword') {
      let rowFound = -1;
      for (let i = 1; i < userData.length; i++) {
        if (userData[i][0] === username && userData[i][1] === password) { rowFound = i + 1; break; }
      }
      if (rowFound !== -1) {
        sheet.getRange(rowFound, 2).setValue(data.newPassword);
        return ContentService.createTextOutput(JSON.stringify({ success: true, message: 'Password berhasil diganti' })).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Password lama salah' })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'save') {
      let rowFound = -1;
      for (let i = 1; i < userData.length; i++) {
        if (userData[i][0] === username && userData[i][1] === password) { rowFound = i + 1; break; }
      }
      if (rowFound !== -1) {
        sheet.getRange(rowFound, 3).setValue(data.saveData);
        return ContentService.createTextOutput(JSON.stringify({ success: true, message: 'Game tersimpan ke cloud' })).setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Gagal menyimpan' })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Aksi tidak dikenal' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Server Error: ' + err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}