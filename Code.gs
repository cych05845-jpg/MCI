/**
 * TTAS 快速檢傷 × Google Sheet 後端（v2）
 * ---------------------------------------------------------
 * 試算表分頁：
 *   檢傷紀錄  - 檢傷工具送出的完整紀錄（含「去向」欄，可由管理面板更新）
 *   傷患登錄  - 社工/輔助人員快速登錄的名單
 *   人員      - 登入名單：A欄=工號, B欄=姓名, C欄=角色(檢傷/社工/行政)
 *               （之後可改由人資系統匯出貼入，工具端不需改）
 *
 * 安裝：試算表 → 擴充功能 → Apps Script → 貼上本檔 → 部署為網頁應用程式
 *      （執行身分：我；存取權：任何人）
 * 金鑰：專案設定 → 指令碼屬性 → 新增 API_KEY（與工具設定內金鑰一致；不設則不驗證）
 *
 * 網址用法：
 *   {URL}                → 公開傷患清單看板（唯讀，10秒自動更新）
 *   {URL}?view=admin     → 行政管理面板（清單＋去向下拉即時寫回＋統計）
 *   {URL}?action=staff   → 人員名單 JSON（登入畫面用）
 */

const T_SHEET = '檢傷紀錄';
const R_SHEET = '傷患登錄';
const S_SHEET = '人員';

const T_HEADERS = ['收件時間','抵達時間','大量傷患編號','病歷號/流水號','患者','性別','年齡','屬性',
  '到院方式','活動狀態','意識(快速)','GCS','無法測量原因',
  '呼吸','脈搏','收縮壓','舒張壓','SpO2','體溫','微血管回充',
  '摘要','判斷依據','依據全文','綜合評級','最終級數','修改分級','備註','人員','模式','去向','已轉錄','事件'];
const R_HEADERS = ['收件時間','大量傷患編號','姓名','性別','生日','身分證/護照','國籍','地址/省市',
  '家用電話','手機','緊急聯絡人','關係','聯絡電話','完成連繫','家屬已到',
  '檢傷級數','傷情簡述','口述摘要','備註','登錄人員','去向','病歷號','事件','掛號建檔'];
function adminOk_(k){ const A = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || ''; return !!A && k === A; }
function activeMs_(){ const d = Number(PropertiesService.getScriptProperties().getProperty('EVENT_ACTIVE_DAYS') || 3); return d*86400000; }
function rowTs_(r){ return (r[0] && typeof r[0].getTime==='function') ? r[0].getTime() : Date.now(); }
function mask_(n){ n=String(n||''); if(!n)return n; return n.length<=1?n:(n[0]+'○'.repeat(Math.max(1,n.length-1))); }

const DISPOSITIONS = ['','檢傷區','急救區','治療區','觀察區','送檢(檢查)','送OR(手術)','送病房(住院)','留觀','出院','轉院','死亡','其他'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const KEY = PropertiesService.getScriptProperties().getProperty('API_KEY') || '';
    if (KEY && d.apiKey !== KEY) return json_({ ok:false, error:'unauthorized' });

    // 事件歸屬：預設以「當天日期」自動開檔；管理者手動設定 CURRENT_EVENT 時蓋過日期
    const CE = PropertiesService.getScriptProperties().getProperty('CURRENT_EVENT') || '';
    const effEvent = CE || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd');
    if (!adminOk_(d.adminKey)) d.event = effEvent;      // 一般人員：強制歸入
    else if (!d.event) d.event = effEvent;              // 管理者：未指定時亦帶預設

    if (d.mode === 'setevent') {                       // 開設/切換/清除目前事件（營運操作：一般金鑰即可）
      PropertiesService.getScriptProperties().setProperty('CURRENT_EVENT', String(d.value||'').trim());
      return json_({ ok:true, currentEvent: String(d.value||'').trim() });
    }
    if (d.mode === 'register') {                       // 大量傷患資料表登錄（同編號同事件→更新既有列，不新增）
      const sh = sheet_(R_SHEET, R_HEADERS);
      const Rh = function(h){ return R_HEADERS.indexOf(h); };
      const ser = String(d.serial||'').trim();
      const evv = String(d.event||'').trim();
      let targetRow = 0;
      if (ser) {
        const last = sh.getLastRow();
        if (last >= 2) {
          const vals = sh.getRange(2, 1, last-1, R_HEADERS.length).getValues();
          for (let i = vals.length-1; i >= 0; i--) {          // 由新到舊找同編號同事件
            if (String(vals[i][Rh('大量傷患編號')]).trim() === ser &&
                String(vals[i][Rh('事件')]||'').trim() === evv &&
                String(vals[i][Rh('去向')]||'').trim() !== '作廢') { targetRow = i+2; break; }
          }
        }
      }
      if (targetRow) {
        // 更新既有列：只覆蓋社工負責的欄位，保留掛號建檔、病歷號(若本次未帶)等他人欄位
        const setIf = function(h, val, always){
          const c = Rh(h)+1;
          if (always || (val !== undefined && val !== null && String(val) !== '')) sh.getRange(targetRow, c).setValue(val);
        };
        setIf('姓名', d.name); setIf('性別', d.sex); setIf('生日', d.birth);
        setIf('身分證/護照', d.natId); setIf('國籍', d.nation); setIf('地址/省市', d.addr);
        setIf('家用電話', d.telHome); setIf('手機', d.telMobile);
        setIf('緊急聯絡人', d.cName); setIf('關係', d.cRel); setIf('聯絡電話', d.cTel);
        setIf('完成連繫', d.cDone?'V':'', true); setIf('家屬已到', d.cArr?'V':'', true);
        setIf('傷情簡述', d.injury); setIf('口述摘要', d.summary); setIf('備註', d.memo);
        setIf('登錄人員', d.staff, true);
        if (String(d.chartNo||'').trim()) setIf('病歷號', d.chartNo);   // 有帶才覆蓋，不洗掉掛號填的
        return json_({ ok:true, row: targetRow, ev: evv, updated:true });
      }
      sh.appendRow([new Date(), d.serial||'', d.name||'', d.sex||'', d.birth||'',
        d.natId||'', d.nation||'', d.addr||'', d.telHome||'', d.telMobile||'',
        d.cName||'', d.cRel||'', d.cTel||'', d.cDone?'V':'', d.cArr?'V':'',
        d.triage||'', d.injury||'', d.summary||'', d.memo||'',
        d.staff||'', d.disposition||'', d.chartNo||'', d.event||'', '']);
      return json_({ ok:true, row: sh.getLastRow(), ev: d.event||'' });
    }
    if (d.mode === 'ping') {                           // 端點體檢
      const sp = PropertiesService.getScriptProperties();
      return json_({ ok:true, version:'v2025-09-01', hasApiKey: !!sp.getProperty('API_KEY'),
                     hasGemini: !!sp.getProperty('GEMINI_API_KEY'),
                     model: sp.getProperty('GEMINI_MODEL') || 'gemini-3.6-flash' });
    }
    if (d.mode === 'transcribe') {                     // 精準語音：Gemini 多模態
      const GKEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
      if (!GKEY) return json_({ ok:false, error:'請於「專案設定 → 指令碼屬性」新增 GEMINI_API_KEY' });
      const model = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-3.6-flash';
      const REG_PROMPT = '你是台灣急診大量傷患現場的資料登錄助手。請將這段台灣華語（可能夾雜台語）錄音轉為逐字稿，並擷取欄位。只輸出 JSON，格式：{"transcript":"逐字稿","fields":{"name":"病人姓名","sex":"男|女|","birth":"生日(民xx/mm/dd 或 yyyy/mm/dd)","natId":"身分證或護照號碼","telMobile":"手機","telHome":"市話","addr":"地址","nation":"本國|中國|其它|","cName":"緊急聯絡人姓名","cRel":"與病人關係","cTel":"聯絡人電話","cDone":false,"cArr":false}}。cDone=是否已聯繫上家屬、cArr=家屬是否已到院。聽不清或未提及的欄位一律給空字串或 false，絕不編造。';
      const TRI_PROMPT = '你是台灣急診檢傷站的紀錄助手。這段錄音是檢傷人員或病人的口述（台灣華語，可能夾雜台語與醫療用語）。請完成兩件事：(1) 精確逐字稿；(2) 將內容改寫為台灣急診檢傷紀錄慣用的精簡醫學摘要 note：涵蓋口述中提到的受傷/發病機轉、部位、傷型（如撕裂傷/穿刺傷/鈍傷）、症狀與程度（疼痛分數）、發生時間、出血與止血狀況、意識等重點；慣用句式如「自訴…」「現…」；保留所有數值與單位（公分、分、小時）。嚴格限制：note 只能整理口述中明確提到的資訊，絕不可推測、診斷或補充未提及的內容；聽不清處以「（不清）」標示。只輸出 JSON：{"transcript":"逐字稿","note":"醫學摘要（1-3句）"}。';
      const payload = {
        contents: [{ parts: [
          { inline_data: { mime_type: d.mime || 'audio/webm', data: d.audio } },
          { text: d.context === 'reg' ? REG_PROMPT : TRI_PROMPT }
        ]}],
        generationConfig: { response_mime_type: 'application/json', temperature: 0 }
      };
      const res = UrlFetchApp.fetch(
        'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + GKEY,
        { method:'post', contentType:'application/json', payload: JSON.stringify(payload), muteHttpExceptions:true });
      let j;
      try { j = JSON.parse(res.getContentText()); }
      catch(e) { return json_({ ok:false, error:'Gemini 回應無法解析 (HTTP ' + res.getResponseCode() + ')' }); }
      if (j.error) return json_({ ok:false, error:'Gemini: ' + j.error.message + '（音檔格式 ' + (d.mime||'?') + '）' });
      let txt = '';
      try { txt = j.candidates[0].content.parts[0].text; } catch (e) {
        var fr = ''; try { fr = j.candidates[0].finishReason || ''; } catch(e2) {}
        var pf = ''; try { pf = JSON.stringify(j.promptFeedback||''); } catch(e3) {}
        return json_({ ok:false, error:'Gemini 未回傳文字（finishReason=' + fr + ' promptFeedback=' + pf + '）' });
      }
      txt = String(txt).replace(/^\s*```(?:json)?/i, '').replace(/```\s*$/,'').trim();
      let out; try { out = JSON.parse(txt); } catch (e) { out = { transcript: txt }; }
      return json_({ ok:true, result: out });
    }
    if (d.mode === 'analyze') {                        // MCI 現場數據 AI 重點分析
      const GKA = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
      if (!GKA) return json_({ ok:false, error:'請於指令碼屬性新增 GEMINI_API_KEY' });
      const mdlA = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-3.6-flash';
      const AP = '你是大量傷患事件(MCI)的現場指揮參謀。以下是「去識別化」的檢傷統計數據(JSON)。請以繁體中文，用條列給出指揮官最該注意的重點，聚焦：1)傷情嚴重度分布與重症負荷 2)到量趨勢與尖峰 3)資源與去向瓶頸(未定去向、待登錄、家屬未連繫) 4)風險提醒與建議下一步。務實精簡，每點一行，不超過8點；只根據數據，不杜撰個案細節，不含任何姓名或病歷號。\n\n數據：' + (d.payload||'');
      const payA = { contents:[{parts:[{text:AP}]}], generationConfig:{temperature:0.2} };
      const resA = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/'+mdlA+':generateContent?key='+GKA,
        { method:'post', contentType:'application/json', payload:JSON.stringify(payA), muteHttpExceptions:true });
      let jA; try{ jA=JSON.parse(resA.getContentText()); }catch(e){ return json_({ ok:false, error:'HTTP '+resA.getResponseCode() }); }
      if (jA.error) return json_({ ok:false, error:'Gemini: '+jA.error.message });
      let tA=''; try{ tA=jA.candidates[0].content.parts[0].text; }catch(e){ return json_({ ok:false, error:'Gemini 未回傳文字' }); }
      return json_({ ok:true, text:tA });
    }
    if (d.mode === 'chart_ocr') {                      // 條碼/病歷號圖片 → 數字
      const GKc = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
      if (!GKc) return json_({ ok:false, error:'請於指令碼屬性新增 GEMINI_API_KEY' });
      const mdlc = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-3.6-flash';
      const CP = '這是台灣醫院大量傷患標籤/手圈照片。標籤格式常為「病歷號-檢查碼 大量NNN」，例如「1497834-0 大量023」代表病歷號1497834、大量傷患編號023。請讀出：病歷號（開頭那串數字，7-8碼，不含-後檢查碼）、大量傷患編號（「大量」後面的數字，如023）。只輸出 JSON：{"chartNo":"病歷號數字或null","mciNo":"大量傷患編號數字或null","barcode":"條碼下方完整數字或null"}。只讀實際可見的，絕不推測。';
      const payc = { contents:[{parts:[{inline_data:{mime_type:d.mime||'image/jpeg',data:d.image}},{text:CP}]}],
        generationConfig:{response_mime_type:'application/json',temperature:0} };
      const resc = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/'+mdlc+':generateContent?key='+GKc,
        { method:'post', contentType:'application/json', payload:JSON.stringify(payc), muteHttpExceptions:true });
      let jc; try{ jc=JSON.parse(resc.getContentText()); }catch(e){ return json_({ ok:false, error:'HTTP '+resc.getResponseCode() }); }
      if (jc.error) return json_({ ok:false, error:'Gemini: '+jc.error.message });
      let tc=''; try{ tc=jc.candidates[0].content.parts[0].text; }catch(e){ return json_({ ok:false, error:'Gemini 未回傳文字' }); }
      tc = String(tc).replace(/^\s*```(?:json)?/i,'').replace(/```\s*$/,'').trim();
      let oc; try{ oc=JSON.parse(tc); }catch(e){ oc={ }; }
      return json_({ ok:true, result:oc });
    }
    if (d.mode === 'doc_ocr') {                        // 到院文件照片 → 摘要＋生命徵象＋基本資料
      const GKd = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
      if (!GKd) return json_({ ok:false, error:'請於指令碼屬性新增 GEMINI_API_KEY' });
      const mdld = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-3.6-flash';
      const DP = '這是台灣急診檢傷/演習病歷卡或傷病患隨身文件照片（可能含姓名年齡性別、生日、身分證、住址、送達方式、主訴、理學檢查、生命徵象、GCS、毛細管回填、可否行走、慢性病史、過敏、血型、緊急聯絡人）。請完整擷取並輸出 JSON：{"note":"將主訴與理學檢查整理為台灣急診檢傷慣用醫學摘要(1-4句，含發病時間/症狀/次數/理學發現)","vitals":{"sbp":收縮壓數字或null,"dbp":舒張壓數字或null,"hr":null,"rr":null,"spo2":null,"bt":體溫數字或null,"gcs":GCS總分或null,"capRefill":毛細管回填秒數或null,"walk":"是|否|"},"patient":{"name":"","sex":"男|女|","age":年齡數字或null,"birth":"yyyy/mm/dd或民式","natId":"身分證/護照","addr":"住址","arrival":"送達方式"},"history":{"chronic":"慢性病史","allergy":"過敏(如NKDA)","bloodType":"血型"},"contact":{"name":"緊急聯絡人姓名","relation":"關係","phone":"電話"},"rawText":"文件可辨識全文"}。嚴格限制：只擷取文件明確記載的，絕不推測杜撰；未記載給null或空字串；數值只保留數字(如BP:114/72→sbp:114,dbp:72)。';
      const payd = { contents:[{parts:[{inline_data:{mime_type:d.mime||'image/jpeg',data:d.image}},{text:DP}]}],
        generationConfig:{response_mime_type:'application/json',temperature:0} };
      const resd = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/'+mdld+':generateContent?key='+GKd,
        { method:'post', contentType:'application/json', payload:JSON.stringify(payd), muteHttpExceptions:true });
      let jd; try{ jd=JSON.parse(resd.getContentText()); }catch(e){ return json_({ ok:false, error:'HTTP '+resd.getResponseCode() }); }
      if (jd.error) return json_({ ok:false, error:'Gemini: '+jd.error.message });
      let td=''; try{ td=jd.candidates[0].content.parts[0].text; }catch(e){ return json_({ ok:false, error:'Gemini 未回傳文字' }); }
      td = String(td).replace(/^\s*```(?:json)?/i,'').replace(/```\s*$/,'').trim();
      let od; try{ od=JSON.parse(td); }catch(e){ od={ note:td }; }
      return json_({ ok:true, result:od });
    }
    if (d.mode === 'vitals_ocr') {                     // 生命徵象照片讀值
      const GK = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
      if (!GK) return json_({ ok:false, error:'請於指令碼屬性新增 GEMINI_API_KEY' });
      const mdl = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-3.6-flash';
      const VP = '這是一張生命徵象監視器、血壓計或體溫計的照片。請只讀取螢幕/顯示器上實際顯示的數值，絕不可推測或補充。輸出 JSON：{"sbp":收縮壓或null,"dbp":舒張壓或null,"hr":每分鐘心跳或null,"rr":每分鐘呼吸或null,"spo2":血氧飽和度%或null,"bt":體溫°C或null}。判讀提示：監視器上心跳常為綠色大字、SpO2 常為藍/青色、血壓格式為「收縮壓/舒張壓」、呼吸為 RR 或 Resp。看不清或未顯示的項目一律 null。';
      const vpay = { contents:[{parts:[{inline_data:{mime_type:d.mime||'image/jpeg',data:d.image}},{text:VP}]}],
        generationConfig:{response_mime_type:'application/json',temperature:0} };
      const vres = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/'+mdl+':generateContent?key='+GK,
        { method:'post', contentType:'application/json', payload:JSON.stringify(vpay), muteHttpExceptions:true });
      let vj; try{ vj=JSON.parse(vres.getContentText()); }catch(e){ return json_({ ok:false, error:'HTTP '+vres.getResponseCode() }); }
      if (vj.error) return json_({ ok:false, error:'Gemini: '+vj.error.message });
      let vt=''; try{ vt=vj.candidates[0].content.parts[0].text; }catch(e){ return json_({ ok:false, error:'Gemini 未回傳文字' }); }
      vt = String(vt).replace(/^\s*```(?:json)?/i,'').replace(/```\s*$/,'').trim();
      let vo; try{ vo=JSON.parse(vt); }catch(e){ return json_({ ok:false, error:'讀值解析失敗' }); }
      return json_({ ok:true, result:vo });
    }
    if (d.mode === 'polish') {                         // 文字整理為醫學摘要
      const GKEY2 = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
      if (!GKEY2) return json_({ ok:false, error:'請於指令碼屬性新增 GEMINI_API_KEY' });
      const model2 = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-3.6-flash';
      const NP = '你是台灣急診檢傷站的紀錄助手。請將以下口述/筆記文字改寫為台灣急診檢傷紀錄慣用的精簡醫學摘要：涵蓋文中提到的受傷/發病機轉、部位、傷型、症狀與程度（疼痛分數）、時間、出血與止血、意識等重點；慣用句式如「自訴…」「現…」；保留所有數值與單位。嚴格限制：只能整理文中明確提到的資訊，絕不可推測、診斷或補充。只輸出 JSON：{"note":"醫學摘要（1-3句）"}。\n\n文字：' + (d.text || '');
      const pay2 = { contents:[{parts:[{text:NP}]}], generationConfig:{response_mime_type:'application/json',temperature:0} };
      const res2 = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/'+model2+':generateContent?key='+GKEY2,
        { method:'post', contentType:'application/json', payload:JSON.stringify(pay2), muteHttpExceptions:true });
      let j2; try{ j2 = JSON.parse(res2.getContentText()); }catch(e){ return json_({ ok:false, error:'HTTP '+res2.getResponseCode() }); }
      if (j2.error) return json_({ ok:false, error:'Gemini: '+j2.error.message });
      let t2=''; try{ t2=j2.candidates[0].content.parts[0].text; }catch(e){ return json_({ ok:false, error:'Gemini 未回傳文字' }); }
      t2 = String(t2).replace(/^\s*```(?:json)?/i,'').replace(/```\s*$/,'').trim();
      let o2; try{ o2=JSON.parse(t2); }catch(e){ o2={ note:t2 }; }
      return json_({ ok:true, result:o2 });
    }
    if (d.mode === 'regedit') {                        // 掛號/社工補正基本資料（白名單欄位）
      const shE = sheet_(R_SHEET, R_HEADERS);
      const mapE = { name:'姓名', sex:'性別', birth:'生日', natId:'身分證/護照', telMobile:'手機' };
      Object.keys(mapE).forEach(function(k){
        if (d[k] !== undefined) shE.getRange(Number(d.row), R_HEADERS.indexOf(mapE[k]) + 1).setValue(String(d[k]||''));
      });
      return json_({ ok:true });
    }
    if (d.mode === 'regmark') {                        // 掛號建檔完成確認
      const sh = sheet_(R_SHEET, R_HEADERS);
      var mk = d.value ? ('V ' + Utilities.formatDate(new Date(),'Asia/Taipei','HH:mm')) : '';
      sh.getRange(Number(d.row), R_HEADERS.indexOf('掛號建檔') + 1).setValue(mk);
      return json_({ ok:true });
    }
    if (d.mode === 'regchart') {                       // 掛號回填病歷號（登錄表）
      const sh = sheet_(R_SHEET, R_HEADERS);
      sh.getRange(Number(d.row), R_HEADERS.indexOf('病歷號') + 1).setValue(d.value || '');
      return json_({ ok:true });
    }
    if (d.mode === 'void') {                           // 作廢一筆（不刪列，避免位移）
      const isR = d.sheet === 'R';
      const sh = isR ? sheet_(R_SHEET, R_HEADERS) : sheet_(T_SHEET, T_HEADERS);
      const H = isR ? R_HEADERS : T_HEADERS;
      sh.getRange(Number(d.row), H.indexOf('去向') + 1).setValue('作廢');
      if (!isR) sh.getRange(Number(d.row), H.indexOf('已轉錄') + 1).setValue('V');
      return json_({ ok:true });
    }
    if (d.mode === 'edit') {                           // 事後修正（白名單欄位）
      const map = { mciNo:'大量傷患編號', level:'最終級數', basis:'依據全文', summary:'摘要' };
      if (!map[d.field]) return json_({ ok:false, error:'不允許修改此欄位' });
      const sh = sheet_(T_SHEET, T_HEADERS);
      sh.getRange(Number(d.row), T_HEADERS.indexOf(map[d.field]) + 1).setValue(d.value || '');
      return json_({ ok:true });
    }
    if (d.mode === 'chartno') {                        // 補登/更新病歷號（檢傷紀錄）
      const sh = sheet_(T_SHEET, T_HEADERS);
      sh.getRange(Number(d.row), T_HEADERS.indexOf('病歷號/流水號') + 1).setValue(d.value || '');
      return json_({ ok:true });
    }
    if (d.mode === 'mark') {                           // 標記已轉錄至正式檢傷
      const sh = sheet_(T_SHEET, T_HEADERS);
      sh.getRange(Number(d.row), T_HEADERS.indexOf('已轉錄') + 1).setValue(d.value ? 'V' : '');
      return json_({ ok:true });
    }
    if (d.mode === 'disposition') {                    // 去向更新（行政）
      const sh = d.sheet === 'R' ? sheet_(R_SHEET, R_HEADERS) : sheet_(T_SHEET, T_HEADERS);
      const col = (d.sheet === 'R' ? R_HEADERS : T_HEADERS).indexOf('去向') + 1;
      sh.getRange(Number(d.row), col).setValue(d.value || '');
      return json_({ ok:true });
    }
    // 檢傷紀錄
    const sh = sheet_(T_SHEET, T_HEADERS);
    const v = d.vitals||{}, c = d.consciousness||{};
    sh.appendRow([new Date(), d.arrTime||'', d.mciNo||'', d.chartNo||'', d.name||'', d.sex||'', d.age,
      d.attr||'', d.arrival||'', d.mobility||'',
      ({A:'清醒',D:'嗜睡',V:'聲音反應',P:'疼痛反應',U:'無反應'})[c.quick]||'',
      (c.gcs&&c.gcs.total)||'', v.unmeasurable?(v.reason||'未填原因'):'',
      v.rr, v.hr, v.sbp, v.dbp, v.spo2, v.bt, v.capRefill||'',
      d.summary||'', (d.criteria||[]).map(x=>'['+x.level+']'+x.text).join('；'),
      d.basis||'', d.autoLevel, d.finalLevel, d.override, d.memo||'',
      d.staff||'', d.mode||'', d.disposition||'', '', d.event||'']);
    return json_({ ok:true, row: sh.getLastRow(), ev: d.event||'' });
  } catch(err) { return json_({ ok:false, error:String(err) }); }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'staff')  return json_(getStaff_());
  if (p.action === 'lookup') return json_(lookup_(p.serial, p.key, p.chartNo, p.event||'', p.adminkey||''));
  if (p.action === 'list')   return json_(list_(p.key, p.event||'', p.adminkey||'', p.from||'', p.to||''));
  if (p.action === 'events') return json_(events_(p.key, p.adminkey||'', p.from||'', p.to||''));
  if (p.action === 'hislookup') return json_({ ok:false, error:'HIS/掛號系統介接尚未啟用（預留接口：由資訊室實作病歷號查詢後回傳 {ok:true,name,sex,birth,natId}）' });
  if (p.view === 'admin')    return adminPage_(p.event||'', p.adminkey||'');
  if (p.view === 'app')      return appPage_();
  return boardPage_(p.event||'');
}

/* ---------- 依大量傷患編號查詢（檢傷↔登錄互通） ---------- */
function lookup_(serial, key, chartNo, ev, adminkey) {
  const KEY = PropertiesService.getScriptProperties().getProperty('API_KEY') || '';
  if (KEY && key !== KEY) return { ok:false, error:'unauthorized' };
  serial = String(serial||'').trim(); chartNo = String(chartNo||'').trim();
  const act = Date.now() - activeMs_(), yr1 = Date.now() - 365*86400000;
  const tOk = function(r){ return ev ? rowTs_(r) >= yr1 : rowTs_(r) >= act; };
  if (!serial && !chartNo) return { ok:false, error:'no serial' };
  const T = h => T_HEADERS.indexOf(h), R = h => R_HEADERS.indexOf(h);
  const tv = data_(sheet_(T_SHEET,T_HEADERS), T_HEADERS.length);
  let triage = null;
  for (let k = tv.length-1; k >= 0; k--) {
    var evT = String(tv[k][T('事件')]||'').trim();
    if (tOk(tv[k]) && (!ev || evT === ev) &&
        ((serial && String(tv[k][T('大量傷患編號')]).trim() === serial) ||
        (chartNo && String(tv[k][T('病歷號/流水號')]).trim() === chartNo))) {
      const r = tv[k];
      triage = { name:r[T('患者')], sex:r[T('性別')], chartNo:r[T('病歷號/流水號')],
        mciNo:r[T('大量傷患編號')], time:fmtT_(r[T('抵達時間')]),
        level:r[T('最終級數')], criteria:r[T('判斷依據')], basis:r[T('依據全文')],
        summary:r[T('摘要')] };
      break;
    }
  }
  const rv = data_(sheet_(R_SHEET,R_HEADERS), R_HEADERS.length);
  let register = null;
  for (let k = rv.length-1; k >= 0; k--) {
    var evR = String(rv[k][R('事件')]||'').trim();
    if (tOk(rv[k]) && (!ev || evR === ev) &&
        ((serial && String(rv[k][R('大量傷患編號')]).trim() === serial) ||
        (chartNo && String(rv[k][R('病歷號')]).trim() === chartNo))) {
      const r = rv[k];
      register = { name:r[R('姓名')], sex:r[R('性別')], birth:fmtD_(r[R('生日')]),
        natId:r[R('身分證/護照')], telMobile:r[R('手機')], chartNo:r[R('病歷號')] };
      break;
    }
  }
  return { ok:true, serial:serial, triage:triage, register:register };
}

/* ---------- 事件結案歸檔（試算表選單） ---------- */
/* 開啟試算表 → 上方選單「MCI系統 → 歸檔（結案）事件」
   將該事件於「檢傷紀錄」「傷患登錄」的所有資料列，搬移到一份全新的
   歸檔試算表（MCI歸檔_事件_日期），主表隨之清空該事件。
   ⚠ 請於事件結束、確認無人正在系統作業時執行（歸檔會使主表列號重排）。
   ⚠ 歸檔檔案產生後，請至雲端硬碟將其存取權設為「僅限管理者」。 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('MCI系統')
    .addItem('歸檔（結案）事件…', 'archiveEventUi_')
    .addToUi();
}
function archiveEventUi_() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('事件結案歸檔',
    '輸入要歸檔的事件代號（該事件資料將搬移至新的歸檔試算表，主表清空該事件）：',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const ev = r.getResponseText().trim();
  if (!ev) { ui.alert('未輸入事件代號。'); return; }
  const c = ui.alert('確認', '確定歸檔事件「' + ev + '」？\n請確認現場已無人作業。', ui.ButtonSet.YES_NO);
  if (c !== ui.Button.YES) return;
  const res = archiveEvent_(ev);
  ui.alert(res);
}
function archiveEvent_(ev) {
  const arch = SpreadsheetApp.create('MCI歸檔_' + ev + '_' +
    Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmm'));
  let moved = 0;
  [[T_SHEET, T_HEADERS], [R_SHEET, R_HEADERS]].forEach(function(cfg) {
    const sh = sheet_(cfg[0], cfg[1]);
    const evI = cfg[1].indexOf('事件');
    const last = sh.getLastRow();
    const dst = arch.insertSheet(cfg[0]);
    dst.appendRow(cfg[1]);
    dst.getRange(1,1,1,cfg[1].length).setFontWeight('bold').setBackground('#DCDCD2');
    if (last < 2) return;
    const vals = sh.getRange(2, 1, last - 1, cfg[1].length).getValues();
    const keep = [], move = [];
    vals.forEach(function(row) {
      (String(row[evI] || '').trim() === ev ? move : keep).push(row);
    });
    if (!move.length) return;
    dst.getRange(2, 1, move.length, cfg[1].length).setValues(move);
    moved += move.length;
    sh.getRange(2, 1, last - 1, cfg[1].length).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, cfg[1].length).setValues(keep);
  });
  const s1 = arch.getSheets()[0];
  if (arch.getSheets().length > 2) arch.deleteSheet(s1);
  if (!moved) return '找不到事件「' + ev + '」的資料，未做任何變更。';
  return '已歸檔 ' + moved + ' 筆 → 「' + arch.getName() + '」\n' + arch.getUrl() +
    '\n\n請至雲端硬碟將該歸檔檔案的存取權設為僅限管理者。';
}

/* ---------- 事件瀏覽 API ---------- */
function events_(key, adminkey, from, to) {
  const KEY = PropertiesService.getScriptProperties().getProperty('API_KEY') || '';
  if (KEY && key !== KEY && !adminOk_(adminkey)) return { ok:false, error:'unauthorized' };
  const elev = adminOk_(adminkey);
  const now = Date.now(), yr = now - 365*86400000, act = now - activeMs_();
  const map = {};
  [[T_SHEET,T_HEADERS],[R_SHEET,R_HEADERS]].forEach(function(cfg){
    const H = cfg[1], evI = H.indexOf('事件');
    data_(sheet_(cfg[0],H), H.length).forEach(function(r){
      const e = String(r[evI]||'').trim(); if(!e) return;
      const t = rowTs_(r);
      const m = map[e] = map[e] || { name:e, start:t, end:t, count:0 };
      m.count++; if(t<m.start)m.start=t; if(t>m.end)m.end=t;
    });
  });
  let evs = Object.keys(map).map(function(k){return map[k];}).filter(function(m){return m.end>=yr;});
  if (!from && !to) evs = evs.filter(function(m){return m.end>=act;});   // 無查詢：僅進行中（近N天）
  if (from) { const f=new Date(from).getTime(); if(!isNaN(f)) evs=evs.filter(function(m){return m.end>=f;}); }
  if (to)   { const t2=new Date(to).getTime()+86400000; if(!isNaN(t2)) evs=evs.filter(function(m){return m.start<=t2;}); }
  evs.sort(function(a,b){return b.end-a.end;});
  return { ok:true, admin:elev, activeDays:activeMs_()/86400000,
    pinned: !!PropertiesService.getScriptProperties().getProperty('CURRENT_EVENT'),
    currentEvent: PropertiesService.getScriptProperties().getProperty('CURRENT_EVENT') ||
      Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd'),
    events: evs.map(function(m){return {
    name:m.name, count:m.count,
    start:Utilities.formatDate(new Date(m.start),'Asia/Taipei','yyyy/MM/dd'),
    end:Utilities.formatDate(new Date(m.end),'Asia/Taipei','yyyy/MM/dd') };}) };
}

/* ---------- 檢傷工具本體（由 Apps Script 直接提供） ---------- */
/* 使用方式：Apps Script 編輯器左側「檔案 ＋」→ HTML → 命名 index
   → 把工具的 index.html 全部內容貼入該檔案 → 儲存 → 部署新版本
   → 開啟 {exec網址}?view=app 即為完整工具，端點與金鑰自動注入。 */
function appPage_() {
  let html;
  try {
    html = HtmlService.createHtmlOutputFromFile('index').getContent();
  } catch (e) {
    return HtmlService.createHtmlOutput(
      '<meta charset="utf-8"><body style="font-family:sans-serif;padding:20px">' +
      '<h3>尚未安裝工具頁面</h3>' +
      '<p>請在 Apps Script 編輯器：左側「檔案」旁 ＋ → HTML → 命名 <b>index</b>，' +
      '把檢傷工具 index.html 的全部內容貼入後儲存，並以新版本重新部署。</p></body>');
  }
  const sp = PropertiesService.getScriptProperties();
  const selfUrl = ScriptApp.getService().getUrl();
  html = html
    .replace('api: ""', 'api: "' + selfUrl + '"')
    .replace('key: ""', 'key: "' + (sp.getProperty('API_KEY') || '') + '"');
  return HtmlService.createHtmlOutput(html)
    .setTitle('急診檢傷系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- 清單 API（主畫面用） ---------- */
function list_(key, ev, adminkey, from, to) {
  const KEY = PropertiesService.getScriptProperties().getProperty('API_KEY') || '';
  if (KEY && key !== KEY) return { ok:false, error:'unauthorized' };
  const T = h => T_HEADERS.indexOf(h), R = h => R_HEADERS.indexOf(h);
  const notHdr = function(o){ return String(o.r[0]||'') !== '收件時間'; };  // 略過殘留標題列
  const tvAll = data_(sheet_(T_SHEET,T_HEADERS), T_HEADERS.length).map(function(r,i){return {r:r,row:i+2};}).filter(notHdr);
  const rvAll = data_(sheet_(R_SHEET,R_HEADERS), R_HEADERS.length).map(function(r,i){return {r:r,row:i+2};}).filter(notHdr);
  const evSet = {};
  tvAll.forEach(function(o){ const e=String(o.r[T('事件')]||'').trim(); if(e) evSet[e]=1; });
  rvAll.forEach(function(o){ const e=String(o.r[R('事件')]||'').trim(); if(e) evSet[e]=1; });
  const evMatch = function(v){ v=String(v||'').trim(); return !ev ? true : v===ev; };  // 嚴格分批：空白事件僅於「全部事件」可見
  const now = Date.now(), yr = now-365*86400000, act = now-activeMs_();
  let fT = 0, tT = Infinity;
  if (from) { const f=new Date(from).getTime(); if(!isNaN(f)) fT=f; }
  if (to)   { const t2=new Date(to).getTime()+86400000; if(!isNaN(t2)) tT=t2; }
  const hasRange = !!(from||to);
  const timeOk = function(o){ const t=rowTs_(o.r);
    if (ev) return t>=yr;                                   // 指定事件夾：一年內皆可見
    if (hasRange) return t>=Math.max(yr,fT) && t<=tT;       // 時間查詢：一年內區間
    return t>=act; };                                       // 預設：進行中窗口（近N天）
  const tv = tvAll.filter(function(o){ return evMatch(o.r[T('事件')]) && timeOk(o); });
  const rv = rvAll.filter(function(o){ return evMatch(o.r[R('事件')]) && timeOk(o); });
  const regSerials = {}, regCharts = {};
  rv.forEach(function(o){
    const s=String(o.r[R('大量傷患編號')]).trim(); if(s) regSerials[s]=true;
    const c=String(o.r[R('病歷號')]).trim(); if(c) regCharts[c]=true;
  });
  const triage = tv.map(function(o){ const r=o.r; return {
    row:o.row, mciNo:r[T('大量傷患編號')], chartNo:r[T('病歷號/流水號')],
    time:fmtT_(r[T('抵達時間')]), name:r[T('患者')], sex:r[T('性別')], attr:r[T('屬性')],
    level:r[T('最終級數')], criteria:r[T('判斷依據')], basis:r[T('依據全文')], summary:r[T('摘要')],
    gcs:r[T('GCS')], spo2:r[T('SpO2')], hr:r[T('脈搏')], sbp:r[T('收縮壓')], dbp:r[T('舒張壓')],
    rr:r[T('呼吸')], bt:r[T('體溫')], mobility:r[T('活動狀態')], consciousness:r[T('意識(快速)')],
    dispo:r[T('去向')], registered:(!!regSerials[String(r[T('大量傷患編號')]).trim()] || !!regCharts[String(r[T('病歷號/流水號')]).trim()]),
    done:!!r[T('已轉錄')], event:r[T('事件')]||'', autoLv:r[T('綜合評級')] }; }).slice(-300);
  // 同一大量傷患編號：僅保留收件時間最新一筆（避免多筆登錄造成各畫面抓到不同版本）
  const regBest = {};
  rv.forEach(function(o){ const r=o.r;
    const s = String(r[R('大量傷患編號')]||'').trim();
    const ts = rowTs_(r);
    if (!s) { regBest['__row'+o.row] = {o:o, ts:ts}; return; }   // 無編號者各自獨立
    if (!regBest[s] || ts >= regBest[s].ts) regBest[s] = {o:o, ts:ts};
  });
  const register = Object.keys(regBest).map(function(k){ const o=regBest[k].o, r=o.r; return {
    row:o.row, serial:r[R('大量傷患編號')], name:r[R('姓名')], sex:r[R('性別')],
    tri:r[R('檢傷級數')], injury:r[R('傷情簡述')], cDone:r[R('完成連繫')],
    staff:r[R('登錄人員')], dispo:r[R('去向')], chartNo:r[R('病歷號')], event:r[R('事件')]||'',
    regDone:String(r[R('掛號建檔')]||''), hasId:!!String(r[R('身分證/護照')]||'').trim() }; }).slice(-300);
  return { ok:true, dispositions:DISPOSITIONS, events:Object.keys(evSet).sort(),
    pinned: !!PropertiesService.getScriptProperties().getProperty('CURRENT_EVENT'),
    currentEvent: PropertiesService.getScriptProperties().getProperty('CURRENT_EVENT') ||
      Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd'),
    triage:triage, register:register };
}

/* ---------- 人員名單 ---------- */
function getStaff_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(S_SHEET);
  if (!sh) {                                          // 初次自動建立示範名單
    sh = ss.insertSheet(S_SHEET);
    sh.appendRow(['工號','姓名','角色']);
    sh.appendRow(['N001','（請在「人員」分頁維護名單）','檢傷']);
    sh.getRange(1,1,1,3).setFontWeight('bold').setBackground('#DCDCD2');
    sh.setFrozenRows(1);
  }
  const last = sh.getLastRow();
  const rows = last>1 ? sh.getRange(2,1,last-1,3).getValues() : [];
  return { ok:true,
    pinned: !!PropertiesService.getScriptProperties().getProperty('CURRENT_EVENT'),
    currentEvent: PropertiesService.getScriptProperties().getProperty('CURRENT_EVENT') ||
      Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd'),
    staff: rows.filter(r=>r[1]).map(r=>({id:String(r[0]),name:String(r[1]),role:String(r[2]||'檢傷')})) };
}

/* ---------- 公開看板（唯讀） ---------- */
function boardPage_(ev) {
  const sh = sheet_(T_SHEET, T_HEADERS);
  const evI = T_HEADERS.indexOf('事件');
  const act = Date.now() - activeMs_();
  const MASK = PropertiesService.getScriptProperties().getProperty('BOARD_MASK') === '1';
  const rows = data_(sh, T_HEADERS.length).filter(function(r){ var e=String(r[evI]||'').trim(); return String(r[0]||'')!=='收件時間' && (!ev?true:e===ev) && rowTs_(r)>=act; });
  const iLv=T_HEADERS.indexOf('最終級數'), iNm=T_HEADERS.indexOf('患者'),
        iNo=T_HEADERS.indexOf('病歷號/流水號'), iAt=T_HEADERS.indexOf('屬性'),
        iTm=T_HEADERS.indexOf('抵達時間'), iDp=T_HEADERS.indexOf('去向'),
        iMc=T_HEADERS.indexOf('大量傷患編號');
  const cnt={1:0,2:0,3:0,4:0,5:0};
  rows.forEach(r=>{ if(cnt[r[iLv]]!==undefined) cnt[r[iLv]]++; });
  let html = head_('傷患清單', true) +
    '<h2>傷患清單'+(ev?'｜事件 '+ev:'')+'（共 '+rows.length+' 人，10秒自動更新）</h2>'+sumBar_(cnt)+
    '<table><tr><th>級</th><th>編號</th><th>抵達時間</th><th>患者</th><th>屬性</th><th>去向</th></tr>';
  rows.slice().sort((a,b)=>(a[iLv]||9)-(b[iLv]||9)).forEach(r=>{
    html += '<tr>'+lvTd_(r[iLv])+'<td><b>'+(r[iMc]||r[iNo]||'—')+'</b></td><td>'+fmtT_(r[iTm])+'</td><td>'+(MASK?mask_(r[iNm]):r[iNm])+
            '</td><td>'+r[iAt]+'</td><td>'+(r[iDp]||'—')+'</td></tr>';
  });
  return HtmlService.createHtmlOutput(html+'</table></body></html>').setTitle('傷患清單');
}

/* ---------- 行政管理面板（去向可編輯） ---------- */
function adminPage_(ev, adminkey) {
  const AK = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || '';
  if (AK && adminkey !== AK) {
    return HtmlService.createHtmlOutput('<meta charset="utf-8"><body style="font-family:sans-serif;padding:24px">'+
      '<h3>需要管理權限</h3><p>本頁含病人個資，僅限各科管理權限者。請於網址加上 <b>&adminkey=管理金鑰</b>，或由系統內行政畫面開啟。</p></body>');
  }
  const sh = sheet_(T_SHEET, T_HEADERS);
  const evI = T_HEADERS.indexOf('事件'), evIR = R_HEADERS.indexOf('事件');
  const evOk = function(v){ v=String(v||'').trim(); return !ev ? true : v===ev; };
  const rowsO = data_(sh, T_HEADERS.length).map(function(r,i){return {r:r,row:i+2};}).filter(function(o){return String(o.r[0]||'')!=='收件時間' && evOk(o.r[evI]);});
  const rows = rowsO.map(function(o){return o.r;});
  const rsh = sheet_(R_SHEET, R_HEADERS);
  const regsAll = data_(rsh, R_HEADERS.length).map(function(r,i){return {r:r,row:i+2};}).filter(function(o){return String(o.r[0]||'')!=='收件時間' && evOk(o.r[evIR]);});
  // 同一編號僅保留收件時間最新一筆
  const rSerIdx = R_HEADERS.indexOf('大量傷患編號');
  const regBestA = {};
  regsAll.forEach(function(o){ const s=String(o.r[rSerIdx]||'').trim(); const ts=rowTs_(o.r);
    if(!s){ regBestA['__r'+o.row]={o:o,ts:ts}; return; }
    if(!regBestA[s]||ts>=regBestA[s].ts) regBestA[s]={o:o,ts:ts}; });
  const regsO = Object.keys(regBestA).map(function(k){return regBestA[k].o;});
  const regs = regsO.map(function(o){return o.r;});
  // 編號→登錄姓名 對照（補檢傷清單患者欄）
  const nameByReg = {};
  regs.forEach(function(r){ const s=String(r[rSerIdx]||'').trim(); const nm=String(r[R_HEADERS.indexOf('姓名')]||'').trim(); if(s&&nm) nameByReg[s]=nm; });
  const iLv=T_HEADERS.indexOf('最終級數'), iNm=T_HEADERS.indexOf('患者'),
        iNo=T_HEADERS.indexOf('病歷號/流水號'), iTm=T_HEADERS.indexOf('抵達時間'),
        iDp=T_HEADERS.indexOf('去向'), iCr=T_HEADERS.indexOf('判斷依據'),
        iMc=T_HEADERS.indexOf('大量傷患編號');
  const cnt={1:0,2:0,3:0,4:0,5:0}, dcnt={};
  rows.forEach(r=>{ if(cnt[r[iLv]]!==undefined)cnt[r[iLv]]++;
                    const d=r[iDp]||'未定'; dcnt[d]=(dcnt[d]||0)+1; });
  const opt = v => DISPOSITIONS.map(o=>'<option'+(o===v?' selected':'')+'>'+o+'</option>').join('');
  let html = head_('行政管理面板', false) +
    '<h2>行政管理面板'+(ev?'｜事件 '+ev:'')+'</h2>'+sumBar_(cnt)+
    '<div class="sum2">去向統計：'+Object.keys(dcnt).map(k=>k+' '+dcnt[k]).join('｜')+'</div>'+
    '<h3>檢傷清單（去向可直接修改，會寫回試算表）</h3>'+
    '<table><tr><th>級</th><th>編號</th><th>抵達時間</th><th>病歷號</th><th>患者</th><th>判斷依據</th><th>去向</th></tr>';
  rows.forEach((r,idx)=>{
    html += '<tr>'+lvTd_(r[iLv])+'<td><b>'+(r[iMc]||'—')+'</b></td><td>'+fmtT_(r[iTm])+'</td><td>'+r[iNo]+'</td><td>'+(String(r[iNm]||'').trim()||nameByReg[String(r[iMc]||'').trim()]||'')+
      '</td><td>'+String(r[iCr]).substring(0,50)+'</td>'+
      '<td><select onchange="setD('+rowsO[idx].row+',this)">'+opt(r[iDp]||'')+'</select></td></tr>';
  });
  const R = h => R_HEADERS.indexOf(h);
  html += '</table><h3>大量傷患登錄清單（社工/掛號/輔助人員）</h3>'+
    '<table><tr><th>時間</th><th>編號</th><th>姓名</th><th>性別</th><th>生日</th><th>身分證/護照</th>'+
    '<th>手機</th><th>緊急聯絡人</th><th>聯繫</th><th>級</th><th>傷情簡述</th><th>登錄人員</th><th>去向</th></tr>';
  regs.forEach((r,idx)=>{
    html += '<tr><td>'+fmt_(r[0])+'</td><td>'+r[R('大量傷患編號')]+'</td><td>'+r[R('姓名')]+'</td><td>'+r[R('性別')]+
      '</td><td>'+fmtD_(r[R('生日')])+'</td><td>'+r[R('身分證/護照')]+'</td><td>'+r[R('手機')]+
      '</td><td>'+r[R('緊急聯絡人')]+(r[R('關係')]?'('+r[R('關係')]+')':'')+
      '</td><td>'+(r[R('完成連繫')]?'✔連繫':'')+(r[R('家屬已到')]?' ✔到院':'')+
      '</td><td>'+r[R('檢傷級數')]+'</td><td>'+String(r[R('傷情簡述')]).substring(0,30)+
      '</td><td>'+r[R('登錄人員')]+
      '</td><td><select onchange="setRD('+regsO[idx].row+',this)">'+opt(r[R('去向')]||'')+'</select></td></tr>';
  });
  html += '</table><p style="color:#666;font-size:12px">修改去向後儲存格會即時寫回；重新整理可取得最新清單。</p>'+
    '<script>'+
    'function setD(row,sel){sel.disabled=true;google.script.run.withSuccessHandler(function(){sel.disabled=false;sel.style.background="#C8E6C9";}).withFailureHandler(function(){sel.disabled=false;sel.style.background="#FFCDD2";}).setDisposition(row,sel.value);}'+
    'function setRD(row,sel){sel.disabled=true;google.script.run.withSuccessHandler(function(){sel.disabled=false;sel.style.background="#C8E6C9";}).withFailureHandler(function(){sel.disabled=false;sel.style.background="#FFCDD2";}).setRegDisposition(row,sel.value);}'+
    '</script></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('行政管理面板');
}
function setDisposition(row, value){
  sheet_(T_SHEET, T_HEADERS).getRange(row, T_HEADERS.indexOf('去向')+1).setValue(value);
  return true;
}
function setRegDisposition(row, value){
  sheet_(R_SHEET, R_HEADERS).getRange(row, R_HEADERS.indexOf('去向')+1).setValue(value);
  return true;
}

/* ---------- 共用 ---------- */
function sheet_(name, headers){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if(!sh) sh = ss.insertSheet(name);
  if(sh.getLastRow()===0){
    sh.appendRow(headers);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#DCDCD2');
    sh.setFrozenRows(1);
  } else {
    // 標題自動遷移：新版本新增的尾端欄位自動補標題（共享者無需手動維護）
    const cur = sh.getRange(1,1,1,headers.length).getValues()[0];
    let dirty = false;
    for (let i = 0; i < headers.length; i++) {
      if (String(cur[i]||'') !== headers[i]) { cur[i] = headers[i]; dirty = true; }
    }
    if (dirty) {
      sh.getRange(1,1,1,headers.length).setValues([cur])
        .setFontWeight('bold').setBackground('#DCDCD2');
    }
  }
  return sh;
}
function data_(sh, w){ const last=sh.getLastRow(); return last>1?sh.getRange(2,1,last-1,w).getValues():[]; }
function fmt_(d){ try{return Utilities.formatDate(new Date(d),'Asia/Taipei','MM/dd HH:mm');}catch(e){return d;} }
function fmtD_(v){ return (v && typeof v.getTime === 'function') ? Utilities.formatDate(v,'Asia/Taipei','yyyy/MM/dd') : v; }
function fmtT_(v){ return (v && typeof v.getTime === 'function') ? Utilities.formatDate(v,'Asia/Taipei','MM/dd HH:mm') : v; }
function head_(title, refresh){
  return '<!DOCTYPE html><html><head><meta charset="utf-8">'+
    (refresh?'<meta http-equiv="refresh" content="10">':'')+
    '<meta name="viewport" content="width=device-width,initial-scale=1"><title>'+title+'</title><style>'+
    'body{font-family:"Microsoft JhengHei",sans-serif;background:#EFEFE7;margin:10px}'+
    '.sum{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:6px}'+
    '.sum div{padding:8px 16px;font-size:20px;font-weight:bold;border:1px solid #888}'+
    '.sum2{margin-bottom:10px;font-weight:bold}'+
    'table{border-collapse:collapse;width:100%;background:#fff;font-size:14px;margin-bottom:14px}'+
    'th,td{border:1px solid #999;padding:4px 8px}th{background:#DCDCD2}'+
    '.lv{font-weight:bold;text-align:center;width:3em}select{font-size:13px}'+
    '</style></head><body>';
}
function sumBar_(cnt){
  const colors={1:'#D32F2F',2:'#F48FD8',3:'#FFF200',4:'#1E9E35',5:'#2036E8'};
  const fg={1:'#fff',2:'#7A0000',3:'#000',4:'#fff',5:'#fff'};
  return '<div class="sum">'+[1,2,3,4,5].map(l=>'<div style="background:'+colors[l]+';color:'+fg[l]+'">'+l+'級：'+cnt[l]+'</div>').join('')+'</div>';
}
function lvTd_(l){
  const colors={1:'#D32F2F',2:'#F48FD8',3:'#FFF200',4:'#1E9E35',5:'#2036E8'};
  const fg={1:'#fff',2:'#7A0000',3:'#000',4:'#fff',5:'#fff'};
  return '<td class="lv" style="background:'+(colors[l]||'#DDD')+';color:'+(fg[l]||'#333')+'">'+(l||'—')+'</td>';
}
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
