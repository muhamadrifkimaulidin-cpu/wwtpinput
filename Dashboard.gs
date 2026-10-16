/**
 * ============================================
 * Dashboard.gs — Data Aggregation v3.0
 * 4-group Presentasi + Softener + PAM Monthly
 * ============================================
 */
function getDashboardData(options) {
  try {
    var sOpt = options && options.startDate ? options.startDate : 'auto';
    var eOpt = options && options.endDate ? options.endDate : 'auto';
    var cacheKey = 'dash_' + sOpt + '_' + eOpt;
    var cached = cacheGet_(cacheKey);
    if (cached) return cached;

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataOps');
    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    var allData = sheet.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    if (allData.length <= 1) return { success:true, data:getEmptyDashboard_() };

    var startDate = options && options.startDate ? new Date(options.startDate) : null;
    var endDate = options && options.endDate ? new Date(options.endDate) : null;
    if (!startDate) { startDate = new Date(); startDate.setDate(startDate.getDate()-30); }
    if (!endDate) { endDate = new Date(); }
    startDate.setHours(0,0,0,0); endDate.setHours(23,59,59,999);

    var colIdx = {};
    for (var h=0;h<headers.length;h++) colIdx[headers[h]]=h;

    var allRecords = [];
    for (var i=1;i<allData.length;i++) {
      var row=allData[i], tanggal=row[colIdx['tanggal']], ts, tsh;
      if (tanggal instanceof Date) { ts=Utilities.formatDate(tanggal,tz,'yyyy-MM-dd'); tsh=Utilities.formatDate(tanggal,tz,'dd/MM'); }
      else if (typeof tanggal==='string' && tanggal.match(/^\d{4}-\d{2}-\d{2}/)) { ts=tanggal.substring(0,10); tsh=ts.substring(8,10)+'/'+ts.substring(5,7); }
      else if (typeof tanggal==='string' && tanggal.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
        var p=tanggal.split('/'); ts=p[2]+'-'+('0'+p[1]).slice(-2)+'-'+('0'+p[0]).slice(-2); tsh=('0'+p[0]).slice(-2)+'/'+('0'+p[1]).slice(-2);
      } else continue;
      var obj={}; for(var j=0;j<headers.length;j++) obj[headers[j]]=row[j];
      obj.tanggal_str=ts; obj.tanggal_short=tsh; allRecords.push(obj);
    }
    var records=allRecords.filter(function(r){var d=new Date(r.tanggal_str);return d>=startDate&&d<=endDate;});

    var totalRecords=records.length, todayStr=Utilities.formatDate(new Date(),tz,'yyyy-MM-dd');
    var todayRecords=records.filter(function(r){return r.tanggal_str===todayStr;}).length;
    var totalHydrofloc=0,totalPolymer=0;
    records.forEach(function(r){totalHydrofloc+=parseFloat(r.chem_hydrofloc_kg)||0;totalPolymer+=parseFloat(r.chem_polymer_g)||0;});

    var allDailyClr={};
    allRecords.forEach(function(r){
      var d=r.tanggal_str; if(!allDailyClr[d]) allDailyClr[d]={bak:0,sungai:0,mg:0};
      var bB=parseFloat(r.clr_bak_sirkulasi_before)||0,bA=parseFloat(r.clr_bak_sirkulasi_after)||0;
      var sB=parseFloat(r.clr_sungai_before)||0,sA=parseFloat(r.clr_sungai_after)||0;
      var mB=parseFloat(r.clr_mg17_18_before)||0,mA=parseFloat(r.clr_mg17_18_after)||0;
      if(bB>0&&bA>0&&bA>=bB) allDailyClr[d].bak+=(bA-bB);
      if(sB>0&&sA>0&&sA>=sB) allDailyClr[d].sungai+=(sA-sB);
      if(mB>0&&mA>0&&mA>=mB) allDailyClr[d].mg+=(mA-mB);
    });
    Object.keys(allDailyClr).forEach(function(d){var c=allDailyClr[d];c.bak=rd2_(c.bak);c.sungai=rd2_(c.sungai);c.mg=rd2_(c.mg);c.total=rd2_(c.bak+c.sungai+c.mg);});

    var allDailyFm={};
    allRecords.forEach(function(r){
      var d=r.tanggal_str; if(!allDailyFm[d]) allDailyFm[d]={sm1:0,sm2:0,sm3:0,sm4:0,sm5:0,sm6:0,pam:0,softener:0};
      var dm=allDailyFm[d];
      dm.sm1=maxVal_(dm.sm1,parseFloat(r.fm_sm1)||0); dm.sm2=maxVal_(dm.sm2,parseFloat(r.fm_sm2)||0);
      dm.sm3=maxVal_(dm.sm3,parseFloat(r.fm_sm3)||0); dm.sm4=maxVal_(dm.sm4,parseFloat(r.fm_sm4)||0);
      dm.sm5=maxVal_(dm.sm5,parseFloat(r.fm_sm5)||0); dm.sm6=maxVal_(dm.sm6,parseFloat(r.fm_sm6)||0);
      dm.pam=maxVal_(dm.pam,parseFloat(r.fm_pam)||0); dm.softener=maxVal_(dm.softener,parseFloat(r.fm_softener_lt3)||0);
    });
    var allFmDates=Object.keys(allDailyFm).sort(), allFmDiffs={};
    for(var i=1;i<allFmDates.length;i++){
      var dT=allFmDates[i],dY=allFmDates[i-1],t=allDailyFm[dT],y=allDailyFm[dY];
      allFmDiffs[dT]={sm1:rd2_(diffPos_(t.sm1,y.sm1)),sm2:rd2_(diffPos_(t.sm2,y.sm2)),sm3:rd2_(diffPos_(t.sm3,y.sm3)),
        sm4:rd2_(diffPos_(t.sm4,y.sm4)),sm5:rd2_(diffPos_(t.sm5,y.sm5)),sm6:rd2_(diffPos_(t.sm6,y.sm6)),
        pam:rd2_(diffPos_(t.pam,y.pam)),softener:rd2_(diffPos_(t.softener,y.softener))};
    }

    var allDatesInRange=[], cur=new Date(startDate);
    while(cur<=endDate){allDatesInRange.push(Utilities.formatDate(cur,tz,'yyyy-MM-dd'));cur.setDate(cur.getDate()+1);}

    var clarifierDaily=[],wellDaily=[],softenerDaily=[],pamDaily=[];
    allDatesInRange.forEach(function(d){
      var c=allDailyClr[d]||{bak:0,sungai:0,mg:0,total:0};
      clarifierDaily.push({date:d.substring(8,10)+'/'+d.substring(5,7),date_full:d,bak_sirkulasi:c.bak||0,sungai:c.sungai||0,mg17_18:c.mg||0,total:c.total||0});
      var f=allFmDiffs[d]||{sm1:0,sm2:0,sm3:0,sm4:0,sm5:0,sm6:0,pam:0,softener:0};
      wellDaily.push({date:d.substring(8,10)+'/'+d.substring(5,7),date_full:d,sm1:f.sm1||0,sm2:f.sm2||0,sm3:f.sm3||0,sm4:f.sm4||0,sm5:f.sm5||0,sm6:f.sm6||0});
      softenerDaily.push({date:d.substring(8,10)+'/'+d.substring(5,7),date_full:d,usage:f.softener||0});
      pamDaily.push({date:d.substring(8,10)+'/'+d.substring(5,7),date_full:d,usage:f.pam||0});
    });

    var clrT={total:[],bak:[],sungai:[],mg:[]};
    clarifierDaily.forEach(function(d){clrT.total.push(d.total);clrT.bak.push(d.bak_sirkulasi);clrT.sungai.push(d.sungai);clrT.mg.push(d.mg17_18);});
    var clarifierAvg={avg_total:rd2_(avg_(clrT.total)),avg_bak:rd2_(avg_(clrT.bak)),avg_sungai:rd2_(avg_(clrT.sungai)),avg_mg:rd2_(avg_(clrT.mg))};

    var wK=['sm1','sm2','sm3','sm4','sm5','sm6'],wS={sm1:0,sm2:0,sm3:0,sm4:0,sm5:0,sm6:0},wC={sm1:0,sm2:0,sm3:0,sm4:0,sm5:0,sm6:0};
    allDatesInRange.forEach(function(d){var diff=allFmDiffs[d];if(!diff)return;wK.forEach(function(k){if(diff[k]>0){wS[k]+=diff[k];wC[k]++;}});});
    var wellAvgUsage=wK.map(function(k){return wC[k]>0?rd2_(wS[k]/wC[k]):0;});

    var pamShiftByMonth={};
    allRecords.forEach(function(r){
      var mk=r.tanggal_str.substring(0,7);
      var v=parseFloat(r.fm_pam_shift)||0;
      if(v>0){
        if(!pamShiftByMonth[mk]) pamShiftByMonth[mk]={min:v,max:v};
        else{ if(v<pamShiftByMonth[mk].min) pamShiftByMonth[mk].min=v; if(v>pamShiftByMonth[mk].max) pamShiftByMonth[mk].max=v; }
      }
    });
    function getPamShiftDiff(month){ var m=pamShiftByMonth[month]; return m ? rd2_(m.max-m.min) : 0; }
    var pamMonthly=[
      {month:'Sep 25',usage:1332},{month:'Okt 25',usage:3250},
      {month:'Nov 25',usage:1461},{month:'Des 25',usage:838},
      {month:'Jan 26',usage:getPamShiftDiff('2026-01')},
      {month:'Feb 26',usage:getPamShiftDiff('2026-02')},
      {month:'Mar 26',usage:getPamShiftDiff('2026-03')}
    ];

    var chemMap={},dms={};
    records.forEach(function(r){var k=r.tanggal_short;dms[k]=r.tanggal_str;if(!chemMap[k])chemMap[k]={h:[],p:[]};chemMap[k].h.push(parseFloat(r.chem_hydrofloc_kg)||0);chemMap[k].p.push(parseFloat(r.chem_polymer_g)||0);});
    var ss=Object.keys(dms).sort(function(a,b){return dms[a].localeCompare(dms[b]);});
    var chemicalTrend=[];ss.forEach(function(d){if(chemMap[d])chemicalTrend.push({date:d,hydrofloc:sum_(chemMap[d].h),polymer:sum_(chemMap[d].p)});});
    var chkP={mg17_18:0,clarifier:0,fe_removal:0,bak_sirkulasi_besar:0};
    var labS={clarifier:0,mg17_18:0,bak_sirkulasi:0,demin:0,pam:0,sm1:0,sm2:0,sm3:0,sm4:0,sm5:0,sm6:0};
    records.forEach(function(r){
      if(r.chk_mg17_18===true||r.chk_mg17_18==='true')chkP.mg17_18++;if(r.chk_clarifier===true||r.chk_clarifier==='true')chkP.clarifier++;
      if(r.chk_fe_removal===true||r.chk_fe_removal==='true')chkP.fe_removal++;if(r.chk_bak_sirkulasi_besar===true||r.chk_bak_sirkulasi_besar==='true')chkP.bak_sirkulasi_besar++;
      if(r.lab_clarifier===true||r.lab_clarifier==='true')labS.clarifier++;if(r.lab_mg17_18===true||r.lab_mg17_18==='true')labS.mg17_18++;
      if(r.lab_bak_sirkulasi===true||r.lab_bak_sirkulasi==='true')labS.bak_sirkulasi++;if(r.lab_demin===true||r.lab_demin==='true')labS.demin++;
      if(r.lab_pam===true||r.lab_pam==='true')labS.pam++;for(var w=1;w<=6;w++){if(r['lab_sm'+w]===true||r['lab_sm'+w]==='true')labS['sm'+w]++;}
    });
    var tot=totalRecords||1;
    var checklist={process:{labels:['MG 17-18','Clarifier','Fe Removal','Bak Sirk. Besar'],values:[Math.round(chkP.mg17_18/tot*100),Math.round(chkP.clarifier/tot*100),Math.round(chkP.fe_removal/tot*100),Math.round(chkP.bak_sirkulasi_besar/tot*100)]},
      lab:{labels:['Clarifier','MG 17-18','Bak Sirk.','Demin','PAM','Sm1','Sm2','Sm3','Sm4','Sm5','Sm6'],values:[Math.round(labS.clarifier/tot*100),Math.round(labS.mg17_18/tot*100),Math.round(labS.bak_sirkulasi/tot*100),Math.round(labS.demin/tot*100),Math.round(labS.pam/tot*100),Math.round(labS.sm1/tot*100),Math.round(labS.sm2/tot*100),Math.round(labS.sm3/tot*100),Math.round(labS.sm4/tot*100),Math.round(labS.sm5/tot*100),Math.round(labS.sm6/tot*100)]}};
    var shiftMap={};records.forEach(function(r){var k=r.tanggal_short;if(!shiftMap[k])shiftMap[k]={s1:0,s2:0,s3:0};var s=parseInt(r.shift);if(s===1)shiftMap[k].s1++;else if(s===2)shiftMap[k].s2++;else if(s===3)shiftMap[k].s3++;});
    var shiftActivity=[];ss.forEach(function(d){if(shiftMap[d])shiftActivity.push({date:d,shift1:shiftMap[d].s1,shift2:shiftMap[d].s2,shift3:shiftMap[d].s3});});

    var result = {success:true,data:{
      summary:{totalRecords:totalRecords,todayRecords:todayRecords,totalHydrofloc:rd2_(totalHydrofloc),totalPolymer:rd2_(totalPolymer)},
      presentasi:{clarifierDaily:clarifierDaily,clarifierAvg:clarifierAvg,wellLabels:['Sm 1','Sm 2','Sm 3','Sm 4','Sm 5','Sm 6'],wellMax:[60,75,75,72,72,18],
        wellAvgUsage:wellAvgUsage,wellDaily:wellDaily,softenerDaily:softenerDaily,pamDaily:pamDaily,pamMonthly:pamMonthly},
      lainnya:{chemicalTrend:chemicalTrend,checklist:checklist,shiftActivity:shiftActivity}
    }};
    cachePut_(cacheKey, result, CACHE_TTL_);
    cacheTrackKey_('dash', cacheKey);
    return result;
  } catch(e){return{success:false,message:'Dashboard error: '+e.message,data:getEmptyDashboard_()};}
}
function getEmptyDashboard_(){return{summary:{totalRecords:0,todayRecords:0,totalHydrofloc:0,totalPolymer:0},presentasi:{clarifierDaily:[],clarifierAvg:{avg_total:0,avg_bak:0,avg_sungai:0,avg_mg:0},wellLabels:[],wellMax:[],wellAvgUsage:[],wellDaily:[],softenerDaily:[],pamDaily:[],pamMonthly:[]},lainnya:{chemicalTrend:[],checklist:{process:{labels:[],values:[]},lab:{labels:[],values:[]}},shiftActivity:[]}};}
function maxVal_(a,b){return b>a?b:a;}
function diffPos_(t,y){if(t<=0||y<=0)return 0;return Math.max(0,t-y);}
function rd2_(v){return Math.round(v*100)/100;}
function avg_(a){if(!a||a.length===0)return 0;var s=0;a.forEach(function(v){s+=v;});return s/a.length;}
function sum_(a){var s=0;a.forEach(function(v){s+=v;});return rd2_(s);}
