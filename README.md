<!doctype html>
<html lang="id">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Catatan Keuangan</title>
  <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js"></script>
</head>
<body class="bg-gray-100 min-h-screen">
  <div class="max-w-6xl mx-auto p-6">
    <h1 class="text-3xl font-bold text-center text-indigo-600 mb-8">📊 Catatan Keuangan Harian</h1>

    <!-- Form Input -->
    <div class="bg-white shadow-md rounded-2xl p-6 mb-8">
      <form id="entryForm" class="grid grid-cols-1 md:grid-cols-5 gap-4">
        <select id="type" class="border rounded-lg p-2">
          <option value="income">Pemasukan</option>
          <option value="expense">Pengeluaran</option>
        </select>
        <input id="date" type="date" class="border rounded-lg p-2" required />
        <input id="amount" type="number" placeholder="Jumlah" class="border rounded-lg p-2" required />
        <input id="note" placeholder="Nama Akun / Keterangan" class="border rounded-lg p-2" />
        <button type="submit" class="bg-indigo-500 hover:bg-indigo-600 text-white rounded-lg px-4 py-2">Tambah</button>
      </form>
      <div class="flex flex-wrap gap-2 mt-4">
        <button onclick="exportExcel()" class="bg-green-500 hover:bg-green-600 text-white rounded-lg px-4 py-2">Export Excel</button>
        <button onclick="exportJSON()" class="bg-yellow-500 hover:bg-yellow-600 text-white rounded-lg px-4 py-2">Export JSON</button>
        <input type="file" id="importFile" accept=".json" class="hidden" onchange="importJSON(event)" />
        <button onclick="document.getElementById('importFile').click()" class="bg-blue-500 hover:bg-blue-600 text-white rounded-lg px-4 py-2">Import JSON</button>
        <button onclick="resetAll()" class="bg-red-500 hover:bg-red-600 text-white rounded-lg px-4 py-2">Reset Semua</button>
      </div>
    </div>

    <!-- Ringkasan -->
    <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
      <div class="bg-white shadow rounded-2xl p-6 text-center">
        <h3 class="text-gray-500">Total Masuk</h3>
        <p id="totalIncome" class="text-2xl font-bold text-green-600">Rp 0</p>
      </div>
      <div class="bg-white shadow rounded-2xl p-6 text-center">
        <h3 class="text-gray-500">Total Keluar</h3>
        <p id="totalExpense" class="text-2xl font-bold text-red-600">Rp 0</p>
      </div>
      <div class="bg-white shadow rounded-2xl p-6 text-center">
        <h3 class="text-gray-500">Saldo</h3>
        <p id="balance" class="text-2xl font-bold text-indigo-600">Rp 0</p>
      </div>
    </div>

    <!-- Grafik -->
    <div class="bg-white shadow-md rounded-2xl p-6 mb-8">
      <canvas id="chartPie"></canvas>
    </div>

    <!-- Daftar Transaksi -->
    <div class="bg-white shadow-md rounded-2xl p-6 mb-8">
      <h3 class="text-xl font-semibold text-gray-700 mb-4">Daftar Transaksi</h3>
      <div class="overflow-x-auto">
        <table class="w-full text-sm">
          <thead class="bg-gray-100">
            <tr>
              <th class="px-3 py-2 text-left">Tanggal & Waktu</th>
              <th class="px-3 py-2 text-left">Jenis</th>
              <th class="px-3 py-2 text-left">Akun</th>
              <th class="px-3 py-2 text-right">Jumlah</th>
              <th class="px-3 py-2 text-center">Aksi</th>
            </tr>
          </thead>
          <tbody id="list"></tbody>
        </table>
      </div>
    </div>

    <!-- Rekapan -->
    <div class="bg-white shadow-md rounded-2xl p-6">
      <h3 class="text-xl font-semibold text-gray-700 mb-4">Rekapan Bulanan</h3>
      <div id="rekap"></div>
    </div>
  </div>

<script>
const STORAGE_KEY = 'finance_full_v10';
let entries = [];
let chart;

// utils
function load(){ entries = JSON.parse(localStorage.getItem(STORAGE_KEY))||[] }
function save(){ localStorage.setItem(STORAGE_KEY, JSON.stringify(entries)) }
function toRp(n){ return 'Rp ' + Number(n).toLocaleString() }

// tambah transaksi
const form = document.getElementById('entryForm');
form.addEventListener('submit',e=>{
  e.preventDefault();
  const type=document.getElementById('type').value;
  const date=document.getElementById('date').value;
  const amount=Number(document.getElementById('amount').value);
  const note=document.getElementById('note').value;
  if(!amount || !date){ alert('Isi tanggal dan jumlah'); return; }
  const now=new Date();
  const time=now.toTimeString().slice(0,5); // HH:mm
  const datetime=date+" "+time;
  entries.unshift({id:Date.now(),type,amount,note,dateTime:datetime});
  save();render();form.reset();
});

// hapus & reset
function deleteEntry(id){ entries=entries.filter(e=>e.id!==id); save();render(); }
function resetAll(){ if(confirm("Yakin hapus semua?")){ entries=[]; save();render(); }}

// export excel
function exportExcel(){
  const ws = XLSX.utils.json_to_sheet(entries);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Transaksi");
  XLSX.writeFile(wb, "keuangan.xlsx");
}

// export/import json
function exportJSON(){
  const dataStr=JSON.stringify(entries,null,2);
  const blob=new Blob([dataStr],{type:"application/json"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a");
  a.href=url;a.download="keuangan.json";a.click();
  URL.revokeObjectURL(url);
}
function importJSON(event){
  const file=event.target.files[0]; if(!file) return;
  const reader=new FileReader();
  reader.onload=function(e){
    try{
      const imported=JSON.parse(e.target.result);
      if(Array.isArray(imported)){
        entries=imported; save(); render();
        alert("✅ Data berhasil diimport!");
      } else alert("❌ Format salah!");
    }catch(err){ alert("❌ Gagal baca file!"); }
  };
  reader.readAsText(file);
}

// render
function render(){
  const listEl=document.getElementById('list');
  listEl.innerHTML='';
  let income=0,expense=0;
  entries.forEach(e=>{
    if(e.type==='income') income+=e.amount; else expense+=e.amount;
    const tr=document.createElement('tr');
    tr.className='border-b hover:bg-gray-50';
    tr.innerHTML=`<td class='px-3 py-2'>${e.dateTime}</td>
      <td class='px-3 py-2'>${e.type==='income'?'Pemasukan':'Pengeluaran'}</td>
      <td class='px-3 py-2'>${e.note||''}</td>
      <td class='px-3 py-2 text-right'>${toRp(e.amount)}</td>
      <td class='px-3 py-2 text-center'>
        <button class='text-red-500 hover:text-red-700' onclick='deleteEntry(${e.id})'>Hapus</button>
      </td>`;
    listEl.appendChild(tr);
  });
  document.getElementById('totalIncome').textContent=toRp(income);
  document.getElementById('totalExpense').textContent=toRp(expense);
  document.getElementById('balance').textContent=toRp(income-expense);

  // chart
  const ctx=document.getElementById('chartPie').getContext('2d');
  if(chart) chart.destroy();
  chart=new Chart(ctx,{type:'pie',
    data:{labels:['Masuk','Keluar'],
      datasets:[{data:[income,expense],backgroundColor:['#4ade80','#f87171']}]
    }
  });

  // rekap bulanan per akun
  const rekapEl=document.getElementById('rekap');
  rekapEl.innerHTML='';
  const grouped={};
  entries.forEach(e=>{
    const month=e.dateTime.slice(0,7); // YYYY-MM
    if(!grouped[month]) grouped[month]={};
    const akun=e.note||'Lainnya';
    if(!grouped[month][akun]) grouped[month][akun]={in:0,out:0};
    if(e.type==='income') grouped[month][akun].in+=e.amount;
    else grouped[month][akun].out+=e.amount;
  });
  Object.keys(grouped).sort().forEach(month=>{
    const div=document.createElement('div');
    div.className="mb-6";
    let totalMonth=0;
    let html=`<h4 class='font-semibold text-indigo-600 mb-2'>${month}</h4>
      <table class='w-full text-sm mb-2'>
      <thead><tr class='bg-gray-100'>
        <th class='px-2 py-1 text-left'>Akun</th>
        <th class='px-2 py-1 text-right'>Masuk</th>
        <th class='px-2 py-1 text-right'>Keluar</th>
        <th class='px-2 py-1 text-right'>Total Akun</th>
      </tr></thead><tbody>`;
    Object.keys(grouped[month]).forEach(akun=>{
      const data=grouped[month][akun];
      const total=data.in-data.out; totalMonth+=total;
      html+=`<tr class='border-b'>
        <td class='px-2 py-1'>${akun}</td>
        <td class='px-2 py-1 text-right'>${toRp(data.in)}</td>
        <td class='px-2 py-1 text-right'>${toRp(data.out)}</td>
        <td class='px-2 py-1 text-right font-medium'>${toRp(total)}</td>
      </tr>`;
    });
    html+=`<tr class='bg-gray-200 font-semibold'>
      <td class='px-2 py-1 text-right' colspan='3'>Total Akhir ${month}</td>
      <td class='px-2 py-1 text-right'>${toRp(totalMonth)}</td>
    </tr></tbody></table>`;
    div.innerHTML=html; rekapEl.appendChild(div);
  });
}

window.deleteEntry=deleteEntry;
window.exportExcel=exportExcel;
window.exportJSON=exportJSON;
window.importJSON=importJSON;
window.resetAll=resetAll;

load();render();
</script>
</body>
</html>
