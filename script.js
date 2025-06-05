let data = [];

document.getElementById('fileInput').addEventListener('change', function(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function(evt) {
    const binary = evt.target.result;
    const workbook = XLSX.read(binary, { type: 'binary' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    data = XLSX.utils.sheet_to_json(sheet);
    alert(`โหลดข้อมูล ${data.length} แถวสำเร็จ`);
  };

  reader.readAsBinaryString(file);
});

function handlePrint() {
  if (data.length === 0) {
    alert("กรุณาอัปโหลดไฟล์ก่อน");
    return;
  }

  // ตัวอย่าง: สร้างหน้าต่างพิมพ์
  let printWindow = window.open('', '', 'height=600,width=800');
  let html = '<html><head><title>พิมพ์สติกเกอร์</title></head><body>';
  
  data.forEach((row, i) => {
    html += `<div style="border:1px solid black; padding:10px; margin:10px;">
               <b>ชื่อ:</b> ${row.Fullname || '-'}<br>
               <b>HN:</b> ${row.HN || '-'}<br>
             </div>`;
    if ((i+1) % 4 === 0) html += '<hr>'; // แบ่งกลุ่มละ 4
  });

  html += '</body></html>';
  printWindow.document.write(html);
  printWindow.document.close();
  printWindow.print();
}
