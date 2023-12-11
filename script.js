document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('excel1Input').addEventListener('change', function(evt) {
        readFile(evt.target.files[0], function(workbook) {
            excel1 = workbook;
            updateProcessButtonVisibility();
        });
    }, false);

    document.getElementById('excel2Input').addEventListener('change', function(evt) {
        readFile(evt.target.files[0], function(workbook) {
            excel2 = workbook;
            updateProcessButtonVisibility();
        });
    }, false);

    document.getElementById('processButton').addEventListener('click', processFiles);
});

function readFile(file, callback) {
    const reader = new FileReader();
    reader.onload = function(event) {
        const data = event.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        callback(workbook);
    };
    reader.readAsBinaryString(file);
}

function processFiles() {
    if (!excel1 || !excel2) {
        alert('두 개의 파일을 모두 업로드해주세요.');
        return;
    }

    const sheet1 = excel1.Sheets[excel1.SheetNames[0]];
    const sheet2 = excel2.Sheets[excel2.SheetNames[0]];

    // 첫 행을 제외하고 데이터를 읽음
    const data1 = XLSX.utils.sheet_to_json(sheet1, { defval: "", range: 1 });
    const data2 = XLSX.utils.sheet_to_json(sheet2, { defval: ""});

    data1.forEach(row1 => {
        const formattedPhone = String(row1['수취인연락처1']).replace(/-/g, '').trim();
        const matchingRow = data2.find(row2 => String(row2['받는분전화번호']).trim() === formattedPhone);
        if (matchingRow) {
            row1['송장번호'] = matchingRow['운송장번호'];
            console.log(`매칭 성공: ${formattedPhone}, 송장번호: ${matchingRow['운송장번호']}`);
        } else {
            console.log(`매칭 실패: ${formattedPhone} `);
            console.log(data1)
            console.log(data2)
            row1['송장번호'] = row1['송장번호'] || "";
        }
    });


    // 새로운 시트 생성 시 첫 행을 포함하지 않음
    const updatedSheet = XLSX.utils.json_to_sheet(data1);

    // 새 워크북 생성 및 시트 추가
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, updatedSheet, "Updated Data");

    // 새 워크북 저장
    XLSX.writeFile(newWorkbook, 'updated_excel1.xlsx');
}

function updateProcessButtonVisibility() {
    if (excel1 && excel2) {
        document.getElementById('processButton').style.display = 'block';
    } else {
        document.getElementById('processButton').style.display = 'none';
    }
}
