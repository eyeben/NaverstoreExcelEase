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

    const data1 = XLSX.utils.sheet_to_json(sheet1, { defval: "", range: 1 });
    const data2 = XLSX.utils.sheet_to_json(sheet2, { defval: "" });

    const updatedData = data1.map(row1 => {
        const formattedPhone = String(row1['수취인연락처1']).replace(/-/g, '').trim();
        const matchingRow = data2.find(row2 => String(row2['받는분전화번호']).trim() === formattedPhone);

        // 필요한 컬럼만 선택하여 새 객체 생성
        return {
            '상품주문번호': row1['상품주문번호'],
            '배송방법': row1['배송방법'],
            '택배사': row1['택배사'] || 'CJ대한통운', // '택배사'가 없는 경우 기본값 사용
            '송장번호': matchingRow ? matchingRow['운송장번호'] : (row1['송장번호'] || "")
        };
    });

    // 새로운 시트 생성
    const updatedSheet = XLSX.utils.json_to_sheet(updatedData);

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
