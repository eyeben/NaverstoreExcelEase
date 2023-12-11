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
        try {
            const workbook = XLSX.read(data, { 
                type: 'binary',
                password: 'aaaa' // 비밀번호 입력
            });

            // 첫 번째 시트의 첫 행을 제거
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            removeFirstRow(worksheet);

            callback(workbook);
        } catch (e) {
            console.error(e.message);
            alert('파일을 열 수 없습니다. 비밀번호를 확인하세요.');
        }
    };
    reader.readAsBinaryString(file);
}

function removeFirstRow(worksheet) {
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let R = range.s.r + 1; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({r: R, c: C});
            const cellAboveAddress = XLSX.utils.encode_cell({r: R - 1, c: C});
            worksheet[cellAboveAddress] = worksheet[cellAddress];
        }
    }
    range.e.r--;
    worksheet['!ref'] = XLSX.utils.encode_range(range);
}

function processFiles() {
    if (!excel1 || !excel2) {
        alert('두 개의 파일을 모두 업로드해주세요.');
        return;
    }

    const sheet1 = excel1.Sheets[excel1.SheetNames[0]];
    const sheet2 = excel2.Sheets[excel2.SheetNames[0]];

    const data1 = XLSX.utils.sheet_to_json(sheet1, { defval: "" });
    const data2 = XLSX.utils.sheet_to_json(sheet2);

    data1.forEach(row1 => {
        const formattedPhone = String(row1['수취인연락처1']).replace(/-/g, '').trim();
        const matchingRow = data2.find(row2 => String(row2['받는분전화번호']).trim() === formattedPhone);
        if (matchingRow) {
            row1['송장번호'] = matchingRow['운송장번호'];
        } else {
            row1['송장번호'] = row1['송장번호'] || "";
        }
    });

    const originalHeaders = Object.keys(XLSX.utils.sheet_to_json(sheet1, { header: 1 })[0]);
    const updatedSheet = XLSX.utils.json_to_sheet(data1, {
        header: originalHeaders,
        skipHeader: true
    });

    XLSX.writeFile(excel1, 'updated_excel1.xlsx');
}

function updateProcessButtonVisibility() {
    if (excel1 && excel2) {
        document.getElementById('processButton').style.display = 'block';
    } else {
        document.getElementById('processButton').style.display = 'none';
    }
}
