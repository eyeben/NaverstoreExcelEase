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

        // 첫 번째 시트의 첫 행을 제거
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        removeFirstRow(worksheet);

        callback(workbook);
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

    // 데이터 처리 로직...
    // 예: 전화번호를 기반으로 매칭하고 송장번호 업데이트

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
