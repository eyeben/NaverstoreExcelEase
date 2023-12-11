let excel1, excel2;

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

    const data1 = XLSX.utils.sheet_to_json(sheet1, { defval: "" });
    const data2 = XLSX.utils.sheet_to_json(sheet2);

    data1.forEach(row1 => {
        const matchingRow = data2.find(row2 => row2['고객주문번호'] === row1['주문번호']);
        console.log('Matching Row:', matchingRow); // 디버깅을 위한 로그
        if (matchingRow) {
            row1['송장번호'] = matchingRow['송장번호'];
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
