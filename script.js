let excel1, excel2;

document.getElementById('excel1Input').addEventListener('change', function(evt) {
    readFile(evt.target.files[0], function(workbook) {
        excel1 = workbook;
    });
}, false);

document.getElementById('excel2Input').addEventListener('change', function(evt) {
    readFile(evt.target.files[0], function(workbook) {
        excel2 = workbook;
    });
}, false);

document.getElementById('processButton').addEventListener('click', processFiles);

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


    XLSX.writeFile(excel1, 'updated_excel1.xlsx');
}
