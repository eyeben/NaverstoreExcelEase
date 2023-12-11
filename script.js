document.getElementById('fileInput').addEventListener('change', handleFileSelect, false);

document.getElementById('downloadButton').addEventListener('click', function() {
    XLSX.writeFile(workbook, 'modified_file.xlsx');
});

let workbook;

function handleFileSelect(evt) {
    const file = evt.target.files[0];
    const reader = new FileReader();

    reader.onload = function(event) {
        const data = event.target.result;
        workbook = XLSX.read(data, {
            type: 'binary'
        });

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        swapRows(worksheet);

        document.getElementById('downloadButton').style.display = 'block';
    };

    reader.readAsBinaryString(file);
}

function swapRows(worksheet) {
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for(let C = range.s.c; C <= range.e.c; ++C) {
        const firstCell = XLSX.utils.encode_cell({c:C, r:0});
        const secondCell = XLSX.utils.encode_cell({c:C, r:1});

        let temp = worksheet[firstCell];
        worksheet[firstCell] = worksheet[secondCell];
        worksheet[secondCell] = temp;
    }
}