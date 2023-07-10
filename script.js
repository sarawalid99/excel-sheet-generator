let table = document.getElementsByClassName("sheet-body")[0],
rows = document.getElementsByClassName("rows")[0],
columns = document.getElementsByClassName("columns")[0]
tableExists = false

const generateTable = () => {
    let rowsNumber = parseInt(rows.value);
    let columnsNumber = parseInt(columns.value);

    if (rowsNumber > 0 && columnsNumber > 0) {
        table.innerHTML = "";

        for (let i = 0; i < rowsNumber; i++) {
            let tableRow = "";
            for (let j = 0; j < columnsNumber; j++) {
                tableRow += `<td contenteditable></td>`;
            }
            table.innerHTML += tableRow;
        }

        tableExists = true;
    } else {
        // Display an alert using SweetAlert.js when the fields are empty
        Swal.fire({
            icon: 'error',
            title: 'Empty Fields',
            text: 'Please enter a valid number of rows and columns.',
        });
    }
}

const ExportToExcel = (type, fn, dl) => {
    if (!tableExists) {
        // Display an alert using SweetAlert.js when there is no generated table
        Swal.fire({
            icon: 'error',
            title: 'No Table to Export',
            text: 'Please generate a table before exporting.',
        });
        return;
    }

    var elt = table
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" })
    return dl ? XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' })
        : XLSX.writeFile(wb, fn || ('MyNewSheet.' + (type || 'xlsx')))
}