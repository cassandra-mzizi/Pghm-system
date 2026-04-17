// Update to swap Total Born and Born Alive display positions
function updateLitterForm(rowData) {
    // D17 was Born Alive, now it reads from rowData[9]
    document.getElementById('D17').innerText = rowData[9];
    // G7 was Total Born, now it reads from rowData[5]
    document.getElementById('G7').innerText = rowData[5];
}

function saveLitterData(data) {
    // Adjusted to read from swapped positions
    const totalBorn = data[5]; // Now totalBorn is read from rowData[5]
    const bornAlive = data[9]; // Now bornAlive is read from rowData[9]
    // Save logic here...
}

function modifyLitter(data) {
    // Adjusted to read from swapped positions
    const totalBorn = data[5]; // Now totalBorn is read from rowData[5]
    const bornAlive = data[9]; // Now bornAlive is read from rowData[9]
    // Modify logic here...
}