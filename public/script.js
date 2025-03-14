document.addEventListener("DOMContentLoaded", async () => {
    const response = await fetch("/data");
    const data = await response.json();

    const container = document.getElementById("hot");
    window.hot = new Handsontable(container, {
        data: data,
        colHeaders: true,
        rowHeaders: true,
        minSpareRows: 1,
        contextMenu: true,
        licenseKey: "non-commercial-and-evaluation",
    });
});

async function saveData() {
    const data = hot.getData();
    const response = await fetch("/save", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data),
    });

    if (response.ok) alert("Data saved!");
}

async function exportPDF() {
    window.open("/export-pdf", "_blank");
}
