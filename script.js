const fileInput = document.getElementById("fileInput");
const fileName = document.getElementById("fileName");
const resetBtn = document.getElementById("resetBtn");

const totalSalesElement = document.getElementById("totalSales");
const totalOrdersElement = document.getElementById("totalOrders");
const averageSalesElement = document.getElementById("averageSales");
const topProductElement = document.getElementById("topProduct");

const tableHead = document.getElementById("tableHead");
const tableBody = document.getElementById("tableBody");
const rowCount = document.getElementById("rowCount");

let productChart = null;
let regionChart = null;


// =========================
// FILE UPLOAD
// =========================

fileInput.addEventListener("change", function (event) {
    const file = event.target.files[0];

    if (!file) {
        return;
    }

    fileName.textContent = "Selected file: " + file.name;

    readExcelFile(file);
});


// =========================
// READ EXCEL / CSV
// =========================

function readExcelFile(file) {

    const reader = new FileReader();

    reader.onload = function (event) {

        try {

            const data = new Uint8Array(event.target.result);

            const workbook = XLSX.read(data, {
                type: "array"
            });

            if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                alert("Excel file me koi sheet nahi mili.");
                return;
            }

            const sheetName = workbook.SheetNames[0];

            const worksheet = workbook.Sheets[sheetName];

            const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                defval: ""
            });

            if (jsonData.length === 0) {
                alert("Excel file empty hai.");
                return;
            }

            processData(jsonData);

        } catch (error) {

            console.error("Excel Error:", error);

            alert(
                "File read nahi ho saki. Please valid Excel ya CSV file upload karein."
            );
        }
    };

    reader.onerror = function () {
        alert("File read karne me problem aa gayi.");
    };

    reader.readAsArrayBuffer(file);
}


// =========================
// PROCESS DATA
// =========================

function processData(data) {

    displayTable(data);

    updateDashboard(data);
}


// =========================
// FIND COLUMN
// =========================

function findColumn(data, possibleNames) {

    if (!data || data.length === 0) {
        return null;
    }

    const columns = Object.keys(data[0]);

    // Exact match
    for (const column of columns) {

        const normalizedColumn = String(column)
            .toLowerCase()
            .trim();

        for (const name of possibleNames) {

            if (normalizedColumn === name) {
                return column;
            }
        }
    }

    // Partial match
    for (const column of columns) {

        const normalizedColumn = String(column)
            .toLowerCase()
            .trim();

        for (const name of possibleNames) {

            if (normalizedColumn.includes(name)) {
                return column;
            }
        }
    }

    return null;
}


// =========================
// NUMBER CONVERTER
// =========================

function getNumber(value) {

    if (value === null || value === undefined || value === "") {
        return 0;
    }

    const number = parseFloat(
        String(value)
            .replace(/,/g, "")
            .replace(/[₹$€£]/g, "")
            .replace(/\s/g, "")
    );

    return isNaN(number) ? 0 : number;
}


// =========================
// UPDATE DASHBOARD
// =========================

function updateDashboard(data) {

    const salesColumn = findColumn(data, [
        "sales",
        "sale",
        "revenue",
        "amount",
        "price",
        "total",
        "income"
    ]);

    const productColumn = findColumn(data, [
        "product",
        "product name",
        "item",
        "item name"
    ]);

    const regionColumn = findColumn(data, [
        "region",
        "area",
        "location",
        "zone",
        "state",
        "city"
    ]);


    // =========================
    // TOTAL SALES
    // =========================

    let totalSales = 0;

    if (salesColumn) {

        data.forEach(function (row) {

            totalSales += getNumber(row[salesColumn]);

        });

    }


    // =========================
    // TOTAL ORDERS
    // =========================

    const totalOrders = data.length;


    // =========================
    // AVERAGE SALES
    // =========================

    const averageSales =
        totalOrders > 0
            ? totalSales / totalOrders
            : 0;


    // =========================
    // TOP PRODUCT
    // =========================

    let topProduct = "-";

    if (productColumn && salesColumn) {

        const productSales = {};

        data.forEach(function (row) {

            const product = String(
                row[productColumn] || ""
            ).trim();

            const sales = getNumber(
                row[salesColumn]
            );

            if (product !== "") {

                if (!productSales[product]) {
                    productSales[product] = 0;
                }

                productSales[product] += sales;
            }

        });


        let highestSales = -Infinity;

        Object.keys(productSales).forEach(function (product) {

            if (productSales[product] > highestSales) {

                highestSales = productSales[product];

                topProduct = product;
            }

        });

    }


    // =========================
    // UPDATE CARDS
    // =========================

    totalSalesElement.textContent =
        formatCurrency(totalSales);

    totalOrdersElement.textContent =
        totalOrders.toLocaleString("en-IN");

    averageSalesElement.textContent =
        formatCurrency(averageSales);

    topProductElement.textContent =
        topProduct;


    // =========================
    // PRODUCT CHART
    // =========================

    if (productColumn && salesColumn) {

        createProductChart(
            data,
            productColumn,
            salesColumn
        );

    } else {

        destroyProductChart();
    }


    // =========================
    // REGION CHART
    // =========================

    if (regionColumn && salesColumn) {

        createRegionChart(
            data,
            regionColumn,
            salesColumn
        );

    } else {

        destroyRegionChart();
    }

}


// =========================
// PRODUCT CHART
// =========================

function createProductChart(
    data,
    productColumn,
    salesColumn
) {

    const productSales = {};

    data.forEach(function (row) {

        const product = String(
            row[productColumn] || ""
        ).trim();

        const sales = getNumber(
            row[salesColumn]
        );

        if (product !== "") {

            if (!productSales[product]) {
                productSales[product] = 0;
            }

            productSales[product] += sales;
        }

    });


    const sortedProducts = Object.entries(productSales)
        .sort(function (a, b) {
            return b[1] - a[1];
        })
        .slice(0, 10);


    const labels = sortedProducts.map(function (item) {
        return item[0];
    });

    const values = sortedProducts.map(function (item) {
        return item[1];
    });


    const canvas = document.getElementById("productChart");

    if (!canvas) {
        return;
    }


    destroyProductChart();


    productChart = new Chart(canvas, {

        type: "bar",

        data: {

            labels: labels,

            datasets: [
                {
                    label: "Sales",

                    data: values,

                    borderWidth: 0,

                    borderRadius: 6
                }
            ]
        },

        options: {

            responsive: true,

            maintainAspectRatio: false,

            plugins: {

                legend: {
                    display: false
                },

                tooltip: {

                    callbacks: {

                        label: function (context) {

                            return "Sales: " +
                                formatCurrency(context.raw);
                        }
                    }
                }
            },

            scales: {

                y: {

                    beginAtZero: true,

                    ticks: {

                        callback: function (value) {

                            return "₹" +
                                Number(value).toLocaleString("en-IN");
                        }
                    }
                },

                x: {

                    grid: {
                        display: false
                    }
                }
            }
        }

    });

}


// =========================
// REGION CHART
// =========================

function createRegionChart(
    data,
    regionColumn,
    salesColumn
) {

    const regionSales = {};

    data.forEach(function (row) {

        const region = String(
            row[regionColumn] || ""
        ).trim();

        const sales = getNumber(
            row[salesColumn]
        );

        if (region !== "") {

            if (!regionSales[region]) {
                regionSales[region] = 0;
            }

            regionSales[region] += sales;
        }

    });


    const labels = Object.keys(regionSales);

    const values = Object.values(regionSales);


    const canvas = document.getElementById("regionChart");

    if (!canvas) {
        return;
    }


    destroyRegionChart();


    regionChart = new Chart(canvas, {

        type: "doughnut",

        data: {

            labels: labels,

            datasets: [
                {
                    label: "Sales",

                    data: values,

                    borderWidth: 2
                }
            ]
        },

        options: {

            responsive: true,

            maintainAspectRatio: false,

            plugins: {

                legend: {
                    position: "bottom"
                },

                tooltip: {

                    callbacks: {

                        label: function (context) {

                            return context.label +
                                ": " +
                                formatCurrency(context.raw);
                        }
                    }
                }
            }
        }

    });

}


// =========================
// DISPLAY TABLE
// =========================

function displayTable(data) {

    if (!data || data.length === 0) {
        return;
    }


    const columns = Object.keys(data[0]);


    tableHead.innerHTML = "";


    const headerRow = document.createElement("tr");


    const numberHeader = document.createElement("th");

    numberHeader.textContent = "#";

    headerRow.appendChild(numberHeader);


    columns.forEach(function (column) {

        const th = document.createElement("th");

        th.textContent = column;

        headerRow.appendChild(th);

    });


    tableHead.appendChild(headerRow);


    tableBody.innerHTML = "";


    data.slice(0, 100).forEach(function (row, index) {

        const tr = document.createElement("tr");


        const numberCell = document.createElement("td");

        numberCell.textContent = index + 1;

        tr.appendChild(numberCell);


        columns.forEach(function (column) {

            const td = document.createElement("td");

            td.textContent =
                row[column] !== undefined
                    ? row[column]
                    : "";

            tr.appendChild(td);

        });


        tableBody.appendChild(tr);

    });


    rowCount.textContent =
        data.length.toLocaleString("en-IN") + " rows";
}


// =========================
// FORMAT CURRENCY
// =========================

function formatCurrency(value) {

    return new Intl.NumberFormat(
        "en-IN",
        {
            style: "currency",
            currency: "INR",
            maximumFractionDigits: 0
        }
    ).format(value || 0);
}


// =========================
// DESTROY PRODUCT CHART
// =========================

function destroyProductChart() {

    if (productChart) {

        productChart.destroy();

        productChart = null;
    }
}


// =========================
// DESTROY REGION CHART
// =========================

function destroyRegionChart() {

    if (regionChart) {

        regionChart.destroy();

        regionChart = null;
    }
}


// =========================
// RESET
// =========================

resetBtn.addEventListener("click", function () {

    fileInput.value = "";

    fileName.textContent = "";


    totalSalesElement.textContent = "₹0";

    totalOrdersElement.textContent = "0";

    averageSalesElement.textContent = "₹0";

    topProductElement.textContent = "-";


    rowCount.textContent = "0 rows";


    tableHead.innerHTML = `
        <tr>
            <th>No data</th>
        </tr>
    `;


    tableBody.innerHTML = `
        <tr>
            <td class="empty">
                Upload an Excel file to see your data here.
            </td>
        </tr>
    `;


    destroyProductChart();

    destroyRegionChart();

});