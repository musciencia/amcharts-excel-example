
/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById('loadDataButton').onclick = getDataFromExcel;
  }
});



async function getDataFromExcel() {
  await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1:B6"); // Adjusted to capture headings and data
      range.load("values");
      await context.sync();

      const data = range.values; // This will hold the 2D array of cell data

      // Extract headings from the first row
      const headings = data[0];
      const chartData = data.slice(1); // Data excluding the headings

      // Pass the data to amCharts
      displayChart(chartData, headings);
  });
}

function displayChart(data, headings) {
  // This function will create a chart with amCharts
  am4core.ready(function() {
      // Create a chart instance
      var chart = am4core.create("chartdiv", am4charts.XYChart);

      // Prepare chart data using headings for X and Y axis labels
      chart.data = data.map(row => {
          return {
              category: row[0], // Category data from the first column
              value: row[1]     // Value data from the second column
          };
      });

      // Create axes
      var categoryAxis = chart.xAxes.push(new am4charts.CategoryAxis());
      categoryAxis.dataFields.category = "category";
      categoryAxis.title.text = headings[0]; // First row (heading) as X-axis title

      var valueAxis = chart.yAxes.push(new am4charts.ValueAxis());
      valueAxis.title.text = headings[1]; // Second row (heading) as Y-axis title

      // Create a series
      var series = chart.series.push(new am4charts.ColumnSeries());
      series.dataFields.valueY = "value";
      series.dataFields.categoryX = "category";

      // Create chart container
      chart.divId = "chartdiv"; // Make sure to have a div with this ID in your HTML
  });
}

