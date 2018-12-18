import * as React from 'react';
// import logo from './logo.svg';
// import './App.css';
import * as am4core from "@amcharts/amcharts4/core";
import * as am4charts from "@amcharts/amcharts4/charts";
import am4themes_animated from "@amcharts/amcharts4/themes/animated";

am4core.useTheme(am4themes_animated);

export default class LineChart extends React.Component<any,any>{
  public componentDidMount() {
    let chart = am4core.create("chartdiv", am4charts.XYChart);

    // Add data
    chart.data = [{
    "country": "USA",
    "visits": 2025
    }, {
    "country": "China",
    "visits": 1882
    }, {
    "country": "Japan",
    "visits": 1809
    }, {
    "country": "Germany",
    "visits": 1322
    }, {
    "country": "UK",
    "visits": 1122
    }, {
    "country": "France",
    "visits": 1114
    }, {
    "country": "India",
    "visits": 984
    }, {
    "country": "Spain",
    "visits": 711
    }, {
    "country": "Netherlands",
    "visits": 665
    }, {
    "country": "Russia",
    "visits": 580
    }, {
    "country": "South Korea",
    "visits": 443
    }, {
    "country": "Canada",
    "visits": 441
    }, {
    "country": "Brazil",
    "visits": 395
    }];

    // Create axes

    let categoryAxis = chart.xAxes.push(new am4charts.CategoryAxis());
    categoryAxis.dataFields.category = "country";
    categoryAxis.renderer.grid.template.location = 0;
    categoryAxis.renderer.minGridDistance = 30;

    // categoryAxis.renderer.labels.template.adapter.add("dy", => (dy, target) {
    // // if (target.dataItem && target.dataItem.index & 2 == 2) {
    //     return dy + 25;
    // // }
    //     return dy;
    // });

    let valueAxis = chart.yAxes.push(new am4charts.ValueAxis());

    // Create series
    let series = chart.series.push(new am4charts.ColumnSeries());
    series.dataFields.valueY = "visits";
    series.dataFields.categoryX = "country";
    series.name = "Visits";
    series.columns.template.tooltipText = "{categoryX}: [bold]{valueY}[/]";
    series.columns.template.fillOpacity = .8;

    let columnTemplate = series.columns.template;
    columnTemplate.strokeWidth = 2;
    columnTemplate.strokeOpacity = 1;
  }

  public render() {
    return (
      <div id="chartdiv" style={{ width: "100%", height: "500px" }}></div>
    );
  }
}
