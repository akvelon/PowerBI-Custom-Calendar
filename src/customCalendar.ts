/*
*  Power BI Visualizations
*
*  Copyright (c) Microsoft Corporation
*  All rights reserved.
*  MIT License
*
*  Permission is hereby granted, free of charge, to any person obtaining a copy
*  of this software and associated documentation files (the ""Software""), to deal
*  in the Software without restriction, including without limitation the rights
*  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
*  copies of the Software, and to permit persons to whom the Software is
*  furnished to do so, subject to the following conditions:
*
*  The above copyright notice and this permission notice shall be included in
*  all copies or substantial portions of the Software.
*
*  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
*  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
*  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
*  THE SOFTWARE.
*/

module powerbi.extensibility.visual {
  "use strict";

  // legend
  import ILegend = powerbi.extensibility.utils.chart.legend.ILegend;
  import LegendData = powerbi.extensibility.utils.chart.legend.LegendData;
  import LegendPosition = powerbi.extensibility.utils.chart.legend.LegendPosition;
  import createLegend = powerbi.extensibility.utils.chart.legend.createLegend;

  // tooltip
  import ITooltipServiceWrapper = powerbi.extensibility.utils.tooltip.ITooltipServiceWrapper;
  import TooltipEventArgs = powerbi.extensibility.utils.tooltip.TooltipEventArgs;
  import createTooltipServiceWrapper = powerbi.extensibility.utils.tooltip.createTooltipServiceWrapper;

  // color
  import ColorHelper = powerbi.extensibility.utils.color.ColorHelper;

  // formatter
  import valueFormatter = powerbi.extensibility.utils.formatting.valueFormatter;

  interface ICalendarViewModel {
    settings: CalendarSettings;
    dataPoints: ICalendarDataPoint[];
  }

  interface ICalendarDataPoint {
    date: string;
    defaultDate: string;
    metric: string;
    hours: any;
    color: string;
    id: string;
    selectionId: powerbi.visuals.ISelectionId;
    index: number;
    metadataColumn: DataViewMetadataColumn;
  }

  interface Metrics {
    metrics: ICalendarMetric[];
  }

  interface ICalendarMetric {
    name: string;
    color: string;
    selectionId: powerbi.visuals.ISelectionId;
  }

  interface ITooltipDataPoint {
    displayName: string;
    value: string;
    color: string;
    header: string;
  }

  export class CustomCalendar implements IVisual {
    private host: IVisualHost;
    private calendarViewModel: ICalendarViewModel;
    private tooltipServiceWrapper: ITooltipServiceWrapper;
    private selectionManager: ISelectionManager;
    private currentViewport: IViewport;

    private settings: CalendarSettings;
    private calendarMetrics: Metrics;

    private rootElement: d3.Selection<any>;
    private visibleGroupContainer: d3.Selection<any>;
    private monthContainer: d3.Selection<SVGElement>;

    private months: string[] = ["January", "February", "March", "April",
      "May", "June", "July", "August",
      "September", "October", "November", "December"];

    private weekDays: string[] = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

    private legend: ILegend;
    private legendObjectProperties: DataViewObject;

    private colorPalette: IColorPalette;

    private tooltips: string[] = [];

    private dataPoints: ICalendarDataPoint[];

    private static selectedCell: any[] = [];
    private static isMultipleSelected: boolean = false;

    private startDate: Date;
    private categoryFormatString: string;

    constructor(options: VisualConstructorOptions) {
      this.init(options);
    }

    private init(options: VisualConstructorOptions): void {
      const element: HTMLElement = options.element;

      this.host = options.host;
      this.selectionManager = options.host.createSelectionManager();
      this.selectionManager.registerOnSelectCallback(
        (ids: any[]) => {
          if (ids.length == 0)
            return;

          let cells: HTMLElement[] = CustomCalendar.getCellBySelectionIds(ids);
          CustomCalendar.clearSelectedCells();
          CustomCalendar.selectCell(d3.select(cells[0]));

          if (cells.length > 1) {
            for (let i = 1; i < cells.length; i++)
              CustomCalendar.selectCell(d3.select(cells[i]), true);
          }
        });

      const visual: d3.Selection<any> = d3.select(options.element);
      const selectionManagerField: ISelectionManager = this.selectionManager;

      visual.on("click", function() {
        if ((d3.event as MouseEvent).toElement.id === "") {
          if (CustomCalendar.selectedCell.length == 0)
            return;
          else {
            for (let i = 0; i < CustomCalendar.selectedCell.length; i++) {
              let cell = document.getElementById(CustomCalendar.selectedCell[i]);
              d3.select(cell).attr("fill", "white");
            }

            CustomCalendar.selectedCell = [];
            selectionManagerField.clear();
          }
        }
      });

      this.colorPalette = this.host.colorPalette;
      this.tooltipServiceWrapper = createTooltipServiceWrapper(
        this.host.tooltipService,
        element
      );

      const svg: d3.Selection<any> = this.rootElement = d3.select(element)
        .style("overflow-y", "auto")
        .style("overflow-x", "auto")
        .attr('drag-resize-disabled', true)
        .classed('visualContainer', true);

      this.legend = createLegend(element, false, null, true);
      this.visibleGroupContainer = svg.append("div").attr("class", "visibleGroup");
    }

    private static getCellBySelectionIds(ids: any[]): HTMLElement[] {
      let cKeysArray: string[] = [];

      for (let i = 0; i < ids.length; i++) {
        let selector = ids[i].getSelector();
        let ckey = selector.metadata ? selector.metadata : "";

        ckey = "a" + ckey.replace( /\//g, "_" );

        if (ckey !== "a" && cKeysArray.indexOf(ckey) === -1)
          cKeysArray.push(ckey);
      }

      let cellArray: HTMLElement[] = [];
      for (let i = 0; i < cKeysArray.length; i++) {
        let cell = document.getElementById(cKeysArray[i]);
        if (cell)
          cellArray.push(cell);
      }

      return cellArray;
    }

    private static clearSelectedCells(): void {
      if (CustomCalendar.selectedCell.length == 0)
        return;
      else {
        for (let i = 0; i < CustomCalendar.selectedCell.length; i++) {
          let cell = document.getElementById(CustomCalendar.selectedCell[i]);
          d3.select(cell).attr("fill", "white");
        }
        CustomCalendar.selectedCell = [];
      }
    }

    public update(options: VisualUpdateOptions) {
      this.visibleGroupContainer.selectAll(".month").remove();
      this.dataPoints = [];

      const width: number = options.viewport.width;
      const height: number = options.viewport.height;

      this.calendarViewModel = this.visualTransform(options, this.host);

      const selectionManagerLegend: ISelectionManager = this.selectionManager;
      if (selectionManagerLegend.getSelectionIds().length == 0) {
        CustomCalendar.clearSelectedCells();
      }

      this.currentViewport = {
        width: Math.max(0, width),
        height: Math.max(0, height)
      };

      if (this.calendarViewModel.dataPoints.length > 0) {
        this.filterMetrics();
      }

      this.renderLegend(this.settings.legendSettings.show);
      this.drawCalendar(width, height);

      d3.selectAll(".legendItem").attr("fill-opacity", 1);

      const legends: d3.Selection<any> = d3.selectAll(".legendItem");
      const calendarMetrics: ICalendarMetric[] = this.calendarMetrics.metrics;

      legends.on("click", function (d) {
        selectionManagerLegend.select(d.identity).then((ids: ISelectionId[]) => {
          const selectedItemsNumber: number = selectionManagerLegend.getSelectionIds().length;

          legends.attr({
            'fill-opacity': ids.length > 0 ? 0.5 : 1
          });

          d3.select(this).attr("fill-opacity", 1);

          for (const metric of calendarMetrics) {
            const fillOpacityValue: number = metric.name !== d.label && selectedItemsNumber > 0 ? 0.3 : 1;

            d3.selectAll("." + metric.name.replace(" ", ""))
              .attr("fill-opacity", fillOpacityValue);
          }
        });
      });
    }

    private visualTransform(options: VisualUpdateOptions, host: IVisualHost): ICalendarViewModel {
      let dataViews: DataView[] = options.dataViews;
      const calendarMetrics: ICalendarMetric[] = [];
      const viewModel: ICalendarViewModel = {
        settings: <CalendarSettings>{},
        dataPoints: <ICalendarDataPoint[]>[]
      };
      const calendarSettings: CalendarSettings = this.settings = CustomCalendar.parseSettings(options.dataViews[0]);

      this.setStartDate();

      if (!dataViews
        || !dataViews[0]
        || !dataViews[0].categorical
        || !dataViews[0].categorical.categories
        || !dataViews[0].categorical.categories[0].source
        || !dataViews[0].categorical.values) {

        return viewModel;
      }

      let index = dataViews[0].categorical.values.length;
      dataViews[0].categorical.values.forEach(value => {
        value.source.index = index;
        index--;
      });

      const categorical: DataViewCategorical = dataViews[0].categorical;
      const category: DataViewCategoryColumn = categorical.categories[0];
      const dataPoints = {
        data: []
      };

      let colorHelper: ColorHelper = new ColorHelper(
        host.colorPalette,
        {
          objectName: "metricsSettings",
          propertyName: "metricColor"
        });

      for (let i = 0, len = Math.min(categorical.values.length, 100); i < len; i++) {
        const groupedValue: DataViewValueColumn = categorical.values[i];
        const objects: DataViewObjects = groupedValue.source.objects;

        const color = colorHelper.getColorForSeriesValue(
          objects,
          this.colorPalette.getColor(groupedValue.source.displayName).value);

        for (let j = 0, len = Math.max(category.values.length, groupedValue.values.length); j < len; j++) {
          if (groupedValue.values[j] != null) {

            const categoryFormatString = valueFormatter.getFormatStringByColumn(category.source);
            this.categoryFormatString = categoryFormatString;
            const value = valueFormatter.format(category.values[j], categoryFormatString);
            const date: Date = <Date>category.values[j];
            const stringDate: string = CustomCalendar.convertDateToString(date);
            const hours: any = <number>groupedValue.values[j];
            const metadataColumn: DataViewMetadataColumn = groupedValue.source;

            if (hours !== null && hours !== undefined) {
              dataPoints.data.push({
                date: stringDate,
                defaultDate: value,
                hours: hours,
                color: color,
                value: groupedValue,
                index: groupedValue.source.index,
                metadataColumn: metadataColumn
              });
            }
          }
        }

        const selectionId: visuals.ISelectionId = host.createSelectionIdBuilder()
          .withSeries(categorical.values, groupedValue)
          .withMeasure(groupedValue.source.queryName)
          .createSelectionId();

        calendarMetrics.push({
          color,
          name: <string>groupedValue.source.displayName,
          selectionId: selectionId
        });
      }

      dataPoints.data.sort(function (second, first) {
        if (first.date == null) {
          return 1;
        } else if (second.date == null) {
          return -1;
        }

        const firstDate = first.date.split("/");
        const secondDate = second.date.split("/");

        if (parseInt(firstDate[2]) < parseInt(secondDate[2])) {
          return 1;
        } else if (parseInt(firstDate[2]) > parseInt(secondDate[2])) {
          return -1;
        }

        if (parseInt(firstDate[0]) < parseInt(secondDate[0])) {
          return 1;
        } else if (parseInt(firstDate[0]) > parseInt(secondDate[0])) {
          return -1;
        }

        if (parseInt(firstDate[1]) <= parseInt(secondDate[1])) {
          return 1;
        } else if (parseInt(firstDate[1]) > parseInt(secondDate[1])) {
          return -1;
        }

        return 0;
      });

      const datesArr = [];
      for (let i = 0; i < dataPoints.data.length - 1; i++) {
        if (dataPoints.data[i].date !== dataPoints.data[i + 1].date) {
          datesArr.push(dataPoints.data[i].date);
        }
      }
      datesArr.push(dataPoints.data[dataPoints.data.length - 1].date);

      for (let j = 0; j < dataPoints.data.length; j++) {
        this.dataPoints.push({
          date: dataPoints.data[j].date,
          defaultDate: dataPoints.data[j].defaultDate,
          metric: dataPoints.data[j].value.source.displayName,
          hours: dataPoints.data[j].hours,
          color: dataPoints.data[j].color,
          id: "#a" + dataPoints.data[j].date,
          selectionId: host.createSelectionIdBuilder()
            .withCategory(dataViews[0].categorical.categories[0],
            datesArr.indexOf(dataPoints.data[j].date))
            .withMeasure(dataPoints.data[j].date)
            .createSelectionId(),
          index: dataPoints.data[j].index,
          metadataColumn: dataPoints.data[j].metadataColumn
        });
      }

      this.calendarMetrics = {
        metrics: calendarMetrics
      };

      let tooltipColumns: string[] = [];

      for (let i = 0; i < categorical.values.length; i++) {
        if(categorical.values[i].source.roles['tooltips']) {
          tooltipColumns.push(categorical.values[i].source.displayName);
        }
      }

      this.tooltips = tooltipColumns.sort();

      return {
        settings: calendarSettings,
        dataPoints: this.dataPoints
      };
    }

    private static GetTooltip(value: any): VisualTooltipDataItem[] {
      if (value) {
        return value;
      }
    }

    private static parseSettings(dataView: DataView): CalendarSettings {
      return CalendarSettings.parse(dataView) as CalendarSettings;
    }

    private renderLegend(legendShown: boolean): void {
      const legendTitleShow: boolean = this.settings.legendSettings.legendTitleShow;
      const legendTitleName: string = this.settings.legendSettings.legendTitleName;
      const legendLabelColor: string = this.settings.legendSettings.legendLabelColor;
      const legendLabelFontSize: number = this.settings.legendSettings.legendLabelFontSize;

      if (!legendShown || (this.calendarViewModel.dataPoints.length === 0)) {
        this.rootElement.style("margin-top", "0px");
        this.rootElement.select('.legend').style('display', 'none');

        return;
      }

      let legendTitle: string;
      if (legendTitleShow) {
        legendTitle = legendTitleName;
      }

      const legendDataTorender: LegendData = {
        fontSize: legendLabelFontSize,
        dataPoints: [],
        title: legendTitle,
        labelColor: legendLabelColor
      };

      for (let i = 0; i < this.calendarMetrics.metrics.length; i++) {
        const metric: ICalendarMetric = this.calendarMetrics.metrics[i];

        legendDataTorender.dataPoints.push({
          label: metric.name,
          color: metric.color,
          icon: powerbi.extensibility.utils.chart.legend.LegendIcon.Box,
          selected: false,
          identity: metric.selectionId
        });
      }

      this.legend.changeOrientation(LegendPosition.Top);

      const legend = powerbi.extensibility.utils.chart.legend;
      const legendData = powerbi.extensibility.utils.chart.legend.data;
      const legendPosition: string = powerbi.extensibility.utils.chart.legend.legendProps.position;

      if (this.legendObjectProperties) {
        legendData.update(legendDataTorender, this.legendObjectProperties);

        const position: string = <string>this.legendObjectProperties[legendPosition];
        if (position) {
          this.legend.changeOrientation(LegendPosition[position]);
        }
      }

      this.legend.drawLegend(legendDataTorender, this.currentViewport);
      legend.positionChartArea(this.rootElement, this.legend);

      this.rootElement.select('.legend').style('top', 0);
      this.rootElement.select('.legend').style("display", "inherit");
    }

    private drawCalendar(width: number, height: number): void {
      const cellSize: number = this.settings.calendarSettings.cellSize;
      const headerColor: string = this.settings.calendarSettings.calendarHeaderColor;
      const headerTitleColor: string = this.settings.calendarSettings.calendarHeaderTitleColor;
      const monthSize: number = cellSize * 7;
      const startDate: Date = this.getStartDate();
      const monthsNumber: number = this.getMonthsNumber();
      const startMonth: number = Number(startDate.getMonth());
      
      this.rootElement.attr("width", width);
      this.rootElement.attr("height", height);
      this.visibleGroupContainer.attr("width", width);
      this.visibleGroupContainer.attr("height", height);

      for (let monthCount = 0; monthCount < monthsNumber; monthCount++) {
        const monthIndex: number = (startMonth + monthCount) % this.months.length;
        const month: string = this.months[monthIndex];
        const year: number = CustomCalendar.getYear(startDate, monthCount);

        this.monthContainer = this.visibleGroupContainer.append("svg").attr("class", "month")
          .attr("width", monthSize + 15)
          .attr("height", monthSize + cellSize + 15)
          .append("g");

        const header: d3.Selection<SVGElement> = this.monthContainer.append("svg").attr("class", "monthHeader")
          .attr("width", monthSize)
          .attr("height", cellSize);

        header.append("rect").attr("x", 1)
          .attr("y", 1)
          .attr("width", cellSize * 7)
          .attr("height", cellSize)
          .attr("fill", headerColor)
          .attr("id", "head" + monthCount);

        header.append("text").text(month + " " + year)
          .attr("x", "50%")
          .attr("y", "65%")
          .attr("text-anchor", "middle")
          .attr("font-size", cellSize / 2 + "px")
          .attr("fill", headerTitleColor);

        this.drawDaysLabels(monthSize);
        this.drawMonthCells(monthSize, month, year, this.getWeekDays(), monthCount);
      }
    }

    private filterMetrics(): void {
      const filteredCalendarMetrics: ICalendarMetric[] = [];
      const drawnMetrics: string[] = [];
      const startDate: Date = this.getStartDate();

      const monthsNumber: number = this.getMonthsNumber();
      const date: string = startDate.toDateString();
      const endDate: Date = new Date(date);

      endDate.setMonth(startDate.getMonth() + monthsNumber);
      endDate.setDate(1);

      for (let dataPoint of this.calendarViewModel.dataPoints) {
        const currentDate: Date = new Date(dataPoint.date);
        if (currentDate >= startDate && currentDate < endDate && dataPoint.metadataColumn.roles['metrics']) {
          if (drawnMetrics.indexOf(dataPoint.metric) === -1) {
            drawnMetrics.push(dataPoint.metric);
          }
        }
      }

      for (let metric of this.calendarMetrics.metrics) {
        if (drawnMetrics.indexOf(metric.name) > -1) {
          filteredCalendarMetrics.push(metric);
        }
      }

      this.calendarMetrics = {
        metrics: filteredCalendarMetrics
      };

    }

    private static getYear(startDate: Date, monthCount: number): number {
      const date: string = startDate.toDateString();
      const newDate: Date = new Date(date);

      newDate.setMonth(startDate.getMonth() + monthCount);

      return newDate.getFullYear();
    }

    private getStartDate(): Date {
      let startDate: Date = new Date(1, 1, 1);
      startDate.setFullYear(this.startDate.getFullYear());
      startDate.setMonth(this.startDate.getMonth());
      startDate.setDate(this.startDate.getDate());
      return startDate;
    }

    private setStartDate(): void {
      const calendarType: number = this.settings.calendarSettings.calendarType;
      //Type is Fixed
      if (calendarType === 0) {
        const startDateSettings: string = this.settings.calendarSettings.startDate;
        this.startDate = CustomCalendar.getValidateDate(startDateSettings);
      }
      //Type is Relative
      else {
        const numOfPreviousMonths: number = this.settings.calendarSettings.numOfPreviousMonths;
        let startDate = new Date();
        startDate.setMonth(startDate.getMonth() - numOfPreviousMonths);
        this.startDate = startDate;
      }
      //The calendar can display data based on the day.To enable this functionality, comment the following line
      this.startDate.setDate(1);
    }

    private static getValidateDate(date: string): Date {
      if ((date.search(/^((0?[1-9]|1[012])[- /.](0?[1-9]|[12][0-9]|3[01])[- /.](19|20)[0-9]{2})*$/) === 0) && (date != "")) {
        const month: number = Number(date.slice(0, date.indexOf("/")));
        const day: number = Number(date.slice(date.indexOf("/") + 1, -5));
        const year: number = Number(date.slice(-4));
        const currentDate = new Date(year, month - 1);
        if (day <= CustomCalendar.getDaysInMonth(currentDate)) {
          currentDate.setDate(day);
          return currentDate;
        }
      }
      return new Date();
    }

    private static getDaysInMonth(date: Date): number {
      return 33 - new Date(date.getFullYear(), date.getMonth(), 33).getDate();
    }

    private getMonthsNumber(): number {
      const calendarType: number = this.settings.calendarSettings.calendarType;
      const numOfMonths: number = this.settings.calendarSettings.numOfMonths;
      const numOfPreviousMonths: number = this.settings.calendarSettings.numOfPreviousMonths;
      const numOfFollowingMonths: number = this.settings.calendarSettings.numOfFollowingMonths;
      return calendarType === 0 ? numOfMonths :
        calendarType === 1 ? numOfPreviousMonths + numOfFollowingMonths :
          12;
    }

    private drawMonthCells(monthSize: number, month: string, year: number,
        newWeek: string[], monthCount): void {
      const cellSize: number = this.settings.calendarSettings.cellSize;
      const cellBorderColor: string = this.settings.calendarSettings.cellBorderColor;
      const dayLabelsColor: string = this.settings.calendarSettings.dayLabelsColor;
      const monthIndex: number = this.months.indexOf(month);

      const cellsContainer: d3.Selection<any> = this.monthContainer.append("svg")
        .attr("class", "cellsContainer")
        .attr("width", monthSize + 10)
        .attr("height", monthSize + cellSize + 10);

      const monthCells: d3.Selection<any>[] = [];

      let numOfDays: number;
      let firstDay: number;
      let newFirstDay: number;
      let count: number = 0;
      let dayNumber: number;
      let cellRowNumber: number;
      
      //Set the parameters for the first month
      if (monthCount === 0 ) {
        const startDate: Date = this.getStartDate();
        firstDay = startDate.getDay();
        cellRowNumber = this.getCountWeek(startDate);
        dayNumber = startDate.getDate();
      }
      //Set the parameters for the remaining months
      else {
        firstDay = new Date(year, monthIndex).getDay();
        cellRowNumber = this.getCountWeek(new Date(year, monthIndex));
        dayNumber = 1;
      }
      newFirstDay = newWeek.indexOf(this.weekDays[firstDay]);
      numOfDays = new Date(year, monthIndex + 1, 0).getDate();

      for (let cellRow = 0; cellRow < cellRowNumber; cellRow++) {
        for (let cellMonthCount = 0; cellMonthCount < 7; cellMonthCount++) {
          const currentX: number = 1 + cellSize * cellMonthCount;
          const currentY: number = 1 + (cellSize * 2) + (cellSize * cellRow);
          const id: string = "a" + (monthIndex + 1) + "_" + dayNumber + "_" + year;

          if (count >= newFirstDay && dayNumber <= numOfDays) {
            
            const cell: d3.Selection<any> = cellsContainer.append("rect").attr("x", currentX)
              .attr("y", currentY)
              .attr("width", cellSize)
              .attr("height", cellSize)
              .attr("class", "cell")
              .attr("id", id)
              .attr("stroke", cellBorderColor)
              .attr("stroke-width", 1)
              .attr("fill", "white");

            for (let i = 0; i < CustomCalendar.selectedCell.length; i++) {
              if (id === CustomCalendar.selectedCell[i]) {
                cell.attr("fill", "darkgrey");
              }
            }

            const label: d3.Selection<any> = cellsContainer.append("text").text(dayNumber)
              .attr("x", currentX + (cellSize / 2))
              .attr("y", currentY + (cellSize / 3))
              .attr("id", id)
              .attr("class", "cell_label")
              .attr("text-anchor", "middle")
              .attr("font-size", cellSize / 4)
              .attr("fill", dayLabelsColor);

            monthCells.push(label);
            
            //Add cards only on days with metrics
            const currentDate: Date = new Date(year, monthIndex, dayNumber);
            if (this.GetIsDayWithMetric(currentDate)) {

              //Adding cell tooltip here
              cell.data(this.GetTooltipsDataCell(this.GetMetricByDate(currentDate), currentDate));
              this.tooltipServiceWrapper.addTooltip(
                cell,
                (tooltipEvent: TooltipEventArgs<number>) => CustomCalendar.GetTooltip(tooltipEvent.data),
                null);

              //Adding label tooltip here
              label.data(this.GetTooltipsDataCell(this.GetMetricByDate(currentDate), currentDate));
              this.tooltipServiceWrapper.addTooltip(
                label,
                (tooltipEvent: TooltipEventArgs<number>) => CustomCalendar.GetTooltip(tooltipEvent.data),
                null);
            }

            dayNumber++;
          }

          count++;
        }
      }

      this.drawMetrics(cellsContainer, monthIndex + 1, year, monthCells, monthCount);
    }

    private drawMetrics(cellsContainer: d3.Selection<any>, monthNumber: number,
      year: number, monthCells: any, monthCount: number): ICalendarDataPoint[] {
      const cellSize: number = <number>this.settings.calendarSettings.cellSize;
      const sortedDaysMetrics: any[] = [];
      let sameMonthDataPoints: ICalendarDataPoint[] = [];
      let dataPoints: ICalendarDataPoint[] = this.dataPoints;
      let index: number = 0;

      for (let i = 0; i < dataPoints.length; i++) {
        if (dataPoints[i].date != null) {
          const dataPoint: ICalendarDataPoint = dataPoints[i];
          const renderedDate: Date = new Date(dataPoint.date);
          const renderedDateString: string = CustomCalendar.convertDateToString(renderedDate);
          const monthString: string = String(monthNumber) + "/";

          if ((renderedDateString.indexOf(monthString) === 0)
            && (renderedDateString.indexOf(String(year)) > -1)) {
            sameMonthDataPoints[index] = dataPoint;
            index++;
          }
        }
      }

      //Since the first month is incomplete, we remove the unnecessary metric
      if (monthCount === 0) {
        const startDay: number = this.getStartDate().getDate();
        let currentDay: number;
        for (let i = 0; i < sameMonthDataPoints.length; i++) {
          currentDay = CustomCalendar.getDayFromDataStr(sameMonthDataPoints[i].date);
          if (currentDay >= startDay) {
            sameMonthDataPoints = sameMonthDataPoints.slice(i);
            break;
          }
          if (i === sameMonthDataPoints.length - 1) {
            sameMonthDataPoints = [];
          }
        }
      }

      const self = this;
      const usedDates: string[] = [];
      let usedDatesIndex: number = 0;

      for (let i = 0; i < sameMonthDataPoints.length; i++) {
        const sameDateDataPoints: ICalendarDataPoint[] = [];
        let count: number = 0;
        for (let j = i + 1; j < sameMonthDataPoints.length; j++) {
          if (sameMonthDataPoints[i].date === sameMonthDataPoints[j].date) {
            sameDateDataPoints[count] = sameMonthDataPoints[j];
            count++;
          }
        }

        if (usedDates.indexOf(sameMonthDataPoints[i].date) === -1) {
          const dataPointsToDraw: ICalendarDataPoint[] = [];

          dataPointsToDraw[0] = sameMonthDataPoints[i];
          for (let j = 0; j < sameDateDataPoints.length; j++) {
            dataPointsToDraw[j + 1] = sameDateDataPoints[j];
          }

          let id: string;
          for (let j = 0; j < dataPointsToDraw.length; j++) {
            const date: string = CustomCalendar.convertDateToString(new Date(dataPointsToDraw[j].date))
              .replace("/", "_");
            const dateFormat: string = date.replace("/", "_");
            id = "#a" + dateFormat;

            dataPointsToDraw[j].id = id;
          }

          dataPointsToDraw.sort((first, second) => {
            if (first.index < second.index) {
              return -1;
            } else {
              return 1;
            }
          });

          let sortedMetrics: ICalendarDataPoint[] = [];

          for (let k = 0; k < dataPointsToDraw.length; k++) {
            if (dataPointsToDraw[k].metadataColumn.roles['metrics']) {
              sortedMetrics.push(dataPointsToDraw[k]);
            }
          }

          sortedDaysMetrics.push(dataPointsToDraw);

          let previousYCoord: number = cellSize;
          for (let j = 0; j < sortedMetrics.length; j++) {
            //If the metric is zero, then simply add a click handler to the cell. Otherwise, we draw the metric and add the click handler to both the cell and the metric
            if (sortedMetrics[j].hours === 0) {
              for (let i = 0; i < monthCells.length; i++) {
                const monthCell: any = monthCells[i][0][0];
                const cell: d3.Selection<any> = d3.select(monthCell.previousSibling);
                const label: d3.Selection<any> = d3.select(monthCell);
                const id: string = "#" + cell.attr("id");

                if (id === sortedMetrics[j].id) {
                  cell.on("click", function () {
                    self.select(cell, [sortedMetrics[0]]);
                  });
                  label.on("click", function () {
                    self.select(cell, [sortedMetrics[0]]);
                  });
                }
              }
            } else {
              const metricCoeff: number = CustomCalendar.getMetricHeight(sortedMetrics, j);
              const width: number = cellSize - 2;
              const height: number = (cellSize - (cellSize / 2.3)) * metricCoeff;
              const xCoord: number = Number(d3.select(id).attr("x")) + 1;
              const yCoord: number = Number(d3.select(id).attr("y")) + previousYCoord - height - 1;
              const metricFormat: string = sortedMetrics[j].metric.replace(" ", "");

              const cellId = id.replace('#', '');
              const cellMetrics: d3.Selection<any> = cellsContainer.append("rect")
                .attr("id", cellId)
                .attr("class", "metric " + metricFormat)
                .attr("fill-opacity", 1)
                .attr("width", width)
                .attr("height", height)
                .attr("x", xCoord)
                .attr("y", yCoord)
                .attr("fill", sortedMetrics[j].color);

              for (let i = 0; i < monthCells.length; i++) {
                const monthCell: any = monthCells[i][0][0];
                const cell: d3.Selection<any> = d3.select(monthCell.previousSibling);
                const label: d3.Selection<any> = d3.select(monthCell);
                const id: string = "#" + cell.attr("id");

                if (id === sortedMetrics[j].id) {
                  cell.on("click", function () {
                    self.select(cell, [sortedMetrics[0]]);
                  });
                  label.on("click", function () {
                    self.select(cell, [sortedMetrics[0]]);
                  });
                  cellMetrics.on("click", function () {
                    self.select(cell, [sortedMetrics[j]]);
                  });
                }
              }
              //Adding tooltip here
              cellMetrics.data(this.GetTooltipsDataMetric(dataPointsToDraw, sortedMetrics[j]));
              this.tooltipServiceWrapper.addTooltip(
                cellMetrics,
                (tooltipEvent: TooltipEventArgs<number>) => CustomCalendar.GetTooltip(tooltipEvent.data),
                null);

              previousYCoord = previousYCoord - height;
            }
          }

          usedDates[usedDatesIndex] = sameMonthDataPoints[i].date;
          usedDatesIndex++;
        }
      }

      return sortedDaysMetrics;
    }

    private GetTooltipsDataCell(allPoints: ICalendarDataPoint[], currentDate: Date): ITooltipDataPoint[] {
      const dataPoints: ICalendarDataPoint[] = CustomCalendar.GetDataPointsWithoutZeroMetric(allPoints);
      let currentTooltipDataPoint: ITooltipDataPoint[] = [];
      let resultTooltipData: any[] = [];
      let tooltips: string[] = this.tooltips.slice();
      let indexInPoints: number = 0;

      if (dataPoints.length === 0) {
        let tooltipData: any = {};
        tooltipData.header = valueFormatter.format(currentDate, this.categoryFormatString);
        currentTooltipDataPoint.push(tooltipData);
      } else {
        for (let i = 0; i < tooltips.length; i++) {
          let tooltipData: any = {};
          tooltipData.header = valueFormatter.format(currentDate, this.categoryFormatString);
          tooltipData.color = this.GetColorByMetricName(tooltips[i]);
          tooltipData.displayName = tooltips[i];
          indexInPoints = CustomCalendar.GetIndexMetricInPointsArray(tooltips[i], dataPoints);
          if (indexInPoints >= 0) {
            tooltipData.value = CustomCalendar.tooltipValue(dataPoints[indexInPoints].metadataColumn, dataPoints[indexInPoints].hours);
          }
          currentTooltipDataPoint.push(tooltipData);
        }
      } 

      resultTooltipData.push(currentTooltipDataPoint);
      return resultTooltipData;
    }

    private GetTooltipsDataMetric(allPoints: ICalendarDataPoint[], point: ICalendarDataPoint): ITooltipDataPoint[] {
      const dataPoints: ICalendarDataPoint[] = CustomCalendar.GetDataPointsWithoutZeroMetric(allPoints);
      let currentTooltipDataPoint: ITooltipDataPoint[] = [];
      let tooltips: string[] = this.tooltips.slice();
      let resultTooltipData: any[] = [];

      //Add the current metric to the tooltip
      let tooltipData: any = {};
      tooltipData.color = point.color;
      tooltipData.displayName = point.metric;
      tooltipData.header = point.defaultDate;
      tooltipData.value = point.hours.toString();
      currentTooltipDataPoint.push(tooltipData);

      //Remove the name of an already added metric from the array tooltips
      if (tooltips.indexOf(point.metric) >= 0) {
        tooltips.splice(tooltips.indexOf(point.metric), 1);
      }

      //Add to the hint all the remaining zero metrics from the array tooltips
      for (let i = 0; i < tooltips.length; i++) {
        let tooltipData: any = {};
        tooltipData.header = point.defaultDate;
        tooltipData.color = this.GetColorByMetricName(tooltips[i]);
        tooltipData.displayName = tooltips[i];

        let indexInPoints: number = CustomCalendar.GetIndexMetricInPointsArray(tooltips[i], dataPoints);
        if (indexInPoints >= 0) {
          tooltipData.value = CustomCalendar.tooltipValue(dataPoints[indexInPoints].metadataColumn, dataPoints[indexInPoints].hours);
        }
        currentTooltipDataPoint.push(tooltipData);
      }
      //Sort the metrics in the order of the display on the tooltip. First specified by the user, then the rest
      currentTooltipDataPoint.sort((first, second) => {
        let sortMetricArr: string[] = this.GetSortMetricArray();
        if (CustomCalendar.GetIndexMetricInArray(first.displayName, sortMetricArr) < CustomCalendar.GetIndexMetricInArray(second.displayName, sortMetricArr)) {
          return -1;
        } else {
          return 1;
        }
      });

      resultTooltipData.push(currentTooltipDataPoint);
      return resultTooltipData;
    }

    private static tooltipValue(metadataColumn: DataViewMetadataColumn, value: PrimitiveValue): any {
      return CustomCalendar.getFormattedValue(metadataColumn, value)
    }

    private static getFormattedValue(column: DataViewMetadataColumn, value: any) {
      let formatString: string = CustomCalendar.getFormatStringFromColumn(column);

      return valueFormatter.format(value, formatString);
    }

    private static getFormatStringFromColumn(column: DataViewMetadataColumn): string {
      if (column) {
        let formatString: string = valueFormatter.getFormatStringByColumn(column, false);

        return formatString || column.format;
      }

      return null;
    }

    private static GetDataPointsWithoutZeroMetric(allPoints: ICalendarDataPoint[]): ICalendarDataPoint[] {
      let dataPoints: ICalendarDataPoint[] = allPoints.slice();
      let i = 0;
      while (i < dataPoints.length) {
        if (dataPoints[i].hours === 0) {
          dataPoints.splice(i, 1);
        } else {
          i = i + 1
        }
      }
      return dataPoints;
    }

    private select(cell: d3.Selection<any>, sortedDataPoints: ICalendarDataPoint[]) {
      if(CustomCalendar.selectCell(cell))
        this.selectMetrics(sortedDataPoints);
      else 
        this.selectionManager.clear();
    }

    private static selectCell(cell: d3.Selection<any>, check: boolean = false): boolean {
      let multipleSelection: boolean = check || ((d3.event as MouseEvent) ? (d3.event as MouseEvent).ctrlKey : false);
      const cellFill: string = cell.attr("fill");
      
      if (multipleSelection) {
        let isAdded: boolean = false;
        let indexDeletedItem: number = 0;
        for (let i = 0; i < CustomCalendar.selectedCell.length; i++) {
          if (CustomCalendar.selectedCell[i] === cell.attr("id")) {
            indexDeletedItem = i;
            isAdded = true;
            break;
          }
        }
        if (isAdded) {
          cell.attr("fill", "white");
          CustomCalendar.selectedCell.splice(indexDeletedItem, 1);
        } else {
          cell.attr("fill", "darkgrey");
          CustomCalendar.selectedCell.push(cell.attr("id"));
        }
        CustomCalendar.isMultipleSelected = true;
      } else {
        d3.selectAll(".cell").attr("fill", "white");
        CustomCalendar.selectedCell = [];
        if (cellFill === "white") {
          cell.attr("fill", "darkgrey");
          CustomCalendar.selectedCell[0] = cell.attr("id");
        } else {
          if (CustomCalendar.isMultipleSelected) {
            cell.attr("fill", "darkgrey");
            CustomCalendar.selectedCell[0] = cell.attr("id");
          }
          else {
            CustomCalendar.isMultipleSelected = false;

            return false;
          }
        }
        CustomCalendar.isMultipleSelected = false;
      }

      return true;
    }

    private selectMetrics(d: ICalendarDataPoint[]) {
      if (!CustomCalendar.isMultipleSelected) {
        this.selectionManager.clear();
      }

      let multipleSelection = CustomCalendar.isMultipleSelected;

      for (let i = 0; i < d.length; i++) {
        if (i != 0 && d.length > 1)
          multipleSelection = true;
        if (d[i].selectionId)
          this.selectionManager.select(d[i].selectionId, multipleSelection);
        else
          this.selectionManager.select(d[i], multipleSelection);
      }
    }

    private static getMetricHeight(dataPoints: ICalendarDataPoint[], dataPointNumber): number {
      const metric: ICalendarDataPoint = dataPoints[dataPointNumber];
      const metricsSum: number = CustomCalendar.getMetricsSum(dataPoints, metric.id);

      return CustomCalendar.isNumeric(metric.hours) ? (metric.hours / metricsSum) : 0;
    }

    private static isNumeric(value: any) : boolean {
      return !isNaN(parseFloat(value)) && isFinite(value);
    }

    private static getMetricsSum(dataPoints: ICalendarDataPoint[], metricId: string): number {
      let metricsSum: number = 0;
      for (let i = 0; i < dataPoints.length; i++) {
        if (dataPoints[i].id === metricId && CustomCalendar.isNumeric(dataPoints[i].hours)) {
          metricsSum = metricsSum + dataPoints[i].hours;
        }
      }

      return metricsSum;
    }

    private GetIsDayWithMetric(date: Date): boolean {
      const dataPoints: ICalendarDataPoint[] = this.calendarViewModel.dataPoints;
      let currentDate: Date;
      for (let i = 0; i < dataPoints.length; i++) {
        currentDate = new Date(dataPoints[i].date);
        if ((date.getFullYear() === currentDate.getFullYear()) && (date.getMonth() === currentDate.getMonth()) && (date.getDate() === currentDate.getDate())) {
          return true;
        }
      }
      return false;
    }

    private static convertDateToString(date: Date): string {
      if (date == null) {
        return null;
      }

      return (date.getMonth() + 1) + "/"
        + (date.getDate()) + "/"
        + (date.getFullYear());
    }

    private drawDaysLabels(monthSize: number): void {
      const cellSize: number = this.settings.calendarSettings.cellSize;
      const dayLabelColor: string = this.settings.calendarSettings.weekDayLabelsColor;
      const newWeek: string[] = this.getWeekDays();
      const dayLabelSize: number = Math.ceil(this.settings.calendarSettings.cellSize / 3);

      const xScale: d3.scale.Ordinal<string, number> = d3.scale.ordinal()
        .domain(newWeek)
        .range(d3.range(0, 6))
        .rangeRoundPoints([Math.round(cellSize / 3),
        monthSize - Math.round(cellSize / 3) - Math.round(cellSize / 10)]);

      const xAxis: d3.svg.Axis = d3.svg.axis().scale(xScale)
        .orient("bottom");

      const lineContainer: d3.Selection<any> = this.monthContainer.append("svg").attr("class", "weekDays")
        .attr("width", "100%")
        .attr("height", "100%")
        .attr("x", 2)
        .attr("y", cellSize);

      lineContainer.append("g").attr("class", "line")
        .attr("transform", "translate(0," + (cellSize / 10) + ")")
        .call(xAxis);

      d3.selectAll(".line").attr("font-family", "wf_standard-font,helvetica,arial,sans-serif")
        .attr("font-size", dayLabelSize + "px")
        .attr("fill", dayLabelColor)
        .attr("text-anchor", "middle");

      //Completely hide the scale
      d3.selectAll(".line .domain")
        .attr("fill", "none");
    }

    private getWeekDays(): string[] {
      const newWeek: string[] = [];
      let firstDay: number = this.settings.calendarSettings.firstDay;

      for (let i = 0; i < 7; i++) {
        newWeek[i] = this.weekDays[firstDay % this.weekDays.length];
        firstDay++;
      }

      return newWeek;
    }

    private getCountWeek(date: Date): number {
      const dayInDate: number = date.getDay();
      const firstDay: number = this.settings.calendarSettings.firstDay;

      const countDayMount: number = new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
      const countShowDayMount: number = countDayMount - date.getDate() + 1;

      let countShowDayFirstWeek: number;
      let countWeek: number;

      //We get the number of days in the first week
      if (firstDay > dayInDate) {
        countShowDayFirstWeek = firstDay - dayInDate;
      } else if (firstDay <= dayInDate) {
        countShowDayFirstWeek = 7 - (dayInDate - firstDay);
      }

      //There is a minimum of one week. We add to it the number of weeks until the end of the month.
      countWeek = 1 + Math.ceil((countShowDayMount - countShowDayFirstWeek) / 7);

      return countWeek;
    }

    private static getDayFromDataStr(dataStr: string): number {
      let dayStr: string;
      dayStr = dataStr.slice(0, -5);
      dayStr = dayStr.slice(dayStr.indexOf("/") + 1);
      return Number(dayStr);
    }

    private GetColorByMetricName(name: string): string {
      const metrics: ICalendarMetric[] = this.calendarMetrics.metrics;
      let color: string = "";
      for (let j = 0; j < metrics.length; j++) {
        if (metrics[j].name === name) {
          color = metrics[j].color;
          break;
        }
      }
      return color;
    }

    private GetMetricByDate(date: Date): any[] {
      const dataPoints: ICalendarDataPoint[] = this.dataPoints;
      let currentDataPoints: ICalendarDataPoint[];
      let isIndexBeginInitialized: boolean = false;
      let indexBegin: number = 0;
      let indexEnd: number = 0;
      let currentDate: Date;

      //Find the index of the first and last metric on the required date
      for (let i = 0; i < dataPoints.length; i++) {
        currentDate = new Date(dataPoints[i].date);
        if ((currentDate.getFullYear() === date.getFullYear()) && (currentDate.getMonth() === date.getMonth()) && (currentDate.getDate() === date.getDate())) {
          if (!isIndexBeginInitialized) {
            indexBegin = i;
            isIndexBeginInitialized = true;
          }
          indexEnd = i;
        }
      }

      //Copy the metrics for the required date
      if (dataPoints.length === indexEnd + 1) {
        currentDataPoints = dataPoints.slice(indexBegin);
      } else {
        currentDataPoints = dataPoints.slice(indexBegin, indexEnd + 1);
      }

      return currentDataPoints;
    }

    private static GetIndexMetricInArray(name: string, arr: any[]): number {
      let index: number = -1;
      for (let i = 0; i < arr.length; i++) {
        if (name === arr[i]) {
          index = i;
        }
      } 
      return index;
    }

    private static GetIndexMetricInPointsArray(name: string, arr: ICalendarDataPoint[]): number {
      let index: number = -1;
      for (let i = 0; i < arr.length; i++) {
        if (name === arr[i].metric) {
          index = i;
        }
      }
      return index;
    }

    private GetSortMetricArray(): any[] {
      let metrics: ICalendarMetric[] = this.calendarMetrics.metrics.slice();
      let tooltips: string[] = this.tooltips.slice();
      //Remove metrics from the metrics array that are in the tooltips array
      for (let i = 0; i < tooltips.length; i++) {
        for (let j = 0; j < metrics.length; j++) {
          if (tooltips[i] === metrics[j].name) {
            metrics.splice(j, 1);
          }
        }
      }
      //All that's left after the deletion is added to the end of the tooltips array
      for (let i = 0; i < metrics.length; i++) {
        tooltips.push(metrics[i].name);
      }
      return tooltips;
    }

    /**
     * This function gets called for each of the objects defined in the capabilities files and allows you to select which of the
     * objects and properties you want to expose to the users in the property pane.
     *
     * Below is a code snippet for a case where you want to expose a single property called "linemonthCountor" from the object called "settings"
     * This object and property should be first defined in the capabilities.json file in the objects section.
     */
    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstanceEnumeration {
      const objectName: string = options.objectName;
      const objectEnumeration: VisualObjectInstance[] = [];

      switch (objectName) {
        case 'metricsSettings':
          for (const metric of this.calendarMetrics.metrics) {
            objectEnumeration.push({
              objectName: objectName,
              displayName: metric.name,
              properties: {
                metricColor: { solid: { color: metric.color } }
              },
              selector: ColorHelper.normalizeSelector((metric.selectionId as powerbi.visuals.ISelectionId).getSelector())
            });
          }
          break;
        case 'legendSettings':
          objectEnumeration.push({
            objectName: objectName,
            displayName: "Legend",
            properties: {
              show: this.settings.legendSettings.show,
              legendLabelColor: this.settings.legendSettings.legendLabelColor,
              legendLabelFontSize: this.settings.legendSettings.legendLabelFontSize,
              legendTitleShow: this.settings.legendSettings.legendTitleShow,
              legendTitleName: this.settings.legendSettings.legendTitleName
            },
            validValues: {
              legendLabelFontSize: {
                numberRange: {
                  min: 4,
                  max: 30
                }
              }
            },
            selector: null
          });
          break;
        case 'calendarSettings':
          if (this.settings.calendarSettings.calendarType === 0) {
            objectEnumeration.push({
              objectName: objectName,
              displayName: "Calendar",
              properties: {
                calendarType: this.settings.calendarSettings.calendarType,
                startDate: this.settings.calendarSettings.startDate,
                numOfMonths: this.settings.calendarSettings.numOfMonths,
                firstDay: this.settings.calendarSettings.firstDay,
                cellSize: this.settings.calendarSettings.cellSize,
                cellBorderColor: this.settings.calendarSettings.cellBorderColor,
                calendarHeaderColor: this.settings.calendarSettings.calendarHeaderColor,
                calendarHeaderTitleColor: this.settings.calendarSettings.calendarHeaderTitleColor,
                weekDayLabelsColor: this.settings.calendarSettings.weekDayLabelsColor,
                dayLabelsColor: this.settings.calendarSettings.dayLabelsColor
              },
              validValues: {
                cellSize: {
                  numberRange: {
                    min: 20,
                    max: 60
                  }
                },
                numOfMonths: {
                  numberRange: {
                    min: 1,
                    max: 60
                  }
                }
              },
              selector: null
            });
          }
          else {
            objectEnumeration.push({
              objectName: objectName,
              displayName: "Calendar",
              properties: {
                calendarType: this.settings.calendarSettings.calendarType,
                numOfPreviousMonths: this.settings.calendarSettings.numOfPreviousMonths,
                numOfFollowingMonths: this.settings.calendarSettings.numOfFollowingMonths,
                firstDay: this.settings.calendarSettings.firstDay,
                cellSize: this.settings.calendarSettings.cellSize,
                cellBorderColor: this.settings.calendarSettings.cellBorderColor,
                calendarHeaderColor: this.settings.calendarSettings.calendarHeaderColor,
                calendarHeaderTitleColor: this.settings.calendarSettings.calendarHeaderTitleColor,
                weekDayLabelsColor: this.settings.calendarSettings.weekDayLabelsColor,
                dayLabelsColor: this.settings.calendarSettings.dayLabelsColor
              },
              validValues: {
                cellSize: {
                  numberRange: {
                    min: 20,
                    max: 60
                  }
                },
                numOfPreviousMonths: {
                  numberRange: {
                    min: 0,
                    max: 60
                  }
                },
                numOfFollowingMonths: {
                  numberRange: {
                    min: 1,
                    max: 60
                  }
                }
              },
              selector: null
            });
          }
          break;
      }
      return objectEnumeration;
    }
  }
}
