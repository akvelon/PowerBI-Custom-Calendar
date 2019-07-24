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

import * as d3 from "d3";
import powerbiVisualsApi from "powerbi-visuals-api";

import IColorPalette = powerbiVisualsApi.extensibility.IColorPalette;
import ISelectionId = powerbiVisualsApi.visuals.ISelectionId;
import ISelectionManager = powerbiVisualsApi.extensibility.ISelectionManager;
import IViewport = powerbiVisualsApi.IViewport;
import IVisual = powerbiVisualsApi.extensibility.visual.IVisual;
import IVisualEventService = powerbiVisualsApi.extensibility.IVisualEventService;
import IVisualHost = powerbiVisualsApi.extensibility.visual.IVisualHost;
import DataView = powerbiVisualsApi.DataView;
import DataViewCategorical = powerbiVisualsApi.DataViewCategorical;
import DataViewCategoryColumn = powerbiVisualsApi.DataViewCategoryColumn;
import DataViewMetadataColumn = powerbiVisualsApi.DataViewMetadataColumn;
import DataViewObject = powerbiVisualsApi.DataViewObject;
import DataViewObjects = powerbiVisualsApi.DataViewObjects;
import DataViewValueColumn = powerbiVisualsApi.DataViewValueColumn;
import EnumerateVisualObjectInstancesOptions = powerbiVisualsApi.EnumerateVisualObjectInstancesOptions;
import PrimitiveValue = powerbiVisualsApi.PrimitiveValue;
import VisualConstructorOptions = powerbiVisualsApi.extensibility.visual.VisualConstructorOptions;
import VisualObjectInstance = powerbiVisualsApi.VisualObjectInstance;
import VisualObjectInstanceEnumeration = powerbiVisualsApi.VisualObjectInstanceEnumeration;
import VisualTooltipDataItem = powerbiVisualsApi.extensibility.VisualTooltipDataItem;
import VisualUpdateOptions = powerbiVisualsApi.extensibility.visual.VisualUpdateOptions;

import {
    Axis,
    axisBottom
} from "d3-axis";

import {
    ScaleBand,
    scaleBand
} from "d3-scale";

import {
    legendInterfaces,
    legend,
    legendData
} from "powerbi-visuals-utils-chartutils";

import { ColorHelper } from "powerbi-visuals-utils-colorutils";

import { valueFormatter as ValueFormatter } from "powerbi-visuals-utils-formattingutils";

import {
    ITooltipServiceWrapper,
    TooltipEventArgs,
    createTooltipServiceWrapper
} from "powerbi-visuals-utils-tooltiputils";

import ILegend = legendInterfaces.ILegend;
import LegendData = legendInterfaces.LegendData;
import LegendPosition = legendInterfaces.LegendPosition;
import createLegend = legend.createLegend;
import legendProps = legendInterfaces.legendProps;

import valueFormatter = ValueFormatter.valueFormatter;

/* tslint:disable:no-relative-imports */
import "../style/visual.less";
import { CalendarSettings } from "./settings";
/* tslint:enable:no-relative-imports */

export type Selection = d3.Selection<any, any, any, any>;

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
    selectionId: ISelectionId;
    index: number;
    metadataColumn: DataViewMetadataColumn;
}

interface Metrics {
    metrics: ICalendarMetric[];
}

interface ICalendarMetric {
    name: string;
    color: string;
    selectionId: ISelectionId;
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
    private events: IVisualEventService;

    private settings: CalendarSettings;
    private calendarMetrics: Metrics;

    private rootElement: Selection;
    private visibleGroupContainer: Selection;
    private monthContainer: Selection;
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
        this.events = options.host.eventService;
        this.selectionManager = options.host.createSelectionManager();
        this.selectionManager.registerOnSelectCallback((ids: any[]) => {
            if (ids.length === 0) return;
            let cells: HTMLElement[] = CustomCalendar.getCellBySelectionIds(ids);
            CustomCalendar.clearSelectedCells();
            CustomCalendar.selectCell(d3.select(cells[0]));

            if (cells.length > 1) {
                for (let i = 1; i < cells.length; i++) {
                    CustomCalendar.selectCell(d3.select(cells[i]), true);
                }
            }
        });

        const visual: Selection = d3.select(options.element);
        const selectionManagerField: ISelectionManager = this.selectionManager;

        visual.on("click", () => {
            if ((<MouseEvent>d3.event).toElement.id === "") {
                if (CustomCalendar.selectedCell.length === 0) return;
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

        const svg: Selection = this.rootElement = d3.select(element)
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

            ckey = "a" + ckey.replace(/\//g, "_");

            if (ckey !== "a" && cKeysArray.indexOf(ckey) === -1) {
                cKeysArray.push(ckey);
            }
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
        if (CustomCalendar.selectedCell.length === 0) return;
        else {
            for (let i = 0; i < CustomCalendar.selectedCell.length; i++) {
                let cell = document.getElementById(CustomCalendar.selectedCell[i]);
                d3.select(cell).attr("fill", "white");
            }
            CustomCalendar.selectedCell = [];
        }
    }

    public update(options: VisualUpdateOptions) {
        this.events.renderingStarted(options);
        this.visibleGroupContainer.selectAll(".month").remove();
        this.dataPoints = [];
        this.calendarViewModel = this.visualTransform(options, this.host);

        const selectionManagerLegend: ISelectionManager = this.selectionManager;
        if (selectionManagerLegend.getSelectionIds().length === 0) {
            CustomCalendar.clearSelectedCells();
        }

        const width: number = options.viewport.width;
        const height: number = options.viewport.height;

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

        const legends: Selection = d3.selectAll(".legendItem");
        const calendarMetrics: ICalendarMetric[] = this.calendarMetrics.metrics;

        legends.on("click", (d, i, legend) => {
            selectionManagerLegend.select(d.identity).then((ids: ISelectionId[]) => {
                const selectedItemsNumber: number = selectionManagerLegend.getSelectionIds().length;

                legends.attr('fill-opacity', ids.length > 0 ? '0.5' : '1');
                d3.select(legend[i]).attr("fill-opacity", 1);

                for (const metric of calendarMetrics) {
                    const fillOpacityValue: number = metric.name !== d.label && selectedItemsNumber > 0 ? 0.3 : 1;

                    d3.selectAll("." + metric.name.replace(" ", ""))
                        .attr("fill-opacity", fillOpacityValue);
                }
            });
        });

        this.events.renderingFinished(options);
    }

    private visualTransform(options: VisualUpdateOptions, host: IVisualHost): ICalendarViewModel {
        let dataViews: DataView[] = options.dataViews;
        const viewModel: ICalendarViewModel = {
            settings: <CalendarSettings>{},
            dataPoints: <ICalendarDataPoint[]>[]
        };
        this.settings = CustomCalendar.parseSettings(options.dataViews[0]);
        this.setStartDate();

        if (!dataViews
            || !dataViews[0]
            || !dataViews[0].categorical
            || !dataViews[0].categorical.values
            || !dataViews[0].categorical.categories
            || !dataViews[0].categorical.categories[0].source) {

            return viewModel;
        }

        let index = dataViews[0].categorical.values.length;
        dataViews[0].categorical.values.forEach(value => {
            value.source.index = index;
            index--;
        });

        const categorical: DataViewCategorical = dataViews[0].categorical;
        const dataPoints = this.getDataPoints(categorical, host);
        CustomCalendar.sortDataPoints(dataPoints);

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
                metric: dataPoints.data[j].metric,
                hours: dataPoints.data[j].hours,
                color: dataPoints.data[j].color,
                id: "#a" + dataPoints.data[j].date,
                selectionId: host.createSelectionIdBuilder()
                    .withCategory(dataViews[0].categorical.categories[0], datesArr.indexOf(dataPoints.data[j].date))
                    .withMeasure(dataPoints.data[j].date)
                    .createSelectionId(),
                index: dataPoints.data[j].index,
                metadataColumn: dataPoints.data[j].metadataColumn
            });
        }

        let tooltipColumns: string[] = [];
        for (let i = 0; i < categorical.values.length; i++) {
            if (categorical.values[i].source.roles['tooltips']) {
                tooltipColumns.push(categorical.values[i].source.displayName);
            }
        }
        this.tooltips = tooltipColumns.sort();

        return {
            settings: this.settings,
            dataPoints: this.dataPoints
        };
    }

    private getDataPoints(categorical, host): any {
        const category: DataViewCategoryColumn = categorical.categories[0];
        const dataPoints = {
            data: []
        };
        const calendarMetrics: ICalendarMetric[] = [];

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
                this.colorPalette.getColor(groupedValue.source.displayName).value
            );

            for (let j = 0, len = Math.max(category.values.length, groupedValue.values.length); j < len; j++) {
                if (category.values[j] !== null && groupedValue.values[j] !== null) {
                    const categoryFormatString = valueFormatter.getFormatStringByColumn(category.source);
                    this.categoryFormatString = categoryFormatString;
                    const date: Date = <Date>category.values[j];
                    const stringDate: string = CustomCalendar.convertDateToString(date);
                    const value = valueFormatter.format(category.values[j], categoryFormatString);
                    const hours: any = <number>groupedValue.values[j];
                    const metric: string = groupedValue.source.displayName;
                    const index: number = groupedValue.source.index;
                    const metadataColumn: DataViewMetadataColumn = groupedValue.source;

                    if (hours !== null && hours !== undefined) {
                        dataPoints.data.push({
                            date: stringDate,
                            defaultDate: value,
                            hours: hours,
                            color: color,
                            metric: metric,
                            index: index,
                            metadataColumn: metadataColumn
                        });
                    }
                }
            }

            const selectionId: ISelectionId = host.createSelectionIdBuilder()
                .withSeries(categorical.values, groupedValue)
                .withMeasure(groupedValue.source.queryName)
                .createSelectionId();

            calendarMetrics.push({
                color,
                name: <string>groupedValue.source.displayName,
                selectionId: selectionId
            });
        }

        this.calendarMetrics = {
            metrics: calendarMetrics
        };

        return dataPoints;
    }

    private static sortDataPoints(dataPoints) {
        dataPoints.data.sort((second, first) => {
            if (first.date === null) {
                return 1;
            } else if (second.date === null) {
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
    }

    private static getTooltip(value: any): VisualTooltipDataItem[] {
        if (value) {
            return value;
        }
    }

    private static parseSettings(dataView: DataView): CalendarSettings {
        return CalendarSettings.parse(dataView);
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

        const legendDataToRender: LegendData = {
            fontSize: legendLabelFontSize,
            dataPoints: [],
            title: legendTitle,
            labelColor: legendLabelColor
        };

        for (let i = 0; i < this.calendarMetrics.metrics.length; i++) {
            const metric: ICalendarMetric = this.calendarMetrics.metrics[i];

            legendDataToRender.dataPoints.push({
                label: metric.name,
                color: metric.color,
                selected: false,
                identity: metric.selectionId
            });
        }

        this.legend.changeOrientation(LegendPosition.Top);

        const legendObjectProperties: DataViewObject = this.legendObjectProperties;

        if (legendObjectProperties) {
            legendData.update(legendDataToRender, legendObjectProperties);

            const position: string = <string>legendObjectProperties[legendProps.position];
            if (position) {
                this.legend.changeOrientation(LegendPosition[position]);
            }
        }

        this.legend.drawLegend(legendDataToRender, this.currentViewport);
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

            this.monthContainer = this.visibleGroupContainer.append("svg")
                .attr("class", "month")
                .attr("width", monthSize + 15)
                .attr("height", monthSize + cellSize + 15)
                .append("g");

            const header: Selection = this.monthContainer.append("svg")
                .attr("class", "monthHeader")
                .attr("width", monthSize)
                .attr("height", cellSize);

            header.append("rect")
                .attr("x", 1)
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

        if (calendarType === 0) {
            const startDateSettings: string = this.settings.calendarSettings.startDate;
            this.startDate = CustomCalendar.getValidateDate(startDateSettings);
        } else {
            const numOfPreviousMonths: number = this.settings.calendarSettings.numOfPreviousMonths;
            let startDate = new Date();
            startDate.setMonth(startDate.getMonth() - numOfPreviousMonths);
            this.startDate = startDate;
        }

        this.startDate.setDate(1);
    }

    private static getValidateDate(date: string): Date {
        if ((date.search(/^((0?[1-9]|1[012])[- /.](0?[1-9]|[12][0-9]|3[01])[- /.](19|20)[0-9]{2})*$/) === 0)
            && (date !== "")) {
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
        const monthCells: Selection[] = [];
        const cellsContainer: Selection = this.monthContainer.append("svg")
            .attr("class", "cellsContainer")
            .attr("width", monthSize + 10)
            .attr("height", monthSize + cellSize + 10);

        let numOfDays: number;
        let firstDay: number;
        let newFirstDay: number;
        let count: number = 0;
        let dayNumber: number;
        let cellRowNumber: number;

        if (monthCount === 0) {
            const startDate: Date = this.getStartDate();
            firstDay = startDate.getDay();
            cellRowNumber = this.getCountWeek(startDate);
            dayNumber = startDate.getDate();
        }
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
                    const cell: Selection = cellsContainer.append("rect")
                        .attr("x", currentX)
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

                    const label: Selection = cellsContainer.append("text").text(dayNumber)
                        .attr("x", currentX + (cellSize / 2))
                        .attr("y", currentY + (cellSize / 3))
                        .attr("id", id)
                        .attr("class", "cell_label")
                        .attr("text-anchor", "middle")
                        .attr("font-size", cellSize / 4)
                        .attr("fill", dayLabelsColor);

                    monthCells.push(label['_groups']);

                    const currentDate: Date = new Date(year, monthIndex, dayNumber);
                    if (this.isDayWithMetric(currentDate)) {
                        cell.data(this.getTooltipsDataCell(this.getMetricByDate(currentDate), currentDate));
                        this.tooltipServiceWrapper.addTooltip(
                            cell,
                            (tooltipEvent: TooltipEventArgs<number>) => CustomCalendar.getTooltip(tooltipEvent.data),
                            null);

                        label.data(this.getTooltipsDataCell(this.getMetricByDate(currentDate), currentDate));
                        this.tooltipServiceWrapper.addTooltip(
                            label,
                            (tooltipEvent: TooltipEventArgs<number>) => CustomCalendar.getTooltip(tooltipEvent.data),
                            null);
                    }

                    dayNumber++;
                }

                count++;
            }
        }

        this.drawMetrics(cellsContainer, monthIndex + 1, year, monthCells, monthCount);
    }

    private drawMetrics(cellsContainer: Selection, monthNumber: number,
                        year: number, monthCells: any, monthCount: number): void {
        let sameMonthDataPoints: ICalendarDataPoint[] = [];
        let dataPoints: ICalendarDataPoint[] = this.dataPoints;
        let index: number = 0;

        for (let i = 0; i < dataPoints.length; i++) {
            if (dataPoints[i].date !== null) {
                const dataPoint: ICalendarDataPoint = dataPoints[i];
                const renderedDate: Date = new Date(dataPoint.date);
                const renderedDateString: string = CustomCalendar.convertDateToString(renderedDate);
                const month: string = String(monthNumber) + "/";

                if ((renderedDateString.indexOf(month) === 0) && (renderedDateString.indexOf(String(year)) > -1)) {
                    sameMonthDataPoints[index] = dataPoint;
                    index++;
                }
            }
        }

        if (monthCount === 0) {
            const startDay: number = this.getStartDate().getDate();
            for (let i = 0; i < sameMonthDataPoints.length; i++) {
                const currentDay: number = CustomCalendar.getDayFromDataStr(sameMonthDataPoints[i].date);
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

                let dataPointId: string;
                for (let j = 0; j < dataPointsToDraw.length; j++) {
                    const date: string = CustomCalendar.convertDateToString(new Date(dataPointsToDraw[j].date))
                        .replace("/", "_");
                    const dateFormat: string = date.replace("/", "_");
                    dataPointId = "#a" + dateFormat;

                    dataPointsToDraw[j].id = dataPointId;
                }

                dataPointsToDraw.sort((first, second) => {
                    if (first.index < second.index) {
                        return -1;
                    } else {
                        return 1;
                    }
                });

                this.drawCellMetrics(self, cellsContainer, monthCells, dataPointId, dataPointsToDraw);

                usedDates[usedDatesIndex] = sameMonthDataPoints[i].date;
                usedDatesIndex++;
            }
        }
    }

    private drawCellMetrics(self, cellsContainer, monthCells, dataPointId, dataPointsToDraw): void {
        const cellSize: number = <number>this.settings.calendarSettings.cellSize;
        let sortedMetrics: ICalendarDataPoint[] = [];
        for (let k = 0; k < dataPointsToDraw.length; k++) {
            if (dataPointsToDraw[k].metadataColumn.roles['metrics']) {
                sortedMetrics.push(dataPointsToDraw[k]);
            }
        }

        let previousYCoord: number = cellSize;
        for (let j = 0; j < sortedMetrics.length; j++) {
            if (sortedMetrics[j].hours === 0) {
                for (let i = 0; i < monthCells.length; i++) {
                    const monthCell: any = monthCells[i][0][0];
                    const cell: Selection = d3.select(monthCell.previousSibling);
                    const label: Selection = d3.select(monthCell);
                    const id: string = "#" + cell.attr("id");

                    if (id === sortedMetrics[j].id) {
                        cell.on("click", () => {
                            self.select(cell, [sortedMetrics[0]]);
                        });
                        label.on("click", () => {
                            self.select(cell, [sortedMetrics[0]]);
                        });
                    }
                }
            } else {
                const metricFactor: number = CustomCalendar.getMetricHeight(sortedMetrics, j);
                const width: number = cellSize - 2;
                const height: number = (cellSize - (cellSize / 2.3)) * metricFactor;
                const xCoord: number = Number(d3.select(dataPointId).attr("x")) + 1;
                const yCoord: number = Number(d3.select(dataPointId).attr("y")) + previousYCoord - height - 1;
                const metricFormat: string = sortedMetrics[j].metric.replace(" ", "");

                const cellId = dataPointId.replace('#', '');
                const cellMetrics: Selection = cellsContainer.append("rect")
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
                    const cell: Selection = d3.select(monthCell.previousSibling);
                    const label: Selection = d3.select(monthCell);
                    const id: string = "#" + cell.attr("id");

                    if (id === sortedMetrics[j].id) {
                        cell.on("click", () => {
                            self.select(cell, [sortedMetrics[0]]);
                        });
                        label.on("click", () => {
                            self.select(cell, [sortedMetrics[0]]);
                        });
                        cellMetrics.on("click", () => {
                            self.select(cell, [sortedMetrics[j]]);
                        });
                    }
                }
                cellMetrics.data(this.getTooltipsDataMetric(dataPointsToDraw, sortedMetrics[j]));
                this.tooltipServiceWrapper.addTooltip(
                    cellMetrics,
                    (tooltipEvent: TooltipEventArgs<number>) => CustomCalendar.getTooltip(tooltipEvent.data),
                    null);

                previousYCoord = previousYCoord - height;
            }
        }
    }

    private getTooltipsDataCell(allPoints: ICalendarDataPoint[], currentDate: Date): ITooltipDataPoint[] {
        const dataPoints: ICalendarDataPoint[] = CustomCalendar.getDataPointsWithoutZeroMetric(allPoints);
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
                tooltipData.color = this.getColorByMetricName(tooltips[i]);
                tooltipData.displayName = tooltips[i];
                indexInPoints = CustomCalendar.getIndexMetricInPointsArray(tooltips[i], dataPoints);

                if (indexInPoints >= 0) {
                    tooltipData.value = CustomCalendar.tooltipValue(
                        dataPoints[indexInPoints].metadataColumn,
                        dataPoints[indexInPoints].hours
                    );
                }
                currentTooltipDataPoint.push(tooltipData);
            }
        }

        resultTooltipData.push(currentTooltipDataPoint);
        return resultTooltipData;
    }

    private getTooltipsDataMetric(allPoints: ICalendarDataPoint[], point: ICalendarDataPoint): ITooltipDataPoint[] {
        const dataPoints: ICalendarDataPoint[] = CustomCalendar.getDataPointsWithoutZeroMetric(allPoints);
        let currentTooltipDataPoint: ITooltipDataPoint[] = [];
        let tooltips: string[] = this.tooltips.slice();
        let resultTooltipData: any[] = [];
        let tooltipData: any = {};

        tooltipData.color = point.color;
        tooltipData.displayName = point.metric;
        tooltipData.header = point.defaultDate;
        tooltipData.value = point.hours.toString();
        currentTooltipDataPoint.push(tooltipData);

        if (tooltips.indexOf(point.metric) >= 0) {
            tooltips.splice(tooltips.indexOf(point.metric), 1);
        }

        for (let i = 0; i < tooltips.length; i++) {
            let tooltipData: any = {};
            tooltipData.header = point.defaultDate;
            tooltipData.color = this.getColorByMetricName(tooltips[i]);
            tooltipData.displayName = tooltips[i];

            let indexInPoints: number = CustomCalendar.getIndexMetricInPointsArray(tooltips[i], dataPoints);
            if (indexInPoints >= 0) {
                tooltipData.value = CustomCalendar.tooltipValue(
                    dataPoints[indexInPoints].metadataColumn,
                    dataPoints[indexInPoints].hours
                );
            }
            currentTooltipDataPoint.push(tooltipData);
        }

        currentTooltipDataPoint.sort((first, second) => {
            let sortMetricArr: string[] = this.getSortMetricArray();
            if (CustomCalendar.getIndexMetricInArray(first.displayName, sortMetricArr) <
                CustomCalendar.getIndexMetricInArray(second.displayName, sortMetricArr)) {
                return -1;
            } else {
                return 1;
            }
        });

        resultTooltipData.push(currentTooltipDataPoint);
        return resultTooltipData;
    }

    private static tooltipValue(metadataColumn: DataViewMetadataColumn, value: PrimitiveValue): any {
        return CustomCalendar.getFormattedValue(metadataColumn, value);
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

    private static getDataPointsWithoutZeroMetric(allPoints: ICalendarDataPoint[]): ICalendarDataPoint[] {
        let dataPoints: ICalendarDataPoint[] = allPoints.slice();
        let i = 0;
        while (i < dataPoints.length) {
            if (dataPoints[i].hours === 0) {
                dataPoints.splice(i, 1);
            } else {
                i++;
            }
        }
        return dataPoints;
    }

    private select(cell: Selection, sortedDataPoints: ICalendarDataPoint[]) {
        if (CustomCalendar.selectCell(cell)) {
            this.selectMetrics(sortedDataPoints);
        } else {
            this.selectionManager.clear();
        }
    }

    private static selectCell(cell: Selection, check: boolean = false): boolean {
        let multipleSelection: boolean = check || ((<MouseEvent>d3.event) ? (<MouseEvent>d3.event).ctrlKey : false);
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
                } else {
                    CustomCalendar.isMultipleSelected = false;

                    return false;
                }
            }
            CustomCalendar.isMultipleSelected = false;
        }

        return true;
    }

    private selectMetrics(dataPoints: ICalendarDataPoint[]) {
        if (!CustomCalendar.isMultipleSelected) {
            this.selectionManager.clear();
        }
        let multipleSelection = CustomCalendar.isMultipleSelected;

        for (let i = 0; i < dataPoints.length; i++) {
            if (i !== 0 && dataPoints.length > 1) {
                multipleSelection = true;
            }
            if (dataPoints[i].selectionId) {
                this.selectionManager.select(dataPoints[i].selectionId, multipleSelection);
            } else {
                this.selectionManager.select(dataPoints[i], multipleSelection);
            }
        }
    }

    private static getMetricHeight(dataPoints: ICalendarDataPoint[], dataPointNumber): number {
        const metric: ICalendarDataPoint = dataPoints[dataPointNumber];
        const metricsSum: number = CustomCalendar.getMetricsSum(dataPoints, metric.id);

        return CustomCalendar.isNumeric(metric.hours) ? (metric.hours / metricsSum) : 0;
    }

    private static isNumeric(value: any): boolean {
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

    private isDayWithMetric(date: Date): boolean {
        const dataPoints: ICalendarDataPoint[] = this.calendarViewModel.dataPoints;
        let currentDate: Date;
        for (let i = 0; i < dataPoints.length; i++) {
            currentDate = new Date(dataPoints[i].date);
            if ((date.getFullYear() === currentDate.getFullYear())
                && (date.getMonth() === currentDate.getMonth())
                && (date.getDate() === currentDate.getDate())) {
                return true;
            }
        }
        return false;
    }

    private static convertDateToString(date: Date): string {
        if (date === null) {
            return null;
        }

        return (date.getMonth() + 1) + "/"
            + (date.getDate()) + "/"
            + (date.getFullYear());
    }

    private drawDaysLabels(monthSize: number): void {
        const cellSize: number = this.settings.calendarSettings.cellSize;
        const weekDayLabelsColor: string = this.settings.calendarSettings.weekDayLabelsColor;
        const newWeek: string[] = this.getWeekDays();
        const dayLabelSize: number = Math.ceil(this.settings.calendarSettings.cellSize / 3);

        const xScale: ScaleBand<string> = scaleBand()
            .domain(newWeek)
            .rangeRound([0, monthSize]);

        const xAxis: Axis<any> = axisBottom(xScale).tickSize(0);

        const lineContainer: Selection = this.monthContainer.append("svg")
            .attr("class", "weekDays")
            .attr("width", "100%")
            .attr("height", "100%")
            .attr("x", 2)
            .attr("y", cellSize * 1.1);

        lineContainer.append("g")
            .attr("class", "line")
            .attr("transform", "translate(0," + (cellSize / 10) + ")")
            .call(xAxis)
            .selectAll("text")
            .attr("fill", weekDayLabelsColor);

        d3.selectAll(".line")
            .attr("font-family", "wf_standard-font,helvetica,arial,sans-serif")
            .attr("font-size", dayLabelSize + "px")
            .attr("text-anchor", "middle");

        d3.select(".domain").remove();
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

        if (firstDay > dayInDate) {
            countShowDayFirstWeek = firstDay - dayInDate;
        } else if (firstDay <= dayInDate) {
            countShowDayFirstWeek = 7 - (dayInDate - firstDay);
        }

        countWeek = 1 + Math.ceil((countShowDayMount - countShowDayFirstWeek) / 7);

        return countWeek;
    }

    private static getDayFromDataStr(dataStr: string): number {
        let dayStr: string;
        dayStr = dataStr.slice(0, -5);
        dayStr = dayStr.slice(dayStr.indexOf("/") + 1);
        return Number(dayStr);
    }

    private getColorByMetricName(name: string): string {
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

    private getMetricByDate(date: Date): any[] {
        const dataPoints: ICalendarDataPoint[] = this.dataPoints;
        let currentDataPoints: ICalendarDataPoint[];
        let isIndexBeginInitialized: boolean = false;
        let indexBegin: number = 0;
        let indexEnd: number = 0;
        let currentDate: Date;

        for (let i = 0; i < dataPoints.length; i++) {
            currentDate = new Date(dataPoints[i].date);
            if ((currentDate.getFullYear() === date.getFullYear())
                && (currentDate.getMonth() === date.getMonth())
                && (currentDate.getDate() === date.getDate())) {
                if (!isIndexBeginInitialized) {
                    indexBegin = i;
                    isIndexBeginInitialized = true;
                }
                indexEnd = i;
            }
        }

        if (dataPoints.length === indexEnd + 1) {
            currentDataPoints = dataPoints.slice(indexBegin);
        } else {
            currentDataPoints = dataPoints.slice(indexBegin, indexEnd + 1);
        }

        return currentDataPoints;
    }

    private static getIndexMetricInArray(name: string, arr: any[]): number {
        let index: number = -1;
        for (let i = 0; i < arr.length; i++) {
            if (name === arr[i]) {
                index = i;
            }
        }
        return index;
    }

    private static getIndexMetricInPointsArray(name: string, arr: ICalendarDataPoint[]): number {
        let index: number = -1;
        for (let i = 0; i < arr.length; i++) {
            if (name === arr[i].metric) {
                index = i;
            }
        }
        return index;
    }

    private getSortMetricArray(): any[] {
        let metrics: ICalendarMetric[] = this.calendarMetrics.metrics.slice();
        let tooltips: string[] = this.tooltips.slice();

        for (let i = 0; i < tooltips.length; i++) {
            for (let j = 0; j < metrics.length; j++) {
                if (tooltips[i] === metrics[j].name) {
                    metrics.splice(j, 1);
                }
            }
        }

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
        let objectEnumeration: VisualObjectInstance[] = [];

        switch (objectName) {
            case 'metricsSettings':
                objectEnumeration = this.getMetricsSettings(objectName);
                break;
            case 'legendSettings':
                objectEnumeration = this.getLegendSettings(objectName);
                break;
            case 'calendarSettings':
                objectEnumeration = this.getCalendarSettings(objectName);
                break;
        }
        return objectEnumeration;
    }

    private getMetricsSettings(objectName) {
        const metricSettings = [];
        for (const metric of this.calendarMetrics.metrics) {
            metricSettings.push({
                objectName: objectName,
                displayName: metric.name,
                properties: {
                    metricColor: {solid: {color: metric.color}}
                },
                selector: ColorHelper.normalizeSelector((<ISelectionId>metric.selectionId).getSelector())
            });
        }

        return metricSettings;
    }

    private getLegendSettings(objectName) {
        const legendSettings = [];
        legendSettings.push({
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

        return legendSettings;
    }

    private getCalendarSettings(objectName) {
        const calendarSettings = [];
        if (this.settings.calendarSettings.calendarType === 0) {
            calendarSettings.push({
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
        } else {
            calendarSettings.push({
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

        return calendarSettings;
    }
}
