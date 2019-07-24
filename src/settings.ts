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

import { dataViewObjectsParser } from "powerbi-visuals-utils-dataviewutils";
import DataViewObjectsParser = dataViewObjectsParser.DataViewObjectsParser;

export class CalendarAppearance {
    public startDate: string = this.getDefaultStartDate();
    public numOfMonths: number = 12;
    public calendarType: number = 0;
    public numOfPreviousMonths: number = 0;
    public numOfFollowingMonths: number = 5;
    public cellSize: number = 50;
    public cellBorderColor: string = "black";
    public calendarHeaderColor: string = "black";
    public calendarHeaderTitleColor: string = "white";
    public firstDay: number = 0;
    public weekDayLabelsColor: string = "black";
    public dayLabelsColor: string = "black";

    private getDefaultStartDate(): string {
        let currentDate = new Date();

        return (currentDate.getMonth() + 1) + "/"
            + (currentDate.getDate()) + "/"
            + (currentDate.getFullYear());
    }
}

export class MetricsAppearance {
    public metricColor: string = "";
}

export class LegendAppearance {
    public show: boolean = false;
    public legendLabelColor: string = "black";
    public legendLabelFontSize: number = 10;
    public legendTitleShow: boolean = false;
    public legendTitleName: string = "Metrics";
}

export class CalendarSettings extends DataViewObjectsParser {
    public calendarSettings: CalendarAppearance = new CalendarAppearance();
    public metricsSettings: MetricsAppearance = new MetricsAppearance();
    public legendSettings: LegendAppearance = new LegendAppearance();
}
