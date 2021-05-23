# ICBC Merge Excel Template Report to Web assignment

## About
Read Excel template report, merge with report data from an XML file and dipslay on Web, ideally display with the styles in Excel

## Technology Stack
Excel, .NET Core 3.1, Razor, NPOI library to read Excel file, Bootstrap, JQuery

## Description
Import Excel report template, merge with report results from XML file, and display report results on a razor web page with format similar to the Excel template

## Backlog
1. UI dropdown for selecting the sheet index/name
2. Add more tests

## Sample XML file with report data
```xml
<?xml version="1.0" encoding="UTF-16"?>
<Reports>
    <Report>
        <Name>F 20.04</Name>
        <ReportVal>
            <ReportRow>10</ReportRow>
            <ReportCol>10</ReportCol>
            <Val>100</Val>
        </ReportVal>
        <ReportVal>
            <ReportRow>10</ReportRow>
            <ReportCol>11</ReportCol>
            <Val>200</Val>
        </ReportVal>
        <ReportVal>
            <ReportRow>10</ReportRow>
            <ReportCol>12</ReportCol>
            <Val>0</Val>
        </ReportVal>
        <ReportVal>
            <ReportRow>20</ReportRow>
            <ReportCol>10</ReportCol>
            <Val>600</Val>
        </ReportVal>
        <ReportVal>
            <ReportRow>20</ReportRow>
            <ReportCol>11</ReportCol>
            <Val>500</Val>
        </ReportVal>
        <ReportVal>
            <ReportRow>20</ReportRow>
            <ReportCol>12</ReportCol>
            <Val>0</Val>
        </ReportVal>
    </Report>
</Reports>
```
