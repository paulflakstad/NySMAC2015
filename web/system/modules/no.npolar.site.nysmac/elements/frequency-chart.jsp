<%-- 
    Document   : frequency-chart
    Description: Reads an Excel file with a specific setup and generates a 
                    Highcharts chart from it.
                    The Excel file's should be constructed like this: 
                    Row 1: Headings
                        A: What (e.g. "Radiometer" or "Land sureveying")
                        B: Who (e.g. "NPI" or "AWI")
                        C: Frequency range start (e.g. "400 MHz" or "24 GHz")
                        D: Frequency range stop (e.g. "406 MHz" or "24 GHz")
                        E: Time intervals (e.g. "camp[aign]" or "cont[inuous]")
                    Rows 2-N: Values
    Created on : Nov 23, 2016
    Author     : Paul-Inge Flakstad, Norwegian Polar Institute <flakstad at npolar.no>
--%>
<%@page import="java.io.*" %>
<%@page import="java.util.*" %>
<%@page import="no.npolar.util.*" %>
<%@page import="org.apache.poi.ss.usermodel.*" %>
<%@page import="org.apache.poi.xssf.usermodel.*" %>
<%@page import="org.opencms.jsp.*" %>
<%@page import="org.opencms.file.*" %>
<%@page import="org.opencms.main.*" %>
<%@page contentType="text/html" pageEncoding="UTF-8"%>
<%@page trimDirectiveWhitespaces="true" %>
<%!
/**
 * Returns the index of the first character that is not numeric or ".".
 * 
 * @param s  A string that starts with a number and is followed by letters.
 */
public int firstNonNumeric(String s) {
    char[] chars = s.toCharArray();
    for (int i = 0; i < chars.length; i++) {
        char c = s.charAt(i);
        if (c == '.') {
            continue;
        }
        try {
            Integer.parseInt(String.valueOf(c));
        } catch (Exception e) {
            return i;
        }
    }
    return -1;
}
/**
 * Converts a human-friendly formatted frequency (like "10 kHz") to its numeric 
 * counterpart (like 10000).
 * <p>
 * The Excel file uses human-friendly format, but Highcharts needs numeric.
 * 
 * @param kiloMegaOrGigaHertz  The human-friendly frequency string.
 * @return double  The numeric counterpart.
 */
public double toHertz(String kiloMegaOrGigaHertz) {
    int splitIndex = firstNonNumeric(kiloMegaOrGigaHertz);
    double num = Double.valueOf(kiloMegaOrGigaHertz.substring(0, splitIndex));
    String label = kiloMegaOrGigaHertz.substring(splitIndex).trim().toLowerCase();
    List<String> labels = Arrays.asList( new String[]{ "hz", "khz", "mhz", "ghz", "thz" } );
    return num * Math.pow(1000, labels.indexOf(label));
}
%>
<%
// We're gonna read the Excel file from the OpenCms VFS, so we need these
CmsAgent cms = new CmsAgent(pageContext, request, response);
CmsObject cmso = cms.getCmsObject();

// The parameter name used for the Excel file URI
final String PARAM_NAME_FILE_URI = "file";
// The parameter name used for chart ID suffixes (explained below)
final String PARAM_NAME_CHART_ID = "id";
// The name of the Excel file's property we'll use as the chart's title
final String CHART_TITLE_PROPERTY = "Title";
// The name of the Excel file's property we'll use as the chart's subtitle
final String CHART_SUBTITLE_PROPERTY = "Description";
// The chart's default title
final String CHART_TITLE_DEFAULT = "Frequency range";
// The chart's default subtitle
final String CHART_SUBTITLE_DEFAULT = "";

// Create the ID for this chart
// ============================
// This script may be included 1-N times on any given page, potentially 
// generating multiple charts.
// => Each chart needs an ID unique to that page.
// => This script does not inherently *know* how many times it is included.
// => We need a supporting mechanism that allows us to ensure unique IDs.
// => For any page that uses multiple charts, each include call should pass a 
//      parameter "id", with a value unique to that page. We'll use this as an 
//      a suffix to the chart ID, to (hopefully) make it unique.
// => For example, including with "&id=1", "&id=2", "&id=3" etc. would work fine.
String chartIdSuffix = request.getParameter(PARAM_NAME_CHART_ID);
if (chartIdSuffix == null || chartIdSuffix.isEmpty()) {
    chartIdSuffix = "1";
}
String hcId = "chart-"+chartIdSuffix;
/*
// This doesn't work:
int chartCount = 0;
try {
    chartCount = Integer.valueOf(
            (String)cms.getRequestContext().getAttribute("chartCount")
    );
    chartCount++;
} catch (Exception e) {
    cms.getRequestContext().setAttribute("chartCount", ++chartCount);
}
String hcId = "chart-"+chartCount;
//*/
%>
<div class="media">
    <div id="<%= hcId %>" style="min-width: 310px; height: 600px; margin: 0 auto"></div>
</div>
<%
// We expect a "file" parameter that tells us which Excel file to use
// e.g. "/practical/frequencies-ny-alesund-short-active-instruments.xlsx";
String filePath = request.getParameter(PARAM_NAME_FILE_URI);

StringBuilder html = new StringBuilder(1024);

StringBuilder chartCategories = new StringBuilder(128);
StringBuilder chartData = new StringBuilder(128);

String chartTitle = cmso.readPropertyObject(filePath, CHART_TITLE_PROPERTY, false).getValue(CHART_TITLE_DEFAULT);
String chartSubtitle = cmso.readPropertyObject(filePath, CHART_SUBTITLE_PROPERTY, false).getValue(CHART_SUBTITLE_DEFAULT);

try {
    ByteArrayInputStream inputStream = new ByteArrayInputStream(cmso.readFile(filePath).getContents());

    Workbook workbook = new XSSFWorkbook(inputStream);
    Sheet firstSheet = workbook.getSheetAt(0);
    Iterator<Row> iRows = firstSheet.iterator();

    html.append("<div class=\"toggleable collapsed\">");
    html.append("<a class=\"toggletrigger\" tabindex=\"0\" aria-controls=\"toggleable-0\">View table</a>");
    html.append("<div class=\"toggletarget\" id=\"toggleable-0\">");
    html.append("<table>");

    while (iRows.hasNext()) {
        html.append("<tr>");
        Row row = iRows.next();                
        Iterator<Cell> iCells = row.cellIterator();
        while (iCells.hasNext()) {
            Cell cell = iCells.next();
            String cellContent = cell.toString();
            try { 
                // Trimming the value is CRITICAL for the chart stuff below here
                // not to crash!
                cellContent = cellContent.trim(); 
            } catch (Exception e) {}

            // Print the table data
            String tableCellType = row.getRowNum() > 0 ? "td" : "th scope=\"col\"";
            html.append("<" + tableCellType + ">"
                            + cellContent
                        + "</" + tableCellType.substring(0,2) + ">");

            if (row.getRowNum() > 0) {
                
                switch (cell.getColumnIndex()) {
                    case 0:
                        chartCategories.append(
                                (chartCategories.length() > 0 ? "," : "") 
                                + "'" + cellContent
                        );
                        break;
                    case 1: 
                        chartCategories.append(" (" + cellContent + ")'");
                        break;
                    case 2:
                        chartData.append(
                                (chartData.length() > 0 ? "," : "")
                                + "[" + String.format("%.0f", toHertz(cellContent))
                        );
                        break;
                    case 3:
                        chartData.append(", " + String.format("%.0f", toHertz(cellContent)) + "]");
                        break;
                }
            }
            /*
            html.append("[" + cell.getColumnIndex() + "]");
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    html.append(cellContent);
                    //out.print(cell.getStringCellValue());
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    html.append(cell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    html.append(cell.getNumericCellValue());
                    break;
            }
            html.append(iCells.hasNext() ? " - " : "<br>\n");
            //*/
        }
        html.append("</tr>");
    }
    inputStream.close();

    html.append("</table>");
    html.append("</div>");
    html.append("</div>");
    
    out.println(html);

} catch (Exception e) {
    out.println("<!-- Error processing file '" + request.getParameter("file") + "' -->");
}

// The next bit of js defines the chart config, and hooks it to the ID of the 
// container div.
//
// Note:
//  - addChart(...), toHumanFreq(...) and renderCharts(...) reside in commons.js
//  - the main template, nysmac.jsp, invokes renderCharts(...)
//
// The implementention is like this so we can load Highcharts *only* when 
// necessary, and (if true) we load it *only once*.
%>
<script type="text/javascript">
addChart('<%= hcId %>', {
    chart: {
        type: 'columnrange',
        inverted: true,
        zoomType: 'y'
    },
    title: {
        text: '<%= chartTitle %>'
    },
    subtitle: {
        text: '<%= chartSubtitle %>'
    },
    xAxis: {
        categories: [
            <%= chartCategories.toString() %>
        ]
    },
    yAxis: {
        title: {
            text: 'Hertz ( Hz )'
        },
        type: 'logarithmic'
    },
    tooltip: {
        //valueSuffix: '°C'
        formatter: function() {
            return '<b>' + this.x + '<b><br/>Frequency:<br />' + toHumanFreq(this.point.low) +
                    (this.point.high > this.point.low ?
                    (' - ' + toHumanFreq(this.point.high)) : '');
        }
    },
    plotOptions: {
        columnrange: {
            dataLabels: {
                enabled: true,
                formatter: function() {
                    return toHumanFreq(this.y);
                }
            }
            ,minPointLength: 20
            //,allowPointSelect: true
        }
    },
    legend: {
        enabled: false
    },
    series: [{
        name: 'Frequencies',
        data: [
            <%= chartData.toString() %>
        ]
    }]
});
</script>