using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocFlow.Core.Helpers
{
    internal static class ExcelChartHelper
    {
        internal static void InjectChart(MemoryStream xlsx, Models.ChartDefinition chart, int totalDataRows, string sheetName)
        {
            xlsx.Position = 0;
            using (var doc = SpreadsheetDocument.Open(xlsx, true))
            {
                var wbPart = doc.WorkbookPart;
                WorksheetPart wsPart = null;
                foreach (var ws in wbPart.WorksheetParts)
                {
                    wsPart = ws;
                    break;
                }
                if (wsPart == null) return;

                var drawingPart = wsPart.AddNewPart<DrawingsPart>();
                var chartPart  = drawingPart.AddNewPart<ChartPart>();

                string chartXml   = BuildChartXml(chart, totalDataRows, sheetName);
                string drawingXml = BuildDrawingXml(chart, drawingPart.GetIdOfPart(chartPart));

                using (var s = new MemoryStream(Encoding.UTF8.GetBytes(chartXml)))
                    chartPart.FeedData(s);

                using (var s = new MemoryStream(Encoding.UTF8.GetBytes(drawingXml)))
                    drawingPart.FeedData(s);

                string drawingRelId = wsPart.GetIdOfPart(drawingPart);
                wsPart.Worksheet.Append(new Drawing { Id = drawingRelId });
                wsPart.Worksheet.Save();
            }
        }

        private static string BuildChartXml(Models.ChartDefinition chart, int totalDataRows, string sheetName)
        {
            int lastRow = chart.DataStartRow + totalDataRows - chart.DataStartRow;
            int lastCol = chart.LastDataColumn < 0 ? chart.FirstDataColumn : chart.LastDataColumn;
            string labelRange = $"{sheetName}!${Col(chart.LabelColumn)}${chart.DataStartRow}:${Col(chart.LabelColumn)}${lastRow}";

            var series = new StringBuilder();
            for (int col = chart.FirstDataColumn; col <= lastCol; col++)
            {
                int idx = col - chart.FirstDataColumn;
                string hdr  = $"{sheetName}!${Col(col)}${chart.HeaderRow}";
                string vals = $"{sheetName}!${Col(col)}${chart.DataStartRow}:${Col(col)}${lastRow}";
                series.Append(BuildSeriesXml(idx, hdr, labelRange, vals));
            }

            string plotBody;
            switch (chart.ChartType)
            {
                case Models.ChartType.Line:
                    plotBody = $"<c:lineChart><c:grouping val=\"standard\"/>{series}</c:lineChart>";
                    break;
                case Models.ChartType.Pie:
                    plotBody = $"<c:pieChart>{series}</c:pieChart>";
                    break;
                default:
                    plotBody = $"<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>{series}</c:barChart>";
                    break;
            }

            return $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<c:chartSpace xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
  <c:lang val=""en-US""/>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{Xml(chart.Title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=""0""/></c:title>
    <c:autoTitleDeleted val=""0""/>
    <c:plotArea>{plotBody}</c:plotArea>
    <c:legend><c:legendPos val=""b""/></c:legend>
    <c:plotVisOnly val=""1""/>
  </c:chart>
</c:chartSpace>";
        }

        private static string BuildSeriesXml(int idx, string headerCell, string labelRange, string dataRange)
        {
            return $@"<c:ser>
  <c:idx val=""{idx}""/><c:order val=""{idx}""/>
  <c:tx><c:strRef><c:f>{Xml(headerCell)}</c:f></c:strRef></c:tx>
  <c:cat><c:strRef><c:f>{Xml(labelRange)}</c:f></c:strRef></c:cat>
  <c:val><c:numRef><c:f>{Xml(dataRange)}</c:f></c:numRef></c:val>
</c:ser>";
        }

        private static string BuildDrawingXml(Models.ChartDefinition chart, string chartRelId)
        {
            int fc = chart.ChartStartColumn - 1;
            int fr = chart.ChartStartRow - 1;
            int tc = fc + chart.ChartWidthCells;
            int tr = fr + chart.ChartHeightCells;
            return $@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<xdr:wsDr xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
  <xdr:twoCellAnchor moveWithCells=""1"" sizeWithCells=""1"">
    <xdr:from><xdr:col>{fc}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{fr}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>{tc}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{tr}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="""">
      <xdr:nvGraphicFramePr><xdr:cNvPr id=""2"" name=""Chart 1""/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""0"" cy=""0""/></xdr:xfrm>
      <a:graphic><a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
        <c:chart r:id=""{chartRelId}"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""/>
      </a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>";
        }

        private static string Col(int column)
        {
            string result = string.Empty;
            while (column > 0)
            {
                int mod = (column - 1) % 26;
                result = (char)('A' + mod) + result;
                column = (column - 1) / 26;
            }
            return result;
        }

        private static string Xml(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            return value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
        }
    }
}
