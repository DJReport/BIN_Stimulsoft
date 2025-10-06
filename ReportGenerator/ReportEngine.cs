// system
using System;
using System.Data;
using System.Text.RegularExpressions;

// bin interface
using BinInterface;

// stimulsoft
using Stimulsoft.Base;
using Stimulsoft.Report;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Export;

namespace ReportGenerator
{
    internal class ReportEngine : IReportGenerator
    {
        private StiReport _stiReport = new();
        private string? _lastError;
        private DataSet? _reportDataSet;

        public void LoadTemplate(string templatePath, string[] args)
        {
            try
            {
                _stiReport.ScriptLanguage = StiReportLanguageType.CSharp;
                _stiReport.Load(templatePath);
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
            }
        }

        public object CleanAndValidateData(string jsonData, string[] args)
        {
            string regexJson = Regex.Replace(jsonData, @"\r\n?|\n", " ");
            _reportDataSet = StiJsonToDataSetConverterV2.GetDataSet(regexJson);
            return _reportDataSet;
        }

        public void SetLicense(string[] args)
        {
            StiLicense.Key = args[0];
        }

        public void RegisterData(string dataSourceName, object data, string[] args)
        {
            try
            {
                DataSet stiData = (DataSet)data;
                _stiReport.RegData(dataSourceName, stiData);
            }
            catch (Exception registerEx)
            {
                _lastError = registerEx.Message;
            }
        }

        public void Render(string[] args)
        {
            try
            {
                _stiReport.Render();
                string outputType = args[0];
                string templatePath = args[1];
                string dataSourceName = args[2];
                string jsonData = args[3];

                if (outputType.ToLower() != "pdf")
                {

                    StiPage wholePage = _stiReport.Pages[0];

                    // iterating over all page components and find the maximum bottom position over all components
                    double maxBottom = 0;
                    foreach (StiPage page in _stiReport.RenderedPages)
                    {
                        foreach (var component in page.Components)
                        {
                            if (component is StiComponent stiComponent)
                            {
                                double currentBottom = stiComponent.Bottom;
                                maxBottom = Math.Max(maxBottom, currentBottom);
                            }
                        }
                    }

                    wholePage.PageHeight = wholePage.Margins.Top + wholePage.Margins.Bottom + maxBottom;
                    _stiReport = new();
                    StiPage newPage = _stiReport.Pages[0];
                    _stiReport.Load(templatePath);
                    _stiReport.RegData(dataSourceName, _reportDataSet);
                    _stiReport.Pages[0].PageHeight = wholePage.Margins.Top + wholePage.Margins.Bottom + maxBottom;
                    _stiReport.Render();
                }
            }
            catch (Exception renderEx)
            {
                _lastError = renderEx.Message;
            }
        }

        public void Export(string outputType, string outputPath, int reportResolution, string[] args)
        {
            using MemoryStream exportStream = new();
            StiExportFormat format;

            try
            {

                switch (outputType.ToLower())
                {
                    case "csv":
                        var exportCsvSettings = new StiCsvExportSettings();
                        format = StiExportFormat.Csv;
                        _stiReport.ExportDocument(format, exportStream, exportCsvSettings);
                        break;
                    case "data":
                        var exportDataSettings = new StiDataExportSettings();
                        format = StiExportFormat.Data;
                        _stiReport.ExportDocument(format, exportStream, exportDataSettings);
                        break;
                    case "dbf":
                        var exportDbfSettings = new StiDbfExportSettings();
                        format = StiExportFormat.Dbf;
                        _stiReport.ExportDocument(format, exportStream, exportDbfSettings);
                        break;
                    case "dif":
                        var exportSettings = new StiDifExportSettings();
                        format = StiExportFormat.Dif;
                        _stiReport.ExportDocument(format, exportStream, exportSettings);
                        break;
                    case "excel":
                        var exportExcelSettings = new StiExcel2007ExportSettings()
                        {
                            ImageResolution = reportResolution
                        };
                        format = StiExportFormat.Excel;
                        _stiReport.ExportDocument(format, exportStream, exportExcelSettings);
                        break;
                    case "excelxml":
                        var exportExcelXmlSettings = new StiExcelXmlExportSettings()
                        {
                            ImageResolution = reportResolution
                        };
                        format = StiExportFormat.ExcelXml;
                        _stiReport.ExportDocument(format, exportStream, exportExcelXmlSettings);
                        break;
                    case "html":
                        var exportHtmlSettings = new StiHtmlExportSettings()
                        {
                            ImageResolution = reportResolution
                        };
                        format = StiExportFormat.Html;
                        _stiReport.ExportDocument(format, exportStream, exportHtmlSettings);
                        break;
                    case "html5":
                        var exportHtml5Settings = new StiHtml5ExportSettings() { ImageResolution = reportResolution };
                        format = StiExportFormat.Html5;
                        _stiReport.ExportDocument(format, exportStream, exportHtml5Settings);
                        break;
                    case "bmp":
                        var exportBmpSettings = new StiBmpExportSettings() { ImageResolution = reportResolution };
                        format = StiExportFormat.ImageBmp;
                        var a = _stiReport.ExportDocument(format, exportStream, exportBmpSettings);
                        break;
                    case "emf":
                        var exportEmfSettings = new StiEmfExportSettings() { ImageResolution = reportResolution };
                        format = StiExportFormat.ImageEmf;
                        _stiReport.ExportDocument(format, exportStream, exportEmfSettings);
                        break;
                    case "gif":
                        var exportGifSettings = new StiGifExportSettings() { ImageResolution = reportResolution };
                        format = StiExportFormat.ImageGif;
                        _stiReport.ExportDocument(format, exportStream, exportGifSettings);
                        break;
                    case "jpeg":
                        var exportJpegSettings = new StiJpegExportSettings() { ImageResolution = reportResolution };
                        format = StiExportFormat.ImageJpeg;
                        _stiReport.ExportDocument(format, exportStream, exportJpegSettings);
                        break;
                    case "pcx":
                        var exportPcxSettings = new StiPcxExportSettings() { ImageResolution = reportResolution };
                        format = StiExportFormat.ImagePcx;
                        _stiReport.ExportDocument(format, exportStream, exportPcxSettings);
                        break;
                    case "png":
                        var exportPngSettings = new StiPngExportSettings() { ImageResolution = reportResolution };
                        format = StiExportFormat.ImagePng;
                        _stiReport.ExportDocument(format, exportStream, exportPngSettings);
                        break;
                    case "svg":
                        var exportSvgSettings = new StiSvgExportSettings() { ImageResolution = reportResolution };
                        format = StiExportFormat.ImageSvg;
                        _stiReport.ExportDocument(format, exportStream, exportSvgSettings);
                        break;
                    case "svgz":
                        var exportSvgzSettings = new StiSvgzExportSettings() { ImageResolution = reportResolution };
                        format = StiExportFormat.ImageSvgz;
                        _stiReport.ExportDocument(format, exportStream, exportSvgzSettings);
                        break;
                    case "tiff":
                        var exportTiffSettings = new StiTiffExportSettings() { ImageResolution = reportResolution };
                        format = StiExportFormat.ImageTiff;
                        _stiReport.ExportDocument(format, exportStream, exportTiffSettings);
                        break;
                    case "json":
                        var exportJsonSettings = new StiJsonExportSettings();
                        format = StiExportFormat.Json;
                        _stiReport.ExportDocument(format, exportStream, exportJsonSettings);
                        break;
                    case "mht":
                        var exportMhtSettings = new StiMhtExportSettings() { ImageResolution = reportResolution };
                        format = StiExportFormat.Mht;
                        _stiReport.ExportDocument(format, exportStream, exportMhtSettings);
                        break;
                    case "ods":
                        var exportOdsSettings = new StiOdsExportSettings() { ImageResolution = reportResolution };
                        format = StiExportFormat.Ods;
                        _stiReport.ExportDocument(format, exportStream, exportOdsSettings);
                        break;
                    case "odt":
                        var exportOdtSettings = new StiOdtExportSettings() { ImageResolution = reportResolution };
                        format = StiExportFormat.Odt;
                        _stiReport.ExportDocument(format, exportStream, exportOdtSettings);
                        break;
                    case "powerpoint":
                        var exportPowerpointSettings = new StiPowerPointExportSettings() { ImageResolution = reportResolution };
                        format = StiExportFormat.PowerPoint;
                        _stiReport.ExportDocument(format, exportStream, exportPowerpointSettings);
                        break;
                    case "rtf":
                        var exportRtfSettings = new StiRtfExportSettings() { ImageResolution = reportResolution };
                        format = StiExportFormat.Rtf;
                        _stiReport.ExportDocument(format, exportStream, exportRtfSettings);
                        break;
                    case "sylk":
                        var exportSylkSettings = new StiSylkExportSettings();
                        format = StiExportFormat.Sylk;
                        _stiReport.ExportDocument(format, exportStream, exportSylkSettings);
                        break;
                    case "word":
                        var exportWordSettings = new StiWordExportSettings() { ImageResolution = reportResolution };
                        format = StiExportFormat.Word;
                        _stiReport.ExportDocument(format, exportStream, exportWordSettings);
                        break;
                    case "xml":
                        var exportXmlSettings = new StiXmlExportSettings();
                        format = StiExportFormat.Xml;
                        _stiReport.ExportDocument(format, exportStream, exportXmlSettings);
                        break;
                    case "xps":
                        var exportXpsSettings = new StiXpsExportSettings();
                        format = StiExportFormat.Xps;
                        _stiReport.ExportDocument(format, exportStream, exportXpsSettings);
                        break;
                    case "pdf":
                        var exportPdfSettings = new StiPdfExportSettings()
                        {
                            ImageResolution = reportResolution,
                            EmbeddedFonts = true,
                            ExportRtfTextAsImage = true,
                            StandardPdfFonts = false,
                        };
                        format = StiExportFormat.Pdf;
                        _stiReport.ExportDocument(format, exportStream, exportPdfSettings);
                        break;
                    default:
                        var exportDefaultSettings = new StiPdfExportSettings()
                        {
                            ImageResolution = reportResolution
                        };
                        format = StiExportFormat.Pdf;
                        _stiReport.ExportDocument(format, exportStream, exportDefaultSettings);
                        break;
                }
                File.WriteAllBytes(outputPath, exportStream.ToArray());
            }
            catch (Exception exportEx)
            {
                _lastError = exportEx.Message;
            }
        }

        public string GetLastError()
        {
            return _lastError;
        }
    }
}
