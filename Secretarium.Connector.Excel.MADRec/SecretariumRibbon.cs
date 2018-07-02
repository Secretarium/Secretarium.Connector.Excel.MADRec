using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Secretarium.Helpers;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Windows.Forms;
using Xl = Microsoft.Office.Interop.Excel;
using XlCore = Microsoft.Office.Core;

namespace Secretarium.Excel
{
    [ComVisible(true)]
    public class SecretariumRibbon : ExcelRibbon
    {
        private static XlCore.MsoThemeColorIndex[] _colors = {
            XlCore.MsoThemeColorIndex.msoThemeColorAccent6,
            XlCore.MsoThemeColorIndex.msoThemeColorAccent5,
            XlCore.MsoThemeColorIndex.msoThemeColorAccent4,
            XlCore.MsoThemeColorIndex.msoThemeColorAccent2,
            XlCore.MsoThemeColorIndex.msoThemeColorAccent3,
        };
        private CustomTaskPane _x509TaskPane;
        private CustomTaskPane _keysTaskPane;
        
        //icon 0101:imageMso="EquationMatrixGallery"
        //icon send: imageMso="DatabaseEncodeDecode"
        public override string GetCustomUI(string RibbonID)
        {
            return @"
                <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' xmlns:Q='Secretarium' loadImage='LoadImage'>
                  <ribbon>
                    <tabs>
                      <tab idQ='Q:Secretarium' label='Secretarium'>
                        <group idQ='Q:SecretariumGrpConnect' label='Secretarium'>
                          <splitButton id='loadKeySplitBtn' size='large'>
                            <button 
                              id='loadKeyBtn' imageMso='FileDocumentEncrypt' label='Load keys'
                              onAction='ShowX509TaskPane' />
                            <menu id='loadKeySplitBtnMenu' itemSize='large'>
                              <button 
                                id='loadKeyBtn2' label='Load from encoded keys'
                                imageMso='FileDocumentEncrypt'
                                description='Load identification and signature keys from base 64 encoded strings'
                                onAction='ShowBase64KeysTaskPane' />
                              <button 
                                id='loadX509Btn' label='Load from X509'
                                imageMso='SignatureShow'
                                description='Load identification and signature keys from X509 certificate'
                                onAction='ShowX509TaskPane' />
                            </menu>
                          </splitButton>
                          <button 
                            id='connect' label='Connect' 
                            imageMso='ServerConnection' size='large'
                            screentip ='Connect to Secretarium'
                            supertip='Opens a secure connection to Secretarium'
                            onAction='Connect' />
                        </group>
                        <group idQ='Q:SecretariumMADRec' label='MADRec'>
                          <button 
                            id='piechart' label='Create report' 
                            imageMso='ChartTypePieInsertGallery' size='large'
                            screentip ='Create a MADRec report with pie charts'
                            supertip='Fill the active sheet with pie chart for each field retreived in the result'
                            onAction='CreateMADRecReport' />
                        </group>
                      </tab>
                    </tabs>
                  </ribbon>
                </customUI>";
        }

        public void ShowX509TaskPane(IRibbonControl control)
        {
            if (_x509TaskPane == null)
            {
                var x509Loader = new LoadX509();
                x509Loader.InitSecretarium(this);

                _x509TaskPane = CustomTaskPaneFactory.CreateCustomTaskPane(x509Loader, "Secretarium");
                _x509TaskPane.Visible = true;
                _x509TaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _x509TaskPane.Width = 500;
            }
            else
            {
                _x509TaskPane.Visible = true;
            }
        }

        public void ShowBase64KeysTaskPane(IRibbonControl control)
        {
            if (_keysTaskPane == null)
            {
                var keysLoader = new LoadKeys();
                keysLoader.InitSecretarium(this);

                _keysTaskPane = CustomTaskPaneFactory.CreateCustomTaskPane(keysLoader, "Secretarium");
                _keysTaskPane.Visible = true;
                _keysTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _keysTaskPane.Width = 500;
            }
            else
            {
                _keysTaskPane.Visible = true;
            }
        }

        public void Connect(IRibbonControl control)
        {
            ECDsaCng key = null;

            if (string.IsNullOrEmpty(SecretariumFunctions.Swss.PublicKey))
            {
                MessageBox.Show("Please register you identity first", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if(key != null && key.PublicKey().ToBase64String() != SecretariumFunctions.Swss.PublicKey)
                SecretariumFunctions.Swss.Set(key);

            if (SecretariumFunctions.Swss.State.IsClosed())
                SecretariumFunctions.Swss.Connect();
        }
        
        private void AddPieChart(Xl.Range rg, int k, string title, MADRecFieldResult mfr)
        {
            var shape = rg.Worksheet.Shapes.AddChart2(251, Xl.XlChartType.xlDoughnut);
            shape.Select();
            shape.Name = "MADRec_piechart_" + k;
            shape.Title = title;
            var sc = (Xl.SeriesCollection)shape.Chart.SeriesCollection();
            while (sc.Count > 1)
            {
                sc.Item(1).Delete();
            }
            if (sc.Count == 0) sc.NewSeries();
            shape.Chart.FullSeriesCollection(1).Name = "=MADRecChartsArgs!R2C" + (k + 1);
            string serieR1C1 = "='" + rg.Application.ActiveWorkbook.Name + "'!MADREC_CHART_SERIE_" + k;
            shape.Chart.FullSeriesCollection(1).Values = serieR1C1;

            (shape.Chart.ChartGroups(1) as Xl.ChartGroup).DoughnutHoleSize = 70;
            shape.Chart.SetElement(XlCore.MsoChartElementType.msoElementLegendNone);
            shape.Chart.SetElement(XlCore.MsoChartElementType.msoElementChartTitleCenteredOverlay);

            shape.Chart.ApplyDataLabels();
            shape.Chart.FullSeriesCollection(1).DataLabels.Format.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = XlCore.MsoThemeColorIndex.msoThemeColorBackground1;
            
            shape.Top = Convert.ToSingle(rg.Top) + (float)Math.Floor(k / 2.0) * 150.0f;
            shape.Left = Convert.ToSingle(rg.Left) + (k % 2) * 180.0f;
            shape.Height = 140f;
            shape.Width = 170f;

            var series = (Xl.Series)shape.Chart.SeriesCollection(1);
            for (int i = 0; i < mfr.split.Length && i < 20; i++)
            {
                var pt = series.Points(i + 1) as Xl.Point;
                pt.Format.Fill.ForeColor.ObjectThemeColor = _colors[i % 5];
                pt.Format.Fill.ForeColor.Brightness = i < 5 ? 0f : i < 10 ? 0.4f : i < 15 ? 0.6f : 0.8f;
            }

            var textBoxContrib = rg.Worksheet.Shapes.AddTextbox(
                XlCore.MsoTextOrientation.msoTextOrientationHorizontal,
                Convert.ToSingle(rg.Left) + (k % 2) * 180.0f + 85f,
                Convert.ToSingle(rg.Top) + (float)Math.Floor(k / 2.0) * 150.0f + 40f,
                80f, 70f);
            textBoxContrib.Select();
            rg.Application.Selection.Name = "MADRec_text_contrib_" + k;
            rg.Application.Selection.Formula = "=MADRecChartsArgs!" + ExcelHelpers.OffsetAddress("A7", 0, k);
            rg.Application.Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 10;
            rg.Application.Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = XlCore.MsoThemeColorIndex.msoThemeColorText1;
            rg.Application.Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.TintAndShade = 0;
            rg.Application.Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.Brightness = 0.5;
            rg.Application.Selection.ShapeRange.Line.Visible = XlCore.MsoTriState.msoFalse;
            rg.Application.Selection.ShapeRange.TextFrame2.MarginLeft = 0;
            rg.Application.Selection.ShapeRange.TextFrame2.MarginRight = 0;
            rg.Application.Selection.ShapeRange.TextFrame2.MarginTop = 0;
            rg.Application.Selection.ShapeRange.TextFrame2.MarginBottom = 0;

            var textBoxStatus = rg.Worksheet.Shapes.AddTextbox(
                XlCore.MsoTextOrientation.msoTextOrientationHorizontal,
                Convert.ToSingle(rg.Left) + (k % 2) * 180.0f + 5f,
                Convert.ToSingle(rg.Top) + (float)Math.Floor(k / 2.0) * 150.0f + 110f,
                160f, 25f);
            textBoxStatus.Select();
            rg.Application.Selection.Name = "MADRec_text_status_" + k;
            rg.Application.Selection.Formula = "=MADRecChartsArgs!" + ExcelHelpers.OffsetAddress("A8", 0, k);
            rg.Application.Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 14;
            rg.Application.Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = XlCore.MsoThemeColorIndex.msoThemeColorText1;
            rg.Application.Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.TintAndShade = 0;
            rg.Application.Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.Brightness = 0.25;
            rg.Application.Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = XlCore.MsoParagraphAlignment.msoAlignCenter;
            rg.Application.Selection.ShapeRange.Line.Visible = XlCore.MsoTriState.msoFalse;

            shape.Chart.PlotArea.Select();
            rg.Application.Selection.Top = 34d;
            rg.Application.Selection.Left = 8d;
            rg.Application.Selection.Height = 65d;
            rg.Application.Selection.Width = 65d;

            rg.Select();
        }
        private int AddPieCHart(Xl.Range rg, string reportR1C1, Xl.Worksheet wsChartArgs, string prefix, List<MADRecFieldResult> values)
        {
            var i = 0;

            foreach(var v in values)
            {
                if (v.values == null)
                {
                    wsChartArgs.Application.ActiveWorkbook.Names.Add(
                        "MADREC_CHART_SERIE_" + i, "=OFFSET(MADRecChartsArgs!R9C" + (i + 1) + ",0,0,MATCH(1000,MADRecChartsArgs!R9C" + (i + 1) + ":R35C" + (i + 1) + ",1),1)");

                    var xlArgsTopLeft = (Xl.Range)wsChartArgs.Cells[1, i + 1];
                    var xlExtract = wsChartArgs.Range[xlArgsTopLeft, xlArgsTopLeft.Offset[34, 0]];
                    xlExtract.FormulaArray = "=MADRec.Extract.ToPieChartData(" + reportR1C1 + ", \"" + (prefix + v.name) + "\")";
                    xlExtract.Calculate();

                    AddPieChart(rg, i, prefix + v.name, v);
                    i++;
                }
                else
                    i += AddPieCHart(rg, reportR1C1, wsChartArgs, prefix + v.name + "/", v.values);
            }

            return i;
        }
        public void CreateMADRecReport(IRibbonControl control)
        {
            var app = (Xl.Application)ExcelDnaUtil.Application;
            var ws = app.ActiveSheet as Xl.Worksheet;
            var selection  = app.Selection as Xl.Range;
            Xl.Range report;
            Xl.Range topLeft;
            MADRecResult mr;
            if (ws == null || selection == null || selection.Areas.Count != 2)
            {
                report = app.InputBox("Select report range", "Select MADRec report", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                string reportStr = (report.Count > 1 ? report.Value2[1, 1] : report.Value2) as string;
                if (!reportStr.TryDeserializeJsonAs(out mr))
                {
                    MessageBox.Show("Selected report is not a valid MADRec result", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                topLeft = app.InputBox("Select top left range where charts will be drawn", "Select range", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
            }
            else
            {
                var area1Range = selection.Areas[1] as Xl.Range;
                string area1 = (area1Range.Count > 1 ? area1Range.Value2[1, 1] : area1Range.Value2) as string;
                var area2Range = selection.Areas[2] as Xl.Range;
                string area2 = (area2Range.Count > 1 ? area2Range.Value2[1, 1] : area2Range.Value2) as string;
                if (string.IsNullOrEmpty(area1) && string.IsNullOrEmpty(area2))
                {
                    MessageBox.Show("Please select the MADRec report and a top left cell for the charts", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                var reportIs1st = area1.TryDeserializeJsonAs(out mr);
                if (!reportIs1st && !area2.TryDeserializeJsonAs(out mr))
                {
                    MessageBox.Show("Selected report is not a valid MADRec result", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                topLeft = reportIs1st ? area2Range : area1Range;
                report = (reportIs1st ? area1Range : area2Range);
            }
            
            if (!app.ActiveWorkbook.FindWorksheet("MADRecChartsArgs", out Xl.Worksheet wsChartArgs))
            {
                wsChartArgs = app.ActiveWorkbook.Worksheets.Add();
                wsChartArgs.Name = "MADRecChartsArgs";
                wsChartArgs.Visible = Xl.XlSheetVisibility.xlSheetHidden;
            }
            wsChartArgs.UsedRange.Clear();
            
            for(var i = 1; i <= ws.Shapes.Count; i++)
            {
                var s = ws.Shapes.Item(i);
                if (s.Name.StartsWith("MADRec"))
                {
                    s.Delete();
                    i--;
                }
            }

            var reportR1C1 = report.Address[true, true, Xl.XlReferenceStyle.xlR1C1, true];

            AddPieCHart(topLeft, reportR1C1, wsChartArgs, "", mr.values);
        }
    }

    public interface ISecretariumCustomTaskPane
    {
        void InitSecretarium(ExcelRibbon ribbon);
    }
}
