using Bs.SolidWorks.Tools.Interop;
using Bs.SolidWorks.Tools.Logging;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using IConfigurationManager = SolidWorks.Interop.sldworks.IConfigurationManager;
using IDrawingDoc = SolidWorks.Interop.sldworks.IDrawingDoc;
using IModelDoc2 = SolidWorks.Interop.sldworks.IModelDoc2;

namespace Bs.SolidWorks.Tools.Commands {

    internal class ExtractConfigurationsJob {
        const int PaperSize = (int)swDwgPaperSizes_e.swDwgPaperA1size;
        const double ScaleNum = 1.0;
        const double ScaleDenom = 2.0;
        const double dX = 0.01, dY = dX;
        private SldWorks swApp;
        private LightLogger logger;
        private string? gostTemplatePath = null;
        private static readonly string[] standardViews = new string[]
        {
            "Front",
            //"Back",
            //"Left",
            //"Right",
            //"Top",
            //"Bottom"
        };
        public ExtractConfigurationsJob(SldWorks swApp, LightLogger logger) {
            this.swApp = swApp;
            this.logger = logger;
        }
        internal void ExtractAll(string currentDirectory) {
            // Find all *.SLDPRT documents in the current directory
            var files = Directory.GetFiles(currentDirectory, "*.SLDPRT", System.IO.SearchOption.AllDirectories);
            foreach (var file in files) {
                logger.Info($"Processing file: {file}");
                // Open the document
                Extract(file);
            }
        }
        internal void Extract(string fullFileName) {
            //Resources res = Resources.Instance;
            // Open the document in swApp and extract configurations
            ModelDoc2 swModel = null;
            int errors = 0, warnings = 0;

            try {
                logger.Info($"Opening document: {fullFileName}");

                // Открываем деталь
                swModel = swApp.OpenDoc6(fullFileName,
                    (int)swDocumentTypes_e.swDocPART,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                    "", ref errors, ref warnings);

                if (swModel == null) {
                    logger.Error($"Failed to open document: {fullFileName}, errors={errors}, warnings={warnings}");
                    return;
                }
                Array? configNameArr = swModel.GetConfigurationNames() as System.Array;
                if (configNameArr == null || configNameArr.Length == 0) {
                    logger.Warn($"No configurations found in {fullFileName}");
                    return;
                }
                //IPartDoc swPart = (IPartDoc)swModel;
                string modelPath = swModel.GetPathName();
                if (string.IsNullOrEmpty(modelPath)) {
                    logger.Warn("Model must be saved to disk before export. Skipping: " + fullFileName);
                    return;
                }
                string baseFolder = Path.GetDirectoryName(fullFileName);
                string baseName = Path.GetFileNameWithoutExtension(fullFileName);

                IConfigurationManager configMgr = swModel.ConfigurationManager;
                int count = 0;
                ScalePreference? scalePreference = new ScalePreference() {
                    ScaleDecimal = 1.0,
                    UseParentScale = true,
                    UseSheetScale = 0
                };
                ;

                string dwgPath = Path.Combine(baseFolder, $"{baseName}.dwg");
                string drawingTemplate = GetDrawingTemplatePath();
                IModelDoc2 drawModel = swApp.INewDocument2(drawingTemplate, PaperSize, 0.0, 0.0);
                IDrawingDoc swDraw = (DrawingDoc)drawModel;
                double xBase = 0, yBase = 0, xRight = 0, xLeft = 0, xBack = 0, yTop = 0, yBottom = 0, yIsometric;

                BoundingBox boundingFront, boundingBack, boundingRight, boundingLeft, boundingTop, boundingBottom, boundingConfig, boundingIsometric;

                foreach (var name in configNameArr) {
                    string configName = (string)name;
                    List<object> views = new(64);
                    count++;
                    if (string.IsNullOrEmpty(configName) || configName.Contains("FLAT-PATTERN")) {
                        logger.Warn($"Вывод конфигурации {configName} пропущен.");
                        continue;
                    }
                    else {
                        logger.Info($"Вывод конфигурации {configName}.");
                    }

                    // Активируем конфигурацию (чтобы виды были корректны для данной конфигурации)
                    try { swModel.ShowConfiguration2(configName); }
                    catch { }
                    try { swModel.EditRebuild3(); }
                    catch { }
                    swDraw.ActivateSheet("Sheet1");
                    var swSheet = swDraw.GetCurrentSheet() as Sheet;
                    {
                        IView viewFront, viewBack, viewRight, viewLeft, viewTop, viewBottom, viewIsometric;
                        // Create isometric drawing view 
                        {
                            IView view = swDraw.CreateDrawViewFromModelView3(swModel.GetPathName(), "*Front", 0, 0, 0);
                            views.Add(view);
                            viewFront = view;
                            //if (!scalePreference.HasValue) {
                            //    scalePreference = new ScalePreference() {
                            //        ScaleDecimal = view.ScaleDecimal,
                            //        UseParentScale = view.UseParentScale,
                            //        UseSheetScale = view.UseSheetScale
                            //    };
                            //    logger.Info($"Captured scale preference from base view: ScaleDecimal={scalePreference.Value.ScaleDecimal}, UseParentScale={scalePreference.Value.UseParentScale}, UseSheetScale={scalePreference.Value.UseSheetScale}");
                            //}
                            //else {
                            //    // Apply captured scale preference to base view
                            //}
                            view.ScaleDecimal = scalePreference.Value.ScaleDecimal;
                            view.UseParentScale = scalePreference.Value.UseParentScale;
                            view.UseSheetScale = scalePreference.Value.UseSheetScale;
                            boundingFront = new BoundingBox(view.GetOutline() as double[]);
                        }
                        {
                            IView view = swDraw.CreateDrawViewFromModelView3(swModel.GetPathName(), "*Left", 0, 0, 0);
                            views.Add(view);
                            viewLeft = view;
                            view.ScaleDecimal = scalePreference.Value.ScaleDecimal;
                            view.UseParentScale = scalePreference.Value.UseParentScale;
                            view.UseSheetScale = scalePreference.Value.UseSheetScale;
                            drawModel.EditRebuild3();
                            boundingLeft = new BoundingBox(view.GetOutline() as double[]);
                            xLeft = boundingFront.Xmax + dX - boundingLeft.Xmin;
                        }
                        {
                            IView view = swDraw.CreateDrawViewFromModelView3(swModel.GetPathName(), "*Right", 0, 0, 0);
                            views.Add(view);
                            viewRight = view;
                            view.ScaleDecimal = scalePreference.Value.ScaleDecimal;
                            view.UseParentScale = scalePreference.Value.UseParentScale;
                            view.UseSheetScale = scalePreference.Value.UseSheetScale;
                            drawModel.EditRebuild3();
                            boundingRight = new BoundingBox(view.GetOutline() as double[]);
                            xRight = boundingFront.Xmin - dX - boundingRight.Xmax;
                        }
                        {
                            IView view = swDraw.CreateDrawViewFromModelView3(swModel.GetPathName(), "*Bottom", 0, 0, 0);
                            views.Add(view);
                            viewBottom = view;
                            view.ScaleDecimal = scalePreference.Value.ScaleDecimal;
                            view.UseParentScale = scalePreference.Value.UseParentScale;
                            view.UseSheetScale = scalePreference.Value.UseSheetScale;
                            drawModel.EditRebuild3();
                            boundingBottom = new BoundingBox(view.GetOutline() as double[]);
                            yBottom = boundingFront.Ymax + dY - boundingBottom.Ymin;
                        }
                        {
                            IView view = swDraw.CreateDrawViewFromModelView3(swModel.GetPathName(), "*Top", 0, 0, 0);
                            views.Add(view);
                            viewTop = view;
                            view.ScaleDecimal = scalePreference.Value.ScaleDecimal;
                            view.UseParentScale = scalePreference.Value.UseParentScale;
                            view.UseSheetScale = scalePreference.Value.UseSheetScale;
                            drawModel.EditRebuild3();
                            boundingTop = new BoundingBox(view.GetOutline() as double[]);
                            yTop = boundingFront.Ymin - dY - boundingTop.Ymax;
                        }
                        {
                            IView view = swDraw.CreateDrawViewFromModelView3(swModel.GetPathName(), "*Back", 0, 0, 0);
                            views.Add(view);
                            viewBack = view;
                            view.ScaleDecimal = scalePreference.Value.ScaleDecimal;
                            view.UseParentScale = scalePreference.Value.UseParentScale;
                            view.UseSheetScale = scalePreference.Value.UseSheetScale;
                            drawModel.EditRebuild3();
                            boundingBack = new BoundingBox(view.GetOutline() as double[]);
                            xBack = boundingFront.Xmax + 2 * dX + (boundingLeft.Xmax - boundingLeft.Xmin) - boundingBack.Xmin;
                        }
                        {
                            IView view = swDraw.CreateDrawViewFromModelView3(swModel.GetPathName(), "*Isometric", 0, 0, 0);
                            views.Add(view);
                            viewIsometric = view;
                            view.ScaleDecimal = scalePreference.Value.ScaleDecimal;
                            view.UseParentScale = scalePreference.Value.UseParentScale;
                            view.UseSheetScale = scalePreference.Value.UseSheetScale;
                            boundingIsometric = new BoundingBox(view.GetOutline() as double[]);
                            yIsometric = boundingFront.Ymax + dY - boundingIsometric.Ymin;
                        }
                        boundingConfig = new(boundingFront.Xmin - dX - (boundingRight.Xmax - boundingRight.Xmin),
                                              boundingFront.Ymin - dY - (boundingTop.Ymax - boundingTop.Ymin),
                                              boundingFront.Xmax + dX + (boundingLeft.Xmax - boundingLeft.Xmin) + (dX + boundingBack.Xmax - boundingBack.Xmin),
                                              boundingFront.Ymax + dY + Math.Max((boundingBottom.Ymax - boundingBottom.Ymin), (boundingIsometric.Ymax - boundingIsometric.Ymin)));
                        double shiftY = -boundingConfig.Ymin + dY;
                        viewFront.Position = new double[] { xBase, yBase + shiftY };
                        viewBack.Position = new double[] { xBase + xBack, yBase + shiftY };
                        viewRight.Position = new double[] { xBase + xRight, yBase + shiftY };
                        viewLeft.Position = new double[] { xBase + xLeft, yBase + shiftY };
                        viewTop.Position = new double[] { xBase, yBase + shiftY + yTop };
                        viewBottom.Position = new double[] { xBase, yBase + shiftY + yBottom };
                        viewIsometric.Position = new double[] { xBase+ boundingFront.Xmax +2*dX - boundingIsometric.Xmin, yBase + shiftY + yIsometric };
                        swDraw.ActivateView(viewRight.GetName2());
                        #region Добавить надпись.
                        try {
                            try {
                                double textX = xBase + boundingConfig.Xmin + 0.005; // небольшая подушка от края
                                double textY = yBase + shiftY + boundingConfig.Ymax - 0.005;

                                Note textObj = swDraw.ICreateText2($"{baseName}-{configName}", textX, textY, 0, TextHeight: 5 * 0.001, TextAngle: 0); //(Note)drawModel.InsertNote($"{baseName}-{configName}");
                            }
                            catch (Exception exText) {
                                logger.Warn($"Failed to insert simple text for config '{configName}': {exText.Message}");
                            }

                            // После добавления графики/аннотаций — перестроим
                            try { drawModel.EditRebuild3(); }
                            catch { }
                        }
                        catch (Exception exRect) {
                            logger.Error($"Unexpected error while adding rectangle/note for config '{configName}': {exRect.Message}");
                        }

                        #endregion

                        yBase += boundingConfig.Ymax - boundingConfig.Ymin + dY;
                        drawModel.EditRebuild3();
                    }
                    foreach (var comObject in views) {
                        ReleaseCom(comObject);
                    }
                    views.Clear();
                    drawModel.EditRebuild3();
                    swDraw.ViewDisplayShaded();


                }
                swDraw.ForceRebuild();
                try {
                    if (!drawModel.SaveAs(dwgPath)) {
                        logger.Error($"Failed to save DWG: \"{dwgPath}\".");
                    }
                    else {
                        logger.Info($"Saved DWG with uniform-scale projection views: {dwgPath}");
                    }
                    string title = drawModel.GetTitle();
                    if (!string.IsNullOrEmpty(title))
                        swApp.CloseDoc(title);
                }
                catch (Exception) {
                }
            }
            catch (COMException comEx) {
                logger.Error($"COM error while opening {fullFileName}: {comEx.Message}");
            }
            catch (Exception ex) {
                logger.Error($"General error while opening {fullFileName}: {ex.Message}");
            }
            finally {
                // Закрываем документ без сохранения
                if (swModel != null) {
                    try {
                        string title = swModel.GetTitle();
                        if (!string.IsNullOrEmpty(title)) {
                            swApp.CloseDoc(title);
                        }
                    }
                    catch { }
                    ReleaseCom(swModel);
                }
            }
        }

        private string GetDrawingTemplatePath() {
            if (gostTemplatePath == null) {
                Resources res = Resources.Instance;
                if (!File.Exists(res.GostDrwTemplate)) {
                    throw new FileNotFoundException("Шаблон чертежа не найден", res.GostDrwTemplate);
                }
                gostTemplatePath = res.GostDrwTemplate;
            }
            return gostTemplatePath;
        }
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static void ReleaseCom(object comObj) {
            if (comObj == null)
                return;
            try {
                Marshal.FinalReleaseComObject(comObj);
            }
            catch { }
        }
    }
}
