using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using Corel.Interop.VGCore;
using SauloTools.Core;
using CorelApp = Corel.Interop.VGCore.Application;

namespace SauloTools.Services
{
    public class MovPaginasConfig
    {
        public double LarguraObjCm { get; set; } = 5.0;
        public double AlturaObjCm { get; set; } = 5.0;
        public double LarguraMaxCm { get; set; } = 20.0;
        public double AfastamentoCm { get; set; } = 0.5;
        public bool UsarBorda { get; set; } = true;
        public string CorBorda { get; set; } = "Preto";
        public bool ConverterBitmap { get; set; } = false;
    }

    public class QuickService
    {
        private readonly CorelApp _app;

        public QuickService(CorelApp app)
        {
            _app = app;
        }

        public ActionResult ConverterBitmap300Dpi(bool cmyk)
        {
            string label = cmyk ? "CMYK" : "RGB";
            return CorelSession.RunUndoable(_app, $"Saulo Tools — Bitmap {label} 300 DPI", null, true, doc =>
            {
                ShapeRange sr = _app.ActiveSelectionRange;
                if (sr == null || sr.Count == 0)
                    return ActionResult.Warn("Selecione os objetos para converter em Bitmap.");

                cdrImageType mode = cmyk ? cdrImageType.cdrCMYKColorImage : cdrImageType.cdrRGBColorImage;
                sr.ConvertToBitmapEx(mode, true, true, 300, cdrAntiAliasingType.cdrNormalAntiAliasing, true, false, 95);
                return ActionResult.Success($"Convertido em Bitmap {label} 300 DPI com transparência.");
            });
        }

        public ActionResult Rotacionar(double graus)
        {
            return CorelSession.RunUndoable(_app, $"Saulo Tools — Rotacionar {graus}°", null, true, doc =>
            {
                ShapeRange sr = _app.ActiveSelectionRange;
                if (sr == null || sr.Count == 0)
                    return ActionResult.Warn("Selecione os objetos para rotacionar.");

                sr.Rotate(graus);
                return ActionResult.Success($"Seleção rotacionada em {graus}°.");
            });
        }

        public ActionResult DefinirUnidadeCentimetros()
        {
            if (_app?.ActiveDocument == null)
                return ActionResult.Warn("Abra um documento primeiro.");

            _app.ActiveDocument.Unit = cdrUnit.cdrCentimeter;
            return ActionResult.Success("Unidade do documento definida para centímetros.");
        }

        /// <summary>
        /// Migrates Quicks.vba Bot_Click (FACA tag/frame with Normal Red or Laminado Cyan + clipboard copy).
        /// </summary>
        public ActionResult CriarEtiquetaFaca(string nomeFaca, bool modoNormal)
        {
            return CorelSession.RunUndoable(_app, "Saulo Tools — Etiqueta Faca", cdrUnit.cdrInch, true, doc =>
            {
                ShapeRange pref = _app.ActiveSelectionRange;
                if (pref == null || pref.Count == 0)
                    return ActionResult.Warn("Selecione os objetos antes de aplicar a Faca.");

                string facaNome = (nomeFaca ?? "").Trim();
                if (string.IsNullOrEmpty(facaNome))
                    return ActionResult.Warn("Digite o nome/código da Faca.");

                ShapeRange selNovo = _app.CreateShapeRange();

                // Retângulo ao redor da seleção + margens TAM (+2cm topo, +4cm base, +2cm laterais)
                Shape obj = doc.ActiveLayer.CreateRectangle2(0, 0, pref.SizeWidth, pref.SizeHeight);
                obj.CenterX = pref.CenterX;
                obj.CenterY = pref.CenterY;

                if (modoNormal)
                    obj.Outline.Color.CMYKAssign(0, 100, 100, 0); // Vermelho
                else
                    obj.Outline.Color.CMYKAssign(100, 0, 0, 0);   // Ciano

                obj.OrderToFront();

                // TAM: Cima (+2cm), Baixo (+4cm), Laterais (+2cm)
                doc.ReferencePoint = cdrReferencePoint.cdrBottomMiddle;
                obj.SetSize(obj.SizeWidth, obj.SizeHeight + UnitHelper.CmToInches(2.0));

                doc.ReferencePoint = cdrReferencePoint.cdrTopMiddle;
                obj.SetSize(obj.SizeWidth, obj.SizeHeight + UnitHelper.CmToInches(4.0));

                doc.ReferencePoint = cdrReferencePoint.cdrCenter;
                obj.SetSize(obj.SizeWidth + UnitHelper.CmToInches(2.0), obj.SizeHeight);

                // Texto da Faca
                Shape texS = doc.ActiveLayer.CreateArtisticText(0, 0, facaNome);
                texS.Text.Story.Size = 24f;
                texS.Text.Story.Font = "Arial";
                texS.Text.Story.Bold = true;

                if (modoNormal)
                    texS.Fill.UniformColor.CMYKAssign(0, 100, 100, 0);
                else
                    texS.Fill.UniformColor.CMYKAssign(100, 0, 0, 0);

                texS.CenterX = obj.CenterX;
                texS.CenterY = obj.TopY - UnitHelper.CmToInches(0.45);

                selNovo.Add(obj);
                selNovo.Add(texS);

                doc.ClearSelection();
                selNovo.CreateSelection();
                selNovo.Group();

                try
                {
                    string clip = facaNome != "F-" ? facaNome : "Faca não Definida!";
                    System.Windows.Clipboard.SetText(clip);
                }
                catch { }

                return ActionResult.Success($"Faca '{facaNome}' ({(modoNormal ? "NORMAL" : "LAMINADO")}) criada e copiada.");
            });
        }

        /// <summary>
        /// Migrates MOVER_PAGINAS.vba (moves Page 2 content to Page 1 and deletes Page 2).
        /// </summary>
        public ActionResult MoverPagina2Para1()
        {
            return CorelSession.RunUndoable(_app, "Saulo Tools — Mover Página 2 → 1", null, true, doc =>
            {
                if (doc.Pages.Count <= 1)
                    return ActionResult.Warn("O documento precisa ter mais de 1 página.");

                if (doc.ActivePage.Index != 2)
                    return ActionResult.Warn("Ative a Página 2 antes de executar a movimentação.");

                if (doc.ActivePage.Shapes.Count == 0)
                {
                    doc.ActivePage.Delete();
                    return ActionResult.Info("Página 2 estava vazia e foi removida.");
                }

                ShapeRange all = doc.ActivePage.Shapes.All();
                if (all.Count > 1)
                {
                    Shape grp = all.Group();
                    grp.CreateSelection();
                }
                else
                {
                    all.CreateSelection();
                }

                _app.ActiveSelection.Copy();
                doc.Pages[1].Activate();

                StructPasteOptions pasteopt = _app.CreateStructPasteOptions();
                pasteopt.ColorConversionOptions.SourceColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%";
                pasteopt.ColorConversionOptions.TargetColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%";

                doc.Pages[1].ActiveLayer.PasteEx(pasteopt);
                doc.Pages[2].Delete();
                doc.Pages[1].Activate();

                return ActionResult.Success($"Página 2 movida para Página 1 ({doc.Pages.Count} página(s) restantes).");
            });
        }

        /// <summary>
        /// Migrates Quicks.vba CommandButton3_Click (all pages → resize, optional bitmap/border, move to Page 1, arrange grid).
        /// </summary>
        public ActionResult MoverTodasPaginasEmGrade(MovPaginasConfig cfg)
        {
            if (_app?.ActiveDocument == null)
                return ActionResult.Warn("Abra um documento primeiro.");

            Document doc = _app.ActiveDocument;
            if (doc.Pages.Count <= 1)
                return ActionResult.Warn("O documento precisa ter mais de 1 página.");

            if (cfg.LarguraObjCm <= 0 || cfg.AlturaObjCm <= 0 || cfg.LarguraMaxCm <= 0)
                return ActionResult.Warn("Preencha largura, altura e largura máxima válidas.");

            const double espessuraFixaPt = 0.5;
            const int tamanhoLote = 5;

            cdrUnit origUnit = doc.Unit;
            bool origOpt = _app.Optimization;
            int totalPaginas = doc.Pages.Count;
            bool groupOpen = false;

            try
            {
                _app.Optimization = true;
                doc.Unit = cdrUnit.cdrCentimeter;
                doc.BeginCommandGroup("Saulo Tools — Mover Páginas e Organizar");
                groupOpen = true;

                double espessuraCm = UnitHelper.PtToCm(espessuraFixaPt);
                Color corBorda = CreateBorderColor(cfg.CorBorda);

                // Página 1
                doc.Pages[1].Activate();
                if (doc.Pages[1].Shapes.Count > 0)
                {
                    ProcessPageShapes(doc.Pages[1], cfg, espessuraCm, corBorda);
                }

                int contadorPaginas = 1;
                StructPasteOptions pasteopt = _app.CreateStructPasteOptions();
                pasteopt.ColorConversionOptions.SourceColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%";
                pasteopt.ColorConversionOptions.TargetColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%";

                // Páginas 2..N
                while (doc.Pages.Count > 1)
                {
                    doc.Pages[2].Activate();
                    if (doc.Pages[2].Shapes.Count > 0)
                    {
                        Shape ready = ProcessPageShapes(doc.Pages[2], cfg, espessuraCm, corBorda);
                        if (ready != null)
                        {
                            ready.CreateSelection();
                            _app.ActiveSelection.Copy();
                            doc.Pages[1].Activate();
                            doc.Pages[1].ActiveLayer.PasteEx(pasteopt);
                        }
                    }

                    doc.Pages[2].Delete();
                    contadorPaginas++;

                    if (contadorPaginas % tamanhoLote == 0 && doc.Pages.Count > 1)
                    {
                        doc.EndCommandGroup();
                        doc.BeginCommandGroup("Saulo Tools — Mover Páginas (lote)");
                    }
                }

                // Montagem da grade na Página 1
                doc.Pages[1].Activate();
                ShapeRange srPag1 = _app.CreateShapeRange();
                var shapesList = new List<Shape>();
                foreach (Shape s in doc.Pages[1].Shapes)
                {
                    if (s.Type != cdrShapeType.cdrGuidelineShape)
                    {
                        srPag1.Add(s);
                        shapesList.Add(s);
                    }
                }

                if (shapesList.Count > 0)
                {
                    double startX = doc.Pages[1].LeftX;
                    double startY = doc.Pages[1].BottomY + doc.Pages[1].SizeHeight;
                    double curX = startX;
                    double curY = startY;

                    for (int k = 0; k < shapesList.Count; k++)
                    {
                        if (k > 0 && (curX + cfg.LarguraObjCm - startX) > (cfg.LarguraMaxCm + 0.001))
                        {
                            curX = startX;
                            curY = curY - cfg.AlturaObjCm - cfg.AfastamentoCm;
                        }

                        double xb, yb, wb, hb;
                        shapesList[k].GetBoundingBox(out xb, out yb, out wb, out hb, false);
                        double dx = curX - xb;
                        double dy = curY - (yb + hb);
                        shapesList[k].Move(dx, dy);

                        curX += cfg.LarguraObjCm + cfg.AfastamentoCm;
                    }

                    srPag1.CreateSelection();
                }

                return ActionResult.Success($"{totalPaginas} páginas consolidadas e organizadas na Página 1.");
            }
            catch (Exception ex)
            {
                return ActionResult.Fail($"Erro ao mover páginas: {ex.Message}");
            }
            finally
            {
                if (groupOpen) { try { doc.EndCommandGroup(); } catch { } }
                try { doc.Unit = origUnit; } catch { }
                try { _app.Optimization = origOpt; } catch { }
                try { _app.ActiveWindow?.Refresh(); } catch { }
                try { _app.Refresh(); } catch { }
            }
        }

        private Shape ProcessPageShapes(Page page, MovPaginasConfig cfg, double espessuraCm, Color corBorda)
        {
            ShapeRange sr = page.Shapes.All();
            if (sr.Count == 0) return null;

            foreach (Shape s in sr)
            {
                if (s.Locked) s.Locked = false;
            }

            Shape grp = sr.Count == 1 ? sr[1] : sr.Group();
            grp.SetSize(cfg.LarguraObjCm, cfg.AlturaObjCm);

            Shape objProc = cfg.ConverterBitmap
                ? grp.ConvertToBitmapEx(cdrImageType.cdrCMYKColorImage, false, true, 100, cdrAntiAliasingType.cdrNormalAntiAliasing, true, false, 95)
                : grp;

            if (cfg.UsarBorda)
            {
                double x, y, w, h;
                objProc.GetBoundingBox(out x, out y, out w, out h, false);
                Shape rect = page.ActiveLayer.CreateRectangle2(x, y, w, h);
                rect.Fill.ApplyNoFill();
                rect.Outline.SetProperties(espessuraCm, null, corBorda);

                ShapeRange pair = _app.CreateShapeRange();
                pair.Add(objProc);
                pair.Add(rect);
                return pair.Group();
            }

            return objProc;
        }

        private Color CreateBorderColor(string cor)
        {
            switch (cor)
            {
                case "Branco": return _app.CreateCMYKColor(0, 0, 0, 0);
                case "Cinza Claro": return _app.CreateCMYKColor(0, 0, 0, 20);
                default: return _app.CreateCMYKColor(0, 0, 0, 100);
            }
        }

        /// <summary>
        /// Migrates RASTREIO.vba (SepararObjetosPorCor — groups shapes by quantized RGB into layers).
        /// </summary>
        public ActionResult SepararObjetosPorCor()
        {
            return CorelSession.RunUndoable(_app, "Saulo Tools — Separar por Cor", null, true, doc =>
            {
                ShapeRange sel = _app.ActiveSelectionRange;
                if (sel == null || sel.Count == 0)
                    return ActionResult.Warn("Selecione o grupo/curvas do rastreamento primeiro.");

                if (sel.Count == 1 && sel[1].Type == cdrShapeType.cdrGroupShape)
                {
                    sel[1].Ungroup();
                    sel = _app.ActiveSelectionRange;
                }

                const int tolerancia = 10;
                int uniforme = 0, outro = 0;

                var list = new List<Shape>(sel.Count);
                for (int i = 1; i <= sel.Count; i++) list.Add(sel[i]);

                foreach (Shape sh in list)
                {
                    string corKey;
                    if (sh.Fill.Type == cdrFillType.cdrUniformFill)
                    {
                        uniforme++;
                        int rr = (sh.Fill.UniformColor.RGBRed / tolerancia) * tolerancia;
                        int gg = (sh.Fill.UniformColor.RGBGreen / tolerancia) * tolerancia;
                        int bb = (sh.Fill.UniformColor.RGBBlue / tolerancia) * tolerancia;
                        corKey = $"Cor_{rr}_{gg}_{bb}";
                    }
                    else
                    {
                        outro++;
                        corKey = $"Outro_Tipo_{(int)sh.Fill.Type}";
                    }

                    Layer lyr = FindLayer(doc.ActivePage, corKey) ?? doc.ActivePage.CreateLayer(corKey);
                    sh.MoveToLayer(lyr);
                }

                return ActionResult.Success($"Separados {list.Count} objetos ({uniforme} uniformes, {outro} outros) em {doc.ActivePage.Layers.Count} camadas.");
            });
        }

        private static Layer FindLayer(Page page, string name)
        {
            foreach (Layer l in page.Layers)
            {
                if (string.Equals(l.Name, name, StringComparison.OrdinalIgnoreCase))
                    return l;
            }
            return null;
        }

        /// <summary>
        /// Migrates Módulo2.vba (CopiarTextoDoGrupo — recursively extracts text and copies to Clipboard).
        /// </summary>
        public ActionResult CopiarTextoDoGrupo()
        {
            if (_app?.ActiveDocument == null)
                return ActionResult.Warn("Abra um documento primeiro.");

            ShapeRange sel = _app.ActiveSelectionRange;
            if (sel == null || sel.Count == 0)
                return ActionResult.Warn("Selecione um grupo ou texto antes de copiar.");

            var sb = new StringBuilder();
            foreach (Shape s in sel)
            {
                ExtractTextRecursive(s, sb);
            }

            string text = sb.ToString().Trim();
            if (string.IsNullOrEmpty(text))
                return ActionResult.Warn("Nenhum texto editável encontrado na seleção.");

            try
            {
                System.Windows.Clipboard.SetText(text);
                int lines = text.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Length;
                return ActionResult.Success($"Texto copiado para a área de transferência ({lines} linha(s)).");
            }
            catch (Exception ex)
            {
                return ActionResult.Fail($"Erro ao copiar para o clipboard: {ex.Message}");
            }
        }

        private static void ExtractTextRecursive(Shape shape, StringBuilder sb)
        {
            if (shape.Type == cdrShapeType.cdrTextShape)
            {
                sb.AppendLine(shape.Text.Story.Text);
            }
            else if (shape.Type == cdrShapeType.cdrGroupShape)
            {
                foreach (Shape child in shape.Shapes)
                {
                    ExtractTextRecursive(child, sb);
                }
            }
        }
    }
}
