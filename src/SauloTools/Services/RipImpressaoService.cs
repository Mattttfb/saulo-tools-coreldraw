using System;
using System.Collections.Generic;
using System.Text;
using Corel.Interop.VGCore;
using SauloTools.Core;

namespace SauloTools.Services
{
    public class RipImpressaoService
    {
        private readonly Application _app;

        public RipImpressaoService(Application app)
        {
            _app = app;
        }

        /// <summary>
        /// Migrates IMPRESSÃO.gms Intersec.vba & RIP.vba CriarGradeIntersecao.
        /// </summary>
        public ActionResult CriarGradeIntersecao(
            int nCols,
            int nRows,
            double cresceDireitaCm,
            double cresceBaixoCm,
            bool aplicarBorda,
            bool intersecaoNovaPagina)
        {
            if (nCols < 1) nCols = 1;
            if (nRows < 1) nRows = 1;
            if (!aplicarBorda && !intersecaoNovaPagina)
                return ActionResult.Warn("Marque 'Aplicar Borda' ou 'Interseção em Nova Página'.");

            return CorelSession.RunUndoable(_app, "Saulo Tools — Grade de Interseção (RIP)", cdrUnit.cdrInch, true, doc =>
            {
                ShapeRange sel = _app.ActiveSelectionRange;
                if (sel == null || sel.Count == 0)
                    return ActionResult.Warn("Selecione um objeto antes de executar.");
                if (sel.Count != 1)
                    return ActionResult.Warn("Selecione exatamente UM objeto antes de executar.");

                Shape shOriginal = sel[1];
                if (shOriginal.Type == cdrShapeType.cdrGroupShape)
                    return ActionResult.Warn("O objeto selecionado é um GRUPO. Desagrupe (Ctrl+U) ou combine (Ctrl+L) antes.");

                double larguraBorda = UnitHelper.CmToInches(0.01);
                double distExtra = UnitHelper.CmToInches(3.0);
                double extDireita = UnitHelper.CmToInches(cresceDireitaCm);
                double extBaixo = UnitHelper.CmToInches(cresceBaixoCm);

                Page pgeOrigem = doc.ActivePage;
                Layer camadaAtiva = shOriginal.Layer;

                double x, y, w, h;
                shOriginal.GetBoundingBox(out x, out y, out w, out h, false);
                double colW = w / nCols;
                double rowH = h / nRows;

                ShapeRange grupoShapes = _app.CreateShapeRange();
                Shape grpFinal = null;

                if (intersecaoNovaPagina)
                {
                    Page novaPagina = doc.AddPages(1);
                    pgeOrigem.Activate();
                    camadaAtiva = pgeOrigem.ActiveLayer;

                    Shape[,] intersecoes = new Shape[nRows, nCols];
                    Shape[,] textos = new Shape[nRows, nCols];

                    for (int j = 0; j < nRows; j++)
                    {
                        for (int i = 0; i < nCols; i++)
                        {
                            double larguraCel = colW + (i < nCols - 1 ? extDireita : 0.0);
                            double alturaCel = rowH + (j > 0 ? extBaixo : 0.0);
                            double xCel = x + i * colW;
                            double yCel = y + j * rowH - (j > 0 ? extBaixo : 0.0);

                            Shape shNovo = camadaAtiva.CreateRectangle2(xCel, yCel, larguraCel, alturaCel);
                            Shape shInter = null;
                            try
                            {
                                shInter = shNovo.Intersect(shOriginal, false, true);
                            }
                            catch { }

                            if (shInter != null)
                            {
                                if (aplicarBorda)
                                {
                                    shInter.Outline.Width = larguraBorda;
                                    shInter.Outline.Color.CMYKAssign(0, 0, 0, 100);

                                    double ix, iy, iw, ih;
                                    shInter.GetBoundingBox(out ix, out iy, out iw, out ih, false);
                                    Shape shMoldura = camadaAtiva.CreateRectangle2(ix, iy, iw, ih);
                                    shMoldura.Fill.ApplyNoFill();
                                    shMoldura.Outline.Width = larguraBorda;
                                    shMoldura.Outline.Color.CMYKAssign(0, 0, 0, 100);
                                    shMoldura.OrderToFront();

                                    ShapeRange srPar = _app.CreateShapeRange();
                                    srPar.Add(shInter);
                                    srPar.Add(shMoldura);
                                    shInter = srPar.Group();
                                }

                                intersecoes[j, i] = shInter;
                            }

                            try { shNovo.Delete(); } catch { }
                        }
                    }

                    // Etiquetas "Parte_N"
                    int contador = 1;
                    for (int j = nRows - 1; j >= 0; j--)
                    {
                        for (int i = 0; i < nCols; i++)
                        {
                            Shape shInter = intersecoes[j, i];
                            if (shInter == null) continue;

                            double ix, iy, iw, ih;
                            shInter.GetBoundingBox(out ix, out iy, out iw, out ih, false);

                            string nomeParte = $"Parte_{contador}";
                            Shape shTexto = camadaAtiva.CreateArtisticText(0, 0, nomeParte);
                            shTexto.Text.Story.Font = "Courier New";
                            shTexto.Fill.ApplyUniformFill(_app.CreateCMYKColor(0, 100, 100, 0));

                            double ptCalculado = (ih * 72.0 * 0.8) / (nomeParte.Length * 0.45);
                            if (ptCalculado < 10.0) ptCalculado = 10.0;
                            if (ptCalculado > 120.0) ptCalculado = 120.0;
                            shTexto.Text.Story.Size = (float)ptCalculado;
                            shTexto.RotationAngle = 90.0;
                            shTexto.OrderToFront();

                            textos[j, i] = shTexto;
                            contador++;
                        }
                    }

                    // Medição para alinhamento em grade espaçada
                    double[] colWidth = new double[nCols];
                    double[] maxTextWidth = new double[nCols];
                    double[] rowHeight = new double[nRows];

                    for (int i = 0; i < nCols; i++)
                    {
                        for (int j = 0; j < nRows; j++)
                        {
                            if (intersecoes[j, i] != null)
                            {
                                double ix, iy, iw, ih;
                                intersecoes[j, i].GetBoundingBox(out ix, out iy, out iw, out ih, false);
                                if (iw > colWidth[i]) colWidth[i] = iw;
                            }
                            if (textos[j, i] != null)
                            {
                                double tx, ty, tw, th;
                                textos[j, i].GetBoundingBox(out tx, out ty, out tw, out th, false);
                                if (tw > maxTextWidth[i]) maxTextWidth[i] = tw;
                            }
                        }
                        if (colWidth[i] == 0) colWidth[i] = colW;
                    }

                    for (int j = 0; j < nRows; j++)
                    {
                        for (int i = 0; i < nCols; i++)
                        {
                            if (intersecoes[j, i] != null)
                            {
                                double ix, iy, iw, ih;
                                intersecoes[j, i].GetBoundingBox(out ix, out iy, out iw, out ih, false);
                                if (ih > rowHeight[j]) rowHeight[j] = ih;
                            }
                        }
                        if (rowHeight[j] == 0) rowHeight[j] = rowH;
                    }

                    double[] targetX = new double[nCols];
                    targetX[0] = x;
                    for (int i = 1; i < nCols; i++)
                        targetX[i] = targetX[i - 1] + colWidth[i - 1] + distExtra + maxTextWidth[i] + distExtra;

                    double[] targetY = new double[nRows];
                    targetY[0] = y;
                    for (int j = 1; j < nRows; j++)
                        targetY[j] = targetY[j - 1] + rowHeight[j - 1] + distExtra;

                    for (int j = 0; j < nRows; j++)
                    {
                        for (int i = 0; i < nCols; i++)
                        {
                            Shape shInter = intersecoes[j, i];
                            if (shInter == null) continue;

                            double ix, iy, iw, ih;
                            shInter.GetBoundingBox(out ix, out iy, out iw, out ih, false);
                            shInter.Move(targetX[i] - ix, targetY[j] - iy);
                            shInter.GetBoundingBox(out ix, out iy, out iw, out ih, false);

                            Shape shTexto = textos[j, i];
                            if (shTexto != null)
                            {
                                double tx, ty, tw, th;
                                shTexto.GetBoundingBox(out tx, out ty, out tw, out th, false);
                                double dx = (ix - distExtra) - (tx + tw);
                                double dy = (iy + (ih / 2.0)) - (ty + (th / 2.0));
                                shTexto.Move(dx, dy);

                                ShapeRange srPar = _app.CreateShapeRange();
                                srPar.Add(shInter);
                                srPar.Add(shTexto);
                                grupoShapes.Add(srPar.Group());
                            }
                            else
                            {
                                grupoShapes.Add(shInter);
                            }
                        }
                    }

                    if (grupoShapes.Count > 0)
                    {
                        grpFinal = grupoShapes.Group();
                        grpFinal.MoveToLayer(novaPagina.ActiveLayer);
                    }

                    novaPagina.Activate();
                    try
                    {
                        pgeOrigem.Name = "ORIGINAL";
                        novaPagina.Name = "INTERSEÇÃO";
                        grpFinal?.CreateSelection();
                        _app.ActiveWindow?.ActiveView?.ToFitSelection();
                    }
                    catch { }

                    return ActionResult.Success($"Interseção {nCols}×{nRows} concluída na página INTERSEÇÃO ({contador - 1} partes).");
                }
                else
                {
                    camadaAtiva = pgeOrigem.ActiveLayer;
                    for (int j = 0; j < nRows; j++)
                    {
                        for (int i = 0; i < nCols; i++)
                        {
                            double larguraCel = colW + (i < nCols - 1 ? extDireita : 0.0);
                            double alturaCel = rowH + (j > 0 ? extBaixo : 0.0);
                            double xCel = x + i * colW;
                            double yCel = y + j * rowH - (j > 0 ? extBaixo : 0.0);

                            Shape shNovo = camadaAtiva.CreateRectangle2(xCel, yCel, larguraCel, alturaCel);
                            shNovo.Fill.ApplyNoFill();

                            if (aplicarBorda)
                            {
                                shNovo.Outline.Width = larguraBorda;
                                shNovo.Outline.Color.CMYKAssign(0, 0, 0, 100);
                            }
                            else
                            {
                                shNovo.Outline.SetNoOutline();
                            }

                            shNovo.OrderToFront();
                            grupoShapes.Add(shNovo);
                        }
                    }

                    if (grupoShapes.Count > 0)
                    {
                        grpFinal = grupoShapes.Group();
                        try
                        {
                            grpFinal.CreateSelection();
                            _app.ActiveWindow?.ActiveView?.ToFitSelection();
                        }
                        catch { }
                    }

                    return ActionResult.Success($"Grade de bordas {nCols}×{nRows} criada à frente do objeto.");
                }
            });
        }

        /// <summary>
        /// Migrates RIP.vba MontarTetrisSelecao (3-tier shelf packing + identical-size grouping).
        /// </summary>
        public ActionResult MontarTetris(double larguraMaxCm)
        {
            if (larguraMaxCm <= 0)
                return ActionResult.Warn("Informe uma largura máxima de mídia válida (cm).");

            return CorelSession.RunUndoable(_app, "Saulo Tools — Montar Tetris (RIP)", cdrUnit.cdrCentimeter, true, doc =>
            {
                ShapeRange sel = _app.ActiveSelectionRange;
                if (sel == null || sel.Count == 0)
                    return ActionResult.Warn("Selecione os objetos antes de montar o Tetris.");

                int n = sel.Count;
                var shapes = new List<Shape>(n);
                var orderedIndices = new List<int>(n);
                for (int i = 1; i <= n; i++)
                {
                    shapes.Add(sel[i]);
                    orderedIndices.Add(i - 1);
                }

                // Ordena por área decrescente
                orderedIndices.Sort((a, b) =>
                    (shapes[b].SizeWidth * shapes[b].SizeHeight).CompareTo(shapes[a].SizeWidth * shapes[a].SizeHeight));

                int b1 = n / 3;
                int b2 = b1 + ((n - b1) / 2);

                double topoY = doc.ActivePage.SizeHeight;
                double groupOffset = 0.0;
                var skipped = new List<string>();

                ShapeRange srGrandes = ApplyShelfTier(shapes, orderedIndices.GetRange(0, b1), larguraMaxCm, topoY, ref groupOffset, skipped);
                ShapeRange srMedios = ApplyShelfTier(shapes, orderedIndices.GetRange(b1, b2 - b1), larguraMaxCm, topoY, ref groupOffset, skipped);
                ShapeRange srPequenos = ApplyShelfTier(shapes, orderedIndices.GetRange(b2, n - b2), larguraMaxCm, topoY, ref groupOffset, skipped);

                if (srGrandes.Count > 0) ProcessarEAgrupar(srGrandes, "1. GRANDES");
                if (srMedios.Count > 0) ProcessarEAgrupar(srMedios, "2. MÉDIOS");
                if (srPequenos.Count > 0) ProcessarEAgrupar(srPequenos, "3. PEQUENOS");

                if (skipped.Count > 0)
                    return ActionResult.Warn($"Tetris concluído, mas {skipped.Count} objeto(s) excederam {larguraMaxCm:0.#} cm.");

                return ActionResult.Success($"Tetris concluído e agrupado ({n} objetos em {larguraMaxCm:0.#} cm).");
            });
        }

        private ShapeRange ApplyShelfTier(
            List<Shape> shapes,
            List<int> indices,
            double maxWidthCm,
            double topoY,
            ref double groupOffset,
            List<string> skipped)
        {
            ShapeRange srOut = _app.CreateShapeRange();
            if (indices.Count == 0) return srOut;

            var input = new List<Tuple<int, double, double>>(indices.Count);
            foreach (int idx in indices)
            {
                input.Add(Tuple.Create(idx, shapes[idx].SizeWidth, shapes[idx].SizeHeight));
            }

            List<ShelfPlanItem> plan = GeometryMath.PackShelves(input, maxWidthCm, topoY, ref groupOffset);
            foreach (ShelfPlanItem item in plan)
            {
                Shape sh = shapes[item.OriginalIndex];
                if (item.Skipped)
                {
                    skipped.Add($"{sh.Name} ({item.FinalWidth:0.0} cm)");
                    continue;
                }

                if (item.Rotated90) sh.Rotate(90.0);
                sh.LeftX = item.LeftX;
                sh.TopY = item.TopY;
                srOut.Add(sh);
            }

            return srOut;
        }

        private void ProcessarEAgrupar(ShapeRange srGrupo, string nomeMaster)
        {
            if (srGrupo.Count == 0) return;

            const double tolerancia = 0.01;
            int count = srGrupo.Count;
            var items = new List<Shape>(count);
            for (int i = 1; i <= count; i++) items.Add(srGrupo[i]);

            bool[] processados = new bool[count];
            ShapeRange srMaster = _app.CreateShapeRange();

            for (int i = 0; i < count; i++)
            {
                if (processados[i]) continue;
                processados[i] = true;

                ShapeRange srIdenticos = _app.CreateShapeRange();
                srIdenticos.Add(items[i]);

                double w1 = items[i].SizeWidth;
                double h1 = items[i].SizeHeight;

                for (int j = i + 1; j < count; j++)
                {
                    if (processados[j]) continue;
                    double w2 = items[j].SizeWidth;
                    double h2 = items[j].SizeHeight;

                    if ((Math.Abs(w1 - w2) < tolerancia && Math.Abs(h1 - h2) < tolerancia) ||
                        (Math.Abs(w1 - h2) < tolerancia && Math.Abs(h1 - w2) < tolerancia))
                    {
                        srIdenticos.Add(items[j]);
                        processados[j] = true;
                    }
                }

                if (srIdenticos.Count > 1)
                {
                    Shape grpId = srIdenticos.Group();
                    grpId.Name = $"Iguais ({w1:0.0}x{h1:0.0} cm)";
                    srMaster.Add(grpId);
                }
                else
                {
                    srMaster.Add(items[i]);
                }
            }

            if (srMaster.Count > 0)
            {
                Shape grpFinal = srMaster.Count > 1 ? srMaster.Group() : srMaster[1];
                grpFinal.Name = nomeMaster;
            }
        }
    }
}
