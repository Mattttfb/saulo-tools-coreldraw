using System;
using System.Collections.Generic;
using Corel.Interop.VGCore;
using SauloTools.Core;

namespace SauloTools.Services
{
    public class AgruparOrganizarService
    {
        private readonly Application _app;

        public AgruparOrganizarService(Application app)
        {
            _app = app;
        }

        /// <summary>
        /// Migrates AGRUPAR.vba / Agruparr.vba (AgruparPorSobreposicao).
        /// </summary>
        public ActionResult AgruparPorSobreposicao(bool alinharERedimensionar)
        {
            return CorelSession.RunUndoable(_app, "Saulo Tools — Agrupar por Sobreposição", null, true, doc =>
            {
                ShapeRange sr = _app.ActiveSelectionRange;
                if (sr == null || sr.Count == 0)
                    return ActionResult.Warn("Selecione os objetos antes de agrupar.");

                const double margemMm = 1.5;
                double mDoc = UnitHelper.MmToDocumentUnit(margemMm, doc.Unit);

                int count = sr.Count;
                var shapes = new List<Shape>(count);
                for (int i = 1; i <= count; i++) shapes.Add(sr[i]);

                bool[] usado = new bool[count];
                var grupos = new List<ShapeRange>();

                for (int i = 0; i < count; i++)
                {
                    if (usado[i]) continue;
                    Shape shpPai = shapes[i];
                    if (shpPai.Type == cdrShapeType.cdrTextShape) continue;

                    ShapeRange grp = _app.CreateShapeRange();
                    grp.Add(shpPai);
                    var grpIds = new HashSet<int> { shpPai.StaticID };

                    double pX, pY, pW, pH;
                    shpPai.GetBoundingBox(out pX, out pY, out pW, out pH, false);
                    var paiBox = new RectBox(pX, pY, pW, pH);

                    for (int j = 0; j < count; j++)
                    {
                        if (i == j || usado[j]) continue;
                        Shape shpFilho = shapes[j];

                        double fX, fY, fW, fH;
                        shpFilho.GetBoundingBox(out fX, out fY, out fW, out fH, false);
                        var filhoBox = new RectBox(fX, fY, fW, fH);

                        if (GeometryMath.Intersects(paiBox, filhoBox) || GeometryMath.IsInside(filhoBox, paiBox))
                        {
                            if (!grpIds.Contains(shpFilho.StaticID))
                            {
                                if (alinharERedimensionar)
                                {
                                    double newW, newH;
                                    GeometryMath.ComputeProportionalFit(pW, pH, fW, fH, mDoc, out newW, out newH);
                                    if (newW > 0 && newH > 0)
                                        shpFilho.SetSize(newW, newH);

                                    shpFilho.CenterX = shpPai.CenterX;
                                    shpFilho.CenterY = shpPai.CenterY;
                                }

                                grp.Add(shpFilho);
                                grpIds.Add(shpFilho.StaticID);
                                usado[j] = true;
                            }
                        }
                    }

                    grupos.Add(grp);
                }

                int totalGrupos = 0;
                foreach (ShapeRange grp in grupos)
                {
                    if (grp.Count > 1)
                    {
                        grp.Group();
                        totalGrupos++;
                    }
                }

                // Mantém seleção organizada (sem textos soltos)
                ShapeRange currentSel = _app.ActiveSelectionRange;
                ShapeRange novaSel = _app.CreateShapeRange();
                if (currentSel != null)
                {
                    foreach (Shape s in currentSel)
                    {
                        if (s.Type != cdrShapeType.cdrTextShape)
                            novaSel.Add(s);
                    }
                }

                if (novaSel.Count > 0)
                    novaSel.CreateSelection();
                else
                    doc.ClearSelection();

                return ActionResult.Success($"Agrupamento concluído: {totalGrupos} grupo(s) formado(s).");
            });
        }

        /// <summary>
        /// Migrates ORGAN.vba (simple grid wrap with configurable gap and max width in cm).
        /// </summary>
        public ActionResult OrganizarEmGrade(double espacoCm, double larguraMaxCm)
        {
            if (larguraMaxCm <= 0)
                return ActionResult.Warn("Informe uma largura máxima válida (cm).");

            return CorelSession.RunUndoable(_app, "Saulo Tools — Organizar em Grade", cdrUnit.cdrInch, true, doc =>
            {
                ShapeRange sel = _app.ActiveSelectionRange;
                if (sel == null || sel.Count == 0)
                    return ActionResult.Warn("Selecione pelo menos um objeto.");

                double gap = UnitHelper.CmToInches(Math.Max(0, espacoCm));
                double maxWidth = UnitHelper.CmToInches(larguraMaxCm);

                double startX = sel[1].LeftX;
                double startY = sel[1].TopY;
                double x = startX;
                double y = startY;
                double rowHeight = 0;

                for (int i = 1; i <= sel.Count; i++)
                {
                    Shape shp = sel[i];
                    double w = shp.SizeWidth;
                    double h = shp.SizeHeight;

                    if ((x - startX + w) > maxWidth && i > 1)
                    {
                        x = startX;
                        y = y - rowHeight - gap;
                        rowHeight = 0;
                    }

                    shp.SetPosition(x, y);
                    x += w + gap;
                    if (h > rowHeight) rowHeight = h;
                }

                return ActionResult.Success($"{sel.Count} objeto(s) organizados em grade ({larguraMaxCm:0.#} cm).");
            });
        }

        /// <summary>
        /// Migrates Organiza.vba (Organizar_1Metro_Inteligente with smart packing & 90° rotation test).
        /// </summary>
        public ActionResult OrganizarInteligente(double larguraLimiteCm)
        {
            if (larguraLimiteCm <= 0)
                return ActionResult.Warn("Informe uma largura limite válida (cm).");

            return CorelSession.RunUndoable(_app, "Saulo Tools — Organizar Inteligente", cdrUnit.cdrInch, true, doc =>
            {
                ShapeRange sr = _app.ActiveSelectionRange;
                if (sr == null || sr.Count == 0)
                    return ActionResult.Warn("Selecione os objetos para organizar.");

                double larguraLimite = UnitHelper.CmToInches(larguraLimiteCm);
                var list = new List<Shape>(sr.Count);
                for (int i = 1; i <= sr.Count; i++) list.Add(sr[i]);

                // Ordena maiores áreas primeiro
                list.Sort((a, b) => (b.SizeWidth * b.SizeHeight).CompareTo(a.SizeWidth * a.SizeHeight));

                var placed = new List<RectBox>(list.Count);

                // Primeiro objeto em (0, 0)
                list[0].SetPosition(0, 0);
                placed.Add(GetBox(list[0]));

                for (int i = 1; i < list.Count; i++)
                {
                    Shape s = list[i];
                    bool colocado = false;
                    double melhorScore = double.MaxValue;
                    double melhorX = 0, melhorY = 0;
                    bool melhorRot = false;

                    int inicioTeste = Math.Max(0, i - 20);
                    double origW = s.SizeWidth;
                    double origH = s.SizeHeight;

                    for (int j = inicioTeste; j < i; j++)
                    {
                        RectBox t = placed[j];

                        for (int r = 0; r <= 1; r++)
                        {
                            double wObj = r == 1 ? origH : origW;
                            double hObj = r == 1 ? origW : origH;

                            // Candidato 1: à direita de t (TopY alinhado)
                            TryCandidate(t.Right, t.Top, wObj, hObj, larguraLimite, placed, r == 1,
                                ref colocado, ref melhorScore, ref melhorX, ref melhorY, ref melhorRot);

                            // Candidato 2: abaixo de t (LeftX alinhado)
                            TryCandidate(t.Left, t.Bottom, wObj, hObj, larguraLimite, placed, r == 1,
                                ref colocado, ref melhorScore, ref melhorX, ref melhorY, ref melhorRot);
                        }
                    }

                    if (colocado)
                    {
                        if (melhorRot) s.Rotate(90);
                        s.SetPosition(melhorX, melhorY);
                    }
                    else
                    {
                        double menorY = 0;
                        foreach (RectBox pb in placed)
                        {
                            if (pb.Bottom < menorY) menorY = pb.Bottom;
                        }

                        if (s.SizeWidth > s.SizeHeight && s.SizeHeight <= larguraLimite)
                            s.Rotate(90);

                        s.SetPosition(0, menorY);
                    }

                    placed.Add(GetBox(s));
                }

                return ActionResult.Success($"Organização inteligente concluída ({list.Count} objetos, limite {larguraLimiteCm:0.#} cm).");
            });
        }

        private static RectBox GetBox(Shape s)
        {
            return new RectBox(s.LeftX, s.BottomY, s.SizeWidth, s.SizeHeight);
        }

        private static void TryCandidate(
            double xLeft, double yTop,
            double wObj, double hObj,
            double larguraLimite,
            List<RectBox> placed,
            bool rotated,
            ref bool colocado,
            ref double melhorScore,
            ref double melhorX,
            ref double melhorY,
            ref bool melhorRot)
        {
            if (xLeft + wObj > larguraLimite) return;

            var candidate = new RectBox(xLeft, yTop - hObj, wObj, hObj);
            for (int k = 0; k < placed.Count; k++)
            {
                if (GeometryMath.Intersects(candidate, placed[k]))
                    return;
            }

            double score = (Math.Abs(yTop) * 100000.0) + (xLeft * 1000.0) + (larguraLimite - (xLeft + wObj));

            // Prioriza encaixe encostado
            for (int k = 0; k < placed.Count; k++)
            {
                if (Math.Abs(candidate.Left - placed[k].Right) < 0.01 ||
                    Math.Abs(candidate.Top - placed[k].Bottom) < 0.01)
                {
                    score -= 50000.0;
                    break;
                }
            }

            if (score < melhorScore)
            {
                melhorScore = score;
                melhorX = xLeft;
                melhorY = yTop;
                melhorRot = rotated;
                colocado = true;
            }
        }
    }
}
