using System;
using System.Collections.Generic;
using Corel.Interop.VGCore;
using SauloTools.Core;

namespace SauloTools.Services
{
    public enum SangriaForma
    {
        Retangulo,
        Circulo,
        Personalizada
    }

    public class SangriasService
    {
        private readonly Application _app;

        public SangriasService(Application app)
        {
            _app = app;
        }

        public ActionResult CriarSangria(string cor, string espessura, string tipoSangria, SangriaForma forma, bool espacamentoFaca)
        {
            return CorelSession.RunUndoable(_app, "Saulo Tools — Criar Sangria", cdrUnit.cdrInch, true, doc =>
            {
                ShapeRange selection = _app.ActiveSelectionRange;
                if (selection == null || selection.Count == 0)
                    return ActionResult.Warn("Selecione pelo menos 1 objeto.");

                ShapeRange created = _app.CreateShapeRange();

                if (forma == SangriaForma.Personalizada)
                {
                    object[] emptyParams = new object[0];
                    Shape boundary = selection.CustomCommand("Boundary", "CreateBoundary", ref emptyParams) as Shape;
                    if (boundary == null)
                        return ActionResult.Warn("Não foi possível criar contorno personalizado na seleção.");

                    boundary.OrderToFront();
                    ApplyOutlineColor(boundary, cor);
                    created.Add(boundary);
                }
                else
                {
                    foreach (Shape src in selection)
                    {
                        Shape novo = forma == SangriaForma.Circulo
                            ? doc.ActiveLayer.CreateEllipse(0, 0, src.SizeWidth, src.SizeHeight)
                            : doc.ActiveLayer.CreateRectangle2(0, 0, src.SizeWidth, src.SizeHeight);

                        novo.CenterX = src.CenterX;
                        novo.CenterY = src.CenterY;
                        novo.OrderToFront();

                        ApplyOutlineColor(novo, cor);
                        ApplyOutlineWidth(novo, espessura);

                        if (forma == SangriaForma.Retangulo)
                        {
                            ApplyBleedSize(doc, src, novo, tipoSangria);
                        }

                        created.Add(novo);
                    }
                }

                doc.ClearSelection();
                if (created.Count > 0)
                {
                    created.CreateSelection();
                    if (espacamentoFaca)
                    {
                        double newX = created.PositionX + created.SizeWidth + UnitHelper.CmToInches(10.0);
                        double newY = created.PositionY;
                        created.SetPosition(newX, newY);
                    }
                }

                return ActionResult.Success($"Sangria aplicada em {created.Count} objeto(s).");
            });
        }

        public ActionResult ExcluirSelecao()
        {
            return CorelSession.RunUndoable(_app, "Saulo Tools — Excluir Sangria", null, true, doc =>
            {
                ShapeRange selection = _app.ActiveSelectionRange;
                if (selection == null || selection.Count == 0)
                    return ActionResult.Warn("Nenhum objeto selecionado para excluir.");

                int count = selection.Count;
                selection.Delete();
                return ActionResult.Success($"{count} objeto(s) excluído(s).");
            });
        }

        public ActionResult EnfileirarSerpentina(bool startRightToLeft)
        {
            return CorelSession.RunUndoable(_app, "Saulo Tools — Enfileirar Serpentina", cdrUnit.cdrInch, true, doc =>
            {
                ShapeRange selection = _app.ActiveSelectionRange;
                if (selection == null || selection.Count < 2)
                    return ActionResult.Warn("Selecione pelo menos 2 objetos para enfileirar.");

                int total = selection.Count;
                var shapes = new List<Shape>(total);
                var items = new List<PositionedItem>(total);

                for (int i = 1; i <= total; i++)
                {
                    Shape s = selection[i];
                    shapes.Add(s);
                    items.Add(new PositionedItem
                    {
                        Index = i - 1,
                        CenterX = s.CenterX,
                        CenterY = s.CenterY,
                        Height = s.SizeHeight
                    });
                }

                List<int> order = GeometryMath.BuildSerpentineOrder(items, startRightToLeft);
                foreach (int idx in order)
                {
                    shapes[idx].OrderToFront();
                }

                string dir = startRightToLeft ? "← (Direita p/ Esquerda)" : "→ (Esquerda p/ Direita)";
                return ActionResult.Success($"Enfileiramento serpentina {dir} concluído ({total} objetos).");
            });
        }

        private static void ApplyOutlineColor(Shape shape, string cor)
        {
            switch (cor)
            {
                case "Branco":
                    shape.Outline.Color.CMYKAssign(0, 0, 0, 40);
                    break;
                case "Vermelho":
                    shape.Outline.Color.CMYKAssign(0, 100, 100, 0);
                    break;
                default:
                    shape.Outline.Color.CMYKAssign(0, 0, 0, 100);
                    break;
            }
        }

        private static void ApplyOutlineWidth(Shape shape, string espessura)
        {
            switch (espessura)
            {
                case "Esp. Mínima":
                    shape.Outline.Width = 0.0048 / 3.0;
                    break;
                case "0,5 Pt":
                    shape.Outline.Width = 0.021 / 3.0;
                    break;
                case "0,75 Pt":
                    shape.Outline.Width = 0.03121 / 3.0;
                    break;
                case "1,5 Pt":
                    shape.Outline.Width = 0.065 / 3.0;
                    break;
            }
        }

        private static void ApplyBleedSize(Document doc, Shape src, Shape novo, string tipo)
        {
            switch (tipo)
            {
                case "-Perfurado e Envel.":
                    doc.ReferencePoint = cdrReferencePoint.cdrCenter;
                    novo.SetSize(src.SizeWidth + UnitHelper.CmToInches(2.0), src.SizeHeight + UnitHelper.CmToInches(2.0));
                    break;
                case "-Ilhós":
                    doc.ReferencePoint = cdrReferencePoint.cdrCenter;
                    novo.SetSize(src.SizeWidth + UnitHelper.CmToInches(6.0), src.SizeHeight + UnitHelper.CmToInches(6.0));
                    break;
                case "-Cavalete":
                    doc.ReferencePoint = cdrReferencePoint.cdrCenter;
                    novo.SetSize(src.SizeWidth + UnitHelper.CmToInches(7.5), src.SizeHeight + UnitHelper.CmToInches(7.5));
                    break;
                case "-Banners":
                    doc.ReferencePoint = cdrReferencePoint.cdrCenter;
                    novo.SetSize(src.SizeWidth, src.SizeHeight + UnitHelper.CmToInches(9.0));
                    break;
                case "-Faixa":
                    doc.ReferencePoint = cdrReferencePoint.cdrCenter;
                    novo.SetSize(src.SizeWidth + UnitHelper.CmToInches(18.0), src.SizeHeight);
                    break;
                case "-Display_12":
                    doc.ReferencePoint = cdrReferencePoint.cdrTopMiddle;
                    novo.SetSize(src.SizeWidth, src.SizeHeight + UnitHelper.CmToInches(12.0));
                    break;
                case "-Display_6":
                    doc.ReferencePoint = cdrReferencePoint.cdrTopMiddle;
                    novo.SetSize(src.SizeWidth, src.SizeHeight + UnitHelper.CmToInches(8.0));
                    break;
                case "-Lona Rebitada F.":
                    doc.ReferencePoint = cdrReferencePoint.cdrCenter;
                    novo.SetSize(src.SizeWidth + UnitHelper.CmToInches(15.0), src.SizeHeight + UnitHelper.CmToInches(15.0));
                    break;
                default:
                    // -Normal: mantém tamanho original
                    break;
            }
        }
    }
}
