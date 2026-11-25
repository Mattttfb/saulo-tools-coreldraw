using System;
using Corel.Interop.VGCore;
using SauloTools.Core;

namespace SauloTools.Services
{
    public class FacaResult : ActionResult
    {
        public int Faltam { get; set; }
        public int Total { get; set; }
    }

    public class FacaService
    {
        private readonly Application _app;

        public FacaService(Application app)
        {
            _app = app;
        }

        public FacaResult CriarFaca(int quantidade, string espacoOpcao, double larguraCm, bool preencherLinha)
        {
            var result = new FacaResult();
            if (quantidade <= 0)
            {
                result.Ok = false;
                result.Level = StatusLevel.Warning;
                result.Message = "Informe uma quantidade maior que zero.";
                return result;
            }
            if (larguraCm <= 0)
            {
                result.Ok = false;
                result.Level = StatusLevel.Warning;
                result.Message = "Informe uma largura válida (em cm).";
                return result;
            }

            ActionResult run = CorelSession.RunUndoable(_app, "Saulo Tools — Criar Faca", cdrUnit.cdrInch, true, doc =>
            {
                ShapeRange obj = _app.ActiveSelectionRange;
                if (obj == null || obj.Count == 0)
                    return ActionResult.Warn("Selecione pelo menos 1 objeto.");

                double valEspCm;
                switch ((espacoOpcao ?? "").Trim())
                {
                    case "0,0 mm":
                    case "0,0mm":
                        valEspCm = 0.0;
                        break;
                    case "1Cm":
                        valEspCm = 1.0;
                        break;
                    default:
                        valEspCm = 0.5;
                        break;
                }
                double valEspIn = UnitHelper.CmToInches(valEspCm);
                double maxLargIn = UnitHelper.CmToInches(larguraCm);

                double x = 0.0;
                double y = 0.0;
                double alturaLinha = 0.0;
                int totalCriados = 0;

                ShapeRange created = _app.CreateShapeRange();

                for (int s = 1; s <= quantidade; s++)
                {
                    foreach (Shape o in obj)
                    {
                        if (totalCriados >= quantidade) break;

                        if (x + o.SizeWidth > maxLargIn + 1e-6 && x > 0)
                        {
                            x = 0.0;
                            y = y - alturaLinha - valEspIn;
                            alturaLinha = 0.0;
                        }

                        Shape dup = o.Duplicate(0, 0);
                        dup.SetPosition(x, y);

                        x += o.SizeWidth + valEspIn;
                        if (o.SizeHeight > alturaLinha)
                            alturaLinha = o.SizeHeight;

                        totalCriados++;
                        created.Add(dup);
                    }
                    if (totalCriados >= quantidade) break;
                }

                FacaStats stats = GeometryMath.ComputeFacaRowStats(maxLargIn, obj[1].SizeWidth, valEspIn, totalCriados);
                int faltam = stats.RemainingToFillRow;
                int totalFinal = totalCriados;

                if (preencherLinha)
                {
                    while (true)
                    {
                        bool adicionou = false;
                        foreach (Shape t in obj)
                        {
                            if (x + t.SizeWidth > maxLargIn + 1e-6) break;

                            Shape dup = t.Duplicate(0, 0);
                            dup.SetPosition(x, y);

                            x += t.SizeWidth + valEspIn;
                            created.Add(dup);
                            totalFinal++;
                            adicionou = true;
                        }
                        if (!adicionou) break;
                    }
                    faltam = 0;
                }

                doc.ClearSelection();
                if (created.Count > 0)
                {
                    created.CreateSelection();
                    if (created.Count > 1)
                        created.Group();
                }

                result.Faltam = faltam;
                result.Total = totalFinal;
                return ActionResult.Success($"Faca criada: {totalFinal} peças (Faltam: {faltam}).");
            });

            result.Ok = run.Ok;
            result.Level = run.Level;
            result.Message = run.Message;
            return result;
        }

        public ActionResult ExcluirSelecao()
        {
            return CorelSession.RunUndoable(_app, "Saulo Tools — Excluir Faca", null, true, doc =>
            {
                ShapeRange selection = _app.ActiveSelectionRange;
                if (selection == null || selection.Count == 0)
                    return ActionResult.Warn("Nenhum objeto selecionado para excluir.");

                int count = selection.Count;
                selection.Delete();
                return ActionResult.Success($"{count} objeto(s) excluído(s).");
            });
        }
    }
}
