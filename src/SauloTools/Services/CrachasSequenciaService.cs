using System;
using Corel.Interop.VGCore;
using SauloTools.Core;

namespace SauloTools.Services
{
    public class CrachasSequenciaService
    {
        private readonly Application _app;

        public CrachasSequenciaService(Application app)
        {
            _app = app;
        }

        public ActionResult RedimensionarCracha(bool paisagem)
        {
            string orientacao = paisagem ? "Paisagem (8,6 × 5,5 cm)" : "Retrato (5,5 × 8,6 cm)";
            return CorelSession.RunUndoable(_app, $"Saulo Tools — Crachá {orientacao}", cdrUnit.cdrInch, true, doc =>
            {
                ShapeRange sr = _app.ActiveSelectionRange;
                if (sr == null || sr.Count == 0)
                    return ActionResult.Warn("Selecione o crachá antes de redimensionar.");

                doc.ReferencePoint = cdrReferencePoint.cdrCenter;
                if (paisagem)
                    sr.SetSize(3.385827, 2.165354);
                else
                    sr.SetSize(2.165354, 3.385827);

                return ActionResult.Success($"Crachá ajustado para {orientacao}.");
            });
        }

        /// <summary>
        /// Migrates Seque.vba (CriarSeq_Click — creates boxed sequential numbers or fixed text with proportional fit).
        /// </summary>
        public ActionResult CriarSequencia(
            int quantidade,
            string tamanhoOpcao,
            string espacoOpcao,
            bool usarNumeracao,
            string prefixoAntes,
            string textoFixo,
            bool negrito,
            bool semPreenchimento)
        {
            if (quantidade <= 0)
                return ActionResult.Warn("Informe uma quantidade válida (maior que 0).");

            return CorelSession.RunUndoable(_app, "Saulo Tools — Criar Sequência", cdrUnit.cdrInch, true, doc =>
            {
                const double margemCm = 0.2;
                double tamTexCm, largCm;
                switch (tamanhoOpcao)
                {
                    case "3 cm":
                        tamTexCm = 11.5; largCm = 32.4;
                        break;
                    case "8 cm":
                        tamTexCm = 8.0; largCm = 5.217;
                        break;
                    case "15 cm":
                        tamTexCm = 15.0; largCm = 9.516;
                        break;
                    case "20 cm":
                        tamTexCm = 20.0; largCm = 12.688;
                        break;
                    default: // 10 cm
                        tamTexCm = 10.0; largCm = 6.344;
                        break;
                }

                double espacoCm;
                switch (espacoOpcao)
                {
                    case "0,5 mm": espacoCm = 0.05; break;
                    case "2 cm": espacoCm = 2.0; break;
                    case "3 cm": espacoCm = 3.0; break;
                    case "4 cm": espacoCm = 4.0; break;
                    case "5 cm": espacoCm = 5.0; break;
                    default: espacoCm = 1.0; break;
                }

                double espacoIn = UnitHelper.CmToInches(espacoCm);
                double largIn = UnitHelper.CmToInches(largCm);
                double tamTexIn = UnitHelper.CmToInches(tamTexCm);
                double margemIn = UnitHelper.CmToInches(margemCm);
                double limiteLinhaIn = UnitHelper.CmToInches(118.0);

                double x = espacoIn;
                double y = espacoIn;

                ShapeRange created = _app.CreateShapeRange();

                for (int val = 1; val <= quantidade; val++)
                {
                    Shape retang = doc.ActiveLayer.CreateRectangle2(x, y, largIn, tamTexIn);
                    retang.Outline.Color.CMYKAssign(0, 0, 0, 100);

                    string conteudo;
                    if (usarNumeracao)
                    {
                        string pref = (prefixoAntes ?? "").Trim();
                        conteudo = string.IsNullOrEmpty(pref) ? val.ToString() : $"{pref} {val}";
                    }
                    else
                    {
                        conteudo = string.IsNullOrEmpty(textoFixo) ? val.ToString() : textoFixo;
                    }

                    Shape texto = doc.ActiveLayer.CreateArtisticText(retang.LeftX, retang.TopY, conteudo);
                    texto.Text.Story.Bold = negrito;

                    if (semPreenchimento)
                    {
                        texto.Fill.ApplyNoFill();
                        texto.Outline.Color.CMYKAssign(0, 0, 0, 100);
                    }
                    else
                    {
                        texto.Fill.ApplyUniformFill(_app.CreateCMYKColor(0, 0, 0, 100));
                    }

                    double newW, newH;
                    GeometryMath.ComputeProportionalFit(retang.SizeWidth, retang.SizeHeight, texto.SizeWidth, texto.SizeHeight, margemIn, out newW, out newH);
                    if (newW > 0 && newH > 0)
                        texto.SetSize(newW, newH);

                    texto.CenterX = retang.CenterX;
                    texto.CenterY = retang.CenterY;

                    created.Add(retang);
                    created.Add(texto);

                    x += retang.SizeWidth + espacoIn;
                    if (x >= limiteLinhaIn)
                    {
                        x = espacoIn;
                        y -= retang.SizeHeight + espacoIn;
                    }
                }

                doc.ClearSelection();
                if (created.Count > 0)
                    created.CreateSelection();

                return ActionResult.Success($"Sequência de {quantidade} item(ns) criada ({tamanhoOpcao}).");
            });
        }
    }
}
