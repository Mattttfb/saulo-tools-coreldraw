using System;
using System.IO;
using System.Web.Script.Serialization;

namespace SauloTools.Core
{
    public class SauloSettings
    {
        // Accordion expanded states
        public bool ExpSangrias { get; set; } = true;
        public bool ExpFaca { get; set; } = false;
        public bool ExpAgrupar { get; set; } = false;
        public bool ExpQuick { get; set; } = false;
        public bool ExpCrachasSeq { get; set; } = false;
        public bool ExpRip { get; set; } = false;

        // 1. Sangrias
        public string SangriaCor { get; set; } = "Preto";
        public string SangriaEspessura { get; set; } = "0,75 Pt";
        public string SangriaTipo { get; set; } = "-Normal";
        public string SangriaForma { get; set; } = "Retangulo";
        public bool SangriaEspacamentoFaca { get; set; } = false;

        // 2. Criar Faca
        public string FacaQuantidade { get; set; } = "50";
        public string FacaEspaco { get; set; } = "0,5mm";
        public string FacaLargura { get; set; } = "100";
        public bool FacaPreencher { get; set; } = false;

        // 3. Agrupar & Organizar
        public bool AgruparAlinharRedimensionar { get; set; } = true;
        public string OrganizarEspaco { get; set; } = "1";
        public string OrganizarLargura { get; set; } = "100";
        public string OrganizarInteligenteLargura { get; set; } = "150";

        // 4. Quick
        public string QuickFacaNome { get; set; } = "F-";
        public bool QuickFacaNormal { get; set; } = true; // true = NORMAL (Red), false = LAMINADO (Cyan)
        public string QuickMovLargObj { get; set; } = "5";
        public string QuickMovAltObj { get; set; } = "5";
        public string QuickMovLargMax { get; set; } = "20";
        public string QuickMovAfastamento { get; set; } = "0,5";
        public bool QuickMovUsarBorda { get; set; } = true;
        public string QuickMovCorBorda { get; set; } = "Preto";
        public bool QuickMovBitmap { get; set; } = false;

        // 5. Sequência
        public string SeqQuantidade { get; set; } = "10";
        public string SeqTamanho { get; set; } = "10 cm";
        public string SeqEspaco { get; set; } = "1 cm";
        public bool SeqUsarNumeracao { get; set; } = true;
        public string SeqPrefixo { get; set; } = "";
        public string SeqTextoFixo { get; set; } = "";
        public bool SeqNegrito { get; set; } = true;
        public bool SeqSemPreenchimento { get; set; } = false;

        // 6. Impressão / RIP
        public string RipColunas { get; set; } = "2";
        public string RipLinhas { get; set; } = "4";
        public string RipCresceDireita { get; set; } = "5";
        public string RipCresceBaixo { get; set; } = "5";
        public bool RipAplicarBorda { get; set; } = true;
        public bool RipIntersecaoNovaPagina { get; set; } = true;
        public string RipTetrisLargura { get; set; } = "120";
    }

    public static class SettingsStore
    {
        public static string DefaultFilePath =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SauloTools", "settings.json");

        public static SauloSettings Load(string customPath = null)
        {
            string path = customPath ?? DefaultFilePath;
            try
            {
                if (File.Exists(path))
                {
                    string json = File.ReadAllText(path);
                    var serializer = new JavaScriptSerializer();
                    var loaded = serializer.Deserialize<SauloSettings>(json);
                    if (loaded != null) return loaded;
                }
            }
            catch
            {
                // Fallback to defaults if corrupted or unreadable
            }
            return new SauloSettings();
        }

        public static void Save(SauloSettings settings, string customPath = null)
        {
            if (settings == null) return;
            string path = customPath ?? DefaultFilePath;
            try
            {
                string dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                var serializer = new JavaScriptSerializer();
                File.WriteAllText(path, serializer.Serialize(settings));
            }
            catch
            {
                // Never crash CorelDRAW if disk save fails
            }
        }
    }
}
