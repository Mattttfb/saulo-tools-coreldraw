using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Corel.Interop.VGCore;
using SauloTools.Core;
using SauloTools.Services;
using CorelApp = Corel.Interop.VGCore.Application;

namespace SauloTools
{
    public partial class DockerUI : UserControl
    {
        private readonly CorelApp _app;
        private readonly SangriasService _sangrias;
        private readonly FacaService _faca;
        private readonly AgruparOrganizarService _agrupar;
        private readonly QuickService _quick;
        private readonly CrachasSequenciaService _crachasSeq;
        private readonly RipImpressaoService _rip;

        private SauloSettings _settings;
        private bool _isLoading;
        private bool _quickFacaModoNormal = true;

        public DockerUI(object appObj) : this(appObj as CorelApp)
        {
        }

        public DockerUI(CorelApp app)
        {
            _isLoading = true;
            InitializeComponent();

            _app = app;
            _sangrias = new SangriasService(_app);
            _faca = new FacaService(_app);
            _agrupar = new AgruparOrganizarService(_app);
            _quick = new QuickService(_app);
            _crachasSeq = new CrachasSequenciaService(_app);
            _rip = new RipImpressaoService(_app);

            _settings = SettingsStore.Load();
            ApplySettingsToUI();
            _isLoading = false;

            ApplyCorelTheme();
            HookCorelEvents();
            UpdateSelectionInfo();
            SetStatus(ActionResult.Info("Saulo Tools pronto."));
        }

        private void HookCorelEvents()
        {
            if (_app == null) return;
            try
            {
                _app.SelectionChange += App_SelectionChange;
                _app.WindowActivate += App_WindowActivate;
                _app.OnApplicationEvent += App_OnApplicationEvent;
                Unloaded += (s, e) => UnhookCorelEvents();
            }
            catch { }
        }

        private void UnhookCorelEvents()
        {
            if (_app == null) return;
            try
            {
                _app.SelectionChange -= App_SelectionChange;
                _app.WindowActivate -= App_WindowActivate;
                _app.OnApplicationEvent -= App_OnApplicationEvent;
            }
            catch { }
        }

        private void App_SelectionChange()
        {
            Dispatcher.BeginInvoke(new Action(UpdateSelectionInfo));
        }

        private void App_WindowActivate(Document doc, Corel.Interop.VGCore.Window window)
        {
            Dispatcher.BeginInvoke(new Action(UpdateSelectionInfo));
        }

        private void App_OnApplicationEvent(string eventName, ref object[] parameters)
        {
            if (string.Equals(eventName, "WorkspaceChanged", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(eventName, "OnColorSchemeChanged", StringComparison.OrdinalIgnoreCase))
            {
                Dispatcher.BeginInvoke(new Action(ApplyCorelTheme));
            }
        }

        private void UpdateSelectionInfo()
        {
            try
            {
                if (_app == null || _app.ActiveDocument == null)
                {
                    TxtSelectionInfo.Text = "Nenhum documento aberto";
                    return;
                }

                int count = _app.ActiveSelectionRange?.Count ?? 0;
                int pages = _app.ActiveDocument.Pages?.Count ?? 0;
                TxtSelectionInfo.Text = $"Seleção: {count} obj.  |  Páginas: {pages}";
            }
            catch
            {
                TxtSelectionInfo.Text = "Seleção: —";
            }
        }

        private void ApplyCorelTheme()
        {
            string scheme = "Dark";
            try
            {
                if (_app != null)
                {
                    object val = _app.GetApplicationPreferenceValue("WindowScheme", "Colors");
                    if (val != null) scheme = val.ToString();
                }
            }
            catch { }

            if (scheme.IndexOf("Black", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                SetPalette("#1B1B1D", "#232326", "#2A2A2E", "#35353B", "#141416", "#3D3D42", "#2F2F34", "#ECECF0", "#9A9AA6", "#313136", "#3F3F46", "#232327", "#151517");
            }
            else if (scheme.IndexOf("Light", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                SetPalette("#F0F0F2", "#E6E6EA", "#DCDCE0", "#D0D0D6", "#FFFFFF", "#B8B8C0", "#D0D0D6", "#1F1F23", "#5A5A64", "#E2E2E6", "#D4D4DC", "#C6C6CE", "#E4E4E8");
            }
            else if (scheme.IndexOf("Medium", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                SetPalette("#5E5E63", "#68686E", "#707076", "#7C7C84", "#4A4A4F", "#82828A", "#727278", "#F4F4F6", "#D2D2DA", "#6E6E75", "#7D7D85", "#58585E", "#525257");
            }
            else
            {
                // Dark (CorelDRAW 2020 default dark seen in user's workspace)
                SetPalette("#2B2B2E", "#333337", "#3A3A3F", "#45454B", "#1F1F22", "#4A4A50", "#3E3E43", "#E6E6EA", "#A6A6B0", "#3F3F45", "#4D4D54", "#2F2F34", "#232326");
            }
        }

        private void SetPalette(
            string bg, string surface, string header, string headerHover,
            string inputBg, string border, string divider, string text,
            string muted, string btnBg, string btnHover, string btnPressed, string statusBg)
        {
            Resources["CorelBgBrush"] = BrushFromHex(bg);
            Resources["CorelSurfaceBrush"] = BrushFromHex(surface);
            Resources["CorelHeaderBrush"] = BrushFromHex(header);
            Resources["CorelHeaderHoverBrush"] = BrushFromHex(headerHover);
            Resources["CorelInputBgBrush"] = BrushFromHex(inputBg);
            Resources["CorelBorderBrush"] = BrushFromHex(border);
            Resources["CorelDividerBrush"] = BrushFromHex(divider);
            Resources["CorelTextBrush"] = BrushFromHex(text);
            Resources["CorelMutedTextBrush"] = BrushFromHex(muted);
            Resources["CorelButtonBgBrush"] = BrushFromHex(btnBg);
            Resources["CorelButtonHoverBrush"] = BrushFromHex(btnHover);
            Resources["CorelButtonPressedBrush"] = BrushFromHex(btnPressed);
            Resources["CorelStatusBgBrush"] = BrushFromHex(statusBg);
        }

        private static SolidColorBrush BrushFromHex(string hex)
        {
            var b = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString(hex));
            b.Freeze();
            return b;
        }

        // =====================================================================
        // SETTINGS SYNC
        // =====================================================================

        private void ApplySettingsToUI()
        {
            ExpSangrias.IsExpanded = _settings.ExpSangrias;
            ExpFaca.IsExpanded = _settings.ExpFaca;
            ExpAgrupar.IsExpanded = _settings.ExpAgrupar;
            ExpQuick.IsExpanded = _settings.ExpQuick;
            ExpCrachasSeq.IsExpanded = _settings.ExpCrachasSeq;
            ExpRip.IsExpanded = _settings.ExpRip;

            SelectComboByText(CmbSangriaCor, _settings.SangriaCor);
            SelectComboByText(CmbSangriaEsp, _settings.SangriaEspessura);
            SelectComboByText(CmbSangriaTipo, _settings.SangriaTipo);
            RbFormaRet.IsChecked = _settings.SangriaForma == "Retangulo";
            RbFormaCirc.IsChecked = _settings.SangriaForma == "Circulo";
            RbFormaPers.IsChecked = _settings.SangriaForma == "Personalizada";
            ChkEspacFaca.IsChecked = _settings.SangriaEspacamentoFaca;
            UpdateSangriaFormAvailability();

            TxtFacaQuant.Text = _settings.FacaQuantidade;
            SelectComboByText(CmbFacaEspaco, _settings.FacaEspaco);
            TxtFacaLargura.Text = _settings.FacaLargura;
            ChkFacaPreencher.IsChecked = _settings.FacaPreencher;

            ChkAgruparAlinhar.IsChecked = _settings.AgruparAlinharRedimensionar;
            TxtOrgEspaco.Text = _settings.OrganizarEspaco;
            TxtOrgLargura.Text = _settings.OrganizarLargura;
            TxtOrgIntelLargura.Text = _settings.OrganizarInteligenteLargura;

            TxtQuickFacaNome.Text = _settings.QuickFacaNome;
            _quickFacaModoNormal = _settings.QuickFacaNormal;
            UpdateQuickFacaModeVisual();
            TxtMovLargObj.Text = _settings.QuickMovLargObj;
            TxtMovAltObj.Text = _settings.QuickMovAltObj;
            TxtMovLargMax.Text = _settings.QuickMovLargMax;
            TxtMovAfast.Text = _settings.QuickMovAfastamento;
            ChkMovBorda.IsChecked = _settings.QuickMovUsarBorda;
            SelectComboByText(CmbMovCorBorda, _settings.QuickMovCorBorda);
            ChkMovBitmap.IsChecked = _settings.QuickMovBitmap;

            TxtSeqQuant.Text = _settings.SeqQuantidade;
            SelectComboByText(CmbSeqTamanho, _settings.SeqTamanho);
            SelectComboByText(CmbSeqEspaco, _settings.SeqEspaco);
            ChkSeqNumeracao.IsChecked = _settings.SeqUsarNumeracao;
            TxtSeqPrefixo.Text = _settings.SeqPrefixo;
            TxtSeqFixo.Text = _settings.SeqTextoFixo;
            ChkSeqNegrito.IsChecked = _settings.SeqNegrito;
            ChkSeqContorno.IsChecked = _settings.SeqSemPreenchimento;
            UpdateSeqModeAvailability();

            TxtRipCols.Text = _settings.RipColunas;
            TxtRipLinhas.Text = _settings.RipLinhas;
            TxtRipDir.Text = _settings.RipCresceDireita;
            TxtRipBaixo.Text = _settings.RipCresceBaixo;
            ChkRipBorda.IsChecked = _settings.RipAplicarBorda;
            ChkRipIntersecao.IsChecked = _settings.RipIntersecaoNovaPagina;
            TxtRipTetrisLarg.Text = _settings.RipTetrisLargura;
        }

        private void CaptureAndSaveSettings()
        {
            if (_isLoading || _settings == null) return;

            _settings.ExpSangrias = ExpSangrias.IsExpanded;
            _settings.ExpFaca = ExpFaca.IsExpanded;
            _settings.ExpAgrupar = ExpAgrupar.IsExpanded;
            _settings.ExpQuick = ExpQuick.IsExpanded;
            _settings.ExpCrachasSeq = ExpCrachasSeq.IsExpanded;
            _settings.ExpRip = ExpRip.IsExpanded;

            _settings.SangriaCor = GetComboText(CmbSangriaCor, "Preto");
            _settings.SangriaEspessura = GetComboText(CmbSangriaEsp, "0,75 Pt");
            _settings.SangriaTipo = GetComboText(CmbSangriaTipo, "-Normal");
            _settings.SangriaForma = RbFormaCirc.IsChecked == true ? "Circulo" : (RbFormaPers.IsChecked == true ? "Personalizada" : "Retangulo");
            _settings.SangriaEspacamentoFaca = ChkEspacFaca.IsChecked == true;

            _settings.FacaQuantidade = TxtFacaQuant.Text;
            _settings.FacaEspaco = GetComboText(CmbFacaEspaco, "0,5mm");
            _settings.FacaLargura = TxtFacaLargura.Text;
            _settings.FacaPreencher = ChkFacaPreencher.IsChecked == true;

            _settings.AgruparAlinharRedimensionar = ChkAgruparAlinhar.IsChecked == true;
            _settings.OrganizarEspaco = TxtOrgEspaco.Text;
            _settings.OrganizarLargura = TxtOrgLargura.Text;
            _settings.OrganizarInteligenteLargura = TxtOrgIntelLargura.Text;

            _settings.QuickFacaNome = TxtQuickFacaNome.Text;
            _settings.QuickFacaNormal = _quickFacaModoNormal;
            _settings.QuickMovLargObj = TxtMovLargObj.Text;
            _settings.QuickMovAltObj = TxtMovAltObj.Text;
            _settings.QuickMovLargMax = TxtMovLargMax.Text;
            _settings.QuickMovAfastamento = TxtMovAfast.Text;
            _settings.QuickMovUsarBorda = ChkMovBorda.IsChecked == true;
            _settings.QuickMovCorBorda = GetComboText(CmbMovCorBorda, "Preto");
            _settings.QuickMovBitmap = ChkMovBitmap.IsChecked == true;

            _settings.SeqQuantidade = TxtSeqQuant.Text;
            _settings.SeqTamanho = GetComboText(CmbSeqTamanho, "10 cm");
            _settings.SeqEspaco = GetComboText(CmbSeqEspaco, "1 cm");
            _settings.SeqUsarNumeracao = ChkSeqNumeracao.IsChecked == true;
            _settings.SeqPrefixo = TxtSeqPrefixo.Text;
            _settings.SeqTextoFixo = TxtSeqFixo.Text;
            _settings.SeqNegrito = ChkSeqNegrito.IsChecked == true;
            _settings.SeqSemPreenchimento = ChkSeqContorno.IsChecked == true;

            _settings.RipColunas = TxtRipCols.Text;
            _settings.RipLinhas = TxtRipLinhas.Text;
            _settings.RipCresceDireita = TxtRipDir.Text;
            _settings.RipCresceBaixo = TxtRipBaixo.Text;
            _settings.RipAplicarBorda = ChkRipBorda.IsChecked == true;
            _settings.RipIntersecaoNovaPagina = ChkRipIntersecao.IsChecked == true;
            _settings.RipTetrisLargura = TxtRipTetrisLarg.Text;

            SettingsStore.Save(_settings);
        }

        private static void SelectComboByText(ComboBox cmb, string value)
        {
            if (cmb == null || cmb.Items.Count == 0) return;
            for (int i = 0; i < cmb.Items.Count; i++)
            {
                var item = cmb.Items[i] as ComboBoxItem;
                string txt = item?.Content?.ToString() ?? cmb.Items[i]?.ToString();
                if (string.Equals(txt, value, StringComparison.OrdinalIgnoreCase))
                {
                    cmb.SelectedIndex = i;
                    return;
                }
            }
            cmb.SelectedIndex = 0;
        }

        private static string GetComboText(ComboBox cmb, string fallback)
        {
            var item = cmb?.SelectedItem as ComboBoxItem;
            return item?.Content?.ToString() ?? cmb?.SelectedItem?.ToString() ?? fallback;
        }

        private void SettingControl_Changed(object sender, RoutedEventArgs e)
        {
            CaptureAndSaveSettings();
        }

        private void Expander_StateChanged(object sender, RoutedEventArgs e)
        {
            CaptureAndSaveSettings();
        }

        private void SetStatus(ActionResult res)
        {
            if (res == null) return;
            TxtStatus.Text = res.Message;
            switch (res.Level)
            {
                case StatusLevel.Success:
                    StatusDot.Fill = BrushFromHex("#3FA65C");
                    break;
                case StatusLevel.Warning:
                    StatusDot.Fill = BrushFromHex("#E2A93B");
                    break;
                case StatusLevel.Error:
                    StatusDot.Fill = BrushFromHex("#D94A4A");
                    break;
                default:
                    StatusDot.Fill = BrushFromHex("#2D89EF");
                    break;
            }
            UpdateSelectionInfo();
        }

        // =====================================================================
        // NAVIGATION BAR
        // =====================================================================

        private void NavJump_Click(object sender, RoutedEventArgs e)
        {
            string tag = (sender as FrameworkElement)?.Tag?.ToString();
            Expander target = null;
            switch (tag)
            {
                case "Sangrias": target = ExpSangrias; break;
                case "Faca": target = ExpFaca; break;
                case "Agrupar": target = ExpAgrupar; break;
                case "Quick": target = ExpQuick; break;
                case "CrachasSeq": target = ExpCrachasSeq; break;
                case "Rip": target = ExpRip; break;
            }
            if (target == null) return;

            target.IsExpanded = true;
            target.BringIntoView();
        }

        private void BtnToggleAll_Click(object sender, RoutedEventArgs e)
        {
            bool anyExpanded = ExpSangrias.IsExpanded || ExpFaca.IsExpanded || ExpAgrupar.IsExpanded ||
                               ExpQuick.IsExpanded || ExpCrachasSeq.IsExpanded || ExpRip.IsExpanded;
            bool next = !anyExpanded;

            _isLoading = true;
            ExpSangrias.IsExpanded = next;
            ExpFaca.IsExpanded = next;
            ExpAgrupar.IsExpanded = next;
            ExpQuick.IsExpanded = next;
            ExpCrachasSeq.IsExpanded = next;
            ExpRip.IsExpanded = next;
            _isLoading = false;

            CaptureAndSaveSettings();
        }

        // =====================================================================
        // SEÇÃO 1: SANGRIAS
        // =====================================================================

        private void RbForma_Changed(object sender, RoutedEventArgs e)
        {
            UpdateSangriaFormAvailability();
            CaptureAndSaveSettings();
        }

        private void UpdateSangriaFormAvailability()
        {
            if (CmbSangriaTipo == null || CmbSangriaEsp == null) return;
            bool isRet = RbFormaRet?.IsChecked == true;
            bool isPers = RbFormaPers?.IsChecked == true;
            CmbSangriaTipo.IsEnabled = isRet;
            CmbSangriaEsp.IsEnabled = !isPers;
        }

        private void BtnSangriaCriar_Click(object sender, RoutedEventArgs e)
        {
            SangriaForma forma = RbFormaCirc.IsChecked == true
                ? SangriaForma.Circulo
                : (RbFormaPers.IsChecked == true ? SangriaForma.Personalizada : SangriaForma.Retangulo);

            var res = _sangrias.CriarSangria(
                GetComboText(CmbSangriaCor, "Preto"),
                GetComboText(CmbSangriaEsp, "0,75 Pt"),
                GetComboText(CmbSangriaTipo, "-Normal"),
                forma,
                ChkEspacFaca.IsChecked == true);
            SetStatus(res);
        }

        private void BtnSangriaExcluir_Click(object sender, RoutedEventArgs e)
        {
            SetStatus(_sangrias.ExcluirSelecao());
        }

        private void BtnSerpentinaEsq_Click(object sender, RoutedEventArgs e)
        {
            SetStatus(_sangrias.EnfileirarSerpentina(startRightToLeft: true));
        }

        private void BtnSerpentinaDir_Click(object sender, RoutedEventArgs e)
        {
            SetStatus(_sangrias.EnfileirarSerpentina(startRightToLeft: false));
        }

        // =====================================================================
        // SEÇÃO 2: CRIAR FACA
        // =====================================================================

        private void BtnFacaCriar_Click(object sender, RoutedEventArgs e)
        {
            int quant;
            double larg;
            if (!UnitHelper.TryParseInt(TxtFacaQuant.Text, out quant))
            {
                SetStatus(ActionResult.Warn("Quantidade inválida."));
                return;
            }
            if (!UnitHelper.TryParseDouble(TxtFacaLargura.Text, out larg))
            {
                SetStatus(ActionResult.Warn("Largura inválida."));
                return;
            }

            FacaResult res = _faca.CriarFaca(quant, GetComboText(CmbFacaEspaco, "0,5mm"), larg, ChkFacaPreencher.IsChecked == true);
            if (res.Ok)
            {
                TxtFacaFaltam.Text = $"Faltam: {res.Faltam}";
                TxtFacaTotal.Text = $"Total: {res.Total}";
            }
            SetStatus(res);
        }

        private void BtnFacaExcluir_Click(object sender, RoutedEventArgs e)
        {
            var res = _faca.ExcluirSelecao();
            if (res.Ok)
            {
                TxtFacaFaltam.Text = "Faltam: 0";
                TxtFacaTotal.Text = "Total: 0";
            }
            SetStatus(res);
        }

        // =====================================================================
        // SEÇÃO 3: AGRUPAR & ORGANIZAR
        // =====================================================================

        private void BtnAgrupar_Click(object sender, RoutedEventArgs e)
        {
            SetStatus(_agrupar.AgruparPorSobreposicao(ChkAgruparAlinhar.IsChecked == true));
        }

        private void BtnOrgGrade_Click(object sender, RoutedEventArgs e)
        {
            double esp, larg;
            if (!UnitHelper.TryParseDouble(TxtOrgEspaco.Text, out esp) || !UnitHelper.TryParseDouble(TxtOrgLargura.Text, out larg))
            {
                SetStatus(ActionResult.Warn("Informe espaço e largura válidos (cm)."));
                return;
            }
            SetStatus(_agrupar.OrganizarEmGrade(esp, larg));
        }

        private void BtnOrgInteligente_Click(object sender, RoutedEventArgs e)
        {
            double larg;
            if (!UnitHelper.TryParseDouble(TxtOrgIntelLargura.Text, out larg))
            {
                SetStatus(ActionResult.Warn("Informe uma largura limite válida (cm)."));
                return;
            }
            SetStatus(_agrupar.OrganizarInteligente(larg));
        }

        // =====================================================================
        // SEÇÃO 4: QUICK
        // =====================================================================

        private void BtnQuickFacaModo_Click(object sender, RoutedEventArgs e)
        {
            _quickFacaModoNormal = !_quickFacaModoNormal;
            UpdateQuickFacaModeVisual();
            CaptureAndSaveSettings();
        }

        private void UpdateQuickFacaModeVisual()
        {
            if (BtnQuickFacaModo == null) return;
            if (_quickFacaModoNormal)
            {
                BtnQuickFacaModo.Content = "NORMAL";
                BtnQuickFacaModo.Foreground = BrushFromHex("#FF6B6B");
            }
            else
            {
                BtnQuickFacaModo.Content = "LAMINADO";
                BtnQuickFacaModo.Foreground = BrushFromHex("#38D9E8");
            }
        }

        private void BtnQuickAplicarFaca_Click(object sender, RoutedEventArgs e)
        {
            SetStatus(_quick.CriarEtiquetaFaca(TxtQuickFacaNome.Text, _quickFacaModoNormal));
        }

        private void BtnBitmapRgb_Click(object sender, RoutedEventArgs e)
        {
            SetStatus(_quick.ConverterBitmap300Dpi(cmyk: false));
        }

        private void BtnBitmapCmyk_Click(object sender, RoutedEventArgs e)
        {
            SetStatus(_quick.ConverterBitmap300Dpi(cmyk: true));
        }

        private void BtnRotEsq_Click(object sender, RoutedEventArgs e)
        {
            SetStatus(_quick.Rotacionar(90.0));
        }

        private void BtnRotDir_Click(object sender, RoutedEventArgs e)
        {
            SetStatus(_quick.Rotacionar(-90.0));
        }

        private void BtnUnidadeCm_Click(object sender, RoutedEventArgs e)
        {
            SetStatus(_quick.DefinirUnidadeCentimetros());
        }

        private void BtnSepararCor_Click(object sender, RoutedEventArgs e)
        {
            SetStatus(_quick.SepararObjetosPorCor());
        }

        private void BtnCopiarTexto_Click(object sender, RoutedEventArgs e)
        {
            SetStatus(_quick.CopiarTextoDoGrupo());
        }

        private void BtnMoverPag2_Click(object sender, RoutedEventArgs e)
        {
            SetStatus(_quick.MoverPagina2Para1());
        }

        private void BtnMoverTodasPaginas_Click(object sender, RoutedEventArgs e)
        {
            double lObj, aObj, lMax, afast;
            if (!UnitHelper.TryParseDouble(TxtMovLargObj.Text, out lObj) ||
                !UnitHelper.TryParseDouble(TxtMovAltObj.Text, out aObj) ||
                !UnitHelper.TryParseDouble(TxtMovLargMax.Text, out lMax) ||
                !UnitHelper.TryParseDouble(TxtMovAfast.Text, out afast))
            {
                SetStatus(ActionResult.Warn("Verifique os valores de largura, altura, máximo e gap (cm)."));
                return;
            }

            var cfg = new MovPaginasConfig
            {
                LarguraObjCm = lObj,
                AlturaObjCm = aObj,
                LarguraMaxCm = lMax,
                AfastamentoCm = afast,
                UsarBorda = ChkMovBorda.IsChecked == true,
                CorBorda = GetComboText(CmbMovCorBorda, "Preto"),
                ConverterBitmap = ChkMovBitmap.IsChecked == true
            };

            SetStatus(_quick.MoverTodasPaginasEmGrade(cfg));
        }

        // =====================================================================
        // SEÇÃO 5: CRACHÁS & SEQUÊNCIA
        // =====================================================================

        private void BtnCrachaPaisagem_Click(object sender, RoutedEventArgs e)
        {
            SetStatus(_crachasSeq.RedimensionarCracha(paisagem: true));
        }

        private void BtnCrachaRetrato_Click(object sender, RoutedEventArgs e)
        {
            SetStatus(_crachasSeq.RedimensionarCracha(paisagem: false));
        }

        private void ChkSeqNumeracao_Changed(object sender, RoutedEventArgs e)
        {
            UpdateSeqModeAvailability();
            CaptureAndSaveSettings();
        }

        private void UpdateSeqModeAvailability()
        {
            if (TxtSeqFixo == null || TxtSeqPrefixo == null || ChkSeqNumeracao == null) return;
            bool seq = ChkSeqNumeracao.IsChecked == true;
            TxtSeqPrefixo.IsEnabled = seq;
            TxtSeqFixo.IsEnabled = !seq;
        }

        private void BtnCriarSeq_Click(object sender, RoutedEventArgs e)
        {
            int quant;
            if (!UnitHelper.TryParseInt(TxtSeqQuant.Text, out quant) || quant <= 0)
            {
                SetStatus(ActionResult.Warn("Informe uma quantidade válida para a sequência."));
                return;
            }

            var res = _crachasSeq.CriarSequencia(
                quant,
                GetComboText(CmbSeqTamanho, "10 cm"),
                GetComboText(CmbSeqEspaco, "1 cm"),
                ChkSeqNumeracao.IsChecked == true,
                TxtSeqPrefixo.Text,
                TxtSeqFixo.Text,
                ChkSeqNegrito.IsChecked == true,
                ChkSeqContorno.IsChecked == true);

            SetStatus(res);
        }

        // =====================================================================
        // SEÇÃO 6: IMPRESSÃO / RIP
        // =====================================================================

        private void BtnRipIntersecao_Click(object sender, RoutedEventArgs e)
        {
            int cols, rows;
            double dirCm, baixoCm;
            if (!UnitHelper.TryParseInt(TxtRipCols.Text, out cols) ||
                !UnitHelper.TryParseInt(TxtRipLinhas.Text, out rows) ||
                !UnitHelper.TryParseDouble(TxtRipDir.Text, out dirCm) ||
                !UnitHelper.TryParseDouble(TxtRipBaixo.Text, out baixoCm))
            {
                SetStatus(ActionResult.Warn("Preencha colunas, linhas e emendas (cm) válidos."));
                return;
            }

            var res = _rip.CriarGradeIntersecao(
                cols, rows, dirCm, baixoCm,
                ChkRipBorda.IsChecked == true,
                ChkRipIntersecao.IsChecked == true);
            SetStatus(res);
        }

        private void BtnRipTetris_Click(object sender, RoutedEventArgs e)
        {
            double larg;
            if (!UnitHelper.TryParseDouble(TxtRipTetrisLarg.Text, out larg))
            {
                SetStatus(ActionResult.Warn("Informe uma largura máxima de mídia válida (cm)."));
                return;
            }
            SetStatus(_rip.MontarTetris(larg));
        }
    }
}
