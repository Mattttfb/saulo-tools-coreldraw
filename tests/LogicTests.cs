using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using SauloTools;
using SauloTools.Core;

namespace SauloTools.Tests
{
    public static class Program
    {
        private static int _passed = 0;
        private static int _failed = 0;

        [STAThread]
        public static int Main(string[] args)
        {
            Run("UnitHelper conversions & parsing", TestUnits);
            Run("GeometryMath Intersects & IsInside", TestOverlap);
            Run("GeometryMath ProportionalFit (normal + tiny box fallback)", TestProportionalFit);
            Run("GeometryMath BuildSerpentineOrder (RTL and LTR)", TestSerpentine);
            Run("GeometryMath ComputeFacaRowStats", TestFacaStats);
            Run("GeometryMath PackShelves (Tetris shelf packing + rotation + skip)", TestTetrisShelves);
            Run("SettingsStore JSON round-trip", TestSettingsPersistence);
            Run("WPF DockerUI headless instantiation & 6 accordions", TestWpfDockerInstantiation);

            Console.WriteLine();
            Console.WriteLine(string.Format("RESULT: {0} passed, {1} failed.", _passed, _failed));
            return _failed == 0 ? 0 : 1;
        }

        private static void Run(string name, Action test)
        {
            try
            {
                test();
                _passed++;
                Console.WriteLine(string.Format("[PASS] {0}", name));
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine(string.Format("[FAIL] {0}: {1}", name, ex.Message));
            }
        }

        private static void Assert(bool cond, string msg)
        {
            if (!cond) throw new Exception(msg);
        }

        private static void AssertNear(double actual, double expected, double eps, string msg)
        {
            if (Math.Abs(actual - expected) > eps)
                throw new Exception(string.Format("{0} (expected {1}, got {2})", msg, expected, actual));
        }

        private static void TestUnits()
        {
            AssertNear(UnitHelper.CmToInches(2.54), 1.0, 1e-6, "2.54cm == 1in");
            AssertNear(UnitHelper.MmToInches(25.4), 1.0, 1e-6, "25.4mm == 1in");
            AssertNear(UnitHelper.PtToCm(72.0), 2.54, 1e-6, "72pt == 2.54cm");

            double v;
            Assert(UnitHelper.TryParseDouble("0,75", out v) && Math.Abs(v - 0.75) < 1e-6, "Comma decimal parse");
            Assert(UnitHelper.TryParseDouble("12.5", out v) && Math.Abs(v - 12.5) < 1e-6, "Dot decimal parse");
            Assert(!UnitHelper.TryParseDouble("abc", out v), "Invalid double rejected");
        }

        private static void TestOverlap()
        {
            var outer = new RectBox(0, 0, 10, 10);
            var inner = new RectBox(2, 2, 4, 4);
            var partial = new RectBox(8, 8, 5, 5);
            var outside = new RectBox(10, 0, 5, 5);

            Assert(GeometryMath.IsInside(inner, outer), "inner inside outer");
            Assert(GeometryMath.Intersects(outer, inner), "outer intersects inner");
            Assert(GeometryMath.Intersects(outer, partial), "outer intersects partial");
            Assert(!GeometryMath.Intersects(outer, outside), "touching edge does not intersect");
        }

        private static void TestProportionalFit()
        {
            double w, h;
            GeometryMath.ComputeProportionalFit(10, 10, 20, 10, 1.0, out w, out h);
            AssertNear(w, 8.0, 1e-6, "fit width with 1.0 margin");
            AssertNear(h, 4.0, 1e-6, "proportional height");

            // Tiny box where 2*margin >= parent size -> 80% fallback
            GeometryMath.ComputeProportionalFit(2, 2, 4, 2, 1.5, out w, out h);
            AssertNear(w, 1.6, 1e-6, "80% fallback width");
            AssertNear(h, 0.8, 1e-6, "80% fallback height");
        }

        private static void TestSerpentine()
        {
            // 2 rows x 2 cols:
            // Row 1 (bottom, y=0): idx 0 (x=0), idx 1 (x=10)
            // Row 2 (top, y=10):   idx 2 (x=0), idx 3 (x=10)
            var items = new List<PositionedItem>
            {
                new PositionedItem { Index = 0, CenterX = 0, CenterY = 0, Height = 5 },
                new PositionedItem { Index = 1, CenterX = 10, CenterY = 0, Height = 5 },
                new PositionedItem { Index = 2, CenterX = 0, CenterY = 10, Height = 5 },
                new PositionedItem { Index = 3, CenterX = 10, CenterY = 10, Height = 5 },
            };

            // RTL start: Row1 [1, 0], Row2 [2, 3]
            var rtl = GeometryMath.BuildSerpentineOrder(items, startRightToLeft: true);
            Assert(string.Join(",", rtl) == "1,0,2,3", "RTL serpentine order mismatch: " + string.Join(",", rtl));

            // LTR start: Row1 [0, 1], Row2 [3, 2]
            var ltr = GeometryMath.BuildSerpentineOrder(items, startRightToLeft: false);
            Assert(string.Join(",", ltr) == "0,1,3,2", "LTR serpentine order mismatch: " + string.Join(",", ltr));
        }

        private static void TestFacaStats()
        {
            // maxWidth=100cm, item=9.5cm, gap=0.5cm -> 10 per row
            var st = GeometryMath.ComputeFacaRowStats(100, 9.5, 0.5, 23);
            Assert(st.PerRow == 10, "PerRow expected 10, got " + st.PerRow);
            Assert(st.RemainingToFillRow == 7, "Remaining expected 7, got " + st.RemainingToFillRow);

            var exact = GeometryMath.ComputeFacaRowStats(100, 9.5, 0.5, 20);
            Assert(exact.RemainingToFillRow == 0, "Exact multiple should have 0 remaining");

            // Zero gap (0,0 mm): 10 items of 10cm inside 100cm
            var zeroGap = GeometryMath.ComputeFacaRowStats(100.0 / 2.54, 10.0 / 2.54, 0.0, 23);
            Assert(zeroGap.PerRow == 10, "Zero gap PerRow expected 10, got " + zeroGap.PerRow);
            Assert(zeroGap.RemainingToFillRow == 7, "Zero gap Remaining expected 7, got " + zeroGap.RemainingToFillRow);
        }

        private static void TestTetrisShelves()
        {
            var items = new List<Tuple<int, double, double>>
            {
                Tuple.Create(0, 60.0, 40.0),
                Tuple.Create(1, 40.0, 50.0),
                Tuple.Create(2, 130.0, 130.0) // exceeds 120cm both orientations -> skipped
            };

            double offset = 0.0;
            var plan = GeometryMath.PackShelves(items, 120.0, 200.0, ref offset);
            Assert(plan.Count == 3, "Plan count == 3");
            Assert(!plan[0].Skipped, "Item 0 placed");
            Assert(!plan[1].Skipped, "Item 1 placed");
            Assert(plan[2].Skipped, "Oversized item 2 skipped");
            Assert(offset > 0, "Group offset advanced");
        }

        private static void TestSettingsPersistence()
        {
            string tmp = Path.Combine(Path.GetTempPath(), "saulotools_test_" + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                var s = new SauloSettings
                {
                    SangriaCor = "Vermelho",
                    FacaQuantidade = "77",
                    RipColunas = "3",
                    ExpRip = true
                };
                SettingsStore.Save(s, tmp);
                var loaded = SettingsStore.Load(tmp);
                Assert(loaded.SangriaCor == "Vermelho", "SangriaCor persisted");
                Assert(loaded.FacaQuantidade == "77", "FacaQuantidade persisted");
                Assert(loaded.RipColunas == "3", "RipColunas persisted");
                Assert(loaded.ExpRip, "ExpRip persisted");
            }
            finally
            {
                if (File.Exists(tmp)) File.Delete(tmp);
            }
        }

        private static void TestWpfDockerInstantiation()
        {
            var ui = new DockerUI((object)null);
            Assert(ui != null, "DockerUI instantiated");
            Assert(ui.FindName("ExpSangrias") != null, "ExpSangrias exists");
            Assert(ui.FindName("ExpFaca") != null, "ExpFaca exists");
            Assert(ui.FindName("ExpAgrupar") != null, "ExpAgrupar exists");
            Assert(ui.FindName("ExpQuick") != null, "ExpQuick exists");
            Assert(ui.FindName("ExpCrachasSeq") != null, "ExpCrachasSeq exists");
            Assert(ui.FindName("ExpRip") != null, "ExpRip exists");
            var cmb = ui.FindName("CmbFacaEspaco") as System.Windows.Controls.ComboBox;
            Assert(cmb != null && cmb.Items.Count == 3, "CmbFacaEspaco should have 3 items (0,0 mm / 0,5mm / 1Cm)");
            var firstItem = cmb.Items[0] as System.Windows.Controls.ComboBoxItem;
            Assert(firstItem != null && firstItem.Content.ToString() == "0,0 mm", "First option must be 0,0 mm");
        }
    }
}
