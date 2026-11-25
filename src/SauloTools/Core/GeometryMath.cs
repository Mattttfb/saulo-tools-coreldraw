using System;
using System.Collections.Generic;

namespace SauloTools.Core
{
    public struct RectBox
    {
        public double Left;
        public double Bottom;
        public double Width;
        public double Height;

        public double Right => Left + Width;
        public double Top => Bottom + Height;
        public double CenterX => Left + (Width / 2.0);
        public double CenterY => Bottom + (Height / 2.0);
        public double Area => Width * Height;

        public RectBox(double left, double bottom, double width, double height)
        {
            Left = left;
            Bottom = bottom;
            Width = width;
            Height = height;
        }
    }

    public struct PositionedItem
    {
        public int Index;
        public double CenterX;
        public double CenterY;
        public double Height;
    }

    public struct ShelfPlanItem
    {
        public int OriginalIndex;
        public double LeftX;
        public double TopY;
        public bool Rotated90;
        public bool Skipped;
        public double FinalWidth;
        public double FinalHeight;
    }

    public struct FacaStats
    {
        public int PerRow;
        public int RemainingToFillRow;
    }

    /// <summary>
    /// Pure geometric & layout math extracted from the VBA macros so it can be unit-tested
    /// deterministically outside CorelDRAW.
    /// </summary>
    public static class GeometryMath
    {
        public static bool Intersects(RectBox a, RectBox b)
        {
            if (a.Right <= b.Left) return false;
            if (a.Left >= b.Right) return false;
            if (a.Top <= b.Bottom) return false;
            if (a.Bottom >= b.Top) return false;
            return true;
        }

        public static bool IsInside(RectBox inner, RectBox outer)
        {
            return inner.Left >= outer.Left &&
                   inner.Right <= outer.Right &&
                   inner.Top <= outer.Top &&
                   inner.Bottom >= outer.Bottom;
        }

        public static void ComputeProportionalFit(
            double parentW, double parentH,
            double childW, double childH,
            double marginDoc,
            out double newW, out double newH)
        {
            newW = childW;
            newH = childH;
            if (childW <= 0 || childH <= 0) return;

            double maxW = parentW - (2.0 * marginDoc);
            double maxH = parentH - (2.0 * marginDoc);
            if (maxW <= 0 || maxH <= 0)
            {
                maxW = parentW * 0.8;
                maxH = parentH * 0.8;
            }

            double scale = Math.Min(maxW / childW, maxH / childH);
            newW = childW * scale;
            newH = childH * scale;
        }

        /// <summary>
        /// Computes serpentine Z-order (bottom-to-top rows, alternating horizontal scan direction).
        /// </summary>
        public static List<int> BuildSerpentineOrder(IList<PositionedItem> items, bool startRightToLeft)
        {
            var result = new List<int>();
            if (items == null || items.Count == 0) return result;

            double refHeight = items[0].Height;
            if (refHeight <= 0.0001) refHeight = 1.0;

            var sorted = new List<PositionedItem>(items);
            sorted.Sort((a, b) => a.CenterY.CompareTo(b.CenterY));

            int i = 0;
            int rowNumber = 0;
            while (i < sorted.Count)
            {
                long currentRow = (long)Math.Round(sorted[i].CenterY / refHeight, MidpointRounding.AwayFromZero);
                rowNumber++;

                int start = i;
                int end = i;
                while (end + 1 < sorted.Count)
                {
                    long nextRow = (long)Math.Round(sorted[end + 1].CenterY / refHeight, MidpointRounding.AwayFromZero);
                    if (nextRow != currentRow) break;
                    end++;
                }

                // Sort this row by CenterX ascending
                sorted.Sort(start, end - start + 1, Comparer<PositionedItem>.Create((a, b) => a.CenterX.CompareTo(b.CenterX)));

                bool rightToLeft = (rowNumber % 2 == 1) ? startRightToLeft : !startRightToLeft;
                if (rightToLeft)
                {
                    for (int j = end; j >= start; j--) result.Add(sorted[j].Index);
                }
                else
                {
                    for (int j = start; j <= end; j++) result.Add(sorted[j].Index);
                }

                i = end + 1;
            }

            return result;
        }

        public static FacaStats ComputeFacaRowStats(double maxWidthInches, double firstItemWidthInches, double gapInches, int totalCreated)
        {
            double step = firstItemWidthInches + gapInches;
            int perRow = step > 0.00001 ? (int)Math.Floor((maxWidthInches + 1e-6) / step) : 1;
            if (perRow < 1) perRow = 1;

            int rem = perRow - (totalCreated % perRow);
            if (rem == perRow) rem = 0;

            return new FacaStats { PerRow = perRow, RemainingToFillRow = rem };
        }

        private class Shelf
        {
            public double TopY;
            public double Height;
            public double UsedWidth;
        }

        /// <summary>
        /// Pure shelf-packing algorithm matching RIP.vba MontarGrupoPrateleira.
        /// </summary>
        public static List<ShelfPlanItem> PackShelves(
            IList<Tuple<int, double, double>> items,
            double maxWidthCm,
            double pageTopYCm,
            ref double groupOffsetCm)
        {
            var shelves = new List<Shelf>();
            var plan = new List<ShelfPlanItem>(items.Count);

            foreach (var item in items)
            {
                int origIndex = item.Item1;
                double w = item.Item2;
                double h = item.Item3;

                int bestShelf = -1;
                bool useRot = false;
                double minWaste = double.MaxValue;

                for (int k = 0; k < shelves.Count; k++)
                {
                    double remaining = maxWidthCm - shelves[k].UsedWidth;
                    if (remaining >= w)
                    {
                        double waste = remaining - w;
                        if (waste < minWaste)
                        {
                            minWaste = waste;
                            bestShelf = k;
                            useRot = false;
                        }
                    }

                    double wRot = h;
                    if (remaining >= wRot)
                    {
                        double waste = remaining - wRot;
                        if (waste < minWaste)
                        {
                            minWaste = waste;
                            bestShelf = k;
                            useRot = true;
                        }
                    }
                }

                if (bestShelf >= 0)
                {
                    double finalW = useRot ? h : w;
                    double finalH = useRot ? w : h;
                    var shelf = shelves[bestShelf];

                    plan.Add(new ShelfPlanItem
                    {
                        OriginalIndex = origIndex,
                        LeftX = shelf.UsedWidth,
                        TopY = shelf.TopY,
                        Rotated90 = useRot,
                        Skipped = false,
                        FinalWidth = finalW,
                        FinalHeight = finalH
                    });

                    shelf.UsedWidth += finalW;
                    if (finalH > shelf.Height) shelf.Height = finalH;
                }
                else
                {
                    double wRot = h;
                    double hRot = w;
                    bool rotateNew = false;

                    if (w > maxWidthCm && wRot <= maxWidthCm)
                        rotateNew = true;
                    else if (w <= maxWidthCm && wRot <= maxWidthCm && hRot < h)
                        rotateNew = true;

                    double finalW = rotateNew ? wRot : w;
                    double finalH = rotateNew ? hRot : h;

                    if (finalW > maxWidthCm)
                    {
                        plan.Add(new ShelfPlanItem
                        {
                            OriginalIndex = origIndex,
                            Rotated90 = rotateNew,
                            Skipped = true,
                            FinalWidth = finalW,
                            FinalHeight = finalH
                        });
                        continue;
                    }

                    double topY = shelves.Count == 0
                        ? pageTopYCm - groupOffsetCm
                        : shelves[shelves.Count - 1].TopY - shelves[shelves.Count - 1].Height;

                    var newShelf = new Shelf
                    {
                        TopY = topY,
                        Height = finalH,
                        UsedWidth = finalW
                    };
                    shelves.Add(newShelf);

                    plan.Add(new ShelfPlanItem
                    {
                        OriginalIndex = origIndex,
                        LeftX = 0,
                        TopY = topY,
                        Rotated90 = rotateNew,
                        Skipped = false,
                        FinalWidth = finalW,
                        FinalHeight = finalH
                    });
                }
            }

            if (shelves.Count > 0)
            {
                var last = shelves[shelves.Count - 1];
                double groupHeight = (pageTopYCm - groupOffsetCm) - (last.TopY - last.Height);
                groupOffsetCm += groupHeight;
            }

            return plan;
        }
    }
}
