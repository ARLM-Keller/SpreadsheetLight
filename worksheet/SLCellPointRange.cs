using System;
using System.Collections.Generic;

namespace SpreadsheetLight
{
    internal struct SLCellPointRange
    {
        internal int StartRowIndex;
        internal int StartColumnIndex;
        internal int EndRowIndex;
        internal int EndColumnIndex;

        internal SLCellPointRange(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            this.StartRowIndex = StartRowIndex;
            this.StartColumnIndex = StartColumnIndex;
            this.EndRowIndex = EndRowIndex;
            this.EndColumnIndex = EndColumnIndex;
        }
    }

    internal class SLCellPointRangeComparer : IComparer<SLCellPointRange>
    {
        public int Compare(SLCellPointRange pt1, SLCellPointRange pt2)
        {
            if (pt1.StartRowIndex < pt2.StartRowIndex)
            {
                return -1;
            }
            else if (pt1.StartRowIndex > pt2.StartRowIndex)
            {
                return 1;
            }
            else
            {
                if (pt1.StartColumnIndex < pt2.StartColumnIndex)
                {
                    return -1;
                }
                else if (pt1.StartColumnIndex > pt2.StartColumnIndex)
                {
                    return 1;
                }
                else
                {
                    if (pt1.EndRowIndex < pt2.EndRowIndex)
                    {
                        return -1;
                    }
                    else if (pt1.EndRowIndex > pt2.EndRowIndex)
                    {
                        return 1;
                    }
                    else
                    {
                        return pt1.EndColumnIndex.CompareTo(pt2.EndColumnIndex);
                    }
                }
            }
        }
    }
}
