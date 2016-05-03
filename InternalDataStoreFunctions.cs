﻿using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLight
{
    partial class SLDocument
    {
        internal int SaveToStylesheet(string Hash)
        {
            int index = 0;
            if (dictStyleHash.ContainsKey(Hash))
            {
                index = dictStyleHash[Hash];
            }
            else
            {
                index = listStyle.Count;
                listStyle.Add(Hash);
                dictStyleHash[Hash] = index;
            }

            return index;
        }

        internal int ForceSaveToStylesheet(string Hash)
        {
            int index = listStyle.Count;
            listStyle.Add(Hash);
            dictStyleHash[Hash] = index;

            return index;
        }

        internal int SaveToStylesheetNumberingFormat(string Hash)
        {
            int index = 0;
            if (dictStyleNumberingFormatHash.ContainsKey(Hash))
            {
                index = dictStyleNumberingFormatHash[Hash];
            }
            else if (dictBuiltInNumberingFormatHash.ContainsKey(Hash))
            {
                index = dictBuiltInNumberingFormatHash[Hash];
            }
            else
            {
                index = NextNumberFormatId;
                ++NextNumberFormatId;
                dictStyleNumberingFormat[index] = Hash;
                dictStyleNumberingFormatHash[Hash] = index;
            }

            return index;
        }

        internal int ForceSaveToStylesheetNumberingFormat(int index, string Hash)
        {
            dictStyleNumberingFormat[index] = Hash;
            dictStyleNumberingFormatHash[Hash] = index;

            return index;
        }

        internal int SaveToStylesheetFont(string Hash)
        {
            int index = 0;
            if (dictStyleFontHash.ContainsKey(Hash))
            {
                index = dictStyleFontHash[Hash];
            }
            else
            {
                index = listStyleFont.Count;
                listStyleFont.Add(Hash);
                dictStyleFontHash[Hash] = index;
            }

            return index;
        }

        internal int ForceSaveToStylesheetFont(string Hash)
        {
            int index = listStyleFont.Count;
            listStyleFont.Add(Hash);
            dictStyleFontHash[Hash] = index;

            return index;
        }

        internal int SaveToStylesheetFill(string Hash)
        {
            int index = 0;
            if (dictStyleFillHash.ContainsKey(Hash))
            {
                index = dictStyleFillHash[Hash];
            }
            else
            {
                index = listStyleFill.Count;
                listStyleFill.Add(Hash);
                dictStyleFillHash[Hash] = index;
            }

            return index;
        }

        internal int ForceSaveToStylesheetFill(string Hash)
        {
            int index = listStyleFill.Count;
            listStyleFill.Add(Hash);
            dictStyleFillHash[Hash] = index;

            return index;
        }

        internal int SaveToStylesheetBorder(string Hash)
        {
            int index = 0;
            if (dictStyleBorderHash.ContainsKey(Hash))
            {
                index = dictStyleBorderHash[Hash];
            }
            else
            {
                index = listStyleBorder.Count;
                listStyleBorder.Add(Hash);
                dictStyleBorderHash[Hash] = index;
            }

            return index;
        }

        internal int ForceSaveToStylesheetBorder(string Hash)
        {
            int index = listStyleBorder.Count;
            listStyleBorder.Add(Hash);
            dictStyleBorderHash[Hash] = index;

            return index;
        }

        internal int SaveToStylesheetCellStylesFormat(string Hash)
        {
            int index = 0;
            if (dictStyleCellStyleFormatHash.ContainsKey(Hash))
            {
                index = dictStyleCellStyleFormatHash[Hash];
            }
            else
            {
                index = listStyleCellStyleFormat.Count;
                listStyleCellStyleFormat.Add(Hash);
                dictStyleCellStyleFormatHash[Hash] = index;
            }

            return index;
        }

        internal int ForceSaveToStylesheetCellStylesFormat(string Hash)
        {
            int index = listStyleCellStyleFormat.Count;
            listStyleCellStyleFormat.Add(Hash);
            dictStyleCellStyleFormatHash[Hash] = index;

            return index;
        }

        internal int SaveToStylesheetCellStyle(string Hash)
        {
            int index = 0;
            if (dictStyleCellStyleHash.ContainsKey(Hash))
            {
                index = dictStyleCellStyleHash[Hash];
            }
            else
            {
                index = listStyleCellStyle.Count;
                listStyleCellStyle.Add(Hash);
                dictStyleCellStyleHash[Hash] = index;
            }

            return index;
        }

        internal int ForceSaveToStylesheetCellStyle(string Hash)
        {
            int index = listStyleCellStyle.Count;
            listStyleCellStyle.Add(Hash);
            dictStyleCellStyleHash[Hash] = index;

            return index;
        }

        internal int SaveToStylesheetDifferentialFormat(string Hash)
        {
            int index = 0;
            if (dictStyleDifferentialFormatHash.ContainsKey(Hash))
            {
                index = dictStyleDifferentialFormatHash[Hash];
            }
            else
            {
                index = listStyleDifferentialFormat.Count;
                listStyleDifferentialFormat.Add(Hash);
                dictStyleDifferentialFormatHash[Hash] = index;
            }

            return index;
        }

        internal int ForceSaveToStylesheetDifferentialFormat(string Hash)
        {
            int index = listStyleDifferentialFormat.Count;
            listStyleDifferentialFormat.Add(Hash);
            dictStyleDifferentialFormatHash[Hash] = index;

            return index;
        }

        internal int SaveToSharedStringTable(string Hash)
        {
            int index = 0;
            if (dictSharedStringHash.ContainsKey(Hash))
            {
                index = dictSharedStringHash[Hash];
            }
            else
            {
                index = listSharedString.Count;
                listSharedString.Add(Hash);
                dictSharedStringHash[Hash] = index;
            }

            return index;
        }

        internal void ForceSaveToSharedStringTable(SharedStringItem ssi)
        {
            int index = listSharedString.Count;
            string sHash = SLTool.RemoveNamespaceDeclaration(ssi.InnerXml);
            listSharedString.Add(sHash);
            dictSharedStringHash[sHash] = index;
        }

        internal int DirectSaveToSharedStringTable(string Data)
        {
            int index = 0;
            string sHash;
            if (SLTool.ToPreserveSpace(Data))
            {
                sHash = string.Format("<x:t xml:space=\"preserve\">{0}</x:t>", Data);
            }
            else
            {
                sHash = string.Format("<x:t>{0}</x:t>", Data);
            }

            if (dictSharedStringHash.ContainsKey(sHash))
            {
                index = dictSharedStringHash[sHash];
            }
            else
            {
                index = listSharedString.Count;
                listSharedString.Add(sHash);
                dictSharedStringHash[sHash] = index;
            }

            return index;
        }

        internal int DirectSaveToSharedStringTable(InlineString Data)
        {
            int index = 0;
            string sHash = SLTool.RemoveNamespaceDeclaration(Data.InnerXml);
            if (dictSharedStringHash.ContainsKey(sHash))
            {
                index = dictSharedStringHash[sHash];
            }
            else
            {
                index = listSharedString.Count;
                listSharedString.Add(sHash);
                dictSharedStringHash[sHash] = index;
            }

            return index;
        }
    }
}
