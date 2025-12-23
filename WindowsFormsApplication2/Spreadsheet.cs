using System;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXMLSpreadsheet;
using Zibs.Configuration;
using Zibs.ExtensionClasses;
using Zibs.tekstManager;

namespace Zibs
{
    namespace ZibExtraction
    {
        partial class ZIB
        {

            public void saveToSpeadsheet(string fileName, string zibText)
            {
                string subject = "";
                string description = "";
                //                File.WriteAllText(fileName.Replace("xlsx", "txt"), zibText);
                using (SpreadsheetDocument document = openXML.CreateSpreadsheet(fileName))
                {
                    if (document == null) return;

                    textManager tm = new textManager("ZibExtractionLabels.cfg");
                    if (tm.dictionaryOK)
                    {
                        tm.Language = Settings.zibcontext.pubLanguage;
                        subject = tm.getLabel("xlsFileSubject");
                        description = tm.getLabel("xlsFileDescription");
                    }
                    document.AddApplicationInfo("ZibExtraction", "2.0",  "Nictiz");
                    document.AddDocumentInfo(zibName.fileName(Fullname, currentLanguage), subject, description, "Zib2Xlsx", this.Status);
                    string aLine, control, command;
                    string innerColor, textColor, textWeight;
                    string[] commandParts;
                    string imageFile;
                    Worksheet mySheet = null;
                    bool newSheet = false;
                    uint iRow = 0;
                    uint nCells;
                    uint[] iOrigin = new uint[2] { 1, 1 };
                    object after = Type.Missing;
                    Dictionary<string, uint[]> sheetsInBook = new Dictionary<string, uint[]>();
                    float imageTop = 50;
                    float imageLeft = 50;

                    using (StringReader sr = new StringReader(zibText))
                    {
                        while (true)
                        {
                            aLine = sr.ReadLine();
                            if (!string.IsNullOrEmpty(aLine))
                            {
                                if (controlInfo(aLine, out control, out command))
                                {
                                    switch (control)
                                    {
                                        case "Sheet":
                                            if (mySheet != null) postprocessSheet(mySheet, iOrigin);
                                            commandParts = command.Split(';');
                                            string sheetName = commandParts[0];

                                            // 17-04-25 : patch voor te lange xls sheet namen
                                            // afkappen werd al in AddWorksheet gedaan maar niet in de sheetsInBook dictionary waardoor een
                                            // net bestaande key gezocht werd.
                                            if (sheetName.Length > 31)
                                            {
                                                OnNewMessage(new ErrorMessageEventArgs(ErrorType.warning, "Naam "+ commandParts[0] + " te lang (>31) voor tabnaam, tabnaam ingekort."));
                                                sheetName = sheetName.Substring(0, 31);
                                            }

                                            mySheet = document.AddWorksheet(sheetName,ref newSheet );
                                            if (newSheet)
                                            {
                                                iOrigin = new uint[2] { 1, 1 };
                                                if (commandParts.Count() == 3)
                                                {
                                                    iOrigin[0] = uint.TryParse(commandParts[1], out iOrigin[0]) ? iOrigin[0] : 1;
                                                    iOrigin[1] = uint.TryParse(commandParts[2], out iOrigin[1]) ? iOrigin[1] : 1;
                                                }
                                                else
                                                {
                                                    iOrigin[0] = 1;
                                                    iOrigin[1] = 1;
                                                }
                                                sheetsInBook.Add(sheetName, iOrigin);
                                                iRow = iOrigin[0] - 1;
                                                OnNewMessage(new ErrorMessageEventArgs(ErrorType.information, "Sheet: " + mySheet.Name()));
                                            }
                                            else
                                            {
                                                iOrigin = sheetsInBook[mySheet.Name()];
                                                iRow = GetIndex(mySheet, SpecialCell.endRow);
                                            }
                                            break;
                                        case "Merge":
                                            string[] merge = command.Split(';');
                                            uint iCol = iOrigin[1];
                                            foreach (string m in merge)
                                                if (uint.TryParse(m, out nCells))
                                                {
                                                    mySheet.MergeCells(iRow, iCol, iRow, iCol + nCells - 1);
                                                    iCol += nCells;
                                                }
                                            break;
                                        case "RowColor":
                                            string[] colorInfo = command.Split(';');
                                            innerColor = colorInfo[0].Trim().Replace("#","FF");
                                            textColor = colorInfo[1].Trim().Replace("#", "FF");
                                            textWeight = colorInfo[2].Trim();
                                            setRowDecoration(ref mySheet, iRow, iOrigin[1], innerColor, textColor, textWeight, true);
                                            break;
                                        case "Header": 
                                            innerColor = "#000099".Replace("#", "FF");
                                            textColor = "#FFFFFF".Replace("#", "FF");
                                            textWeight = "Bold";
                                            setRowDecoration(ref mySheet, iRow, iOrigin[1], innerColor, textColor, textWeight);
                                            break;
                                        case "Subheader":
                                            innerColor = "#D3D3D3".Replace("#","FF");
                                            textColor = "#000000".Replace("#","FF");
                                            textWeight = "Bold";
                                            setRowDecoration(ref mySheet, iRow, iOrigin[1], innerColor, textColor, textWeight);
                                            break;
                                        case "ColWidth":
                                            string[] width = command.Split(';');
                                            if (width[0].ToLower() == "auto")
                                                mySheet.SetColumnWidth(iOrigin[1], iOrigin[1] + 5, 25, autofit: true);
                                            else
                                                for (uint i = 0; i < width.Count(); i++)
                                                    mySheet.SetColumnWidth(iOrigin[1] + i, iOrigin[1] + i, double.Parse(width[i]));
                                            break;
                                        case "Image": // Absolute plaatsing
                                            commandParts = command.Split(';');
                                            imageFile = Path.Combine(commandParts[1], commandParts[0]);
                                            if (File.Exists(imageFile))
                                            {
                                                mySheet.InsertImageAbsolute(imageFile, (long)imageLeft, (long)imageTop, null, null);
                                            }
                                            else
                                            {
                                                fillRow("Bestand niet gevonden: " + imageFile, ref iRow, iOrigin[1], ref mySheet);
                                                OnNewMessage(new ErrorMessageEventArgs(ErrorType.warning, "Bestand niet gevonden: " + imageFile));
                                            }
                                            //imageFile = Path.Combine(Settings.userPreferences.DiagramLocation, command);
                                            break;
                                        case "ImageRelative": //onCellAnchor relative plaatsing tov opgegeven cel
                                            int result, imageRow, imageCol;
                                            long longResult;
                                            long? imageWidth = null;
                                            long? imageHeight = null;
                                            long offsetX, offsetY;
                                            commandParts = command.Split(';');
                                            imageFile = Path.Combine(commandParts[1], commandParts[0]);
                                            imageRow = commandParts.Count() > 2 ? (int.TryParse(commandParts[2], out result) ? result : 1) : 1;
                                            imageCol = commandParts.Count() > 3 ? (int.TryParse(commandParts[3], out result) ? result : 1) : 1;
                                            offsetX = commandParts.Count() > 4 ? (int.TryParse(commandParts[4], out result) ? result : 0) : 0;
                                            offsetY = commandParts.Count() > 5 ? (int.TryParse(commandParts[5], out result) ? result : 0) : 0;
                                            if (commandParts.Count() > 6 && long.TryParse(commandParts[6], out longResult))
                                            {
                                                imageWidth = longResult;
                                            }
                                            if (commandParts.Count() > 7 && long.TryParse(commandParts[7], out longResult))
                                            {
                                                imageHeight = longResult;
                                            }
                                            if (File.Exists(imageFile))
                                            {
                                                mySheet.InsertImageRelative(imageFile, imageRow, imageCol, offsetX, offsetY, imageWidth, imageHeight);
                                            }
                                            else
                                            {
                                                fillRow("Bestand niet gevonden: " + imageFile, ref iRow, iOrigin[1], ref mySheet);
                                                OnNewMessage(new ErrorMessageEventArgs(ErrorType.warning, "Bestand niet gevonden: " + imageFile));
                                            }
                                            break;
                                        case "IgnoreNumberAsText":
                                            uint sRow = 0, sCol = 0, eRow = 0, eCol = 0;
                                            string[] dimParts = command.Split(';');
                                            if (dimParts.Count() < 4) break;
                                            if (uint.TryParse(dimParts[0], out sRow) && uint.TryParse(dimParts[1], out sCol) && uint.TryParse(dimParts[2], out eRow) && uint.TryParse(dimParts[3], out eCol) != true) break;
                                            string ignoreRange = Range.ToReference(iOrigin[0] + sRow - 1, iOrigin[1] + sCol - 1, iOrigin[0] + eRow - 1, iOrigin[1] + eCol - 1);
                                            mySheet.AddIgnoredError(ignoreRange, IgnoredErrorProperty.numberStoredAsText, set: true);
                                            break;
                                        default:
                                            string message = "Onbekende opdracht: " + control + command;
                                            fillRow(message, ref iRow, iOrigin[1], ref mySheet);
                                            break;
                                    }

                                }                                else
                                    fillRow(aLine, ref iRow, iOrigin[1], ref mySheet);
                            }
                            else
                                if (aLine == null) break;
                        }
                    }
                    //             Laatste sheet bijwerken
                    if (mySheet != null) postprocessSheet(mySheet, iOrigin);
                    document.SelectWorksheet("Data");
                    document.Dispose();
                }
            }


            private void postprocessSheet(Worksheet thisSheet, uint[] iOrigin)
            {
//                string cellText;
                string _rangeReference = thisSheet.SheetDimension.Reference;
                IEnumerable<Cell> _range = thisSheet.GetCells(ref _rangeReference);
                IEnumerable<Cell> conceptRange;

                thisSheet.SetTextAlignment(_rangeReference, VerticalAlignmentValues.Top, true); 
                BorderTypeArray borderTypeArray = new BorderTypeArray();
                borderTypeArray.All = BorderType.Thin;
                thisSheet.SetBorder(_rangeReference, borderTypeArray);
                // special processing
                if (thisSheet.Name() == "Data")
                {
                    string conceptRangeReference = Range.ToReference(iOrigin[0], iOrigin[1], GetIndex(thisSheet, SpecialCell.endRow), iOrigin[1] + (uint)maxIndent);  // -1
                    conceptRange = thisSheet.GetCells(ref conceptRangeReference);
                    thisSheet.SetTextWrap(conceptRangeReference, false);
                    //Set the thin lines style.
                    borderTypeArray.Reset();
                    borderTypeArray.Outside = BorderType.Thin;
                    borderTypeArray.Horizontal = BorderType.Thin;
                    borderTypeArray.Vertical = BorderType.None;
                    thisSheet.SetBorder(conceptRangeReference, borderTypeArray);
                    Row row;
                    for (uint i = GetIndex(thisSheet, SpecialCell.startRow); i < GetIndex(thisSheet, SpecialCell.endRow) + 1; i++)
                    {
                        row = thisSheet.GetRow(i);
                        if (row.Height != null && row.Height > 50)
                            row.Height = 50;
                    }
                }
            }

            private void setRowDecoration(ref Worksheet _sheet, uint iRow, uint startCol, string innerColor, string textColor, string textWeight, bool mergeExpand = false)
            {
                uint endRow = iRow;
                bool Bold = false, Italic = false;
                if (textWeight.ToLower() == "bold")
                    Bold = true;
                else if (textWeight.ToLower() == "italic")
                    Italic = true;

                uint endCol = GetIndex(_sheet, SpecialCell.endColumn);
                // breidt als de rij verticale mergecellen heeft de selectie uit, indien met 'mergeExpand' aangegeven is dat dat gewenst is
                if (mergeExpand)
                {
                    for (uint col = startCol; col <= endCol; col++)
                    {
                        bool merged = _sheet.IsMerged(Range.ToReference(iRow, col), out string mergeRangeRef);
                        if (merged)
                        {
                            Range.FromReference(mergeRangeRef, out uint sR, out uint sC, out uint eR, out uint eC);
                            if (sR < iRow) iRow = sR;
                            if (eR > endRow) endRow = eR;
                        }
                    }
                }
                string range = Range.ToReference(iRow, startCol, endRow, endCol);
                _sheet.SetDecoration(range, "Calibri", 10, textColor, innerColor, Bold, Italic);
            } 

            private bool controlInfo(string aLine, out string control, out string command)
            {
                bool controlLine = false;
                if (aLine.Length > 2 && aLine.Substring(0, 2) == "<<" && aLine.Substring(2).Contains(">>"))
                {
                    controlLine = true;
                    control = aLine.Substring(2, aLine.IndexOf(">>") - 2);
                    command = aLine.Substring(aLine.IndexOf(">>") + 2).Trim();
                }
                else
                {
                    control = "";
                    command = "";
                }
                return controlLine;
            }

            private void fillRow(string aLine, ref uint rowNumber, uint startColumn, ref Worksheet thisSheet)
            {
                int mergeRows = 1;
                bool[] splittedColums = new bool[1000];
                string[] splitText = null;
                string[] splitCell = null;
                rowNumber++;
                splitText = aLine.Split(';');
                List<Cell> _range = thisSheet.GetCells(rowNumber, startColumn, rowNumber, startColumn + (uint)splitText.Count() - 1).ToList();
                for (int j = 0; j < splitText.Length; j++)
                {
                    string xv = splitText[j].decodeForXLS().Trim();         // 10-02-2021 decodeforXLS toegevoegd
                    if (xv.Length > 0)
                    {
                        splitCell = Regex.Split(xv, Settings.xlsSplitCell);
                        if (splitCell.Count() > 1)
                        {
                            mergeRows = splitCell.Count() > mergeRows ? splitCell.Count() : mergeRows;
                            splittedColums[startColumn + j] = true;
                        }
                        for (int k = 0; k < splitCell.Count(); k++)
                        {
                            string xv2 = splitCell[k].Trim();
                            int s = xv2.IndexOf("[Hyperlink:");
                            if (s > -1)
                            {
                                int e = xv2.IndexOf("]");
                                string target = xv2.Substring(s + 11, e - s - 11);
                                xv2 = xv2.Substring(e + 1);
                                if (target.IndexOf("http") > -1)
                                {
                                    Uri uri = new Uri(target, UriKind.Absolute);
                                    thisSheet.AddHyperlink(_range[j].CellReference, uri);
                                }
                                else
                                    thisSheet.AddHyperlink(_range[j].CellReference, target);
                            }
                            //                        Range[j].SetValue(xv, CellValues.String);  xv.decodeForXLS()
                            _range[j].SetRTFValue(xv2.decodeForXLS());

                            if (k != splitCell.Count() - 1)
                            {
                                _range = thisSheet.AddCopiedRow(ref rowNumber, true).ToList();
                                //rowNumber++;
                                //_range = thisSheet.GetCells(rowNumber, startColumn, rowNumber, startColumn + (uint)splitText.Count() - 1).ToList();
                                
                            }
                        }
                        //                            Range[j].SetRTFValue(xv.decodeForXLS());
                    }
                }
                if (mergeRows > 1)
                {
                    //for (uint j = startColumn + (uint)maxIndent + 1; j < startColumn + (uint)splitText.Length; j++)
                    for (uint j = startColumn; j < startColumn + (uint)splitText.Length; j++)
                    {
                        if (!splittedColums[j])
                        {
                            var x = thisSheet.GetFirstChild<MergeCells>()?.ChildElements.Where(y => ((MergeCell)y).Reference.Value.Contains(Extensions.ColumnNameFromIndex(j)))?? Enumerable.Empty<MergeCell>(); 
                            if (x.Count() > 0)
                            {
                                var range = (MergeCell)(x.First());
                            }
                            thisSheet.MergeCells(rowNumber - (uint)mergeRows + 1, j, rowNumber, j);
                        }
                    }
                    BorderTypeArray borderTypeArray = new BorderTypeArray();
                    borderTypeArray.Outside = BorderType.Thin;
                    borderTypeArray.Horizontal = BorderType.Thin;
                    borderTypeArray.Vertical = BorderType.None;
                    string mergeRange = Range.ToReference(rowNumber - (uint)mergeRows + 1, startColumn + (uint)maxIndent + 1, rowNumber, startColumn + (uint)splitText.Length - 1);  // -1
                    thisSheet.SetBorder(mergeRange, borderTypeArray);
                }
            }

            private uint GetIndex(Worksheet sheet, SpecialCell s)
            {
                uint sRow, sCol, eRow, eCol, index =0 ;
                Range.FromReference(sheet.SheetDimension.Reference, out sRow, out sCol, out eRow, out eCol);
                switch (s)
                {
                    case SpecialCell.startRow:
                        index = sRow;
                        break;
                    case SpecialCell.startColumn:
                        index = sCol;
                        break;
                    case SpecialCell.endRow:
                        index = eRow;
                        break;
                    case SpecialCell.endColumn:
                        index = eCol;
                        break;
                }
                return index;
            }

            public enum SpecialCell
            {
                startRow,
                startColumn,
                endRow,
                endColumn
            }
        }
    }
}
