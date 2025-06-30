using System;
using System.Text;
using System.Linq;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using Zibs.ExtensionClasses;
using Spire.Doc;
using Spire.Doc.Fields;
using Spire.Doc.Documents;



namespace Zibs
{
    namespace ZibExtraction
    {
        public class zibExample
        {
            int informationModelWidth = 1500; // breedte van de informatie model tabel (als maximale breedte)

            public delegate void NewMessageEventHandler(object sender, ErrorMessageEventArgs e);
            public event NewMessageEventHandler NewMessage;

            protected virtual void OnNewMessage(ErrorMessageEventArgs e)
            {
                NewMessage?.Invoke(this, e);
            }

            public string ReadContent(string filename)
            {
                string error;
                string wikiText = "";
                Document thisDocument = null;

                try
                {
                    if (!File.Exists(filename)) throw new FileNotFoundException();

                    thisDocument = new Document();
                    thisDocument.LoadFromFileInReadMode(filename, FileFormat.Docx);  // was vroeger FileFormat.Doc
                    foreach (Table thisTable in thisDocument.Sections[0].Tables)
                    {
                        wikiText += ReadTable(thisTable, out error);
                        if (error != "") OnNewMessage(new ErrorMessageEventArgs(ErrorType.error, "Fouten in voorbeeld generatie: " + error));
                    }
                    OnNewMessage(new ErrorMessageEventArgs(ErrorType.information, "Voorbeeldtabel aangemaakt. Aantal: " + thisDocument.Sections[0].Tables.Count));
                }
                catch (Exception e) //COMException
                {
                    wikiText = "Voorbeeld file fout: " + e.Message + " : " + filename;
                    OnNewMessage(new ErrorMessageEventArgs(ErrorType.error, "Voorbeeld file fout: " + e.Message + " : " + filename));
                }
                finally
                {
                    if (thisDocument != null)
                    {
                        thisDocument.Close();
                        thisDocument.Dispose();
                    }
                }
                return wikiText;
            }

            private int columnCount(Table table)
            {
                var j = table.Rows.OfType<TableRow>().Select(x => x.Cells.Count).Max();
                return j;
            }

            public string ReadTable(Table thisTable, out string error)
            {
                StringBuilder content = new StringBuilder();
                int iRow = 0;
                int iCol = 0;
                int iRowTemp;
                int colSpan, rowSpan;
                string wikiCode;

                error = "\r\n";
                float thisTableWidth = thisTable.Width;

                int maxRow = thisTable.Rows.Count;
                int maxColumn = columnCount(thisTable);

                //         Bepaal de tabel breedte
                int tableWidth = (int)Math.Round(thisTableWidth * 1.2d / 0.75d); // 1.2d: 20% groter gemaakt
                tableWidth = tableWidth > informationModelWidth ? informationModelWidth : tableWidth;

                content.AppendLine("{|class=\"wikitable\" width=\"" + tableWidth.ToString() + "px\" style= \"font-size: 9.5pt;\""); 
                // nieuw
                for (iRow = 0; iRow < thisTable.Rows.Count; iRow++)
                {
                    content.AppendLine("|-style=vertical-align:top;");
                    for (iCol = 0; iCol < thisTable.Rows[iRow].Cells.Count; iCol++)
                    {
                        TableCell thisCell = thisTable.Rows[iRow].Cells[iCol];
                        rowSpan = 1;
                        if (thisCell.CellFormat.VerticalMerge == CellMerge.Start)
                        {
                            iRowTemp = iRow;
                            do
                            {
                                iRowTemp++;
                            } while (iRowTemp < thisTable.Rows.Count  && thisTable.Rows[iRowTemp].Cells[iCol].CellFormat.VerticalMerge == CellMerge.Continue && iRowTemp <= thisTable.Rows.Count); // AND clausule toegevoegd voor end of table
                            rowSpan = iRowTemp - iRow;
                        }
                        else if (thisCell.CellFormat.VerticalMerge == CellMerge.Continue)
                            continue;

                        colSpan = thisCell.GridSpan;

                        wikiCode = "|";

                        //                  Hier de Wiki cell info

                        if (colSpan > 1) wikiCode = wikiCode + "colspan=\"" + colSpan.ToString() + "\"";
                        if (colSpan > 1 && rowSpan > 1) wikiCode = wikiCode + " ";
                        if (rowSpan > 1) wikiCode = wikiCode + "rowspan=\"" + rowSpan.ToString() + "\"";

                        Color CellColor = thisCell.CellFormat.BackColor.IsEmpty?Color.Empty : thisCell.CellFormat.BackColor;
                        string strCellStyle = "";
                        if (CellColor != Color.Empty) strCellStyle = strCellStyle + "background-color: " + ColorTranslator.ToHtml(CellColor) + "; ";
                        //strCellStyle = strCellStyle + "width: " + (int)Math.Round(thisCell.Width / thisTableWidth * 100f) + "%; "; 15-2-23 Spire update 10.8.0 width obsolete
                        strCellStyle = strCellStyle + "width: " + (int)Math.Round(thisCell.GetCellWidth() / thisTableWidth * 100f) + "%; ";
                        
                        string strCell = "";
                        Type childType;
                        Color presentTextColor = Color.Empty;
                        bool colorTagOpen = false;
                        string textPartPrefix;
                        for (int i = 0; i < thisCell.Paragraphs.Count; i++)
                        {
                            for (int j = 0; j < thisCell.Paragraphs[i].ChildObjects.Count; j++)
                            {
                                childType = thisCell.Paragraphs[i].ChildObjects[j].GetType();
                                if (childType.Equals(typeof(TextRange)) || childType.Equals(typeof(MergeField)))
                                {
                                    TextRange childr = (TextRange)thisCell.Paragraphs[i].ChildObjects[j];
                                    string textPart = childr.CharacterFormat.Bold ? "<b>" + childr.Text + "</b>" : childr.Text;
                                    textPart = childr.CharacterFormat.Italic ? "<i>" + textPart + "</i>" : textPart;
                                    textPart = childr.CharacterFormat.IsSmallCaps ? "<small>" + textPart + "</small>" : textPart;
                                    textPart = childr.CharacterFormat.UnderlineStyle != UnderlineStyle.None ? "<u>" + textPart + "</u>" : textPart;
                                    textPart = childr.CharacterFormat.SubSuperScript == SubSuperScript.SubScript ? "<sub>" + textPart + "</sub>" : textPart;
                                    textPart = childr.CharacterFormat.SubSuperScript == SubSuperScript.SuperScript ? "<sup>" + textPart + "</sup>" : textPart;
                                    Color textColor = getColor(childr.CharacterFormat.TextColor.IsEmpty ? Color.Empty : childr.CharacterFormat.TextColor, CellColor);
                                    textPartPrefix = "";
                                    if (textColor != presentTextColor)
                                    {
                                        if (colorTagOpen)
                                        {
                                            textPartPrefix = "</font>";
                                            colorTagOpen = false;
                                        }
                                        if (textColor.ToArgb() != Color.Black.ToArgb())
                                        {
                                            textPartPrefix += "<font color=" + ColorTranslator.ToHtml(textColor) + ">";
                                            colorTagOpen = true;
                                        }
                                    }
                                    strCell += (textPartPrefix + textPart);
                                }
                                else if (childType.Equals(typeof(Break)))
                                {
                                    strCell += "\f"; // BreakType.PageBreak 
                                }
                                else if (childType.Equals(typeof(BookmarkStart)) || childType.Equals(typeof(BookmarkEnd)) || childType.Equals(typeof(Field)) || childType.Equals(typeof(FieldMark)))
                                {
                                    // Bekende childObjecten die geen actie vereisen
                                }
                                else
                                    error += "Onbekende childobject: " + thisCell.Paragraphs[i].ChildObjects[j].GetType().ToString() + " \r\n";
                            }
                            strCell += "\v";
                        }
                        strCell = strCell.Substring(0, strCell.Length - 1);  // was 2
                        strCell += (colorTagOpen ? "</font>" : ""); // staat er nog een font tag open?

                        if (strCell.Length > 0)
                        {
                            if (strCell.Substring(0,1)=="\f")
                            {
                                strCell = strCell.Substring(1, strCell.Length - 1); // beetje en noodgreep: de tweede tabel start soms met een pagebreak; /f wordt als 1 char gezien
                            }

                            //                  Wat wiki specials vervangen

                            if (strCell.Substring(0, 1) == "-" || strCell.Substring(0, 1) == "+")
                                    strCell = strCell.Substring(0, 1).wikiEncode() + strCell.Substring(1, strCell.Length - 1);
                            strCell = strCell.toWiki(); 
                            strCell = strCell.toWikiTable();
                        } // strCell.Length >0

                            //                  Bepaal de celattributen

                        if (strCellStyle.Length > 0)
                            wikiCode = wikiCode + " style=\"" + strCellStyle + "\"";
                        if (wikiCode.Length > 1)
                            wikiCode = wikiCode + "|";
                        content.AppendLine(wikiCode + strCell);
                    }
                }
                content.AppendLine("|}");
                error = error.Substring(0, error.Length - 2); // Haal de laatste crlf eraf
                return content.ToString();
            }

            private Color getColor(Color textColor, Color cellColor)
            {
                Color _color;
                if (textColor == Color.Empty)
                    _color =(cellColor != Color.Empty) ? (Luminance(cellColor) < 0.5 ? Color.White : Color.Black) : Color.Black;
                else
                    _color = textColor;
                return _color;
            }

            private double Luminance(Color color)
            {
                double temp = 0;
                temp = 0.2126 * color.R / 255d + 0.7152 * color.G / 255d + 0.0722 * color.B / 255d;
                return temp;
            }
        }
    }
}
