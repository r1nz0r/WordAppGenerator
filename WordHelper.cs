using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace AppendixGenConsole
{
    class WordHelper
    {
        public Application App;
        public Document Doc;
        public Range Rng;
        public Table TableWord;
        public static object oEndOfDoc = "\\endofdoc";

        public void CreateDoc()
        {
            if (App == null) OpenApplication();
            Doc = App.Documents.Add();

            Console.WriteLine("Все хорошо, начинаем процедуру генерации Word-документа...\nЭто может занять некоторое время.");
        }
        public void OpenApplication()
        {
            App = new Application();
        }
        public void SetVisibilityOfApp(bool bIsVisible)
        {
            App.Visible = bIsVisible;
        }
        public void AddLabelCaptionToDoc(string _LabelCaption, ref object _oLabelCaption)
        {
            App.CaptionLabels.Add(_LabelCaption);
            _oLabelCaption = App.CaptionLabels[_LabelCaption];
        }
        public void AddTextToDoc(string _Text)
        {
            App.Selection.Text = _Text;
            MoveRightBySelectionSentence(1);
            App.Selection.InsertParagraphAfter();
            MoveRightBySelectionSentence(1);
        }
        public void InsertPicture(string _ImgPath)
        {
            MoveToEndOfLine();
            var pic = App.Selection.InlineShapes.AddPicture(_ImgPath);
            App.Selection.Paragraphs.SpaceAfter = 6;
            App.Selection.Paragraphs.SpaceBefore = 6;
            App.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            MoveRightBySelectionSentence(1);
            App.Selection.InsertParagraphAfter();
        }
        public void InsertPicture(string _ImgPath, float _Height, float _Width)
        {
            MoveToEndOfLine();
            var pic = App.Selection.InlineShapes.AddPicture(_ImgPath);
            pic.Height = App.CentimetersToPoints(_Height);
            pic.Width = App.CentimetersToPoints(_Width);
            App.Selection.Paragraphs.SpaceAfter = 6;
            App.Selection.Paragraphs.SpaceBefore = 6;
            MoveToEndOfLine();
            App.Selection.InsertParagraphAfter();
        }
        public void MoveRightBySelectionSentence(int _Num)
        {
            App.Selection.MoveRight(WdUnits.wdSentence, _Num, WdMovementType.wdMove);
        }
        public void MoveToEndOfLine()
        {
            App.Selection.EndKey(WdUnits.wdStory, WdMovementType.wdMove);
        }
        public void MoveToBeginningOfTheDoc()
        {
            App.Selection.HomeKey(WdUnits.wdStory, WdMovementType.wdMove);
        }
        public void FormatCaption()
        {
            App.Selection.MoveEnd(WdUnits.wdParagraph, 1);
            App.Selection.Font.Name = "Times New Roman";
            App.Selection.Font.Size = 12;
            App.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            App.Selection.Font.Bold = 0;
            App.Selection.Font.Italic = 0;
            App.Selection.Font.Color = WdColor.wdColorBlack;
            App.Selection.Paragraphs.SpaceAfter = 0;
            MoveToEndOfLine();
        }
        public void RemoveSpacesFromLabel(string _LabelCaption)
        {
            //Удаляем пробел перед номером рисунка в названии (Рисунок A. 1 -> Рисунок A.1)
            Range rangeForReplaceCapt = App.ActiveDocument.Content;
            rangeForReplaceCapt.Find.Execute(FindText: (_LabelCaption + " "), ReplaceWith: _LabelCaption, Replace: WdReplace.wdReplaceAll);
        }
        public void ReplaceUoMToSuperScript(int _Offset)
        {
            int FieldValueCharNum = 19; // Для учета кол-ва символов в номере рисунка, т.к. ворд использует целиком весь ключ { SEQ Fig_A. \* ARABIC }

            foreach (Paragraph paragraph in Doc.Paragraphs)
            {
                do
                {
                    Range range = paragraph.Range;
                    if (range.Text.Contains("^"))
                    {
                        range.Start = range.Start + FieldValueCharNum + _Offset + range.Text.IndexOf("^");
                        range.End = range.Start + 2;
                        range.Select();
                        range.Font.Superscript = 1;
                    }
                    else break;

                    range.Find.Execute(FindText: ("^"), ReplaceWith: "", Replace: WdReplace.wdReplaceOne);
                }
                while (true);
            }
        }
        public void ReplaceUoMToSuperScript()
        {
            foreach (Paragraph paragraph in Doc.Paragraphs)
            {
                do
                {
                    Range range = paragraph.Range;
                    if (range.Text.Contains("^"))
                    {
                        range.Start = range.Start + range.Text.IndexOf("^");
                        range.End = range.Start + 2;
                        range.Select();
                        range.Font.Superscript = 1;
                    }
                    else break;

                    range.Find.Execute(FindText: ("^"), ReplaceWith: "", Replace: WdReplace.wdReplaceOne);
                }
                while (true);
            }
        }
        public void InsertCaptionForPic(object _oLabelCaption, object _LabelCaption)
        {
            MoveToEndOfLine();
            App.Selection.Range.InsertCaption(_oLabelCaption, _LabelCaption);

            FormatCaption();

            App.Selection.InsertParagraphAfter();
            MoveToEndOfLine();
        }
        public void FormatRange()
        {
            Rng.Font.Name = "Times New Roman";
            Rng.Font.Size = 12;
            Rng.Font.Color = WdColor.wdColorBlack;
            Rng.Font.Bold = 0;
            Rng.Font.Italic = 0;
            Rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphDistribute;
        }
        public void FormatRange(int fontSize, string fontName, WdParagraphAlignment paragraphAlignment)
        {
            Rng.Font.Size = fontSize;
            Rng.Font.Name = fontName;
            Rng.ParagraphFormat.Alignment = paragraphAlignment;
        }
        public void FormatRange(int fontSize, string fontName, WdParagraphAlignment paragraphAlignment, int spaceBefore, int spaceAfter)
        {
            Rng.Font.Name = fontName;
            Rng.Font.Size = fontSize;
            Rng.ParagraphFormat.Alignment = paragraphAlignment;
            Rng.ParagraphFormat.SpaceBefore = spaceBefore;
            Rng.ParagraphFormat.SpaceAfter = spaceAfter;
        }
        public void InsertCaptionForPic(object oLabelCaption, string labelCaption, string pictureNamePt1, string pictureNamePt2)
        {
            var rng = Doc.Bookmarks[oEndOfDoc].Range;
            App.Selection.EndKey(WdUnits.wdStory, WdMovementType.wdMove);

            rng.InsertCaption(oLabelCaption, pictureNamePt1);
            App.Selection.EndKey(WdUnits.wdStory, WdMovementType.wdMove);
            App.Selection.Paragraphs.SpaceBefore = 6;
            App.Selection.Range.InsertParagraphAfter();
            App.Selection.MoveRight(WdUnits.wdCharacter, 1, WdMovementType.wdMove);
            App.Selection.Text = pictureNamePt2;
            App.Selection.Paragraphs.SpaceBefore = 0;
            App.Selection.MoveUp(WdUnits.wdParagraph, 2, WdMovementType.wdExtend);

            FormatCaption(labelCaption);

            App.Selection.EndKey(WdUnits.wdStory, WdMovementType.wdMove);
            App.Selection.InsertNewPage();
        }
        public void InsertCaptionForPic(object oLabelCaption, string labelCaption, string pictureName)
        {
            var rng = Doc.Bookmarks[oEndOfDoc].Range;
            App.Selection.EndKey(WdUnits.wdStory, WdMovementType.wdMove);

            rng.InsertCaption(oLabelCaption, pictureName);
            App.Selection.EndKey(WdUnits.wdStory, WdMovementType.wdMove);
            App.Selection.Paragraphs.SpaceBefore = 6;
            App.Selection.MoveUp(WdUnits.wdParagraph, 1, WdMovementType.wdExtend);

            FormatCaption(labelCaption);

            App.Selection.EndKey(WdUnits.wdStory, WdMovementType.wdMove);
            App.Selection.InsertNewPage();
        }
        public void FormatCaption(string labelCaption)
        {
            App.Selection.MoveEnd(WdUnits.wdParagraph, 1);
            App.Selection.Font.Name = "Times New Roman";
            App.Selection.Font.Size = 12;
            App.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            App.Selection.Font.Bold = 0;
            App.Selection.Font.Italic = 0;
            App.Selection.Font.Color = WdColor.wdColorBlack;
            App.Selection.Paragraphs.SpaceAfter = 0;
            Range rangeForReplaceCapt = App.ActiveDocument.Content;
            rangeForReplaceCapt.Find.Execute(FindText: (labelCaption + " "), ReplaceWith: labelCaption, Replace: WdReplace.wdReplaceAll);
        }
        public void InsertTable(int rows, int columns)
        {          
            Range wrdRng = Doc.Bookmarks[ref oEndOfDoc].Range;
            wrdRng.Select();
            TableWord = Doc.Tables.Add(wrdRng, rows, columns);
            TableWord.Columns[1].SetWidth(App.CentimetersToPoints(13.50f), RulerStyle: WdRulerStyle.wdAdjustNone);
            TableWord.Columns[2].SetWidth(App.CentimetersToPoints(3.50f), RulerStyle: WdRulerStyle.wdAdjustNone);
            TableWord.Borders.Enable = 1;
            TableWord.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
        }
        public void InsertTableFRS(int rows, int columns)
        {
            TableWord = Doc.Tables.Add(App.Selection.Range, rows + 1, columns);
            TableWord.Columns[1].SetWidth(App.CentimetersToPoints(3.50f), RulerStyle: WdRulerStyle.wdAdjustNone);
            TableWord.Columns[2].SetWidth(App.CentimetersToPoints(3.50f), RulerStyle: WdRulerStyle.wdAdjustNone);
            TableWord.Columns[3].SetWidth(App.CentimetersToPoints(3.00f), RulerStyle: WdRulerStyle.wdAdjustNone);
            TableWord.Columns[4].SetWidth(App.CentimetersToPoints(3.50f), RulerStyle: WdRulerStyle.wdAdjustNone);
            TableWord.Columns[5].SetWidth(App.CentimetersToPoints(3.50f), RulerStyle: WdRulerStyle.wdAdjustNone);
            TableWord.Borders.Enable = 1;
            TableWord.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
            int count = 0;
            for (int j = 2; j <= rows + 1; j++)
            {
                for (int i = 3; i <= columns; i++)
                {
                    var cell = TableWord.Cell(j + count, i);
                    Rng = cell.Range;
                    cell.Split(3, 1);
                }
                count += 2;
            }


        }
        public void InsertTable(int rows, int columns, float columnOneHeight, float columnTwoHeight)
        {
            TableWord = Doc.Tables.Add(App.Selection.Range, rows, columns);
            TableWord.Columns[1].SetWidth(App.CentimetersToPoints(columnOneHeight), RulerStyle: WdRulerStyle.wdAdjustNone);
            TableWord.Columns[2].SetWidth(App.CentimetersToPoints(columnTwoHeight), RulerStyle: WdRulerStyle.wdAdjustNone);
            TableWord.Borders.Enable = 1;
            TableWord.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
        }
        public void InsertFormatedTextInCell(string text)
        {
            Rng = App.Selection.Range;            
            App.Selection.Text = text;
            App.Selection.Range.Font.Bold = 1;
            App.Selection.Range.Font.Name = "Times New Roman";
            App.Selection.Range.Font.Size = 12;
            App.Selection.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            App.Selection.MoveRight(WdUnits.wdCell, 1);
        }
        public void FormatTableForFRS(string[] caption, string accelPic, string displPic)
        {
            TableWord.Rows[1].SetHeight(App.CentimetersToPoints(19.0f), HeightRule: WdRowHeightRule.wdRowHeightExactly);
            TableWord.Columns[1].SetWidth(App.CentimetersToPoints(13.50f), RulerStyle: WdRulerStyle.wdAdjustNone);
            TableWord.Columns[2].SetWidth(App.CentimetersToPoints(3.50f), RulerStyle: WdRulerStyle.wdAdjustNone);
            FillTableWithContent(caption, accelPic, displPic);
        }
        public void FillTableWithContent(string[] caption, string imgPath1, string imgPath2)
        {
            var cell = TableWord.Cell(1, 2);
            Rng = cell.Range;
            App.Selection.MoveRight(WdUnits.wdCell, 1);

            foreach (var captionPart in caption)
            {
                App.Selection.Text = captionPart;
                App.Selection.MoveRight(WdUnits.wdSentence, 1, WdMovementType.wdMove);
                App.Selection.Paragraphs.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                App.Selection.InsertParagraphAfter();
                App.Selection.ParagraphFormat.SpaceBefore = 0;
                App.Selection.ParagraphFormat.SpaceAfter = 0;
                App.Selection.MoveRight(WdUnits.wdCharacter, 1, WdMovementType.wdMove);
            }

            FormatRange(11, "Times New Roman", WdParagraphAlignment.wdAlignParagraphLeft);

            cell = TableWord.Cell(1, 1);
            Rng = cell.Range;
            Rng.ParagraphFormat.SpaceBefore = 6;
            Rng.ParagraphFormat.SpaceAfter = 6;
            App.Selection.MoveLeft(WdUnits.wdCell, 1, WdMovementType.wdMove);
            Rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            InsertPictureToTable(imgPath1);
            InsertPictureToTable(imgPath2);
        } 
        public void InsertPictureToTable(string imgPath)
        {
            var pic = App.Selection.InlineShapes.AddPicture(imgPath);
            pic.Height = App.CentimetersToPoints(9.1f);
            pic.Width = App.CentimetersToPoints(13.0f);
        }        
        public void SaveDoc(string path, string name)
        {
            try
            {
                Doc.SaveAs2(path, WdSaveFormat.wdFormatDocumentDefault);
                Doc.Close();
                App.Quit();

                Console.Clear();
                            
                Console.Write($"Документ готов и сохранен с названием: ");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write($"{FileManager.GlobalPath + $"{name}.docx"}");
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine();
                Console.WriteLine("Для выхода из программы нажмите любую клавишу...");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public void InsertInfoTextInTable(string[] captionForColumn)
        {
            foreach (var captionPart in captionForColumn)
            {
                App.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                App.Selection.MoveRight(WdUnits.wdSentence, 1, WdMovementType.wdMove);
                App.Selection.InsertParagraphAfter();
                App.Selection.Font.Name = "Times New Roman";
                App.Selection.Font.Size = 12;
                App.Selection.Text = captionPart;
                App.Selection.MoveRight(WdUnits.wdSentence, 1, WdMovementType.wdMove);
                App.Selection.Paragraphs.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                App.Selection.ParagraphFormat.FirstLineIndent = App.CentimetersToPoints(1.5f);
                App.Selection.InsertParagraphAfter();
                App.Selection.ParagraphFormat.SpaceBefore = 0;
                App.Selection.ParagraphFormat.SpaceAfter = 0;
                App.Selection.MoveRight(WdUnits.wdCharacter, 1, WdMovementType.wdMove);
            }
        }
        public void FillHeadOfFRSTable()
        {
            var cell = TableWord.Cell(1, 1);
            Rng = cell.Range;

            InsertFormatedTextInCell("Отметка");
            InsertFormatedTextInCell("Узлы");
            InsertFormatedTextInCell("Компонента");
            InsertFormatedTextInCell("Направление");
            InsertFormatedTextInCell("Номер рисунка");
        }
        public void FillBodyOfFRSTable(List<string> elevations, ref object oLabelCaption)
        {
            int numCells = elevations.Count / 3;
            var cell = TableWord.Cell(2, 1);
            Rng = cell.Range;
            Rng.Select();

            for (int i = 0; i < numCells; i++)
            {
                if (i == 0)
                {
                    cell = TableWord.Cell(2, 1);
                }
                else
                {
                    cell = TableWord.Cell(i*3 + 2, 1);
                }

                Rng = cell.Range;
                Rng.Select();
                App.Selection.Text = elevations[i*3];
                FormatRange(12, "Times New Roman", WdParagraphAlignment.wdAlignParagraphJustify);

                for (int j = 1; j <= 3; j++)
                {
                    cell = TableWord.Cell(i * 3 + j + 1, 3);
                    Rng = cell.Range;
                    Rng.Select();

                    switch (j)
                    {
                        case 1:
                            App.Selection.Text = "X";
                            break;
                        case 2:
                            App.Selection.Text = "Y";
                            break;
                        case 3:
                            App.Selection.Text = "Z";
                            break;
                        default:
                            return;
                    }
                    FormatRange(12, "Times New Roman", WdParagraphAlignment.wdAlignParagraphCenter);

                    cell = TableWord.Cell(i * 3 + j + 1, 4);
                    Rng = cell.Range;
                    Rng.Select();

                    switch (j)
                    {
                        case 1:
                            App.Selection.Text = "Горизонтальное";
                            break;
                        case 2:
                            App.Selection.Text = "Горизонтальное";
                            break;
                        case 3:
                            App.Selection.Text = "Вертикальное";
                            break;
                        default:
                            return;
                    }

                    FormatRange(12, "Times New Roman", WdParagraphAlignment.wdAlignParagraphCenter);
                }                
            }

            for (int index = 1; index <= elevations.Count; index++)
            {
                cell = TableWord.Cell(index + 1, 5);
                Rng = cell.Range;
                Rng.Select();
                Rng.Collapse(WdCollapseDirection.wdCollapseStart); // без этого метода почему-то не вставляется в нужное место ссылка.
                Rng.InsertCrossReference(ref oLabelCaption, WdReferenceKind.wdOnlyLabelAndNumber, index, InsertAsHyperlink: true);
                FormatRange(12, "Times New Roman", WdParagraphAlignment.wdAlignParagraphCenter);
                Rng.Font.Italic = 0;
            }
        }
        public void InsertCrossReferenceFRS(List<string> accelerationList, ref object oLabelCaption)
        {
            for (int index = 1; index <= accelerationList.Count; index++)
            {
                var userParagrah = Doc.Paragraphs.Add();
                var userRange = Doc.Bookmarks[oEndOfDoc].Range;
                userRange.InsertCrossReference(ref oLabelCaption, WdReferenceKind.wdEntireCaption, index, true);
                userRange.Move(WdUnits.wdParagraph, 1);
                userRange.Text = ";";
                userRange.Font.Italic = 0;
                userRange.Font.Bold = 0;
                userParagrah.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            }

            App.Selection.EndKey(WdUnits.wdStory, WdMovementType.wdMove);
            App.Selection.InsertNewPage();
        }
        public void InsertCrossReferenceInTablePSA(List<PictureData> pictureData, ref object oLabelCaption)
        {
            var rng = TableWord.Cell(1, 1).Range;

            for (int index = 1; index <= pictureData.Count; index++)
            {
                var cell = TableWord.Cell(index, 1);
                App.Selection.Range.InsertCrossReference(ref oLabelCaption, WdReferenceKind.wdOnlyLabelAndNumber, index, InsertAsHyperlink: true);                
                App.Selection.MoveRight(WdUnits.wdCell, 1);
                cell = TableWord.Cell(index, 2);
                App.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                App.Selection.Range.InsertCrossReference(ref oLabelCaption, WdReferenceKind.wdOnlyCaptionText, index, InsertAsHyperlink: true);        
                App.Selection.MoveRight(WdUnits.wdCell, 1);
            }
        }
        public void CloseDoc()
        {
            Console.Clear();
            Console.WriteLine("Для выхода из программы и закрытия документа Word нажмите любую кнопку.");
            Console.ReadKey();

            Doc.Close();
            App.Quit();
        }

    }
}
