using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace JsonToDocumentPOC
{
    public partial class GenerateDocument : Form
    {
        private const string KEY_PREFIX = "<<";
        private const string KEY_SUFFIX = ">>";
        private const string SOURCE_FILE_NAME = "OrthoInitialTemplateOriginal.docx";
        private const string OUTPUT_FILE_FOLDER_PATH = @"C:\AdarshData\Project Documents\OrthoInitialImplementation\JsonToDocumentPOC\JsonToDocumentPOC\Data\";
        private const string TEMPORARY_FILE_FOLDER_PATH = @"C:\AdarshData\Project Documents\OrthoInitialImplementation\JsonToDocumentPOC\JsonToDocumentPOC\Data\";
        private const string SOURCE_FILE_FOLDER_PATH = @"C:\AdarshData\Project Documents\OrthoInitialImplementation\JsonToDocumentPOC\JsonToDocumentPOC\Data\";

        public GenerateDocument()
        {
            InitializeComponent();
        }

        private void btnGenerateDocument_Click(object sender, EventArgs e)
        {
            string data = string.Empty;
            string sourcefile = SOURCE_FILE_FOLDER_PATH + SOURCE_FILE_NAME;
            string outputFileName = "OrthoInitialMerged.pdf";

            JsonToDocument(txtJsonData.Text, sourcefile, OUTPUT_FILE_FOLDER_PATH, outputFileName);

            MessageBox.Show("The document has been generated successfully with file name: " + OUTPUT_FILE_FOLDER_PATH + outputFileName);
        }

        #region JsonToDocument
        private void JsonToDocument(string jsonData, string templateFileName, string outputFolderName, string outputFileName)
        {
            object sourcefile = templateFileName;
            string tempWordFileName = TEMPORARY_FILE_FOLDER_PATH + outputFileName.Split('.')[0].ToString() + ".docx";
            string outputPDFFilePath = outputFolderName + outputFileName;

            var jsonLinq = JObject.Parse(jsonData);
            IList<string> keys = jsonLinq.Properties().Select(p => p.Name).ToList();

            object missing = Type.Missing;
            object oFalse = false;
            object oTrue = true;

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(ref sourcefile, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            doc.Activate();

            foreach (string key in keys)
            {
                //If key is a property
                if (jsonLinq[key].GetType().ToString() == "Newtonsoft.Json.Linq.JValue")
                {
                    FindAndReplaceText(doc, key, jsonLinq[key].ToString());
                }
                //if key is a collection
                else if (jsonLinq[key].GetType().ToString() == "Newtonsoft.Json.Linq.JArray")
                {
                    var collection = jsonLinq[key].ToArray();

                    AddTableToDocument(doc, key, collection);
                }
            }

            AddImageToDocument(doc);

            if (System.IO.File.Exists(tempWordFileName))
            {
                System.IO.File.Delete(tempWordFileName);
            }
            doc.SaveAs(tempWordFileName);

            if (System.IO.File.Exists(outputPDFFilePath))
            {
                System.IO.File.Delete(outputPDFFilePath);
            }
            doc.ExportAsFixedFormat(outputPDFFilePath, WdExportFormat.wdExportFormatPDF);

            doc.Close(ref missing, ref missing, ref missing);
            wordApp.Quit(ref missing, ref missing, ref missing);

            if (System.IO.File.Exists(tempWordFileName))
            {
                System.IO.File.Delete(tempWordFileName);
            }
        }
        #endregion

        #region FindAndReplaceText
        private void FindAndReplaceText(Microsoft.Office.Interop.Word.Document doc, string textToFind, string textToReplace)
        {
            object missing = Type.Missing;
            object oFalse = false;
            object oTrue = true;

            foreach (Microsoft.Office.Interop.Word.Range range in doc.StoryRanges)
            {
                Microsoft.Office.Interop.Word.Find find = range.Find;
                object findText = KEY_PREFIX + textToFind + KEY_SUFFIX;
                object replacText = textToReplace;
                object replace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
                object findWrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
                find.Execute(ref findText, ref missing, ref missing, ref missing, ref oFalse, ref missing,
                    ref missing, ref findWrap, ref missing, ref replacText,
                    ref replace, ref missing, ref missing, ref missing, ref missing);

                Marshal.FinalReleaseComObject(find);
            }
        }
        #endregion

        #region AddTableToDocument
        private void AddTableToDocument(Microsoft.Office.Interop.Word.Document doc, string tablenamekey, JToken[] collection)
        {
            int numberOfRows = collection.Length;
            int numberOfColumns = collection[0].Count() - 1;
            int rowCounter = 1;
            int columnCounter = 1;
            object oMissing = Type.Missing;
            object oEndOfDoc = tablenamekey;
            object styleName = "Grid Table 5 Dark";

            Microsoft.Office.Interop.Word.Table oTable;
            if (doc.Bookmarks.Exists(tablenamekey))
            {
                Microsoft.Office.Interop.Word.Range wrdRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                oTable = doc.Tables.Add(wrdRng, numberOfRows, numberOfColumns, ref oMissing, ref oMissing);
                oTable.Range.ParagraphFormat.SpaceAfter = 6;
                oTable.Borders.Enable = 1;
                oTable.set_Style(ref styleName);
                oTable.Columns.DistributeWidth();

                foreach (JObject item in collection)
                {
                    columnCounter = 1;

                    if (item["IsHeader"].ToString() == "1")
                    {
                        //Logic to Add Header Row to table
                        foreach (JProperty column in item.Properties())
                        {
                            // Only include JValue types
                            if (column.Value is JValue && column.Name != "IsHeader")
                            {
                                oTable.Cell(rowCounter, columnCounter).Range.Text = column.Value.ToString();
                                oTable.Cell(rowCounter, columnCounter).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                oTable.Cell(rowCounter, columnCounter).Range.Bold = 1;
                                oTable.Cell(rowCounter, columnCounter).Range.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
                                //oTable.Cell(rowCounter, columnCounter).Column.AutoFit();
                            }
                            columnCounter++;
                        }
                    }
                    else
                    {
                        //Logic to Add Data Row to table
                        foreach (JProperty column in item.Properties())
                        {
                            // Only include JValue types
                            if (column.Value is JValue && column.Name != "IsHeader")
                            {
                                oTable.Cell(rowCounter, columnCounter).Range.Text = column.Value.ToString();
                                oTable.Cell(rowCounter, columnCounter).Range.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
                                oTable.Cell(rowCounter, columnCounter).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                                //oTable.Cell(rowCounter, columnCounter).Column.AutoFit();
                            }
                            columnCounter++;
                        }
                    }
                    rowCounter++;
                }
            }
        }
        #endregion

        #region AddImageToDocument
        private void AddImageToDocument(Microsoft.Office.Interop.Word.Document doc)
        {
            object oEndOfDoc = "SIGNATURE";
            Range imageRange = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            string imagePath = @"C:\AdarshData\Project Documents\OrthoInitialImplementation\JsonToDocumentPOC\JsonToDocumentPOC\Data\Signature.jpg";

            // Create an InlineShape in the InlineShapes collection where the picture should be added later
            // It is used to get automatically scaled sizes.
            InlineShape autoScaledInlineShape = imageRange.InlineShapes.AddPicture(imagePath);
            float scaledWidth = autoScaledInlineShape.Width;
            float scaledHeight = autoScaledInlineShape.Height;
            autoScaledInlineShape.Delete();

            // Create a new Shape and fill it with the picture
            Shape newShape = doc.Shapes.AddShape(1, 0, 0, scaledWidth, scaledHeight);
            newShape.Fill.UserPicture(imagePath);

            // Convert the Shape to an InlineShape and optional disable Border
            InlineShape finalInlineShape = newShape.ConvertToInlineShape();
            finalInlineShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

            // Cut the range of the InlineShape to clipboard
            finalInlineShape.Range.Cut();

            // And paste it to the target Range
            imageRange.Paste();
        }
        #endregion
    }
}
