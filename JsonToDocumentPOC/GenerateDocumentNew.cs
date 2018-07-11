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
    public partial class GenerateDocumentNew : Form
    {
        #region Variable Declaration
        private const string KEY_PREFIX = "<<";
        private const string KEY_SUFFIX = ">>";
        private const string AREA_BOOKMARK_SUFFIX = "AreaBookMark";
        private const string DOCTOR_SIGNATURE_KEY = "DoctorSignature";
        private const string PATIENT_SIGNATURE_KEY = "PatientSignature";
        private const string TECHNICIAN_SIGNATURE_KEY = "TechnicianSignature";
        private const string PROCEDURE_CODE_KEY = "ProcedureCodes";

        private const string SOURCE_FILE_NAME = "OrthoInitialTemplate16_Jan_2018.docx";
        private const string OUTPUT_FILE_FOLDER_PATH = @"C:\AdarshData\Project Documents\OrthoInitialImplementation\JsonToDocumentPOC\JsonToDocumentPOC\Data\";
        private const string TEMPORARY_FILE_FOLDER_PATH = @"C:\AdarshData\Project Documents\OrthoInitialImplementation\JsonToDocumentPOC\JsonToDocumentPOC\Data\";
        private const string SOURCE_FILE_FOLDER_PATH = @"C:\AdarshData\Project Documents\OrthoInitialImplementation\JsonToDocumentPOC\JsonToDocumentPOC\Data\";
        string signatureImagePath = @"C:\AdarshData\Project Documents\OrthoInitialImplementation\JsonToDocumentPOC\JsonToDocumentPOC\Data\Signature.jpg";
        private string TABLE_GRID_STYLE = "Grid Table 1 Light";

        #endregion

        #region Constructor 
        public GenerateDocumentNew()
        {
            InitializeComponent();
        }
        #endregion

        #region btnGenerateDocument_Click
        private void btnGenerateDocument_Click(object sender, EventArgs e)
        {
            string data = string.Empty;
            string sourcefile = SOURCE_FILE_FOLDER_PATH + SOURCE_FILE_NAME;
            string outputFileName = OUTPUT_FILE_FOLDER_PATH + "OrthoInitialTemplate16_Jan_2018_Merged.pdf";

            List<KeyValuePair> procedureCodeList = new List<KeyValuePair>();
            procedureCodeList.Add(new KeyValuePair { Key = "99521", Value = "Some Procedure Code Description 1" });
            procedureCodeList.Add(new KeyValuePair { Key = "98955", Value = "Some Procedure Code Description 2" });

            JsonToDocument(
                txtJsonData.Text, 
                sourcefile, 
                outputFileName,
                signatureImagePath,
                signatureImagePath,
                signatureImagePath, 
                PROCEDURE_CODE_KEY,
                procedureCodeList);

            MessageBox.Show("The document has been generated successfully with file name: " + outputFileName);
        }
        #endregion

        #region JsonToDocument
        public void JsonToDocument(
            string jsonData,
            string templateFileNameWithPath,
            string outputPDFFileNameWithPath,
            string doctorSignatureFilePath,
            string patientSignatureFilePath,
            string technitianSignatureFilePath,
            string procedreCodeKey,
            List<KeyValuePair> procedureCodeList)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            try
            {
                object sourcefile = templateFileNameWithPath;

                System.Random objRandom = new Random();
                string tempWordFileName = TEMPORARY_FILE_FOLDER_PATH + Guid.NewGuid().ToString() + DateTime.Now.ToString("yyyyMMddHHmmssms") + ".docx";

                var jsonLinq = JObject.Parse(jsonData);
                var keyCollection = jsonLinq["notesdata"].ToArray();
                List<KeyValuePair> keyValue = new List<KeyValuePair>();

                object missing = Type.Missing;
                object oFalse = false;
                object oTrue = true;


                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(ref sourcefile, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                doc.Activate();

                foreach (JToken key in keyCollection)
                {
                    var renderKeyName = key["RenderKeyName"].ToString();
                    var renderKeyValue = key["RenderKeyValue"].ToString();
                    var isRender = key["IsRender"].ToString();
                    var dataType = key["DataType"].ToString();
                    var renderType = key["RenderType"].ToString();
                    var renderValue = key["RenderValue"].ToString();
                    var finalRenderValue = "";

                    if (dataType == "Item" && isRender == "1")
                    {
                        finalRenderValue = RenderItemValue(renderKeyName, renderKeyValue, renderType, renderValue);
                        FindAndReplaceText(doc, renderKeyName, finalRenderValue);
                        keyValue.Add(new KeyValuePair { Key = renderKeyName, Value = finalRenderValue });
                    }
                    else if (dataType == "List" && isRender == "1")
                    {
                        var listdata = key["RenderValue"].ToArray();

                        if (renderType == "RenderValueListAsCommaSeperatedString")
                        {
                            JToken jToken = JToken.Parse(renderValue);
                            string finalStringList = RenderValueListAsCommaSeperatedString(jToken);
                            FindAndReplaceText(doc, renderKeyName, finalStringList);
                            keyValue.Add(new KeyValuePair { Key = renderKeyName, Value = finalStringList });
                        }
                        else if (renderType == "RenderValueListToString")
                        {
                            JToken jToken = JToken.Parse(renderValue);
                            string finalStringList = RenderValueListToString(jToken);
                            FindAndReplaceText(doc, renderKeyName, finalStringList);
                            keyValue.Add(new KeyValuePair { Key = renderKeyName, Value = finalStringList });
                        }
                        else if (renderType == "RenderValueListAsTable")
                        {
                            if (TableHasChiledItemToRender(JArray.Parse(renderValue)))
                            {
                                AddTableToDocumentNew(doc, renderKeyName, JArray.Parse(renderValue), TABLE_GRID_STYLE);
                            }
                            else
                            {
                                DeleteBookMarkParagraph(doc, renderKeyName);
                            }
                        }
                        else if (renderType == "RenderValueAsBulletList")
                        {
                            if (HasChiledItemToRender(JArray.Parse(renderValue)))
                            {
                                AddBulletedListToDocument(doc, renderKeyName, JArray.Parse(renderValue));
                            }
                            else
                            {
                                DeleteBookMarkParagraph(doc, renderKeyName);
                            }
                        }
                    }
                    else if (isRender == "0")
                    {
                        //Logic to remove bookmark/datakey tag from merged document
                        DeleteBookMarkParagraph(doc, renderKeyName);
                    }
                }

                AddImageToDocument(doc, DOCTOR_SIGNATURE_KEY, doctorSignatureFilePath);
                AddImageToDocument(doc, PATIENT_SIGNATURE_KEY, patientSignatureFilePath);
                AddImageToDocument(doc, TECHNICIAN_SIGNATURE_KEY, technitianSignatureFilePath);

                AddProcedureCodesTable(doc, procedreCodeKey, procedureCodeList, TABLE_GRID_STYLE);

                //Replace in-progress key values
                foreach (KeyValuePair item in keyValue)
                {
                    if (item.Value != null && item.Value != "")
                    {
                        FindAndReplaceText(doc, item.Key, item.Value);
                    }
                }

                if (System.IO.File.Exists(tempWordFileName))
                {
                    System.IO.File.Delete(tempWordFileName);
                }
                doc.SaveAs(tempWordFileName);

                if (System.IO.File.Exists(outputPDFFileNameWithPath))
                {
                    System.IO.File.Delete(outputPDFFileNameWithPath);
                }
                doc.ExportAsFixedFormat(outputPDFFileNameWithPath, WdExportFormat.wdExportFormatPDF);

                doc.Close(ref missing, ref missing, ref missing);

                if (System.IO.File.Exists(tempWordFileName))
                {
                    System.IO.File.Delete(tempWordFileName);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                wordApp.Quit(SaveChanges: false, OriginalFormat: false, RouteDocument: false);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }
        }
        #endregion

        #region FindAndReplaceText
        private void FindAndReplaceText(Microsoft.Office.Interop.Word.Document doc, string textToFind, string textToReplace)
        {
            object missing = Type.Missing;
            object oFalse = false;
            object oTrue = true;
            string searchKey = KEY_PREFIX + textToFind + KEY_SUFFIX;

            foreach (Microsoft.Office.Interop.Word.Range range in doc.StoryRanges)
            {
                ReplaceTextInRange(range, searchKey, textToReplace);
                //Microsoft.Office.Interop.Word.Find find = range.Find;
                //object findText = KEY_PREFIX + textToFind + KEY_SUFFIX;
                //object replacText = textToReplace;
                //object replace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
                //object findWrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
                //find.Execute(ref findText, ref missing, ref missing, ref missing, ref oFalse, ref missing,
                //    ref missing, ref findWrap, ref missing, ref replacText,
                //    ref replace, ref missing, ref missing, ref missing, ref missing);

                //Marshal.FinalReleaseComObject(find);
            }
        }
        #endregion

        #region ReplaceTextInRange
        private void ReplaceTextInRange(Microsoft.Office.Interop.Word.Range rngStory, string strSearch, string strReplace)
        {
            if (strReplace.Length > 250)
            {
                ReplaceTextInRange(rngStory, strSearch, strSearch + strReplace.Substring(256));

                strReplace = strReplace.Substring(0, 250);
            }

            rngStory.Find.ClearFormatting();
            rngStory.Find.Replacement.ClearFormatting();

            rngStory.Find.Text = strSearch;
            rngStory.Find.Replacement.Text = strReplace.Replace("\n", "");

            rngStory.Find.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;

            rngStory.Find.Execute(
                System.Reflection.Missing.Value, // Find Pattern
                true, //MatchCase
                true, //MatchWholeWord
                System.Reflection.Missing.Value, //MatchWildcards
                false, //MatchSoundsLike
                System.Reflection.Missing.Value, //MatchAllWordForms
                System.Reflection.Missing.Value, //Forward
                System.Reflection.Missing.Value, //Wrap
                System.Reflection.Missing.Value, //Format
                System.Reflection.Missing.Value, //ReplaceWith
                Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne, //Replace
                System.Reflection.Missing.Value, //MatchKashida
                System.Reflection.Missing.Value, //MatchDiacritics
                System.Reflection.Missing.Value, //MatchAlefHamza
                System.Reflection.Missing.Value); //MatchControl
        }
        #endregion

        #region AddTableToDocument
        private void AddTableToDocument(Microsoft.Office.Interop.Word.Document doc, string tablenamekey, JArray collection, string tableGridStyle)
        {
            int numberOfRows = collection.Count;
            int numberOfColumns = collection[0].Count() - 1;
            int rowCounter = 1;
            int columnCounter = 1;
            object oMissing = Type.Missing;
            object oEndOfDoc = tablenamekey;
            object styleName = tableGridStyle;// "Grid Table 5 Dark";

            Microsoft.Office.Interop.Word.Table oTable;
            if (doc.Bookmarks.Exists(tablenamekey))
            {
                Microsoft.Office.Interop.Word.Range wrdRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                oTable = doc.Tables.Add(wrdRng, numberOfRows, numberOfColumns, ref oMissing, ref oMissing);
                oTable.Range.ParagraphFormat.SpaceAfter = 6;
                oTable.Borders.Enable = 1;
                oTable.set_Style(ref styleName);
                oTable.Columns.DistributeWidth();

                foreach (JObject item in collection.Children<JObject>())
                {
                    columnCounter = 1;

                    if (item["IsHeader"]["RenderKeyValue"].ToString() == "1")
                    {
                        //Logic to Add Header Row to table
                        foreach (JProperty column in item.Properties())
                        {
                            // Only include JValue types
                            if (column.Value is JObject && column.Name != "IsHeader")
                            {
                                var renderKeyName = column.Value["RenderKeyName"].ToString();
                                var renderKeyValue = column.Value["RenderKeyValue"].ToString();
                                var isRender = column.Value["IsRender"].ToString();
                                var dataType = column.Value["DataType"].ToString();
                                var renderType = column.Value["RenderType"].ToString();
                                var renderValue = column.Value["RenderValue"].ToString();

                                oTable.Cell(rowCounter, columnCounter).Range.Text = RenderItemValue(renderKeyName, renderKeyValue, renderType, renderValue);
                                oTable.Cell(rowCounter, columnCounter).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                oTable.Cell(rowCounter, columnCounter).Range.Bold = 1;
                                oTable.Cell(rowCounter, columnCounter).Range.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
                                columnCounter++;
                            }
                        }
                    }
                    else
                    {
                        //Logic to Add Data Row to table
                        foreach (JProperty column in item.Properties())
                        {
                            // Only include JValue types
                            if (column.Value is JObject && column.Name != "IsHeader")
                            {
                                var renderKeyName = column.Value["RenderKeyName"].ToString();
                                var renderKeyValue = column.Value["RenderKeyValue"].ToString();
                                var isRender = column.Value["IsRender"].ToString();
                                var dataType = column.Value["DataType"].ToString();
                                var renderType = column.Value["RenderType"].ToString();
                                var renderValue = column.Value["RenderValue"].ToString();

                                string columnValue = string.Empty;

                                if (renderType == "RenderValueListToString")
                                {
                                    JToken jToken = JToken.Parse(renderValue);
                                    columnValue = RenderValueListToString(jToken);
                                }
                                else if (renderType == "RenderValueListAsCommaSeperatedString")
                                {
                                    JToken jToken = JToken.Parse(renderValue);
                                    columnValue = RenderValueListAsCommaSeperatedString(jToken);
                                }
                                else
                                {
                                    columnValue = RenderItemValue(renderKeyName, renderKeyValue, renderType, renderValue);
                                }

                                oTable.Cell(rowCounter, columnCounter).Range.Text = columnValue;
                                oTable.Cell(rowCounter, columnCounter).Range.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
                                oTable.Cell(rowCounter, columnCounter).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                                columnCounter++;
                            }
                            
                        }
                    }
                    rowCounter++;
                }
            }
        }
        #endregion

        #region AddTableToDocumentNew
        private void AddTableToDocumentNew(Microsoft.Office.Interop.Word.Document doc, string tablenamekey, JArray collection, string tableGridStyle)
        {
            int numberOfRows = TableRowCountToRender(collection); //collection.Count;
            int numberOfColumns = collection[0]["Coloums"].Count();
            int rowCounter = 1;
            int columnCounter = 1;
            object oMissing = Type.Missing;
            object oEndOfDoc = tablenamekey;
            object styleName = tableGridStyle;// "Grid Table 5 Dark";

            Microsoft.Office.Interop.Word.Table oTable;
            if (doc.Bookmarks.Exists(tablenamekey))
            {
                Microsoft.Office.Interop.Word.Range wrdRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                oTable = doc.Tables.Add(wrdRng, numberOfRows, numberOfColumns, ref oMissing, ref oMissing);
                oTable.Range.ParagraphFormat.SpaceAfter = 6;
                oTable.Borders.Enable = 1;
                oTable.set_Style(ref styleName);
                oTable.Columns.DistributeWidth();

                foreach (JObject item in collection.Children<JObject>())
                {
                    columnCounter = 1;
                    var columnCollection = item["Coloums"].Children<JObject>();

                    //Logic to Add Header Row to table
                    if (item["IsHeader"]["RenderKeyValue"].ToString() == "1")
                    {
                        if (TableRowHasColumnsToRender(columnCollection))
                        {
                            foreach (JObject columnitem in columnCollection)
                            {
                                var renderKeyName = columnitem["RenderKeyName"].ToString();
                                var renderKeyValue = columnitem["RenderKeyValue"].ToString();
                                var isRender = columnitem["IsRender"].ToString();
                                var dataType = columnitem["DataType"].ToString();
                                var renderType = columnitem["RenderType"].ToString();
                                var renderValue = columnitem["RenderValue"].ToString();

                                oTable.Cell(rowCounter, columnCounter).Range.Text = RenderItemValue(renderKeyName, renderKeyValue, renderType, renderValue);
                                oTable.Cell(rowCounter, columnCounter).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                oTable.Cell(rowCounter, columnCounter).Range.Bold = 1;
                                oTable.Cell(rowCounter, columnCounter).Range.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;

                                columnCounter++;
                            }
                            rowCounter++;
                        }
                    }
                    else
                    {
                        if (TableRowHasColumnsToRender(columnCollection))
                        {
                            //Logic to Add Data Row to table
                            foreach (JObject columnitem in columnCollection)
                            {
                                var renderKeyName = columnitem["RenderKeyName"].ToString();
                                var renderKeyValue = columnitem["RenderKeyValue"].ToString();
                                var isRender = columnitem["IsRender"].ToString();
                                var dataType = columnitem["DataType"].ToString();
                                var renderType = columnitem["RenderType"].ToString();
                                var renderValue = columnitem["RenderValue"].ToString();

                                string columnValue = string.Empty;

                                if (renderType == "RenderValueListToString")
                                {
                                    JToken jToken = JToken.Parse(renderValue);
                                    columnValue = RenderValueListToString(jToken);
                                }
                                else if (renderType == "RenderValueListAsCommaSeperatedString")
                                {
                                    JToken jToken = JToken.Parse(renderValue);
                                    columnValue = RenderValueListAsCommaSeperatedString(jToken);
                                }
                                else
                                {
                                    columnValue = RenderItemValue(renderKeyName, renderKeyValue, renderType, renderValue);
                                }

                                oTable.Cell(rowCounter, columnCounter).Range.Text = columnValue;
                                oTable.Cell(rowCounter, columnCounter).Range.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
                                oTable.Cell(rowCounter, columnCounter).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                                columnCounter++;
                            }
                            rowCounter++;
                        }
                    }
                }
            }
        }
        #endregion

        #region AddImageToDocument
        private void AddImageToDocument(
            Microsoft.Office.Interop.Word.Document doc, 
            string signatureKeyName, 
            string signatureFilePath)
        {
            object oEndOfDoc = signatureKeyName;
            if (doc.Bookmarks.Exists(signatureKeyName))
            {
                Range imageRange = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;

                // Create an InlineShape in the InlineShapes collection where the picture should be added later
                // It is used to get automatically scaled sizes.
                InlineShape autoScaledInlineShape = imageRange.InlineShapes.AddPicture(signatureFilePath);
                float scaledWidth = autoScaledInlineShape.Width;
                float scaledHeight = autoScaledInlineShape.Height;
                autoScaledInlineShape.Delete();

                // Create a new Shape and fill it with the picture
                Shape newShape = doc.Shapes.AddShape(1, 0, 0, scaledWidth, scaledHeight);
                newShape.Fill.UserPicture(signatureFilePath);

                // Convert the Shape to an InlineShape and optional disable Border
                InlineShape finalInlineShape = newShape.ConvertToInlineShape();
                finalInlineShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

                // Cut the range of the InlineShape to clipboard
                finalInlineShape.Range.Cut();

                // And paste it to the target Range
                imageRange.Paste();
            }
            else
            {
                FindAndReplaceText(doc, signatureKeyName, "Signature Image is missig"); ;
            }
        }
        #endregion

        #region RenderItemValue
        private string RenderItemValue(string renderKeyName, string renderKeyValue, string renderType, string renderValue)
        {
            if (renderType == "RenderValue")
            {
                return renderValue;
            }
            if (renderType == "RenderKeyValue")
            {
                return renderKeyValue;
            }
            else if (renderType == "RenderValueWithKeyReplacement")
            {
                return renderValue.Replace(KEY_PREFIX + renderKeyName + KEY_SUFFIX, renderKeyValue);
            }
            else
            {
                return renderKeyValue;
            }
        }
        #endregion

        #region RenderListToString
        private string RenderValueListToString(Newtonsoft.Json.Linq.JToken data)
        {
            string list = "";

            foreach (JToken token in data)
            {
                string finalValue = "";
                var renderKeyName = token["RenderKeyName"].ToString();
                var renderKeyValue = token["RenderKeyValue"].ToString();
                var isRender = token["IsRender"].ToString();
                var dataType = token["DataType"].ToString();
                var renderType = token["RenderType"].ToString();
                var renderValue = token["RenderValue"].ToString();

                if (isRender == "1")
                {
                    if (dataType == "Item")
                    {
                        finalValue = RenderItemValue(renderKeyName, renderKeyValue, renderType, renderValue);
                    }
                    else if (dataType == "List" && isRender == "1" && renderType == "RenderValueListToString")
                    {
                        JToken jToken = JToken.Parse(renderValue);
                        finalValue = RenderValueListToString(jToken);
                    }
                    else if (dataType == "List" && isRender == "1" && renderType == "RenderValueListAsCommaSeperatedString")
                    {
                        JToken jToken = JToken.Parse(renderValue);
                        finalValue = RenderValueListAsCommaSeperatedString(jToken);
                    }

                    list = list + finalValue;
                }
            }

            return list;
        }
        #endregion

        #region RenderValueListAsCommaSeperatedString
        private string RenderValueListAsCommaSeperatedString(Newtonsoft.Json.Linq.JToken data)
        {
            string list = "";

            foreach (JToken token in data)
            {
                string finalValue = "";
                var renderKeyName = token["RenderKeyName"].ToString();
                var renderKeyValue = token["RenderKeyValue"].ToString();
                var isRender = token["IsRender"].ToString();
                var dataType = token["DataType"].ToString();
                var renderType = token["RenderType"].ToString();
                var renderValue = token["RenderValue"].ToString();
                if (isRender == "1")
                {
                    if (dataType == "Item")
                    {
                        finalValue = RenderItemValue(renderKeyName, renderKeyValue, renderType, renderValue);
                    }
                    else if (dataType == "List" && renderType == "RenderValueListToString")
                    {
                        JToken jToken = JToken.Parse(renderValue);
                        finalValue = RenderValueListToString(jToken);
                    }
                    else if (dataType == "List" && renderType == "RenderValueListAsCommaSeperatedString")
                    {
                        JToken jToken = JToken.Parse(renderValue);
                        finalValue = RenderValueListAsCommaSeperatedString(jToken);
                    }

                    if (list.Length == 0)
                    {
                        list = finalValue;
                    }
                    else
                    {
                        list = list + ", " + finalValue;
                    }
                }
            }
            return list;
        }
        #endregion

        #region HasIsChiledItemToRender
        private bool HasChiledItemToRender(JToken data)
        {
            foreach (JToken token in data)
            {
                if (token["IsRender"].ToString() == "1")
                    return true;
            }
            return false;
        }
        #endregion

        #region TableRowHasColumnsToRender
        private bool TableRowHasColumnsToRender(JEnumerable<JObject> data)
        {
            foreach (JObject token in data)
            {
                if (token["IsRender"].ToString() == "1")
                    return true;
            }
            return false;
        }
        #endregion

        #region TableHasChiledItemToRender
        private bool TableHasChiledItemToRender(JToken data)
        {
            foreach (JObject item in data.Children<JObject>())
            {
                var columnCollection = item["Coloums"].Children<JObject>();

                //Logic to Add Header Row to table
                if (item["IsHeader"]["RenderKeyValue"].ToString() != "1")
                {
                    foreach (JObject columnitem in columnCollection)
                    {
                        if(columnitem["IsRender"].ToString() == "1")
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }
        #endregion

        #region TableRowCountToRender
        private int TableRowCountToRender(JToken data)
        {
            int counter = 0;
            foreach (JObject item in data.Children<JObject>())
            {
                var columnCollection = item["Coloums"].Children<JObject>();
                var isRender = false;

                if (item["IsHeader"]["RenderKeyValue"].ToString() == "1")
                {
                    foreach (JObject columnitem in columnCollection)
                    {
                        if (columnitem["IsRender"].ToString() == "1")
                        {
                            isRender = true;
                            break;
                        }
                    }
                }
                else 
                {
                    foreach (JObject columnitem in columnCollection)
                    {
                        if (columnitem["IsRender"].ToString() == "1")
                        {
                            isRender = true;
                            break;
                        }
                    }
                }
                if(isRender)
                {
                    counter++;
                }
            }
            return counter;
        }
        #endregion

        #region AddBulletedListToDocument
        private void AddBulletedListToDocument(
            Microsoft.Office.Interop.Word.Document doc,
            string bulletListKeyName,
            JToken data)
        {
            object oEndOfDoc = bulletListKeyName;
            if (doc.Bookmarks.Exists(bulletListKeyName))
            {
                Range pragraphRange = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                FindAndReplaceText(doc, bulletListKeyName, "");
                Microsoft.Office.Interop.Word.Paragraph bulletParagraph = doc.Content.Paragraphs.Add(pragraphRange);
                bulletParagraph.Range.ListFormat.ApplyBulletDefault();
                string finalValue = "";
                int counter = 1;
                try
                {
                    //Get list of nodes to process in bullet list with IsRender value as true or 1
                    JArray datatoProcess = new JArray();
                    
                    foreach (JToken token in data)
                    {
                        if (token["IsRender"].ToString() == "1")
                        {
                            datatoProcess.Add(token);
                        }
                    }

                    foreach (JToken token in datatoProcess)
                    {
                        var renderKeyName = token["RenderKeyName"].ToString();
                        var renderKeyValue = token["RenderKeyValue"].ToString();
                        var isRender = token["IsRender"].ToString();
                        var dataType = token["DataType"].ToString();
                        var renderType = token["RenderType"].ToString();
                        var renderValue = token["RenderValue"].ToString();

                        if (isRender == "1")
                        {
                            if (dataType == "Item")
                            {
                                finalValue = RenderItemValue(renderKeyName, renderKeyValue, renderType, renderValue);
                            }
                            else if (dataType == "List" && renderType == "RenderValueListToString")
                            {
                                JToken jToken = JToken.Parse(renderValue);
                                finalValue = RenderValueListToString(jToken);
                            }
                            else if (dataType == "List" && renderType == "RenderValueListAsCommaSeperatedString")
                            {
                                JToken jToken = JToken.Parse(renderValue);
                                finalValue = RenderValueListAsCommaSeperatedString(jToken);
                            }

                            if (counter < datatoProcess.Count())
                            {
                                finalValue = finalValue + "\n";
                            }

                            bulletParagraph.Range.InsertBefore(finalValue);
                        }
                        counter++;
                    }

                    // Cut the range of the InlineShape to clipboard
                    //bulletParagraph.Range.Cut();

                    //// And paste it to the target Range
                    //pragraphRange.Paste();
                }
                catch (Exception)
                {
                }
                finally
                {
                    pragraphRange = null;
                    bulletParagraph = null;
                }
            }
            else
            {
                FindAndReplaceText(doc, bulletListKeyName, "");
            }
        }
        #endregion

        #region AddProcedureCodesTable
        private void AddProcedureCodesTable(
            Microsoft.Office.Interop.Word.Document doc, 
            string keyName, List<KeyValuePair> procedureCodes, 
            string tableGridStyle)
        {
            int numberOfRows = procedureCodes.Count + 1;
            int numberOfColumns = 2;
            int rowCounter = 1;
            object oMissing = Type.Missing;
            object oEndOfDoc = keyName;
            object styleName = tableGridStyle; // "Grid Table 5 Dark";

            Microsoft.Office.Interop.Word.Table oTable;
            if (doc.Bookmarks.Exists(keyName))
            {
                if (procedureCodes.Any())
                {
                    Microsoft.Office.Interop.Word.Range wrdRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    oTable = doc.Tables.Add(wrdRng, numberOfRows, numberOfColumns, ref oMissing, ref oMissing);
                    oTable.Range.ParagraphFormat.SpaceAfter = 6;
                    oTable.Borders.Enable = 1;
                    oTable.set_Style(ref styleName);
                    oTable.Columns.DistributeWidth();

                    //Add Header Row
                    oTable.Cell(1, 1).Range.Text = "Procedure Code";
                    oTable.Cell(1, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    oTable.Cell(1, 1).Range.Bold = 1;
                    oTable.Cell(1, 1).Range.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;

                    oTable.Cell(1, 2).Range.Text = "Procedure Decription";
                    oTable.Cell(1, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    oTable.Cell(1, 2).Range.Bold = 1;
                    oTable.Cell(1, 2).Range.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;

                    rowCounter++;

                    foreach (KeyValuePair item in procedureCodes)
                    {
                        oTable.Cell(rowCounter, 1).Range.Text = item.Key;
                        oTable.Cell(rowCounter, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        oTable.Cell(rowCounter, 1).Range.Bold = 1;
                        oTable.Cell(rowCounter, 1).Range.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;

                        oTable.Cell(rowCounter, 2).Range.Text = item.Value;
                        oTable.Cell(rowCounter, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        oTable.Cell(rowCounter, 2).Range.Bold = 1;
                        oTable.Cell(rowCounter, 2).Range.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;

                        rowCounter++;
                    }
                }
                else
                {
                    FindAndReplaceText(doc, keyName, "");
                }
            }
        }
        #endregion

        #region DeleteBookMarkParagraph
        private void DeleteBookMarkParagraph(Microsoft.Office.Interop.Word.Document doc, string keyName)
        {
            string bookMarkKeyArea = keyName + AREA_BOOKMARK_SUFFIX;
            try
            {
                if (doc.Bookmarks.Exists(bookMarkKeyArea))
                {
                    object oEndOfDoc = bookMarkKeyArea;
                    Range pragraphRange = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    pragraphRange.Delete();
                }
                else if (doc.Bookmarks.Exists(keyName))
                {
                    object oEndOfDoc = keyName;
                    Range pragraphRange = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    pragraphRange.Delete();
                }
            }
            catch (Exception)
            {

            }
        }
        #endregion
    }

    #region KeyValuePair
    public class KeyValuePair
    {
        public string Key { get; set; }
        public string Value { get; set; }
    }
    #endregion
}
