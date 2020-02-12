using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Text;
using System.Linq;
using Inventor;
using System.Xml;

namespace DynamoEngine
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    /// 

    public interface _ComClass
    {
        bool LogCreator(string log, string path, string filename, bool open);
        bool STPSaver(string stpPath);
        bool WriteCustomProp(string propname, string propvalue);
        Tuple<bool, string> GetDocType();
        Tuple<bool, string> DrawingViewChecker();
        bool StartTransaction();
        bool EndTransaction();
        bool AbortTransaction();
        void DocSetter(Document doc);
        void DocResetter();
        Document DocGetter();
        Tuple<bool, string> GetTechnicalRequirements();
        Tuple<bool, List<string>> SearchRefDocs(string str);
        bool RefDocsXMLWriter(List<string> list, string filepath);
        Tuple<bool, List<string>, List<string>> GetDrawingViews();
        bool ScaleDelete(List<string> BadViewNames);
        bool ScaleAdd(List<string> GoodViewNames);
        bool ViewRenamer(List<string> BadViewNames, List<string> GoodViewNames);
        bool RotateViewDetector();
        Tuple<bool, string> ViewsAndSheets();
        Tuple<bool, string> FindEmptyBody();
        Tuple<bool, double, double> SearchThreadPitch();
        bool CreateWorkAxis();
        bool CreateWorkPoint();
        bool CreateWallPlane();
        bool CreateSketchPlane();
        bool CreatePointOnEdge();
        bool CreateSecondAxis();
        bool CreatePlanarSketch(double pitch, double diameter);
        bool DeleteTransientAttrib(List<string> atrlist);
    }

    [GuidAttribute("97663284-6f69-4343-ba0e-b021629ab4bb")]
    public class StandardAddInServer : Inventor.ApplicationAddInServer, _ComClass
    {
        // Inventor application object.
        private Inventor.Application m_inventorApplication;

        public StandardAddInServer()
        {
        }

        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            // This method is called by Inventor when it loads the addin.
            // The AddInSiteObject provides access to the Inventor Application object.
            // The FirstTime flag indicates if the addin is loaded for the first time.

            // Initialize AddIn members.
            m_inventorApplication = addInSiteObject.Application;

            // TODO: Add ApplicationAddInServer.Activate implementation.
            // e.g. event initialization, command creation etc.
        }

        public void Deactivate()
        {
            // This method is called by Inventor when the AddIn is unloaded.
            // The AddIn will be unloaded either manually by the user or
            // when the Inventor session is terminated

            // TODO: Add ApplicationAddInServer.Deactivate implementation

            // Release objects.
            m_inventorApplication = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int commandID)
        {
            // Note:this method is now obsolete, you should use the 
            // ControlDefinition functionality for implementing commands.
        }

        public object Automation
        {
            // This property is provided to allow the AddIn to expose an API 
            // of its own to other programs. Typically, this  would be done by
            // implementing the AddIn's API interface in a class and returning 
            // that class object through this property.

            get
            {
                return this;
            }
        }

        public bool LogCreator(string log, string path, string filename, bool open)
        {
            try
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                if (!System.IO.File.Exists(path + filename))
                {
                    using (StreamWriter writer = new StreamWriter(path + filename))
                    {
                        writer.Write(log);
                    }
                }
                else
                {
                    using (StreamWriter writer = new StreamWriter(path + filename, true))
                    {
                        writer.WriteLine(log);
                    }
                }

                if (open)
                {
                    Process.Start(path + filename);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool STPSaver(string stpPath)
        {
            try
            {
                Document doc = null;

                if (Globals.GlobalDoc != null)
                {
                    doc = Globals.GlobalDoc;
                }
                else
                {
                    return false;
                }

                string propName = string.Empty;
                string propValue = string.Empty;
                Inventor.Application invapp = doc.Parent as Inventor.Application;

                PartDocument partdoc = doc as PartDocument;

                TranslatorAddIn oSTEPTranslator = invapp.ApplicationAddIns.ItemById["{90AF7F40-0C01-11D5-8E83-0010B541CD80}"] as TranslatorAddIn;
                TranslationContext oContext = invapp.TransientObjects.CreateTranslationContext();
                NameValueMap oOptions = invapp.TransientObjects.CreateNameValueMap();

                if (oSTEPTranslator.HasSaveCopyAsOptions[partdoc, oContext, oOptions])
                {
                    oOptions.Value["ApplicationProtocolType"] = 3;
                    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;
                    DataMedium oData = invapp.TransientObjects.CreateDataMedium();
                    oData.FileName = stpPath;
                    oSTEPTranslator.SaveCopyAs(partdoc, oContext, oOptions, oData);

                    propName = "STP Path";
                    propValue = oData.FileName;
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool WriteCustomProp(string propname, string propvalue)
        {
            try
            {
                Document doc = null;

                if (Globals.GlobalDoc != null)
                {
                    doc = Globals.GlobalDoc;
                }
                else
                {
                    return false;
                }

                PropertySet custompropertyset = doc.PropertySets["Inventor User Defined Properties"];
                try
                {
                    Property property = custompropertyset[propname];
                    property.Value = propvalue;
                    return true;
                }
                catch
                {
                    custompropertyset.Add(propvalue, propname);
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        public Tuple<bool, string> GetDocType()
        {
            string type = string.Empty;

            try
            {
                Document doc = null;

                if (Globals.GlobalDoc != null)
                {
                    doc = Globals.GlobalDoc;
                }
                else
                {
                    Tuple<bool, string> tupleError = new Tuple<bool, string>(false, type);
                    return tupleError;
                }

                DocumentTypeEnum docType = doc.DocumentType;

                if (docType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    type = "Assembly";
                }
                else if (docType == DocumentTypeEnum.kDrawingDocumentObject)
                {
                    type = "Drawing";
                }
                else if (docType == DocumentTypeEnum.kPartDocumentObject)
                {
                    type = "Part";
                }
                else if (docType == DocumentTypeEnum.kUnknownDocumentObject)
                {
                    type = "UnknownDocument";
                }

                Tuple<bool, string> tuple = new Tuple<bool, string>(true, type);
                return tuple;
            }
            catch
            {
                Tuple<bool, string> tuple = new Tuple<bool, string>(false, type);
                return tuple;
            }
        }

        public Tuple<bool, string> DrawingViewChecker()
        {
            try
            {
                Document doc = null;

                if (Globals.GlobalDoc != null)
                {
                    doc = Globals.GlobalDoc;
                }
                else
                {
                    string emptyLog = string.Empty;
                    Tuple<bool, string> tupleError = new Tuple<bool, string>(false, emptyLog);
                    return tupleError;
                }

                DrawingDocument drawdoc = doc as DrawingDocument;
                Inventor.Application app = doc.Parent as Inventor.Application;
                Color RedColor = app.TransientObjects.CreateColor(255, 0, 0);
                List<string> result = new List<string>();
                StringBuilder sb = new StringBuilder();
                string log = string.Empty;

                foreach (Sheet sheet in drawdoc.Sheets)
                {
                    foreach (DrawingView view in sheet.DrawingViews)
                    {
                        DrawingCurvesEnumerator viewCurves = view.DrawingCurves;
                        if (view.GeneralDimensionType == GeneralDimensionTypeEnum.kTrueGeneralDimension)
                        {
                            result.Add(view.Name);
                            sb.AppendLine(sheet.Name + " - " + view.Name);
                            foreach (DrawingCurve curve in viewCurves)
                            {
                                curve.Color = RedColor;
                                curve.LineWeight = 0.05;
                            }
                        }
                        else
                        {
                            foreach (DrawingCurve curve in viewCurves)
                            {
                                curve.Color = null;
                                curve.LineWeight = 0.05;
                            }
                        }
                    }
                }
                if (result.Count != 0)
                {
                    log = log + "====== Отчет утилиты DrawingViewChecker ======" + System.Environment.NewLine;
                    log = log + $"Документ: {drawdoc.FullFileName}" + System.Environment.NewLine;
                    log = log + "Найдены виды с фактическим типом размеров:" + System.Environment.NewLine;
                    log = log + sb.ToString();
                    log = log + "====== Завершение работы утилиты DrawingViewChecker ======" + System.Environment.NewLine;
                }
                else
                {
                    log = log + "====== Отчет утилиты DrawingViewChecker ======" + System.Environment.NewLine;
                    log = log + $"Документ: {drawdoc.FullFileName}" + System.Environment.NewLine;
                    log = log + "Видов с фактическими размерами не найдено. Все хорошо." + System.Environment.NewLine;
                    log = log + "====== Завершение работы утилиты DrawingViewChecker ======" + System.Environment.NewLine;
                }
                Tuple<bool, string> tuple = new Tuple<bool, string>(true, log);
                return tuple;
            }
            catch
            {
                string log = string.Empty;
                Tuple<bool, string> tuple = new Tuple<bool, string>(false, log);
                return tuple;
            }
        }

        public bool StartTransaction()
        {
            try
            {
                Document doc = null;

                if (Globals.GlobalDoc != null)
                {
                    doc = Globals.GlobalDoc;
                }
                else
                {
                    return false;
                }

                Inventor.Application app = doc.Parent as Inventor.Application;
                Transaction trans = app.TransactionManager.StartTransaction(app.ActiveDocument, "Transaction");
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool EndTransaction()
        {
            try
            {
                Document doc = null;

                if (Globals.GlobalDoc != null)
                {
                    doc = Globals.GlobalDoc;
                }
                else
                {
                    return false;
                }

                Inventor.Application app = doc.Parent as Inventor.Application;
                Transaction trans = app.TransactionManager.CurrentTransaction;
                trans.End();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool AbortTransaction()
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                return false;
            }

            Inventor.Application app = doc.Parent as Inventor.Application;
            Transaction trans = app.TransactionManager.CurrentTransaction;

            if (trans.DisplayName == "Transaction")
            {
                trans.Abort();
                return true;
            }
            else
            {
                return false;
            }
        }

        public void DocSetter(Document doc)
        {
            Globals.GlobalDoc = doc;
        }

        public void DocResetter()
        {
            Globals.GlobalDoc = null;
        }

        public Document DocGetter()
        {
            Document doc = Globals.GlobalDoc;
            return doc;
        }

        public Tuple<bool, string> GetTechnicalRequirements()
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                Tuple<bool, string> tupleerror = new Tuple<bool, string>(false, string.Empty);
                return tupleerror;
            }

            Application app = doc.Parent as Application;
            Documents docs = app.Documents;
            StringBuilder sb = new StringBuilder();

            foreach (Document Doc in docs)
            {
                if (Doc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                {
                    DrawingDocument drawdoc = Doc as DrawingDocument;
                    foreach (Sheet sheet in drawdoc.Sheets)
                    {
                        DrawingSketches sketches = sheet.Sketches;
                        foreach (DrawingSketch sketch in sketches)
                        {
                            if (sketch.Name == "Технические требования" && sketch.TextBoxes.Count != 0)
                            {
                                foreach (TextBox textbox in sketch.TextBoxes)
                                {
                                    sb.Append(textbox.Text + " ");
                                }
                            }
                        }
                    }
                }
            }
            Tuple<bool, string> tuple = new Tuple<bool, string>(true, sb.ToString());
            return tuple;
        }

        public Tuple<bool, List<string>> SearchRefDocs(string str)
        {
            try
            {
                List<string> list = new List<string>();
                Regex ost = new Regex(@"[О][С][Т]\s?[0-9]?\w?\s?\w*\.\d*\.\d*\-?\d*");
                Regex TU = new Regex(@"[Т][У]\s?\w*[\.\-]\w*[\.\-]\w*[\.\-]\w*\-?[0-9]?[0-9]?");

                MatchCollection OSTmatches = ost.Matches(str);
                if (OSTmatches.Count > 0)
                {
                    foreach (Match match in OSTmatches)
                    {
                        if (!list.Contains(match.Value))
                        {
                            list.Add(match.Value);
                        }
                    }
                }

                MatchCollection TUmatches = TU.Matches(str);
                if (TUmatches.Count > 0)
                {
                    foreach (Match match in TUmatches)
                    {
                        if (!list.Contains(match.Value))
                        {
                            list.Add(match.Value);
                        }
                    }
                }

                var trimmedList = list.Select(st => st.TrimEnd('.')).ToList();
                Tuple<bool, List<string>> tuple = new Tuple<bool, List<string>>(true, trimmedList);
                return tuple;
            }
            catch
            {
                Tuple<bool, List<string>> tuple = new Tuple<bool, List<string>>(false, null);
                return tuple;
            }
        }

        public bool RefDocsXMLWriter(List<string> list, string filepath)
        {
            try
            {
                using (XmlWriter writer = XmlWriter.Create(filepath))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("ReferenceDocuments");

                    foreach (string str in list)
                    {
                        writer.WriteElementString("Document", str);
                    }

                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public Tuple<bool, List<string>, List<string>> GetDrawingViews()
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                Tuple<bool, List<string>, List<string>> tupleerror = new Tuple<bool, List<string>, List<string>>(false, null, null);
                return tupleerror;
            }

            try
            {
                DrawingDocument drawdoc = doc as DrawingDocument;
                List<string> BadViewNames = new List<string>();
                List<string> GoodViewNames = new List<string>();
                PropertySet GOSTProp = drawdoc.PropertySets[6];
                string DocScale = GOSTProp["Масштаб"].Expression;

                foreach (Sheet sheet in drawdoc.Sheets)
                {
                    foreach (DrawingView view in sheet.DrawingViews)
                    {
                        if (view.ShowLabel && view.ParentView != null && (DrawingViewTypeEnum)view.Type != DrawingViewTypeEnum.kStandardDrawingViewType)
                        {
                            if (view.ScaleString == DocScale)
                            {
                                BadViewNames.Add(view.Name);
                            }
                            else
                            {
                                GoodViewNames.Add(view.Name);
                            }
                        }
                    }
                }

                Tuple<bool, List<string>, List<string>> tuple = new Tuple<bool, List<string>, List<string>>(true, BadViewNames, GoodViewNames);
                return tuple;
            }
            catch
            {
                Tuple<bool, List<string>, List<string>> tupleerror = new Tuple<bool, List<string>, List<string>>(false, null, null);
                return tupleerror;
            }
        }

        public bool ScaleDelete(List<string> BadViewNames)
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                return false;
            }

            try
            {
                DrawingDocument drawdoc = doc as DrawingDocument;

                foreach (Sheet sheet in drawdoc.Sheets)
                {
                    foreach (DrawingView view in sheet.DrawingViews)
                    {
                        if (BadViewNames.Contains(view.Name))
                        {
                            if (view.Type != ObjectTypeEnum.kSectionDrawingViewObject && view.Label.FormattedText.Contains("DrawingViewScale"))
                            {
                                string userText = view.Label.FormattedText.Replace("<DrawingViewName/> ( <DrawingViewScale/> )", string.Empty);
                                view.Label.FormattedText = "<DrawingViewName/>" + userText;
                            }
                            else if (view.Type == ObjectTypeEnum.kSectionDrawingViewObject && view.Label.FormattedText.Contains("DrawingViewScale"))
                            {
                                string userText = view.Label.FormattedText.Replace("<DrawingViewName/>-<DrawingViewName/> ( <DrawingViewScale/> )", string.Empty);
                                view.Label.FormattedText = "<DrawingViewName/>-<DrawingViewName/>" + userText;
                            }
                        }
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool ScaleAdd(List<string> GoodViewNames)
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                return false;
            }

            try
            {
                DrawingDocument drawdoc = doc as DrawingDocument;

                foreach (Sheet sheet in drawdoc.Sheets)
                {
                    foreach (DrawingView view in sheet.DrawingViews)
                    {
                        if (GoodViewNames.Contains(view.Name))
                        {
                            if (view.Type != ObjectTypeEnum.kSectionDrawingViewObject && !view.Label.FormattedText.Contains("DrawingViewScale"))
                            {
                                string UserText = view.Label.FormattedText.Replace("<DrawingViewName/>", string.Empty);
                                view.Label.FormattedText = "<DrawingViewName/> ( <DrawingViewScale/> )" + UserText;
                            }
                            else if (view.Type == ObjectTypeEnum.kSectionDrawingViewObject && !view.Label.FormattedText.Contains("DrawingViewScale"))
                            {
                                string UserText = view.Label.FormattedText.Replace("<DrawingViewName/>-<DrawingViewName/>", string.Empty);
                                view.Label.FormattedText = "<DrawingViewName/>-<DrawingViewName/> ( <DrawingViewScale/> )" + UserText;
                            }
                        }
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool ViewRenamer(List<string> BadViewNames, List<string> GoodViewNames)
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                return false;
            }

            try
            {
                List<string> ViewNames = new List<string>();

                List<string> Chars = new List<string>();
                Chars.Add("А");
                Chars.Add("Б");
                Chars.Add("В");
                Chars.Add("Г");
                Chars.Add("Д");
                Chars.Add("Е");
                Chars.Add("Ж");
                Chars.Add("И");
                Chars.Add("К");
                Chars.Add("Л");
                Chars.Add("М");
                Chars.Add("Н");
                Chars.Add("П");
                Chars.Add("Р");
                Chars.Add("С");
                Chars.Add("Т");
                Chars.Add("У");
                Chars.Add("Ф");
                Chars.Add("Ц");
                Chars.Add("Ш");
                Chars.Add("Щ");
                Chars.Add("Э");
                Chars.Add("Ю");
                Chars.Add("Я");

                List<string> RomanNumbers = new List<string>();
                RomanNumbers.Add("I"); //0
                RomanNumbers.Add("II"); //1
                RomanNumbers.Add("III"); //2
                RomanNumbers.Add("IV"); //3
                RomanNumbers.Add("V"); //4
                RomanNumbers.Add("VI"); //5
                RomanNumbers.Add("VII"); //6
                RomanNumbers.Add("VIII"); //7
                RomanNumbers.Add("IX"); //8
                RomanNumbers.Add("X"); //9

                DrawingDocument drawdoc = doc as DrawingDocument;
                int DocViewCount = BadViewNames.Count + GoodViewNames.Count;

                if (DocViewCount > 0)
                {
                    int Cycle = -1;

                    int i = 0;
                    int p = 0;
                    while (Cycle < 5)
                    {
                        i++;
                        if (Cycle == -1)
                        {
                            ViewNames.Add(Chars[i - 1]);
                        }
                        else if (Cycle >= 0)
                        {
                            ViewNames.Add(Chars[i - 1] + RomanNumbers[p - 1]);
                        }

                        if (i % 24 == 0)
                        {
                            Cycle++;
                            i = 0;
                            p++;
                        }
                    }

                    if (DocViewCount < ViewNames.Count)
                    {
                        int z = 0;
                        foreach (Sheet sheet in drawdoc.Sheets)
                        {
                            foreach (DrawingView view in sheet.DrawingViews)
                            {
                                if (view.ShowLabel == true)
                                {
                                    view.Name = ViewNames[z];
                                    z++;
                                }
                            }
                        }
                    }
                }
                return true;
            }
            catch
            {
                List<string> errlist = new List<string>();
                return false;
            }
        }

        public bool RotateViewDetector()
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                return false;
            }

            try
            {
                DrawingDocument drawdoc = doc as DrawingDocument;
                const double pi2 = 2 * Math.PI;

                foreach (Sheet sheet in drawdoc.Sheets)
                {
                    foreach (DrawingView view in sheet.DrawingViews)
                    {
                        Debug.Print($"Вид: {view.Name}");
                        double ViewAngle = Math.Abs(view.Rotation);
                        double result = 0;
                        Debug.Print($"ViewAngle = {ViewAngle}");
                        if (ViewAngle != 0)
                        {
                            result = Math.Round(pi2 - ViewAngle, 4);
                        }
                        Debug.Print($"result = {result}");
                        Debug.Print("-------------");
                        string rotated = "<StyleOverride Font='GOST Common'>" + '\ue94e' + "</StyleOverride>";

                        if (result != 0 /*view.Rotation != -0*/ && !view.Label.FormattedText.Contains(rotated))
                        {
                            if (view.Label.FormattedText.Contains("<Br/>"))
                            {
                                int i = view.Label.FormattedText.IndexOf("<Br/>");
                                view.Label.FormattedText = view.Label.FormattedText.Insert(i, rotated);
                            }
                            else
                            {
                                view.Label.FormattedText = view.Label.FormattedText + rotated;
                            }
                        }
                        else if (result == 0 /*view.Rotation == -0*/ && view.Label.FormattedText.Contains(rotated))
                        {
                            view.Label.FormattedText = view.Label.FormattedText.Replace(rotated, string.Empty);
                        }
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public Tuple<bool, string> ViewsAndSheets()
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                Tuple<bool, string> tupleerror = new Tuple<bool, string>(false, string.Empty);
                return tupleerror;
            }

            try
            {
                DrawingDocument drawdoc = doc as DrawingDocument;
                StringBuilder sb = new StringBuilder();
                string result;

                foreach (Sheet sheet in drawdoc.Sheets)
                {
                    foreach (DrawingView view in sheet.DrawingViews)
                    {
                        if (view.ParentView != null && view.Parent.Name != view.ParentView.Parent.Name)
                        {
                            sb.AppendLine($"Вид {view.Name}, ссылка на -> {view.ParentView.Parent.Name}");
                        }
                    }
                }

                string sbresult = sb.ToString();

                if (sbresult == string.Empty)
                {
                    result = "Перемещенных видов на другие листы нет. Все в порядке." + System.Environment.NewLine + "---------------";
                }
                else
                {
                    result = $"На чертеже {drawdoc.FullDocumentName} обнаружены перемещенные на другие листы виды:" + System.Environment.NewLine + sbresult + "---------------";
                }

                Tuple<bool, string> tuple = new Tuple<bool, string>(true, result);
                return tuple;
            }
            catch
            {
                Tuple<bool, string> tupleerror = new Tuple<bool, string>(false, string.Empty);
                return tupleerror;
            }
        }

        public Tuple<bool, string> FindEmptyBody()
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                Tuple<bool, string> tupleerror = new Tuple<bool, string>(false, string.Empty);
                return tupleerror;
            }

            try
            {
                PartDocument partdoc = doc as PartDocument;
                StringBuilder sb = new StringBuilder();
                string result;

                if (partdoc.ComponentDefinition.SurfaceBodies.Count > 1)
                {
                    foreach (SurfaceBody body in partdoc.ComponentDefinition.SurfaceBodies)
                    {
                        if (body.Faces.Count == 0)
                        {
                            sb.AppendLine($"{body.Name};");
                        }
                    }
                }

                string sbresult = sb.ToString();

                if (sbresult == string.Empty)
                {
                    result = "Пустых тел не обнаружено. Все в порядке." + System.Environment.NewLine + "---------------";
                }
                else
                {
                    int i = 0;
                    foreach (SurfaceBody body in partdoc.ComponentDefinition.SurfaceBodies)
                    {
                        if (body.Faces.Count == 0)
                        {
                            body.Name = "_____EMPTY" + i;
                            i++;
                        }
                    }

                    result = $"В детали {partdoc.FullDocumentName} обнаружены пустые тела:" + System.Environment.NewLine + sbresult + "---------------";
                }

                Tuple<bool, string> tuple = new Tuple<bool, string>(true, result);
                return tuple;
            }
            catch
            {
                Tuple<bool, string> tupleerror = new Tuple<bool, string>(false, string.Empty);
                return tupleerror;
            }
        }

        public Tuple<bool, double, double> SearchThreadPitch()
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                Tuple<bool, double, double> tupleerror = new Tuple<bool, double, double>(false, 0.0, 0.0);
                return tupleerror;
            }

            try
            {
                PartDocument partdoc = doc as PartDocument;

                if (partdoc.SelectSet.Count == 2)
                {
                    AttributeManager attribManager = partdoc.AttributeManager as AttributeManager;

                    Face cylinder = null;
                    Face wall = null;

                    foreach (Face face in partdoc.SelectSet)
                    {
                        if (face.SurfaceType == SurfaceTypeEnum.kCylinderSurface)
                        {
                            cylinder = face;
                            cylinder.AttributeSets.AddTransient("cylinderFace");
                        }
                        else if (face.SurfaceType == SurfaceTypeEnum.kPlaneSurface)
                        {
                            wall = face;
                            wall.AttributeSets.AddTransient("wallFace");
                        }
                    }

                    if (cylinder == null || wall == null)
                    {
                        Tuple<bool, double, double> tupleerror = new Tuple<bool, double, double>(false, 0.0, 0.0);
                        return tupleerror;
                    }
                    else
                    {
                        string thread = string.Empty;
                        double threadPitch = 0.0;
                        double diameter = 0.0;

                        if (cylinder.ThreadInfos.Count != 0)
                        {
                            ThreadInfo threadinfo = cylinder.ThreadInfos[1] as ThreadInfo;
                            thread = threadinfo.ThreadDesignation;
                            string stringPitch = thread.Substring(thread.IndexOf('x') + 1);
                            stringPitch = stringPitch.Replace('.', ',');
                            threadPitch = Convert.ToDouble(stringPitch);

                            int i = thread.IndexOf('x');
                            string stringDiameter = thread.Substring(0, i);
                            stringDiameter = stringDiameter.Replace("M", "");
                            stringDiameter = stringDiameter.Replace('.', ',');
                            diameter = Convert.ToDouble(stringDiameter);
                        }

                        Tuple<bool, double, double> tuple = new Tuple<bool, double, double>(true, threadPitch, diameter);
                        return tuple;
                    }
                }
                else
                {
                    Tuple<bool, double, double> tupleerror = new Tuple<bool, double, double>(false, 0.0, 0.0);
                    return tupleerror;
                }
            }
            catch
            {
                Tuple<bool, double, double> tupleerror = new Tuple<bool, double, double>(false, 0.0, 0.0);
                return tupleerror;
            }

        }

        public bool CreateWorkAxis()
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                return false;
            }

            try
            {
                PartDocument partdoc = doc as PartDocument;
                AttributeManager attribManager = partdoc.AttributeManager as AttributeManager;
                ObjectCollection FindAttrib = attribManager.FindObjects("cylinderFace");
                Face cylinder = FindAttrib[1] as Face;
                WorkAxis axis = partdoc.ComponentDefinition.WorkAxes.AddByRevolvedFace(cylinder);
                axis.Visible = false;
                axis.AttributeSets.AddTransient("cylinderAxis");
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool CreateWorkPoint()
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                return false;
            }

            try
            {
                PartDocument partdoc = doc as PartDocument;
                AttributeManager attribManager = partdoc.AttributeManager as AttributeManager;

                ObjectCollection FindAxis = attribManager.FindObjects("cylinderAxis");
                WorkAxis axis = FindAxis[1] as WorkAxis;

                ObjectCollection FindWall = attribManager.FindObjects("wallFace");
                Face wall = FindWall[1] as Face;

                WorkPoint point = partdoc.ComponentDefinition.WorkPoints.AddByCurveAndEntity(axis, wall);
                point.Visible = false;
                point.AttributeSets.AddTransient("wallPoint");
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool CreateWallPlane()
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                return false;
            }

            try
            {
                PartDocument partdoc = doc as PartDocument;
                AttributeManager attribManager = partdoc.AttributeManager as AttributeManager;

                ObjectCollection FindAxis = attribManager.FindObjects("cylinderAxis");
                WorkAxis axis = FindAxis[1] as WorkAxis;

                ObjectCollection Findpoint = attribManager.FindObjects("wallPoint");
                WorkPoint point = Findpoint[1] as WorkPoint;

                WorkPlane plane = partdoc.ComponentDefinition.WorkPlanes.AddByNormalToCurve(axis, point);
                plane.Visible = false;
                plane.AttributeSets.AddTransient("WallPlane");
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool CreateSketchPlane()
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                return false;
            }

            try
            {
                PartDocument partdoc = doc as PartDocument;
                AttributeManager attribManager = partdoc.AttributeManager as AttributeManager;

                ObjectCollection FindSecondAxis = attribManager.FindObjects("SecondAxis");
                WorkAxis SecondAxis = FindSecondAxis[1] as WorkAxis;

                ObjectCollection FindCylinderAxis = attribManager.FindObjects("cylinderAxis");
                WorkAxis CylinderAxis = FindCylinderAxis[1] as WorkAxis;

                WorkPlane SketchPlane = partdoc.ComponentDefinition.WorkPlanes.AddByTwoLines(CylinderAxis, SecondAxis);
                SketchPlane.Visible = false;
                SketchPlane.AttributeSets.AddTransient("SketchPlane");
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool CreatePointOnEdge()
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                return false;
            }

            try
            {
                PartDocument partdoc = doc as PartDocument;
                AttributeManager attribManager = partdoc.AttributeManager as AttributeManager;

                ObjectCollection WallFace = attribManager.FindObjects("wallFace");
                Face wallFace = WallFace[1] as Face;

                Point PointOnEdge = wallFace.Edges[1].PointOnEdge;

                WorkPoint point = partdoc.ComponentDefinition.WorkPoints.AddFixed(PointOnEdge);
                point.Visible = false;
                point.AttributeSets.AddTransient("PointOnEdge");
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool CreateSecondAxis()
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                return false;
            }

            try
            {
                PartDocument partdoc = doc as PartDocument;
                AttributeManager attribManager = partdoc.AttributeManager as AttributeManager;

                ObjectCollection Cylinderaxis = attribManager.FindObjects("cylinderAxis");
                WorkAxis axis = Cylinderaxis[1] as WorkAxis;

                ObjectCollection PointOnEdge = attribManager.FindObjects("PointOnEdge");
                WorkPoint point = PointOnEdge[1] as WorkPoint;

                WorkAxis SecondAxis = partdoc.ComponentDefinition.WorkAxes.AddByLineAndPoint(axis, point);
                SecondAxis.Visible = false;
                SecondAxis.AttributeSets.AddTransient("SecondAxis");
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool CreatePlanarSketch(double pitch, double diameter)
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                return false;
            }

            try
            {
                PartDocument partdoc = doc as PartDocument;
                Application app = partdoc.Parent as Application;
                TransientGeometry oTransGeom = app.TransientGeometry;
                PlanarSketches Sketches = partdoc.ComponentDefinition.Sketches;
                AttributeManager attribManager = partdoc.AttributeManager as AttributeManager;

                ObjectCollection Cylinderaxis = attribManager.FindObjects("cylinderAxis");
                WorkAxis axis = Cylinderaxis[1] as WorkAxis;

                ObjectCollection Findpoint = attribManager.FindObjects("wallPoint");
                WorkPoint point = Findpoint[1] as WorkPoint;

                ObjectCollection FindSketchPlane = attribManager.FindObjects("SketchPlane");
                WorkPlane SketchPlane = FindSketchPlane[1] as WorkPlane;

                ObjectCollection FindWallPlane = attribManager.FindObjects("WallPlane");
                WorkPlane WallPlane = FindWallPlane[1] as WorkPlane;


                //проверка вектора
                ObjectCollection WallFace = attribManager.FindObjects("wallFace");
                Face Face = WallFace[1] as Face;
                Point pointOnFace = Face.PointOnFace;
                SurfaceEvaluator evaluator = Face.Evaluator;
                double[] param = new double[3];
                param[0] = pointOnFace.X;
                param[1] = pointOnFace.Y;
                param[2] = pointOnFace.Z;
                double[] normal = new double[3];
                evaluator.GetNormalAtPoint(param, ref normal);
                UnitVector vector1 = oTransGeom.CreateUnitVector(normal[0], normal[1], normal[2]);
                UnitVector vector2 = axis.Line.Direction;
                bool axisDirection = vector1.IsEqualTo(vector2);
                //конец проверки вектора

                PlanarSketch Sketch = Sketches.AddWithOrientation(SketchPlane, axis, axisDirection, true, point);

                Sketch.Edit();

                #region Анурьев
                double boreDiameter = 0.0;
                double grooveWidth = 0.0;
                double smallRad = 0.0;
                double largeRad = 0.0;

                if (pitch <= 0.4)
                {
                    boreDiameter = 0.6;
                    grooveWidth = 1.0;
                    largeRad = 0.3;
                    smallRad = 0.2;
                }
                else if (pitch == 0.45)
                {
                    boreDiameter = 0.7;
                    grooveWidth = 1.0;
                    largeRad = 0.3;
                    smallRad = 0.2;
                }
                else if (pitch == 0.5)
                {
                    boreDiameter = 0.8;
                    grooveWidth = 1.0;
                    largeRad = 0.3;
                    smallRad = 0.2;
                }
                else if (pitch == 0.6)
                {
                    boreDiameter = 0.9;
                    grooveWidth = 1.0;
                    largeRad = 0.3;
                    smallRad = 0.2;
                }
                else if (pitch == 0.7)
                {
                    boreDiameter = 1.0;
                    grooveWidth = 1.6;
                    largeRad = 0.5;
                    smallRad = 0.3;
                }
                else if (pitch == 0.75)
                {
                    boreDiameter = 1.2;
                    grooveWidth = 1.6;
                    largeRad = 0.5;
                    smallRad = 0.3;
                }
                else if (pitch == 0.8)
                {
                    boreDiameter = 1.2;
                    grooveWidth = 1.6;
                    largeRad = 0.5;
                    smallRad = 0.3;
                }
                else if (pitch == 1.0)
                {
                    boreDiameter = 1.5;
                    grooveWidth = 2.0;
                    largeRad = 0.5;
                    smallRad = 0.3;
                }
                else if (pitch == 1.25)
                {
                    boreDiameter = 1.8;
                    grooveWidth = 2.5;
                    largeRad = 1.0;
                    smallRad = 0.5;
                }
                else if (pitch == 1.5)
                {
                    boreDiameter = 2.2;
                    grooveWidth = 2.5;
                    largeRad = 1.0;
                    smallRad = 0.5;
                }
                else if (pitch == 1.75)
                {
                    boreDiameter = 2.5;
                    grooveWidth = 2.5;
                    largeRad = 1.0;
                    smallRad = 0.5;
                }
                else if (pitch == 2.0)
                {
                    boreDiameter = 3.0;
                    grooveWidth = 3.0;
                    largeRad = 1.0;
                    smallRad = 0.5;
                }
                else if (pitch == 2.5)
                {
                    boreDiameter = 3.5;
                    grooveWidth = 4.0;
                    largeRad = 1.0;
                    smallRad = 0.5;
                }
                else if (pitch == 3.0)
                {
                    boreDiameter = 4.5;
                    grooveWidth = 4.0;
                    largeRad = 1.0;
                    smallRad = 0.5;
                }
                else if (pitch == 3.5)
                {
                    boreDiameter = 5.0;
                    grooveWidth = 5.0;
                    largeRad = 1.6;
                    smallRad = 0.5;
                }
                else if (pitch == 4.0)
                {
                    boreDiameter = 6.0;
                    grooveWidth = 5.0;
                    largeRad = 1.6;
                    smallRad = 0.5;
                }
                else if (pitch == 4.5)
                {
                    boreDiameter = 6.5;
                    grooveWidth = 6.0;
                    largeRad = 1.6;
                    smallRad = 1.0;
                }
                else if (pitch == 5.0)
                {
                    boreDiameter = 7.0;
                    grooveWidth = 6.0;
                    largeRad = 1.6;
                    smallRad = 1.0;
                }
                else if (pitch == 5.5)
                {
                    boreDiameter = 8.0;
                    grooveWidth = 8.0;
                    largeRad = 2.0;
                    smallRad = 1.0;
                }
                else if (pitch >= 6.0)
                {
                    boreDiameter = 9.0;
                    grooveWidth = 8.0;
                    largeRad = 2.0;
                    smallRad = 1.0;
                }
                #endregion

                double halfDiameter = (diameter * 0.1) / 2;

                SketchLine cilinderaxis = Sketch.AddByProjectingEntity(axis) as SketchLine;
                cilinderaxis.Centerline = true;
                cilinderaxis.Construction = true;

                SketchLine wallLine = Sketch.AddByProjectingEntity(WallPlane) as SketchLine;
                wallLine.Construction = true;

                Point2d p1 = oTransGeom.CreatePoint2d(1, 1);
                Point2d p2 = oTransGeom.CreatePoint2d(2, 1);
                Point2d p3 = oTransGeom.CreatePoint2d(2, 2);
                Point2d p4 = oTransGeom.CreatePoint2d(1, 2);

                SketchLine line1 = Sketch.SketchLines.AddByTwoPoints(p1, p2);
                SketchLine line2 = Sketch.SketchLines.AddByTwoPoints(p2, p3);
                SketchLine line3 = Sketch.SketchLines.AddByTwoPoints(p3, p4);
                SketchLine line4 = Sketch.SketchLines.AddByTwoPoints(p4, p1);

                Sketch.GeometricConstraints.AddCoincident(line1.EndSketchPoint as SketchEntity, line2 as SketchEntity);
                Sketch.GeometricConstraints.AddCoincident(line2.StartSketchPoint as SketchEntity, line1 as SketchEntity);
                Sketch.GeometricConstraints.AddCoincident(line2.EndSketchPoint as SketchEntity, line3 as SketchEntity);
                Sketch.GeometricConstraints.AddCoincident(line3.StartSketchPoint as SketchEntity, line2 as SketchEntity);
                Sketch.GeometricConstraints.AddCoincident(line3.EndSketchPoint as SketchEntity, line4 as SketchEntity);
                Sketch.GeometricConstraints.AddCoincident(line4.StartSketchPoint as SketchEntity, line3 as SketchEntity);
                Sketch.GeometricConstraints.AddCoincident(line4.EndSketchPoint as SketchEntity, line1 as SketchEntity);
                Sketch.GeometricConstraints.AddCoincident(line1.StartSketchPoint as SketchEntity, line4 as SketchEntity);
                Sketch.GeometricConstraints.AddCollinear(line4 as SketchEntity, wallLine as SketchEntity);

                DimensionConstraint boreDiameterConstraint = Sketch.DimensionConstraints.AddOffset(line1, cilinderaxis as SketchEntity, p1, true) as DimensionConstraint;
                boreDiameterConstraint.Parameter.Value = (diameter * 0.1) - (boreDiameter * 0.1);

                DimensionConstraint cilinderDiameterConstraint = Sketch.DimensionConstraints.AddOffset(line3, cilinderaxis as SketchEntity, p3, true) as DimensionConstraint;
                cilinderDiameterConstraint.Parameter.Value = (diameter * 0.1) + (5 * 0.1);

                DimensionConstraint grooveWidthConstraint = Sketch.DimensionConstraints.AddTwoPointDistance(line1.StartSketchPoint, line1.EndSketchPoint, DimensionOrientationEnum.kHorizontalDim, p1) as DimensionConstraint;
                grooveWidthConstraint.Parameter.Value = grooveWidth * 0.1;

                SketchArc smallArc = Sketch.SketchArcs.AddByFillet(line1 as SketchEntity, line2 as SketchEntity, smallRad * 0.1, line1.StartSketchPoint.Geometry, line2.EndSketchPoint.Geometry);
                DimensionConstraint smallArcRad = Sketch.DimensionConstraints.AddRadius(smallArc as SketchEntity, p2) as DimensionConstraint;

                SketchArc largeArc = Sketch.SketchArcs.AddByFillet(line1 as SketchEntity, line4 as SketchEntity, largeRad * 0.1, line1.EndSketchPoint.Geometry, line4.StartSketchPoint.Geometry);
                DimensionConstraint largeArcRad = Sketch.DimensionConstraints.AddRadius(largeArc as SketchEntity, p1) as DimensionConstraint;

                DimensionConstraint angle = Sketch.DimensionConstraints.AddTwoLineAngle(line1, line2, p3) as DimensionConstraint;
                angle.Parameter.Value = (45 * Math.PI) / 180;

                Sketch.ExitEdit();

                partdoc.Update();

                Profile Profile = Sketch.Profiles.AddForSolid();

                partdoc.ComponentDefinition.Features.RevolveFeatures.AddFull(Profile, axis, PartFeatureOperationEnum.kCutOperation);

                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool DeleteTransientAttrib(List<string> atrlist)
        {
            Document doc = null;

            if (Globals.GlobalDoc != null)
            {
                doc = Globals.GlobalDoc;
            }
            else
            {
                return false;
            }

            AttributeManager oAttribManager = doc.AttributeManager;
            AttributeSetsEnumerator FindTransientAttrib;

            foreach (string str in atrlist)
            {
                FindTransientAttrib = oAttribManager.FindAttributeSets(str);

                if (FindTransientAttrib.Count != 0)
                {
                    foreach (AttributeSet item in FindTransientAttrib)
                    {
                        item.Delete();
                    }
                }
            }

            return true;
        }
    }
}
