using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using KodeMagd.Reporter;

namespace KodeMagd.WorkbookAnalysis
{
    class ClsGenerateFlowDiagramme
    {
        //private const int ciHorSpacing = 80;
        //private const int ciVertSpacing = 100;

        //public enum enumSymbolType
        //{
        //    //eSym_Input_FromSheet, /*rectangle with S-bottom line*/
        //    //eSym_Input_FromUser,/*Rectangle with top sloped at a angle*/
        //    //eSym_Input_FromOther,/**/
        //    //eSym_Output_ToSheet,/*rectangle with S-bottom line*/
        //    //eSym_Output_ToOther,/**/
        //    //eSym_SubRoutineFunction, /*Rectangle with double vertical lines*/
        //    //eSym_BeginningOfRepeatativeStructure, /*hexagon*/

        //    eSym_VDU,/*curved shape looks like a CRT*/
        //    eSym_Data, /*parallelagram*/
        //    eSym_Decision, /*diamond*/
        //    eSym_Terminator, /*Round left and right flat top and bottom*/
        //    eSym_MagneticDisk,/*cylinder*/
        //    eSym_Process, /*rectangle*/
        //    eSym_Document,/*rectangle with S-bottom line*/
        //    eSym_PredefinedProcess,/*rectangle with double vertical sides*/
        //    eSym_ManualInput,
        //    eSym_AlternativeProcess,
        //    eSym_MultiDocument,
        //    eSym_InternalStorage,
        //    eSym_Sort,
        //    eSym_ConnectingFlows
        //}

        //public enum enumSize
        //{
        //    eSmall,
        //    eMedium,
        //    eLarge,
        //}

        //public enum enumLoopType
        //{
        //    eCondition_Before,
        //    eCondition_After
        //}

        //public enum enumConnectionNodePosition
        //{
        //    eTop,
        //    eBottom,
        //    eLeft,
        //    eRight
        //}

        //public struct strConnectionNode
        //{
        //    public enumConnectionNodePosition ePos;
        //    public int iHor;/*relative to shape*/
        //    public int iVert;/*relative to shape*/
        //}

        //public struct strSymbol
        //{
        //    public int id;
        //    public List<int> lstSymbolsGoingTo;
        //    public List<int> lstSymbolsComingFrom;
        //    public enumSymbolType eType;
        //    public string sVbaCode;
        //    public string sCaption;
        //    public int iVertPos;
        //    public int iHorPos;
        //    public enumSize eSize;
        //}

        //public struct strLoops
        //{
        //    public int iSymbolStart;
        //    public int iSymbolEnd;
        //    public enumLoopType eType;
        //    public int iMaxDepth; //related to the order of the nexted loops
        //    public int iDepthPx; //related to the pixels required to draw around any symbols or other loops
        //}

        //public struct strIf
        //{
        //    public int iSymbolIF;
        //    public List<int> iSymbolElseIf;
        //    public int iSymbolElse;
        //    public int iSymbolEnd;
        //    public enumLoopType eType;
        //    public int iMaxDepth;
        //}

        //public enum enumLineType 
        //{
        //    eStraight,
        //    eCurved,
        //    eCircle,
        //}

        //public struct strCoordinates 
        //{
        //    public int iHor;
        //    public int iVert;
        //}

        //public struct strStraightLine
        //{
        //    public int order;
        //    public bool freshStart;
        //    public strCoordinates start;
        //    public strCoordinates end;
        //}

        //public struct strCurvedLine
        //{
        //    public int order;
        //    public bool freshStart;
        //    public strCoordinates CentreOfRotation;
        //    public int radiusHor;
        //    public int radiusVert;
        //    public double startAngle; 
        //    public double endAngle; 
        //}

        //public struct strShape 
        //{
        //    public List<strStraightLine> lstStraightLine;
        //    public List<strCurvedLine> lstCurvedLine;
        //    public List<strConnectionNode> lstConnectionNodes;
        //    public enumSymbolType eSymbol;
        //    public enumSize eSize;
        //}

        //public struct strSquare
        //{
        //    public int iLeft;
        //    public int iTop;
        //    public int iHeight;
        //    public int iWidth;
        //    public List<strSquare> lstSquares;
        //    public List<strSymbol> lstSymbols;
        //    public List<strLoops> lstLoops;
        //}

        //List<ClsConfigReporterFlowDiagram.strShape> lstShapes;
        //List<ClsConfigReporterFlowDiagram.strSymbol> lstSymbols = new List<ClsConfigReporterFlowDiagram.strSymbol>();
        //Stack<strLoops> stkLoops = new Stack<strLoops>();
        //List<ClsConfigReporterFlowDiagram.strLoops> lstLoops = new List<ClsConfigReporterFlowDiagram.strLoops>();

        //int iCanvasWidth;
        //int iCanvasHeight;
        string sModuleName;
        string sFunctionName;

        //public int canvasWidth
        //{
        //    get
        //    {
        //        try
        //        {
        //            return iCanvasWidth;
        //        }
        //        catch (Exception ex)
        //        {
        //            MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //            string sMessage = "";

        //            sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //            sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //            sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //            sMessage += ex.Message;

        //            MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

        //            return 0;
        //        }
        //    }
        //}

        //public int canvasHeight
        //{
        //    get
        //    {
        //        try
        //        {
        //            return iCanvasHeight;
        //        }
        //        catch (Exception ex)
        //        {
        //            MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //            string sMessage = "";

        //            sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //            sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //            sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //            sMessage += ex.Message;

        //            MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

        //            return 0;
        //        }
        //    }
        //}

        //public ClsGenerateFlowDiagramme()
        //{
        //    try
        //    {
        //        fillShapes();
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        //private void fillShapes()
        //{
        //    try
        //    {
        //        lstShapes = new List<strShape>();

        //        fillShapesTerminator();
        //        fillShapesMagneticDisk();
        //        fillShapesDocument();
        //        fillShapesDecision();
        //        fillShapesVDU();
        //        fillShapesData();
        //        fillShapesProcess();
        //        fillShapesPredefinedProcess();
        //        fillShapesManualInput();
        //        fillShapesAlternativeProcess();
        //        fillShapesMultiDocument();
        //        fillShapesInternalStorage();
        //        fillShapesSort();
        //        fillShapesConnectingFlows();

        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}


        //private void fillShapesTerminator()
        //{
        //    try
        //    {
        //        strShape objShape = new strShape();
        //        strStraightLine objStraightLine = new strStraightLine();
        //        strCurvedLine objCurvedLine = new strCurvedLine();
        //        strConnectionNode objConnectionNode = new strConnectionNode();
        //        int iOrder = 0;

        //        objShape.lstCurvedLine = new List<strCurvedLine>();
        //        objShape.lstStraightLine = new List<strStraightLine>();

        //        /*Start*/
        //        objShape.eSymbol = enumSymbolType.eSym_Terminator;
        //        objShape.eSize = enumSize.eLarge;
        //        objShape.lstConnectionNodes = new List<strConnectionNode>();

        //        objCurvedLine = new strCurvedLine();
        //        iOrder++;
        //        objCurvedLine.order = iOrder;
        //        objCurvedLine.freshStart = true;
        //        objCurvedLine.CentreOfRotation.iHor = 40;
        //        objCurvedLine.CentreOfRotation.iVert = 0;
        //        objCurvedLine.startAngle = -0.5 * Math.PI;
        //        objCurvedLine.endAngle = 0.5 * Math.PI;
        //        objCurvedLine.radiusHor = 15;
        //        objCurvedLine.radiusVert = 15;
        //        objShape.lstCurvedLine.Add(objCurvedLine);

        //        objCurvedLine = new strCurvedLine();
        //        iOrder++;
        //        objCurvedLine.order = iOrder;
        //        objCurvedLine.freshStart = false;
        //        objCurvedLine.CentreOfRotation.iHor = -40;
        //        objCurvedLine.CentreOfRotation.iVert = 0;
        //        objCurvedLine.startAngle = 0.5 * Math.PI;
        //        objCurvedLine.endAngle = 1.5 * Math.PI;
        //        objCurvedLine.radiusHor = 15;
        //        objCurvedLine.radiusVert = 15;
        //        objShape.lstCurvedLine.Add(objCurvedLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = -40;
        //        objStraightLine.start.iVert = -15;
        //        objStraightLine.end.iHor = 40;
        //        objStraightLine.end.iVert = -15;
        //        objShape.lstStraightLine.Add(objStraightLine);


        //        objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = 15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eTop;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = -15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        lstShapes.Add(objShape);
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        //private void fillShapesMagneticDisk()
        //{
        //    try
        //    {
        //        strShape objShape = new strShape();
        //        strStraightLine objStraightLine = new strStraightLine();
        //        strCurvedLine objCurvedLine = new strCurvedLine();
        //        strConnectionNode objConnectionNode = new strConnectionNode();
        //        int iOrder = 0;

        //        objShape.lstCurvedLine = new List<strCurvedLine>();
        //        objShape.lstStraightLine = new List<strStraightLine>();
        //        objShape.lstConnectionNodes = new List<strConnectionNode>();

        //        /*Start*/
        //        objShape.eSymbol = enumSymbolType.eSym_MagneticDisk;
        //        objShape.eSize = enumSize.eLarge;

        //        objCurvedLine = new strCurvedLine();
        //        iOrder++;
        //        objCurvedLine.order = iOrder;
        //        objCurvedLine.freshStart = true;
        //        objCurvedLine.CentreOfRotation.iHor = 0;
        //        objCurvedLine.CentreOfRotation.iVert = -20;
        //        objCurvedLine.startAngle = 0;
        //        objCurvedLine.endAngle = 2 * Math.PI;
        //        objCurvedLine.radiusHor = 64;
        //        objCurvedLine.radiusVert = 32;
        //        objShape.lstCurvedLine.Add(objCurvedLine);

        //        objCurvedLine = new strCurvedLine();
        //        iOrder++;
        //        objCurvedLine.order = iOrder;
        //        objCurvedLine.freshStart = false;
        //        objCurvedLine.CentreOfRotation.iHor = 0;
        //        objCurvedLine.CentreOfRotation.iVert = 20;
        //        objCurvedLine.startAngle = 0;
        //        objCurvedLine.endAngle = Math.PI;
        //        objCurvedLine.radiusHor = 64;
        //        objCurvedLine.radiusVert = 32;
        //        objShape.lstCurvedLine.Add(objCurvedLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = -48;
        //        objStraightLine.start.iVert = -20;
        //        objStraightLine.end.iHor = -48;
        //        objStraightLine.end.iVert = -20;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eTop;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = -15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = 15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eRight;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);
                
        //        lstShapes.Add(objShape);

        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        
        //private void fillShapesConnectingFlows()
        //{
        //    try
        //    {
        //        strShape objShape = new strShape();
        //        strStraightLine objStraightLine = new strStraightLine();
        //        strCurvedLine objCurvedLine = new strCurvedLine();
        //        strConnectionNode objConnectionNode = new strConnectionNode();
        //        int iOrder = 0;

        //        objShape.lstCurvedLine = new List<strCurvedLine>();
        //        objShape.lstStraightLine = new List<strStraightLine>();
        //        objShape.lstConnectionNodes = new List<strConnectionNode>();

        //        /*Start*/
        //        //objShape.eSymbol = enumSymbolType.eSym_Output_ToDatabase;
        //        objShape.eSize = enumSize.eLarge;

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = -15;
        //        objStraightLine.start.iVert = -20;
        //        objStraightLine.end.iHor = 15;
        //        objStraightLine.end.iVert = -20;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 15;
        //        objStraightLine.start.iVert = -20;
        //        objStraightLine.end.iHor = 0;
        //        objStraightLine.end.iVert = 20;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 0;
        //        objStraightLine.start.iVert = 20;
        //        objStraightLine.end.iHor = -15;
        //        objStraightLine.end.iVert = -20;
        //        objShape.lstStraightLine.Add(objStraightLine);


        //        objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = 20;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 20;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eRight;
        //        objConnectionNode.iHor = 15;
        //        objConnectionNode.iVert = 20;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objShape.eSymbol = enumSymbolType.eSym_ConnectingFlows;
        //        lstShapes.Add(objShape);
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}



        //private void fillShapesDocument()
        //{
        //    try
        //    {
        //        strShape objShape = new strShape();
        //        strStraightLine objStraightLine = new strStraightLine();
        //        strCurvedLine objCurvedLine = new strCurvedLine();
        //        strConnectionNode objConnectionNode = new strConnectionNode();
        //        int iOrder = 0;

        //        objShape.lstCurvedLine = new List<strCurvedLine>();
        //        objShape.lstStraightLine = new List<strStraightLine>();
        //        objShape.lstConnectionNodes = new List<strConnectionNode>();

        //        /*Start*/
        //        //objShape.eSymbol = enumSymbolType.eSym_Output_ToDatabase;
        //        objShape.eSize = enumSize.eLarge;

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = -40;
        //        objStraightLine.start.iVert = 32;
        //        objStraightLine.end.iHor = -40;
        //        objStraightLine.end.iVert = -32;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = -40;
        //        objStraightLine.start.iVert = -32;
        //        objStraightLine.end.iHor = 40;
        //        objStraightLine.end.iVert = -32;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = 40;
        //        objStraightLine.start.iVert = -32;
        //        objStraightLine.end.iHor = 40;
        //        objStraightLine.end.iVert = 32;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objCurvedLine = new strCurvedLine();
        //        iOrder++;
        //        objCurvedLine.order = iOrder;
        //        objCurvedLine.freshStart = false;
        //        objCurvedLine.CentreOfRotation.iHor = 20;
        //        objCurvedLine.CentreOfRotation.iVert = 32;
        //        objCurvedLine.startAngle = 0;
        //        objCurvedLine.endAngle = Math.PI;
        //        objCurvedLine.radiusHor = 22;
        //        objCurvedLine.radiusVert = 18;
        //        objShape.lstCurvedLine.Add(objCurvedLine);

        //        objCurvedLine = new strCurvedLine();
        //        iOrder++;
        //        objCurvedLine.order = iOrder;
        //        objCurvedLine.freshStart = false;
        //        objCurvedLine.CentreOfRotation.iHor = -20;
        //        objCurvedLine.CentreOfRotation.iVert = 32;
        //        objCurvedLine.startAngle = 0;
        //        objCurvedLine.endAngle = -1 * Math.PI;
        //        objCurvedLine.radiusHor = 22;
        //        objCurvedLine.radiusVert = 18;
        //        objShape.lstCurvedLine.Add(objCurvedLine);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eTop;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = -15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = 15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eRight;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objShape.eSymbol = enumSymbolType.eSym_Document;
        //        lstShapes.Add(objShape);
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        //private void fillShapesMultiDocument()
        //{
        //    try
        //    {
        //        strShape objShape = new strShape();
        //        strStraightLine objStraightLine = new strStraightLine();
        //        strCurvedLine objCurvedLine = new strCurvedLine();
        //        strConnectionNode objConnectionNode = new strConnectionNode();
        //        int iOrder = 0;

        //        objShape.lstCurvedLine = new List<strCurvedLine>();
        //        objShape.lstStraightLine = new List<strStraightLine>();
        //        objShape.lstConnectionNodes = new List<strConnectionNode>();

        //        /*Start*/
        //        objShape.eSymbol = enumSymbolType.eSym_MultiDocument;
        //        objShape.eSize = enumSize.eLarge;

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = -40;
        //        objStraightLine.start.iVert = 32;
        //        objStraightLine.end.iHor = -40;
        //        objStraightLine.end.iVert = -32;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = -40;
        //        objStraightLine.start.iVert = -32;
        //        objStraightLine.end.iHor = 40;
        //        objStraightLine.end.iVert = -32;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = 40;
        //        objStraightLine.start.iVert = -32;
        //        objStraightLine.end.iHor = 40;
        //        objStraightLine.end.iVert = 32;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objCurvedLine = new strCurvedLine();
        //        iOrder++;
        //        objCurvedLine.order = iOrder;
        //        objCurvedLine.freshStart = false;
        //        objCurvedLine.CentreOfRotation.iHor = 20;
        //        objCurvedLine.CentreOfRotation.iVert = 32;
        //        objCurvedLine.startAngle = 0;
        //        objCurvedLine.endAngle = Math.PI;
        //        objCurvedLine.radiusHor = 22;
        //        objCurvedLine.radiusVert = 18;
        //        objShape.lstCurvedLine.Add(objCurvedLine);

        //        objCurvedLine = new strCurvedLine();
        //        iOrder++;
        //        objCurvedLine.order = iOrder;
        //        objCurvedLine.freshStart = false;
        //        objCurvedLine.CentreOfRotation.iHor = -20;
        //        objCurvedLine.CentreOfRotation.iVert = 32;
        //        objCurvedLine.startAngle = 0;
        //        objCurvedLine.endAngle = -1 * Math.PI;
        //        objCurvedLine.radiusHor = 22;
        //        objCurvedLine.radiusVert = 18;
        //        objShape.lstCurvedLine.Add(objCurvedLine);

        //        /*shadow 1*/
        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = 40;
        //        objStraightLine.start.iVert = 27;
        //        objStraightLine.end.iHor = 45;
        //        objStraightLine.end.iVert = 27;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 45;
        //        objStraightLine.start.iVert = 27;
        //        objStraightLine.end.iHor = 45;
        //        objStraightLine.end.iVert = -37;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 45;
        //        objStraightLine.start.iVert = -37;
        //        objStraightLine.end.iHor = -35;
        //        objStraightLine.end.iVert = -37;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = -35;
        //        objStraightLine.start.iVert = -37;
        //        objStraightLine.end.iHor = -35;
        //        objStraightLine.end.iVert = -32;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        /*shadow 2*/
        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = 45;
        //        objStraightLine.start.iVert = 22;
        //        objStraightLine.end.iHor = 50;
        //        objStraightLine.end.iVert = 22;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 50;
        //        objStraightLine.start.iVert = 22;
        //        objStraightLine.end.iHor = 50;
        //        objStraightLine.end.iVert = -42;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 50;
        //        objStraightLine.start.iVert = -42;
        //        objStraightLine.end.iHor = -30;
        //        objStraightLine.end.iVert = -42;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = -30;
        //        objStraightLine.start.iVert = -42;
        //        objStraightLine.end.iHor = -30;
        //        objStraightLine.end.iVert = -37;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eTop;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = -15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = 15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eRight;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        lstShapes.Add(objShape);
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        //private void fillShapesDecision()
        //{
        //    try
        //    {
        //        strShape objShape = new strShape();
        //        strStraightLine objStraightLine = new strStraightLine();
        //        strCurvedLine objCurvedLine = new strCurvedLine();
        //        strConnectionNode objConnectionNode = new strConnectionNode();
        //        int iOrder = 0;

        //        objShape.lstCurvedLine = new List<strCurvedLine>();
        //        objShape.lstStraightLine = new List<strStraightLine>();
        //        objShape.lstConnectionNodes = new List<strConnectionNode>();

        //        /*Start*/
        //        //objShape.eSymbol = enumSymbolType.eSym_Output_ToDatabase;
        //        objShape.eSize = enumSize.eLarge;

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = -40;
        //        objStraightLine.start.iVert = 0;
        //        objStraightLine.end.iHor = 0;
        //        objStraightLine.end.iVert = -40;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 0;
        //        objStraightLine.start.iVert = -40;
        //        objStraightLine.end.iHor = 40;
        //        objStraightLine.end.iVert = 0;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 40;
        //        objStraightLine.start.iVert = 0;
        //        objStraightLine.end.iHor = 0;
        //        objStraightLine.end.iVert = 40;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 0;
        //        objStraightLine.start.iVert = 40;
        //        objStraightLine.end.iHor = -40;
        //        objStraightLine.end.iVert = 0;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eTop;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = -15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = 15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eRight;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objShape.eSymbol = enumSymbolType.eSym_Decision;
        //        lstShapes.Add(objShape);

        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        //private void fillShapesSort()
        //{
        //    try
        //    {
        //        strShape objShape = new strShape();
        //        strStraightLine objStraightLine = new strStraightLine();
        //        strCurvedLine objCurvedLine = new strCurvedLine();
        //        strConnectionNode objConnectionNode = new strConnectionNode();
        //        int iOrder = 0;

        //        objShape.lstCurvedLine = new List<strCurvedLine>();
        //        objShape.lstStraightLine = new List<strStraightLine>();
        //        objShape.lstConnectionNodes = new List<strConnectionNode>();

        //        /*Start*/
        //        objShape.eSymbol = enumSymbolType.eSym_Sort;
        //        objShape.eSize = enumSize.eLarge;

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = -40;
        //        objStraightLine.start.iVert = 0;
        //        objStraightLine.end.iHor = 0;
        //        objStraightLine.end.iVert = -40;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 0;
        //        objStraightLine.start.iVert = -40;
        //        objStraightLine.end.iHor = 40;
        //        objStraightLine.end.iVert = 0;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 40;
        //        objStraightLine.start.iVert = 0;
        //        objStraightLine.end.iHor = 0;
        //        objStraightLine.end.iVert = 40;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 0;
        //        objStraightLine.start.iVert = 40;
        //        objStraightLine.end.iHor = -40;
        //        objStraightLine.end.iVert = 0;
        //        objShape.lstStraightLine.Add(objStraightLine);
        //        /*
        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = 0;
        //        objStraightLine.start.iVert = 40;
        //        objStraightLine.end.iHor = 0;
        //        objStraightLine.end.iVert = -40;
        //        objShape.lstStraightLine.Add(objStraightLine);
        //        */
        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = 40;
        //        objStraightLine.start.iVert = 0;
        //        objStraightLine.end.iHor = -40;
        //        objStraightLine.end.iVert = 0;
        //        objShape.lstStraightLine.Add(objStraightLine);
                
        //        objConnectionNode.ePos = enumConnectionNodePosition.eTop;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = -40;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = 40;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
        //        objConnectionNode.iHor = -40;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eRight;
        //        objConnectionNode.iHor = -40;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        lstShapes.Add(objShape);

        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        //private void fillShapesData()
        //{
        //    try
        //    {
        //        strShape objShape = new strShape();
        //        strStraightLine objStraightLine = new strStraightLine();
        //        strCurvedLine objCurvedLine = new strCurvedLine();
        //        strConnectionNode objConnectionNode = new strConnectionNode();
        //        int iOrder = 0;

        //        objShape.lstCurvedLine = new List<strCurvedLine>();
        //        objShape.lstStraightLine = new List<strStraightLine>();
        //        objShape.lstConnectionNodes = new List<strConnectionNode>();

        //        /*Start*/
        //        objShape.eSymbol = enumSymbolType.eSym_Data;
        //        objShape.eSize = enumSize.eLarge;

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = -50;
        //        objStraightLine.start.iVert = 25;
        //        objStraightLine.end.iHor = 30;
        //        objStraightLine.end.iVert = 25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 30;
        //        objStraightLine.start.iVert = 25;
        //        objStraightLine.end.iHor = 50;
        //        objStraightLine.end.iVert = -25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 50;
        //        objStraightLine.start.iVert = -25;
        //        objStraightLine.end.iHor = -30;
        //        objStraightLine.end.iVert = -25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = -30;
        //        objStraightLine.start.iVert = -25;
        //        objStraightLine.end.iHor = -50;
        //        objStraightLine.end.iVert = 25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eTop;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = -15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = 15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eRight;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        lstShapes.Add(objShape);
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        //private void fillShapesProcess()
        //{
        //    try
        //    {
        //        strShape objShape = new strShape();
        //        strStraightLine objStraightLine = new strStraightLine();
        //        strCurvedLine objCurvedLine = new strCurvedLine();
        //        strConnectionNode objConnectionNode = new strConnectionNode();
        //        int iOrder = 0;

        //        objShape.lstCurvedLine = new List<strCurvedLine>();
        //        objShape.lstStraightLine = new List<strStraightLine>();
        //        objShape.lstConnectionNodes = new List<strConnectionNode>();

        //        /*Start*/
        //        objShape.eSymbol = enumSymbolType.eSym_Process;
        //        objShape.eSize = enumSize.eLarge;

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = -50;
        //        objStraightLine.start.iVert = 25;
        //        objStraightLine.end.iHor = 50;
        //        objStraightLine.end.iVert = 25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 50;
        //        objStraightLine.start.iVert = 25;
        //        objStraightLine.end.iHor = 50;
        //        objStraightLine.end.iVert = -25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 50;
        //        objStraightLine.start.iVert = -25;
        //        objStraightLine.end.iHor = -50;
        //        objStraightLine.end.iVert = -25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = -50;
        //        objStraightLine.start.iVert = -25;
        //        objStraightLine.end.iHor = -50;
        //        objStraightLine.end.iVert = 25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eTop;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = -15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = 15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eRight;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        lstShapes.Add(objShape);
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}


        //private void fillShapesInternalStorage()
        //{
        //    try
        //    {
        //        strShape objShape = new strShape();
        //        strStraightLine objStraightLine = new strStraightLine();
        //        strCurvedLine objCurvedLine = new strCurvedLine();
        //        strConnectionNode objConnectionNode = new strConnectionNode();
        //        int iOrder = 0;

        //        objShape.lstCurvedLine = new List<strCurvedLine>();
        //        objShape.lstStraightLine = new List<strStraightLine>();
        //        objShape.lstConnectionNodes = new List<strConnectionNode>();

        //        /*Start*/
        //        objShape.eSymbol = enumSymbolType.eSym_InternalStorage;
        //        objShape.eSize = enumSize.eLarge;

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = -40;
        //        objStraightLine.start.iVert = 40;
        //        objStraightLine.end.iHor = 40;
        //        objStraightLine.end.iVert = 40;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 40;
        //        objStraightLine.start.iVert = 40;
        //        objStraightLine.end.iHor = 40;
        //        objStraightLine.end.iVert = -40;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 40;
        //        objStraightLine.start.iVert = -40;
        //        objStraightLine.end.iHor = -40;
        //        objStraightLine.end.iVert = -40;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = -40;
        //        objStraightLine.start.iVert = -40;
        //        objStraightLine.end.iHor = -40;
        //        objStraightLine.end.iVert = 40;
        //        objShape.lstStraightLine.Add(objStraightLine);


        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = -35;
        //        objStraightLine.start.iVert = -40;
        //        objStraightLine.end.iHor = -35;
        //        objStraightLine.end.iVert = 40;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = -40;
        //        objStraightLine.start.iVert = -35;
        //        objStraightLine.end.iHor = 40;
        //        objStraightLine.end.iVert = -35;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eTop;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = -15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = 15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eRight;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        lstShapes.Add(objShape);
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        //private void fillShapesAlternativeProcess()
        //{
        //    try
        //    {
        //        strShape objShape = new strShape();
        //        strStraightLine objStraightLine = new strStraightLine();
        //        strCurvedLine objCurvedLine = new strCurvedLine();
        //        strConnectionNode objConnectionNode = new strConnectionNode();
        //        int iOrder = 0;

        //        objShape.lstCurvedLine = new List<strCurvedLine>();
        //        objShape.lstStraightLine = new List<strStraightLine>();
        //        objShape.lstConnectionNodes = new List<strConnectionNode>();

        //        /*Start*/
        //        objShape.eSymbol = enumSymbolType.eSym_AlternativeProcess;
        //        objShape.eSize = enumSize.eLarge;

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = -40;
        //        objStraightLine.start.iVert = 25;
        //        objStraightLine.end.iHor = 40;
        //        objStraightLine.end.iVert = 25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objCurvedLine = new strCurvedLine();
        //        iOrder++;
        //        objCurvedLine.order = iOrder;
        //        objCurvedLine.freshStart = false;
        //        objCurvedLine.CentreOfRotation.iHor = 40;
        //        objCurvedLine.CentreOfRotation.iVert = 15;
        //        objCurvedLine.startAngle = 0.5 * Math.PI;
        //        objCurvedLine.endAngle = 0 * Math.PI;
        //        objCurvedLine.radiusHor = 10;
        //        objCurvedLine.radiusVert = 10;
        //        objShape.lstCurvedLine.Add(objCurvedLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 50;
        //        objStraightLine.start.iVert = 15;
        //        objStraightLine.end.iHor = 50;
        //        objStraightLine.end.iVert = -15;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objCurvedLine = new strCurvedLine();
        //        iOrder++;
        //        objCurvedLine.order = iOrder;
        //        objCurvedLine.freshStart = false;
        //        objCurvedLine.CentreOfRotation.iHor = 40;
        //        objCurvedLine.CentreOfRotation.iVert = -15;
        //        objCurvedLine.startAngle = 0 * Math.PI;
        //        objCurvedLine.endAngle = -0.5 * Math.PI;
        //        objCurvedLine.radiusHor = 10;
        //        objCurvedLine.radiusVert = 10;
        //        objShape.lstCurvedLine.Add(objCurvedLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 40;
        //        objStraightLine.start.iVert = -25;
        //        objStraightLine.end.iHor = -40;
        //        objStraightLine.end.iVert = -25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objCurvedLine = new strCurvedLine();
        //        iOrder++;
        //        objCurvedLine.order = iOrder;
        //        objCurvedLine.freshStart = false;
        //        objCurvedLine.CentreOfRotation.iHor = -40;
        //        objCurvedLine.CentreOfRotation.iVert = -15;
        //        objCurvedLine.startAngle = -0.5 * Math.PI;
        //        objCurvedLine.endAngle = -1 * Math.PI;
        //        objCurvedLine.radiusHor = 10;
        //        objCurvedLine.radiusVert = 10;
        //        objShape.lstCurvedLine.Add(objCurvedLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = -50;
        //        objStraightLine.start.iVert = -15;
        //        objStraightLine.end.iHor = -50;
        //        objStraightLine.end.iVert = 15;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objCurvedLine = new strCurvedLine();
        //        iOrder++;
        //        objCurvedLine.order = iOrder;
        //        objCurvedLine.freshStart = false;
        //        objCurvedLine.CentreOfRotation.iHor = -40;
        //        objCurvedLine.CentreOfRotation.iVert = 15;
        //        objCurvedLine.startAngle = -1 * Math.PI;
        //        objCurvedLine.endAngle = -1.5 * Math.PI;
        //        objCurvedLine.radiusHor = 10;
        //        objCurvedLine.radiusVert = 10;
        //        objShape.lstCurvedLine.Add(objCurvedLine);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eTop;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = -15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = 15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eRight;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        lstShapes.Add(objShape);
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        //private void fillShapesManualInput()
        //{
        //    try
        //    {
        //        strShape objShape = new strShape();
        //        strStraightLine objStraightLine = new strStraightLine();
        //        strCurvedLine objCurvedLine = new strCurvedLine();
        //        strConnectionNode objConnectionNode = new strConnectionNode();
        //        int iOrder = 0;

        //        objShape.lstCurvedLine = new List<strCurvedLine>();
        //        objShape.lstStraightLine = new List<strStraightLine>();
        //        objShape.lstConnectionNodes = new List<strConnectionNode>();

        //        /*Start*/
        //        objShape.eSymbol = enumSymbolType.eSym_ManualInput;
        //        objShape.eSize = enumSize.eLarge;

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = -50;
        //        objStraightLine.start.iVert = 25;
        //        objStraightLine.end.iHor = 50;
        //        objStraightLine.end.iVert = 25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 50;
        //        objStraightLine.start.iVert = 25;
        //        objStraightLine.end.iHor = 50;
        //        objStraightLine.end.iVert = -25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 50;
        //        objStraightLine.start.iVert = -25;
        //        objStraightLine.end.iHor = -50;
        //        objStraightLine.end.iVert = -10;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = -50;
        //        objStraightLine.start.iVert = -10;
        //        objStraightLine.end.iHor = -50;
        //        objStraightLine.end.iVert = 25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eTop;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = -15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = 15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eRight;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        lstShapes.Add(objShape);
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        //private void fillShapesPredefinedProcess()
        //{
        //    try
        //    {
        //        strShape objShape = new strShape();
        //        strStraightLine objStraightLine = new strStraightLine();
        //        strCurvedLine objCurvedLine = new strCurvedLine();
        //        strConnectionNode objConnectionNode = new strConnectionNode();
        //        int iOrder = 0;

        //        objShape.lstCurvedLine = new List<strCurvedLine>();
        //        objShape.lstStraightLine = new List<strStraightLine>();
        //        objShape.lstConnectionNodes = new List<strConnectionNode>();

        //        /*Start*/
        //        objShape.eSymbol = enumSymbolType.eSym_PredefinedProcess;
        //        objShape.eSize = enumSize.eLarge;

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = -50;
        //        objStraightLine.start.iVert = 25;
        //        objStraightLine.end.iHor = 50;
        //        objStraightLine.end.iVert = 25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 50;
        //        objStraightLine.start.iVert = 25;
        //        objStraightLine.end.iHor = 50;
        //        objStraightLine.end.iVert = -25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = 50;
        //        objStraightLine.start.iVert = -25;
        //        objStraightLine.end.iHor = -50;
        //        objStraightLine.end.iVert = -25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = false;
        //        objStraightLine.start.iHor = -50;
        //        objStraightLine.start.iVert = -25;
        //        objStraightLine.end.iHor = -50;
        //        objStraightLine.end.iVert = 25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = -40;
        //        objStraightLine.start.iVert = -25;
        //        objStraightLine.end.iHor = -40;
        //        objStraightLine.end.iVert = 25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objStraightLine = new strStraightLine();
        //        iOrder++;
        //        objStraightLine.order = iOrder;
        //        objStraightLine.freshStart = true;
        //        objStraightLine.start.iHor = 40;
        //        objStraightLine.start.iVert = -25;
        //        objStraightLine.end.iHor = 40;
        //        objStraightLine.end.iVert = 25;
        //        objShape.lstStraightLine.Add(objStraightLine);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eTop;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = -15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = 15;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eRight;
        //        objConnectionNode.iHor = -15;
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        lstShapes.Add(objShape);
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        //private void fillShapesVDU()
        //{
        //    try
        //    {
        //        strShape objShape = new strShape();
        //        strStraightLine objStraightLine = new strStraightLine();
        //        strCurvedLine objCurvedLine = new strCurvedLine();
        //        strConnectionNode objConnectionNode = new strConnectionNode();
        //        int iOrder = 0;

        //        objShape.lstCurvedLine = new List<strCurvedLine>();
        //        objShape.lstStraightLine = new List<strStraightLine>();
        //        objShape.lstConnectionNodes = new List<strConnectionNode>();

        //        /*Start*/
        //        //objShape.eSymbol = enumSymbolType.eSym_Output_ToDatabase;
        //        objShape.eSize = enumSize.eLarge;

        //        objCurvedLine = new strCurvedLine();
        //        iOrder++;
        //        objCurvedLine.order = iOrder;
        //        objCurvedLine.freshStart = true;
        //        objCurvedLine.CentreOfRotation.iHor = 0;
        //        objCurvedLine.CentreOfRotation.iVert = 16;
        //        objCurvedLine.startAngle = Math.PI + Math.Sin(2.0 / 5.0);
        //        objCurvedLine.endAngle = 1.5 * Math.PI;
        //        objCurvedLine.radiusHor = 40;
        //        objCurvedLine.radiusVert = 40;
        //        objShape.lstCurvedLine.Add(objCurvedLine);

        //        objCurvedLine = new strCurvedLine();
        //        iOrder++;
        //        objCurvedLine.order = iOrder;
        //        objCurvedLine.freshStart = false;
        //        objCurvedLine.CentreOfRotation.iHor = -15;
        //        objCurvedLine.CentreOfRotation.iVert = 0;
        //        objCurvedLine.startAngle = -1.0 * Math.Sin(3.0 / 4.0);
        //        objCurvedLine.endAngle = Math.Sin(3.0 / 4.0);
        //        objCurvedLine.radiusHor = 40;
        //        objCurvedLine.radiusVert = 40;
        //        objShape.lstCurvedLine.Add(objCurvedLine);

        //        objCurvedLine = new strCurvedLine();
        //        iOrder++;
        //        objCurvedLine.order = iOrder;
        //        objCurvedLine.freshStart = false;
        //        objCurvedLine.CentreOfRotation.iHor = 0;
        //        objCurvedLine.CentreOfRotation.iVert = -16;
        //        objCurvedLine.startAngle = 0.5*Math.PI;
        //        objCurvedLine.endAngle = Math.PI-Math.Sin(2.0 / 5.0);
        //        objCurvedLine.radiusHor = 40;
        //        objCurvedLine.radiusVert = 40;
        //        objShape.lstCurvedLine.Add(objCurvedLine);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
        //        objConnectionNode.iHor = (int)(Math.Sqrt((40 ^ 2) - (16 ^ 2)));
        //        objConnectionNode.iVert = 0;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = 30;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objConnectionNode.ePos = enumConnectionNodePosition.eTop;
        //        objConnectionNode.iHor = 0;
        //        objConnectionNode.iVert = -30;
        //        objShape.lstConnectionNodes.Add(objConnectionNode);

        //        objShape.eSymbol = enumSymbolType.eSym_VDU;
        //        lstShapes.Add(objShape);

        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        //private void drawShape(strSymbol objSymbol, ref List<string> lstHtml)
        //{
        //    try
        //    {
        //        if (lstShapes.Exists(x => x.eSymbol == objSymbol.eType && x.eSize == objSymbol.eSize))
        //        {
        //            string sLine = "";
        //            strShape objShape = lstShapes.Find(x => x.eSymbol == objSymbol.eType && x.eSize == objSymbol.eSize);

        //            int iOrderMin = 0;
        //            int iOrderMax = 0;

        //            if (objShape.lstCurvedLine.Count() > 0)
        //            {
        //                iOrderMin = objShape.lstCurvedLine.Min(x => x.order);
        //                iOrderMax = objShape.lstCurvedLine.Max(x => x.order);
        //            }

        //            if (objShape.lstStraightLine.Count() > 0)
        //            {
        //                int iOrderStriaghtMin = objShape.lstStraightLine.Min(x => x.order);
        //                int iOrderStriaghtMax = objShape.lstStraightLine.Max(x => x.order);

        //                if (iOrderMin > iOrderStriaghtMin)
        //                { iOrderMin = iOrderStriaghtMin; }

        //                if (iOrderMax < iOrderStriaghtMax)
        //                { iOrderMax = iOrderStriaghtMax; }
        //            }

        //            bool bIsContinuing = false;
        //            int iPreviousHor = 0;
        //            int iPreviousVert = 0;

        //            for (int iOrderCounter = iOrderMin; iOrderCounter <= iOrderMax; iOrderCounter++)
        //            {

        //                sLine = "";
        //                foreach (strCurvedLine objCurvedLine in objShape.lstCurvedLine.FindAll(x => x.order == iOrderCounter))
        //                {
        //                    sLine = "";
        //                    double dHorScale = 1;
        //                    double dVertScale = 1;
        //                    double dHorUnScale = 1;
        //                    double dVertUnScale = 1;

        //                    /* Scale */
        //                    if (objCurvedLine.radiusHor != objCurvedLine.radiusVert)
        //                    {
        //                        dHorScale = 0;
        //                        dVertScale = 0;

        //                        if (objCurvedLine.radiusHor > objCurvedLine.radiusVert)
        //                        {
        //                            dHorScale = 1.0;
        //                            dVertScale = (double)objCurvedLine.radiusVert / (double)objCurvedLine.radiusHor;
        //                            dHorUnScale = 1.0;
        //                            dVertUnScale = (double)objCurvedLine.radiusHor / (double)objCurvedLine.radiusVert;
        //                        }
        //                        else
        //                        {
        //                            dHorScale = (double)objCurvedLine.radiusHor / (double)objCurvedLine.radiusVert;
        //                            dVertScale = 1.0;
        //                            dHorUnScale = (double)objCurvedLine.radiusVert / (double)objCurvedLine.radiusHor;
        //                            dVertUnScale = 1.0;
        //                        }

        //                        sLine += "ctx.scale(" + dHorScale.ToString() + ", " + dVertScale.ToString() + ");\n"; 
        //                    }

        //                    if (objCurvedLine.freshStart)
        //                    {
        //                        int iStartHor = (int)((dHorUnScale * ((double)objCurvedLine.CentreOfRotation.iHor + (double)objSymbol.iHorPos)) + (double)(Math.Cos(objCurvedLine.startAngle) * (double)(objCurvedLine.radiusHor + objCurvedLine.radiusVert) / 2.0));
        //                        int iStartVert = (int)((dVertUnScale * ((double)objCurvedLine.CentreOfRotation.iVert + (double)objSymbol.iVertPos)) + (double)(Math.Sin(objCurvedLine.startAngle) * (double)(objCurvedLine.radiusHor + objCurvedLine.radiusVert) / 2.0));

        //                        sLine += "ctx.moveTo(" + iStartHor.ToString() + ", " + iStartVert.ToString() + ");\n";
        //                    }

        //                    /*arc*/
        //                    sLine += "ctx.arc(" + (dHorUnScale * (objCurvedLine.CentreOfRotation.iHor + objSymbol.iHorPos)).ToString() + ", " + (dVertUnScale * (objCurvedLine.CentreOfRotation.iVert + objSymbol.iVertPos)).ToString() + ",";
        //                    sLine += ((objCurvedLine.radiusHor + objCurvedLine.radiusVert) / 2).ToString() + ",";
        //                    sLine += objCurvedLine.startAngle.ToString() + "," + objCurvedLine.endAngle.ToString();

        //                    if (objCurvedLine.startAngle < objCurvedLine.endAngle)
        //                    { sLine += ");\n"; }
        //                    else
        //                    { sLine += ", true);\n"; }

        //                    /* Un-scale */
        //                    if (objCurvedLine.radiusHor != objCurvedLine.radiusVert)
        //                    {
        //                        sLine += "ctx.scale(" + dHorUnScale.ToString() + ", " + dVertUnScale.ToString() + ");\n";
        //                    }

        //                    lstHtml.Add(sLine);
        //                }

        //                sLine = "";
        //                foreach (strStraightLine objStraightLine in objShape.lstStraightLine.FindAll(x => x.order == iOrderCounter))
        //                {
        //                    sLine = "";
        //                    if (objStraightLine.freshStart)
        //                    { sLine += "ctx.moveTo(" + (objStraightLine.start.iHor + objSymbol.iHorPos).ToString() + ", " + (objStraightLine.start.iVert + objSymbol.iVertPos).ToString() + ");\n"; }
        //                    sLine += "ctx.lineTo(" + (objStraightLine.end.iHor + objSymbol.iHorPos).ToString() + ", " + (objStraightLine.end.iVert + objSymbol.iVertPos).ToString() + ");\n";

        //                    lstHtml.Add(sLine);
        //                }
        //            }
        //            sLine = "ctx.fillText('" + objSymbol.sCaption + "', " + objSymbol.iHorPos + ", " + objSymbol.iVertPos + ");";
        //            lstHtml.Add(sLine);
        //            sLine = "ctx.stroke();\n";
        //            lstHtml.Add(sLine);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        public string moduleName
        {
            get 
            {
                try
                {
                    return sModuleName;
                }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = "";

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                    return string.Empty;
                }
            }
            set 
            {
                try
                {
                    sModuleName = value;
                }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = "";

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
                }
            }
        }


        public string functionName
        {
            get
            {
                try
                {
                    return sFunctionName;
                }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = "";

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                    return string.Empty;
                }
            }
            set
            {
                try
                {
                    sFunctionName = value;
                }
                catch (Exception ex)
                {
                    MethodBase mbTemp = MethodBase.GetCurrentMethod();

                    string sMessage = "";

                    sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                    sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                    sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                    sMessage += ex.Message;

                    MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
                }
            }
        }

        public void generate(ref List<string> lstHtml, List<ClsCodeMapper.strLine> lstLines)
        {
            try
            {
                ClsConfigReporterFlowDiagram cConfigReporterFlowDiagram = new ClsConfigReporterFlowDiagram();

                calculateSymbols(ref lstLines, ref cConfigReporterFlowDiagram);

                //cConfigReporterFlowDiagram.groupIntoSquares();

                cConfigReporterFlowDiagram.reposition();

                cConfigReporterFlowDiagram.GenerateHtml(ref lstHtml);
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }


        private void addToLoopList(ref List<ClsConfigReporterFlowDiagram.strLoops> lstLoops, int iSymbolId, ClsConfigReporterFlowDiagram.enumLoopType eLoopType)
        {
            try
            {
                ClsConfigReporterFlowDiagram.strLoops objLoopNew = new ClsConfigReporterFlowDiagram.strLoops();

                objLoopNew.iSymbolStart = iSymbolId;
                objLoopNew.eType = eLoopType;
                objLoopNew.iMaxDepth = 0;

                lstLoops.Add(objLoopNew);

                for (int iPos = 0; iPos < lstLoops.Count(); iPos++)
                {
                    ClsConfigReporterFlowDiagram.strLoops objLoop = lstLoops[iPos];

                    if (objLoop.iMaxDepth < lstLoops.Count() - iPos)
                    { objLoop.iMaxDepth = lstLoops.Count() - iPos; }

                    lstLoops[iPos] = objLoop;
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        private ClsConfigReporterFlowDiagram.strLoops removeFromLoopList(ref List<ClsConfigReporterFlowDiagram.strLoops> lstLoops, int iSymbolId)
        {
            try
            {
                ClsConfigReporterFlowDiagram.strLoops objResult = new ClsConfigReporterFlowDiagram.strLoops();

                if (lstLoops.Count > 0)
                {
                    objResult = lstLoops[lstLoops.Count - 1];
                    lstLoops.RemoveAt(lstLoops.Count - 1);
                    objResult.iSymbolEnd = iSymbolId; 
                }

                return objResult;
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

                return new ClsConfigReporterFlowDiagram.strLoops();
            }
        }


        private void calculateSymbols(ref List<ClsCodeMapper.strLine> lstLines, ref ClsConfigReporterFlowDiagram cConfigReporterFlowDiagram)
        {
            try
            {
                /*
                 * Every symbol and loop goes in a square
                 * Every square gets nested in another square
                 * The list lstSquareId will keep track of all the current nested squares this function is dealing with.
                 * 
                 * 
                 */


                //List<strLoops> lstLoops = new List<strLoops>();

                List<ClsConfigReporterFlowDiagram.strLoops> lstLoopsTemp_For = new List<ClsConfigReporterFlowDiagram.strLoops>();
                List<ClsConfigReporterFlowDiagram.strLoops> lstLoopsTemp_DoLoop = new List<ClsConfigReporterFlowDiagram.strLoops>();
                List<ClsConfigReporterFlowDiagram.strLoops> lstLoopsTemp_WhileWend = new List<ClsConfigReporterFlowDiagram.strLoops>();
                List<ClsConfigReporterFlowDiagram.strLoops> lstLoopsTemp_Other = new List<ClsConfigReporterFlowDiagram.strLoops>();
                List<ClsConfigReporterFlowDiagram.strIf> lstIfTemp = new List<ClsConfigReporterFlowDiagram.strIf>();
                ClsConfigReporterFlowDiagram.strSymbol objSymbol = new ClsConfigReporterFlowDiagram.strSymbol();
                ClsConfigReporterFlowDiagram.strSymbol objSymbolPrevious = new ClsConfigReporterFlowDiagram.strSymbol();
                //ClsConfigReporterFlowDiagram.strSquare objSquare = new ClsConfigReporterFlowDiagram.strSquare();
                List<int> lstSquareId = new List<int>();
                int iCurrentSquareId = cConfigReporterFlowDiagram.addSquare();
                cConfigReporterFlowDiagram.rootOfAllSquaresId = iCurrentSquareId;
                //int iCurrentId = 1;

                int iSymbolId = 0;
                int iSymbolPreviousId = 0;

                //iSymbolPreviousId = cConfigReporterFlowDiagram.addSquare();
                lstSquareId.Add(iCurrentSquareId);
                //cConfigReporterFlowDiagram.addSquare();

                foreach (ClsCodeMapper.strLine objLine in lstLines.FindAll(x => x.sText_Orig.Trim() != ""))
                {
                    if (!string.IsNullOrWhiteSpace(objLine.sText_NoComment))
                    {
                        //objSymbol = new strSymbol();

                        objSymbol.id = 0;
                        objSymbol.eSize = ClsConfigReporterFlowDiagram.enumSize.eLarge;
                        objSymbol.lstSymbolsComingFrom = new List<int>();
                        objSymbol.lstSymbolsGoingTo = new List<int>();

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_AssignValue))
                        {
                            objSymbol.eType = ClsConfigReporterFlowDiagram.enumSymbolType.eSym_Process;
                            objSymbol.iHorPos = 0;
                            objSymbol.iVertPos = 0;
                            objSymbol.sCaption = objLine.sText_NoComment.Trim();
                            objSymbol.lstSymbolsComingFrom.Add(objSymbolPrevious.id);
                            objSymbolPrevious.lstSymbolsGoingTo.Add(objSymbol.id);

                            iSymbolId = cConfigReporterFlowDiagram.addSymbol(objSymbol);

                            cConfigReporterFlowDiagram.squareAddSymbol(iCurrentSquareId, iSymbolId);
                            //iCurrentId++;
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_BeginLoop))
                        {
                            if (objLine.sText_NoComment.ToLower().Contains("=") || objLine.sText_NoComment.ToLower().Contains("until") || objLine.sText_NoComment.ToLower().Contains("while"))
                            {
                                objSymbol.eType = ClsConfigReporterFlowDiagram.enumSymbolType.eSym_Decision;
                            }
                            else
                            {
                                objSymbol.eType = ClsConfigReporterFlowDiagram.enumSymbolType.eSym_ConnectingFlows;
                            }
                            
                            objSymbol.iHorPos = 0;
                            objSymbol.iVertPos = 0;
                            objSymbol.lstSymbolsComingFrom.Add(objSymbolPrevious.id);
                            objSymbolPrevious.lstSymbolsGoingTo.Add(objSymbol.id);
                            objSymbol.sCaption = objLine.sText_NoComment.Trim();

                            ClsConfigReporterFlowDiagram.enumLoopType eLoopType;

                            if (objLine.sText_NoComment.Trim().ToLower().StartsWith("for ")
                                || objLine.sText_NoComment.Trim().ToLower().StartsWith("for each ")
                                || objLine.sText_NoComment.Trim().ToLower().StartsWith("do while ")
                                || objLine.sText_NoComment.Trim().ToLower().StartsWith("while ")
                                || objLine.sText_NoComment.Trim().ToLower().StartsWith("until "))
                            { eLoopType = ClsConfigReporterFlowDiagram.enumLoopType.eCondition_Before; }
                            else
                            { eLoopType = ClsConfigReporterFlowDiagram.enumLoopType.eCondition_After; }

                            if (objLine.sText_NoComment.Trim().ToLower().StartsWith("for ")
                                || objLine.sText_NoComment.Trim().ToLower().StartsWith("for each "))
                            { addToLoopList(ref lstLoopsTemp_For, objSymbol.id, eLoopType); }
                            else if (objLine.sText_NoComment.Trim().ToLower().StartsWith("do "))
                            { addToLoopList(ref lstLoopsTemp_DoLoop, objSymbol.id, eLoopType); }
                            else if (objLine.sText_NoComment.Trim().ToLower().StartsWith("while "))
                            { addToLoopList(ref lstLoopsTemp_WhileWend, objSymbol.id, eLoopType); }
                            else
                            { addToLoopList(ref lstLoopsTemp_Other, objSymbol.id, eLoopType); }

                            //iSymbolPreviousId = cConfigReporterFlowDiagram.addSymbol(objSymbol);
                            int iNewSquareId = cConfigReporterFlowDiagram.addSquare();

                            lstSquareId.Add(iNewSquareId);
                            cConfigReporterFlowDiagram.squareAddSquare(iCurrentSquareId, iNewSquareId);

                            iSymbolId = cConfigReporterFlowDiagram.addSymbol(objSymbol);
                            iCurrentSquareId = iNewSquareId;
                            //cConfigReporterFlowDiagram.squareAddSymbol(iCurrentSquareId, iSymbolId);

                            //ClsConfigReporterFlowDiagram.strSquare objSquare = new ClsConfigReporterFlowDiagram.strSquare();

                            //objSquare

                            //objSquare.lstSymbols.Add();

                            //int iSquareId = cConfigReporterFlowDiagram.addSquare();
                            //cConfigReporterFlowDiagram.addSquare(objSquare);

                            //cConfigReporterFlowDiagram.squareAddSymbol(iSquareId, iSymbolPreviousId);
                            /*
                            int iSquareParent;
                            
                            if (lstSquareId.Count == 0)
                            { iSquareParent= 1; }
                            else
                            { iSquareParent = lstSquareId[lstSquareId.Count - 1]; }

                            cConfigReporterFlowDiagram.squareAddSquare(iSquareParent, iSquareId);
                            */

                            ClsConfigReporterFlowDiagram.strLoops objLoop = new ClsConfigReporterFlowDiagram.strLoops();

                            if (objLine.sText_NoComment.Trim().ToLower().StartsWith("for ")
                                || objLine.sText_NoComment.Trim().ToLower().StartsWith("for each ") 
                                || objLine.sText_NoComment.Trim().ToLower().StartsWith("do while ")
                                || objLine.sText_NoComment.Trim().ToLower().StartsWith("while "))
                            { objLoop.eType = ClsConfigReporterFlowDiagram.enumLoopType.eCondition_Before; }
                            else
                            { objLoop.eType = ClsConfigReporterFlowDiagram.enumLoopType.eCondition_After; }
                            

                            objLoop.iDepthPx = 0;
                            objLoop.iId = 0;
                            objLoop.iMaxDepth = 0;
                            objLoop.iSymbolEnd = 0;
                            objLoop.iSymbolStart = objSymbol.id;

                            cConfigReporterFlowDiagram.addLoop(objLoop);

                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_Call))
                        {
                            objSymbol.eType = ClsConfigReporterFlowDiagram.enumSymbolType.eSym_PredefinedProcess;
                            objSymbol.iVertPos = 0;
                            objSymbol.iHorPos = 0;
                            objSymbol.lstSymbolsComingFrom.Add(objSymbolPrevious.id);
                            objSymbolPrevious.lstSymbolsGoingTo.Add(objSymbol.id);
                            objSymbol.sCaption = objLine.sText_NoComment.Trim();

                            //iSymbolPreviousId = cConfigReporterFlowDiagram.addSymbol(objSymbol);

                            iSymbolId = cConfigReporterFlowDiagram.addSymbol(objSymbol);
                            cConfigReporterFlowDiagram.squareAddSymbol(iCurrentSquareId, iSymbolId);
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_Comment))
                        {
                            /*
                             Later write code to optionally display comments as notes.
                             */
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ContinuedFromAbove))
                        {

                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_DeInitialise))
                        {

                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_Dim))
                        {
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_DllFunctionDeclare))
                        {
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_Else))
                        {
                            objSymbol.iHorPos = 0;
                            objSymbol.iVertPos = 0;

                            ClsConfigReporterFlowDiagram.strIf objIf = lstIfTemp[lstIfTemp.Count - 1];
                            if (lstIfTemp.Count > 0)
                            {
                                objIf = lstIfTemp[lstIfTemp.Count - 1];
                                if (objIf.lstSymbolElseIf.Count == 0)
                                { objSymbol.lstSymbolsComingFrom.Add(objIf.iSymbolIF); }
                                else
                                { objSymbol.lstSymbolsComingFrom.Add(objIf.lstSymbolElseIf[objIf.lstSymbolElseIf.Count - 1]); }
                            }
                            objSymbol.sCaption = objLine.sText_NoComment.Trim();
                            //objSymbolPrevious.lstSymbolsGoingTo.Add(objSymbol.id);
                            iSymbolId = cConfigReporterFlowDiagram.addSymbol(objSymbol);

                            /*Come out of the current square and opena new square*/
                            if (lstSquareId.Count >= 2)
                            {
                                lstSquareId.RemoveAt(lstSquareId.Count - 1);
                                int iParentSquareId = lstSquareId[lstSquareId.Count - 1];
                                int iNewSquareId = cConfigReporterFlowDiagram.addSquare();
                                lstSquareId.Add(iNewSquareId);

                                cConfigReporterFlowDiagram.squareAddSquare(iParentSquareId, iNewSquareId);
                                cConfigReporterFlowDiagram.squareAddSymbol(iNewSquareId, iSymbolId);
                            }
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ElseIF))
                        {
                            objSymbol.iHorPos = 0;
                            objSymbol.iVertPos = 0;

                            ClsConfigReporterFlowDiagram.strIf objIf = lstIfTemp[lstIfTemp.Count - 1];
                            if (lstIfTemp.Count > 0)
                            {
                                objIf = lstIfTemp[lstIfTemp.Count - 1];
                                if (objIf.lstSymbolElseIf.Count == 0)
                                { objSymbol.lstSymbolsComingFrom.Add(objIf.iSymbolIF); }
                                else
                                { objSymbol.lstSymbolsComingFrom.Add(objIf.lstSymbolElseIf[objIf.lstSymbolElseIf.Count - 1]); }
                            }
                            objSymbol.sCaption = objLine.sText_NoComment.Trim();
                            //objSymbolPrevious.lstSymbolsGoingTo.Add(objSymbol.id);
                            iSymbolId = cConfigReporterFlowDiagram.addSymbol(objSymbol);

                            if (lstIfTemp.Count > 0)
                            {
                                objIf = lstIfTemp[lstIfTemp.Count - 1];
                                objIf.lstSymbolElseIf.Add(iSymbolId);
                                lstIfTemp[lstIfTemp.Count - 1] = objIf;
                            }
                            else 
                            {
                                objIf = new ClsConfigReporterFlowDiagram.strIf();
                                objIf.lstSymbolElseIf.Add(iSymbolId);
                                lstIfTemp[lstIfTemp.Count - 1] = objIf;
                            }

                            /*Come out of the current square and opena new square*/
                            if (lstSquareId.Count >= 2)
                            {
                                lstSquareId.RemoveAt(lstSquareId.Count - 1);
                                int iParentSquareId = lstSquareId[lstSquareId.Count - 1];
                                int iNewSquareId = cConfigReporterFlowDiagram.addSquare();
                                lstSquareId.Add(iNewSquareId);

                                cConfigReporterFlowDiagram.squareAddSquare(iParentSquareId, iNewSquareId);
                                cConfigReporterFlowDiagram.squareAddSymbol(iNewSquareId, iSymbolId);
                            }
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_Empty))
                        {
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndFunction))
                        {
                            objSymbolPrevious = objSymbol;
                            objSymbol.id = 0;
                            objSymbol.eType = ClsConfigReporterFlowDiagram.enumSymbolType.eSym_Terminator;
                            objSymbol.sCaption = "End " + lstLines[0].sFunctionName;
                            objSymbol.iHorPos = 0;
                            objSymbol.iVertPos = 0;

                            //iSymbolPreviousId = cConfigReporterFlowDiagram.addSymbol(objSymbol);
                            iSymbolId = cConfigReporterFlowDiagram.addSymbol(objSymbol);
                            cConfigReporterFlowDiagram.squareAddSymbol(iCurrentSquareId, iSymbolId);
                            if (lstSquareId.Count > 0)
                            { lstSquareId.RemoveAt(lstSquareId.Count - 1); }
                            if (lstSquareId.Count > 0)
                            { iCurrentSquareId = lstSquareId[lstSquareId.Count - 1]; }
                            else
                            { iCurrentSquareId = 0; }
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndIf))
                        {
                            objSymbolPrevious = objSymbol;
                            objSymbol.id = 0;
                            objSymbol.eType = ClsConfigReporterFlowDiagram.enumSymbolType.eSym_ConnectingFlows;
                            objSymbol.sCaption = "";
                            objSymbol.iHorPos = 0;
                            objSymbol.iVertPos = 0;

                            iSymbolId = cConfigReporterFlowDiagram.addSymbol(objSymbol);
                            cConfigReporterFlowDiagram.squareAddSymbol(iCurrentSquareId, iSymbolId);

                            if (lstSquareId.Count > 0)
                            { lstSquareId.RemoveAt(lstSquareId.Count - 1); }
                            if (lstSquareId.Count > 0)
                            { iCurrentSquareId = lstSquareId[lstSquareId.Count - 1]; }
                            else
                            { iCurrentSquareId = 0; }
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndLoop))
                        {
                            if (objLine.sText_NoComment.ToLower().Contains("=") || objLine.sText_NoComment.ToLower().Contains("until") || objLine.sText_NoComment.ToLower().Contains("while"))
                            {
                                objSymbolPrevious = objSymbol;
                                objSymbol.id = 0;
                                objSymbol.eType = ClsConfigReporterFlowDiagram.enumSymbolType.eSym_Decision;
                                objSymbol.sCaption = "";
                                objSymbol.iHorPos = 0;
                                objSymbol.iVertPos = 0;

                            }
                            else
                            {
                                objSymbolPrevious = objSymbol;
                                objSymbol.id = 0;
                                objSymbol.eType = ClsConfigReporterFlowDiagram.enumSymbolType.eSym_ConnectingFlows;
                                objSymbol.sCaption = "";
                                objSymbol.iHorPos = 0;
                                objSymbol.iVertPos = 0;
                            }

                            ClsConfigReporterFlowDiagram.strLoops objLoopTemp = new ClsConfigReporterFlowDiagram.strLoops();
                            if (objLine.sText_NoComment.ToLower().Trim().StartsWith("next"))
                            { objLoopTemp = removeFromLoopList(ref lstLoopsTemp_Other, objSymbol.id); }
                            else if (objLine.sText_NoComment.ToLower().Trim().StartsWith("loop"))
                            { objLoopTemp = removeFromLoopList(ref lstLoopsTemp_DoLoop, objSymbol.id); }
                            else if (objLine.sText_NoComment.ToLower().Trim().StartsWith("wend"))
                            { objLoopTemp = removeFromLoopList(ref lstLoopsTemp_WhileWend, objSymbol.id); }

                            //cConfigReporterFlowDiagram.addLoop(objLoopTemp);
                            iSymbolId = cConfigReporterFlowDiagram.addSymbol(objSymbol);
                            cConfigReporterFlowDiagram.squareAddSymbol(iCurrentSquareId, iSymbolId);

                            cConfigReporterFlowDiagram.loopEndIf(objLoopTemp.iSymbolStart, iSymbolId);

                            if (lstSquareId.Count > 0)
                            { lstSquareId.RemoveAt(lstSquareId.Count - 1); }
                            if (lstSquareId.Count > 0)
                            { iCurrentSquareId = lstSquareId[lstSquareId.Count - 1]; }
                            else
                            { iCurrentSquareId = 0; }
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_EndWith))
                        {
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ErrorHandler))
                        {
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ExitFn))
                        {
                            objSymbol.sCaption = objLine.sText_NoComment.Trim();
                            objSymbol.eType = ClsConfigReporterFlowDiagram.enumSymbolType.eSym_Terminator;
                            objSymbol.iHorPos = 0;
                            objSymbol.iVertPos = 0;

                            iSymbolId = cConfigReporterFlowDiagram.addSymbol(objSymbol);
                            cConfigReporterFlowDiagram.squareAddSymbol(iCurrentSquareId, iSymbolId);
                            //iSymbolPreviousId = cConfigReporterFlowDiagram.addSymbol(objSymbol);
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ExitIf))
                        {
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_ExitLoop))
                        {
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_FunctionName))
                        {
                            objSymbolPrevious = objSymbol;
                            objSymbol.eType = ClsConfigReporterFlowDiagram.enumSymbolType.eSym_Terminator;
                            objSymbol.id = 0;
                            objSymbol.sCaption = "Start " + lstLines[0].sFunctionName;
                            objSymbol.iHorPos = 0;
                            objSymbol.iVertPos = 0;

                            iSymbolId = cConfigReporterFlowDiagram.addSymbol(objSymbol);
                            cConfigReporterFlowDiagram.squareAddSymbol(iCurrentSquareId, iSymbolId);
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_Goto))
                        {
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_If))
                        {
                            objSymbolPrevious = objSymbol;
                            objSymbol.id = 0;
                            objSymbol.eType = ClsConfigReporterFlowDiagram.enumSymbolType.eSym_Decision;
                            objSymbol.sCaption = objLine.sText_NoComment;
                            objSymbol.iHorPos = 0;
                            objSymbol.iVertPos = 0;

                            int iNewSquareId = cConfigReporterFlowDiagram.addSquare();
                            cConfigReporterFlowDiagram.squareAddSquare(iCurrentSquareId, iNewSquareId);
                            iCurrentSquareId = iNewSquareId;

                            iSymbolId = cConfigReporterFlowDiagram.addSymbol(objSymbol);
                            cConfigReporterFlowDiagram.squareAddSymbol(iCurrentSquareId, iSymbolId);
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_Initialise))
                        {
                            objSymbolPrevious = objSymbol;
                            objSymbol.id = 0;
                            objSymbol.sCaption = "Start " + lstLines[0].sFunctionName;
                            objSymbol.iHorPos = 0;

                            int iNewSquareId = cConfigReporterFlowDiagram.addSquare();
                            cConfigReporterFlowDiagram.squareAddSquare(iCurrentSquareId, iNewSquareId);
                            iCurrentSquareId = iNewSquareId;

                            iSymbolId = cConfigReporterFlowDiagram.addSymbol(objSymbol);
                            cConfigReporterFlowDiagram.squareAddSymbol(iCurrentSquareId, iSymbolId);
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_OnError))
                        {
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_Options))
                        {
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_Unknown))
                        {
                        }

                        if (objLine.lstLineType.Contains(ClsCodeMapper.enumLineType.eLineType_With))
                        {
                        }

                        //objSymbolPrevious = objSymbol;
                        //lstSymbols.Add(objSymbol);
                    }
                }
            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }

        //private void GenerateHtml(ref List<string> lstHtml)
        //{
        //    try
        //    {
        //        //strShape objShape = new strShape();
        //        enumSize eSize = enumSize.eLarge;
        //        //int iHor = 100;
        //        //int iVert = 100;

        //        enumSymbolType eSymbolType = enumSymbolType.eSym_Terminator;
        //        //objShape.

        //        string sLine = "";
        //        sLine = "<!DCOTYPE html>\n";
        //        sLine += "<html>\n";
        //        sLine += "<head>\n";
        //        sLine += "<meta http-equiv='X-UA-Compatible' content='IE=9' >\n";
        //        sLine += "</head>\n";
        //        sLine += "<body>\n";
        //        sLine += "<canvas id='myCanvas' width=" + iCanvasWidth.ToString() + " height=" + iCanvasHeight.ToString() + " style='border:1px solid #d3d3d3;'>Your browser does not support HTML5 canvas.</canvas>\n";
        //        sLine += "<script type='text/javascript'>\n";
        //        sLine += "var c=document.getElementById('myCanvas');\n";
        //        sLine += "var ctx=c.getContext('2d');\n";

        //        lstHtml.Add(sLine);
                
        //        sLine = "ctx.beginPath();\n";
        //        lstHtml.Add(sLine);


        //        foreach (strSymbol objSymbol in lstSymbols)
        //        {
        //            //drawShape(objSymbol.eType, enumSize.eLarge, objSymbol.iHorPos, objSymbol.iVertPos, objSymbol.sCaption, ref lstHtml);
        //            drawShape(objSymbol, ref lstHtml);
        //        }

        //        sLine = "ctx.stroke();\n";
        //        sLine += "</script>\n";
        //        sLine += "</body>\n";
        //        sLine += "</html>\n";
        //        lstHtml.Add(sLine);

        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}
        /*
        private void GenerateHtmlTest(ref List<string> lstHtml)
        {
            try
            {
                strShape objShape = new strShape();
                enumSize eSize = enumSize.eLarge;
                int iHor = 100;
                int iVert = 100;

                enumSymbolType eSymbolType = enumSymbolType.eSym_Terminator;
                //objShape.

                string sLine = "";
                sLine = "<!DCOTYPE html>\n";
                sLine += "<html>\n";
                sLine += "<head>\n";
                sLine += "<meta http-equiv='X-UA-Compatible' content='IE=9' >\n";
                sLine += "</head>\n";
                sLine += "<body>\n";
                sLine += "<canvas id='myCanvas' width=1000 height=2000 style='border:1px solid #d3d3d3;'>Your browser does not support HTML5 canvas.</canvas>\n";
                sLine += "<script type='text/javascript'>\n";
                sLine += "var c=document.getElementById('myCanvas');\n";
                sLine += "var ctx=c.getContext('2d');\n";

                lstHtml.Add(sLine);

                sLine = "ctx.beginPath();\n";
                lstHtml.Add(sLine);

                drawShape(enumSymbolType.eSym_Terminator, eSize, iHor, iVert, ref lstHtml);

                iVert = iVert + 100;

                drawShape(enumSymbolType.eSym_MagneticDisk, eSize, iHor, iVert, ref lstHtml);

                iVert = iVert + 100;

                drawShape(enumSymbolType.eSym_Document, eSize, iHor, iVert, ref lstHtml);

                iVert = iVert + 100;

                drawShape(enumSymbolType.eSym_Decision, eSize, iHor, iVert, ref lstHtml);

                iVert = iVert + 100;

                drawShape(enumSymbolType.eSym_VDU, eSize, iHor, iVert, ref lstHtml);

                iVert = iVert + 100;

                drawShape(enumSymbolType.eSym_Data, eSize, iHor, iVert, ref lstHtml);

                iVert = iVert + 100;

                drawShape(enumSymbolType.eSym_PredefinedProcess, eSize, iHor, iVert, ref lstHtml);

                iVert = iVert + 100;

                drawShape(enumSymbolType.eSym_ManualInput, eSize, iHor, iVert, ref lstHtml);

                iVert = iVert + 100;

                drawShape(enumSymbolType.eSym_AlternativeProcess, eSize, iHor, iVert, ref lstHtml);

                iVert = iVert + 100;

                drawShape(enumSymbolType.eSym_MultiDocument, eSize, iHor, iVert, ref lstHtml);

                iVert = iVert + 100;

                drawShape(enumSymbolType.eSym_InternalStorage, eSize, iHor, iVert, ref lstHtml);

                iVert = iVert + 100;

                drawShape(enumSymbolType.eSym_Sort, eSize, iHor, iVert, ref lstHtml);

                sLine = "ctx.stroke();\n";
                sLine += "</script>\n";
                sLine += "</body>\n";
                sLine += "</html>\n";
                lstHtml.Add(sLine);

            }
            catch (Exception ex)
            {
                MethodBase mbTemp = MethodBase.GetCurrentMethod();

                string sMessage = "";

                sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
                sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
                sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
                sMessage += ex.Message;

                MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
            }
        }
        */
        //private void reposition()
        //{
        //    try
        //    {
        //        int iMaxHor = lstSymbols.Max(x => x.iHorPos);
        //        int iMinHor = lstSymbols.Min(x => x.iHorPos);

        //        int iMaxVert = lstSymbols.Max(x => x.iVertPos);
        //        int iMinVert = lstSymbols.Min(x => x.iVertPos);

        //        for (int iIndex = 0; iIndex < lstSymbols.Count; iIndex++)
        //        {
        //            ClsConfigReporterFlowDiagram.strSymbol objSymbol = lstSymbols[iIndex];
        //            objSymbol.iHorPos += ciHorSpacing - iMinHor;
        //            objSymbol.iVertPos += ciVertSpacing - iMinVert;
        //            lstSymbols[iIndex] = objSymbol;
        //        }

        //        iCanvasWidth = iMaxHor - iMinHor + 2 * ciHorSpacing;
        //        iCanvasHeight = iMaxVert - iMinVert + 2 * ciVertSpacing;
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        //private void CalculateLoopDepths()
        //{
        //    try
        //    {
        //        foreach (ClsConfigReporterFlowDiagram.strLoops objLoop in lstLoops.OrderBy(x => x.iMaxDepth))
        //        {
        //            int iMaxLeft = getSymbolsInLoop(objLoop).Min(x => x.iHorPos);
                
                
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);
        //    }
        //}

        //private List<ClsConfigReporterFlowDiagram.strSymbol> getSymbolsInLoop(ClsConfigReporterFlowDiagram.strLoops objLoop)
        //{
        //    try
        //    {
        //        List<ClsConfigReporterFlowDiagram.strSymbol> lstResult = new List<ClsConfigReporterFlowDiagram.strSymbol>();
        //        bool bDone = false;
        //        List<int> lstIdToAdd = new List<int>();
        //        int iPerviousCount = 0;

        //        lstIdToAdd.Add(objLoop.iSymbolStart);

        //        while (!bDone)
        //        {
        //            foreach (ClsConfigReporterFlowDiagram.strSymbol objSymbol in lstSymbols.FindAll(x => lstIdToAdd.Contains(x.id) && x.id != objLoop.iSymbolEnd))
        //            { lstResult.Add(objSymbol); }

        //            lstResult = lstResult.Distinct().ToList<ClsConfigReporterFlowDiagram.strSymbol>();

        //            foreach (ClsConfigReporterFlowDiagram.strSymbol objSymbol in lstResult)
        //            {
        //                foreach (int iGoingTo in objSymbol.lstSymbolsGoingTo)
        //                { lstIdToAdd.Add(iGoingTo); }
        //            }

        //            lstIdToAdd = lstIdToAdd.Distinct().ToList<int>();
                      
        //            if (iPerviousCount == lstIdToAdd.Count)
        //            { bDone = true; }
        //        }
                
        //        return lstResult;
        //    }
        //    catch (Exception ex)
        //    {
        //        MethodBase mbTemp = MethodBase.GetCurrentMethod();

        //        string sMessage = "";

        //        sMessage += "Add-in: " + mbTemp.ReflectedType.Name + "\n\r";
        //        sMessage += "Module Name: " + mbTemp.Module.Name + "\n\r";
        //        sMessage += "Function Name: " + mbTemp.Name + "\n\r\n\r";
        //        sMessage += ex.Message;

        //        MessageBox.Show(text: sMessage, caption: "Error", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Error);

        //        return new List<ClsConfigReporterFlowDiagram.strSymbol>();
        //    }
        //}

    }
}
