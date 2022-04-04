using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;

namespace KodeMagd.Reporter
{
    class ClsConfigReporterFlowDiagram : ClsConfigReporter
    {
        private const int ciHorSpacing = 80;
        private const int ciVertSpacing = 100;

        public enum enumSymbolType
        {
            //eSym_Input_FromSheet, /*rectangle with S-bottom line*/
            //eSym_Input_FromUser,/*Rectangle with top sloped at a angle*/
            //eSym_Input_FromOther,/**/
            //eSym_Output_ToSheet,/*rectangle with S-bottom line*/
            //eSym_Output_ToOther,/**/
            //eSym_SubRoutineFunction, /*Rectangle with double vertical lines*/
            //eSym_BeginningOfRepeatativeStructure, /*hexagon*/

            eSym_VDU,/*curved shape looks like a CRT*/
            eSym_Data, /*parallelagram*/
            eSym_Decision, /*diamond*/
            eSym_Terminator, /*Round left and right flat top and bottom*/
            eSym_MagneticDisk,/*cylinder*/
            eSym_Process, /*rectangle*/
            eSym_Document,/*rectangle with S-bottom line*/
            eSym_PredefinedProcess,/*rectangle with double vertical sides*/
            eSym_ManualInput,
            eSym_AlternativeProcess,
            eSym_MultiDocument,
            eSym_InternalStorage,
            eSym_Sort,
            eSym_ConnectingFlows
        }

        public enum enumSize
        {
            eSmall,
            eMedium,
            eLarge,
        }

        public enum enumLoopType
        {
            eCondition_Before,
            eCondition_After
        }

        public enum enumConnectionNodePosition
        {
            eTop,
            eBottom,
            eLeft,
            eRight
        }

        public struct strConnectionNode
        {
            public enumConnectionNodePosition ePos;
            public int iHor;/*relative to shape*/
            public int iVert;/*relative to shape*/
        }

        public struct strSymbol
        {
            public int id;
            public List<int> lstSymbolsGoingTo;
            public List<int> lstSymbolsComingFrom;
            public enumSymbolType eType;
            public string sVbaCode;
            public string sCaption;
            public int iVertPos;
            public int iHorPos;
            public enumSize eSize;
        }

        public struct strLoops
        {
            public int iId;
            public int iSymbolStart;
            public int iSymbolEnd;
            public enumLoopType eType;
            public int iMaxDepth; //related to the order of the nexted loops
            public int iDepthPx; //related to the pixels required to draw around any symbols or other loops
        }

        public struct strIf
        {
            public int iSymbolIF;
            public List<int> lstSymbolElseIf;
            public int iSymbolElse;
            public int iSymbolEnd;
            public int iMaxDepth;
        }

        public enum enumLineType
        {
            eStraight,
            eCurved,
            eCircle,
        }

        public struct strCoordinates
        {
            public int iHor;
            public int iVert;
        }

        public struct strStraightLine
        {
            public int order;
            public bool freshStart;
            public strCoordinates start;
            public strCoordinates end;
        }

        public struct strCurvedLine
        {
            public int order;
            public bool freshStart;
            public strCoordinates CentreOfRotation;
            public int radiusHor;
            public int radiusVert;
            public double startAngle;
            public double endAngle;
        }

        public struct strShape
        {
            public List<strStraightLine> lstStraightLine;
            public List<strCurvedLine> lstCurvedLine;
            public List<strConnectionNode> lstConnectionNodes;
            public enumSymbolType eSymbol;
            public enumSize eSize;
        }

        public struct strSquare
        {
            public int iId;
            public int iLeft;
            public int iTop;
            public int iHeight;
            public int iWidth;
            public int iStartSymbol;
            public int iEndSymbol;
            public List<int> lstSquares;
            public List<int> lstSymbols;
            public List<int> lstLoops;
        }

        int iCanvasWidth;
        int iCanvasHeight;
        int iRootOfAllSquaresId;

        List<strShape> lstShapes;
        List<ClsConfigReporterFlowDiagram.strSymbol> lstSymbols = new List<ClsConfigReporterFlowDiagram.strSymbol>();
        List<ClsConfigReporterFlowDiagram.strLoops> lstLoops = new List<ClsConfigReporterFlowDiagram.strLoops>();
        List<ClsConfigReporterFlowDiagram.strSquare> lstSquares = new List<ClsConfigReporterFlowDiagram.strSquare>();

        public int rootOfAllSquaresId
        {
            get
            {
                try
                {
                    return iRootOfAllSquaresId;
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

                    return 0;
                }
            }
            set
            {
                try
                {
                    iRootOfAllSquaresId = value; 
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

        public int canvasWidth
        {
            get
            {
                try
                {
                    return iCanvasWidth;
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

                    return 0;
                }
            }
        }

        public int canvasHeight
        {
            get
            {
                try
                {
                    return iCanvasHeight;
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

                    return 0;
                }
            }
        }

        public int squareStartSymbol(int iSquareID)
        {
            try
            {
                int iIndex = lstSquares.FindIndex(x => x.iId == iSquareID);
                int iResult;

                if (iIndex > 0)
                { iResult = lstSquares[iIndex].iStartSymbol; }
                else
                { iResult = -1; }

                return iResult;
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

                return -1;
            }
        }

        public ClsConfigReporterFlowDiagram()
        {
            try
            {
                fillShapes();
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

        public void GenerateHtml(ref List<string> lstHtml)
        {
            try
            {
                //strShape objShape = new strShape();
                enumSize eSize = enumSize.eLarge;
                //int iHor = 100;
                //int iVert = 100;

                enumSymbolType eSymbolType = enumSymbolType.eSym_Terminator;
                //objShape.

                string sLine = "";
                sLine = "<!DCOTYPE html>\n";
                sLine += "<html>\n";
                sLine += "<head>\n";
                sLine += "<meta http-equiv='X-UA-Compatible' content='IE=9' >\n";
                sLine += "</head>\n";
                sLine += "<body>\n";
                sLine += "<canvas id='myCanvas' width=" + iCanvasWidth.ToString() + " height=" + iCanvasHeight.ToString() + " style='border:1px solid #d3d3d3;'>Your browser does not support HTML5 canvas.</canvas>\n";
                sLine += "<script type='text/javascript'>\n";
                sLine += "var c=document.getElementById('myCanvas');\n";
                sLine += "var ctx=c.getContext('2d');\n";

                lstHtml.Add(sLine);

                sLine = "ctx.beginPath();\n";
                lstHtml.Add(sLine);


                foreach (strSymbol objSymbol in lstSymbols)
                {
                    //drawShape(objSymbol.eType, enumSize.eLarge, objSymbol.iHorPos, objSymbol.iVertPos, objSymbol.sCaption, ref lstHtml);
                    drawShape(objSymbol, ref lstHtml);
                }

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
        /*
        public void groupIntoSquares()
        {
            try
            {
                int iMaxLoopDepth = lstLoops.Max(x => x.iMaxDepth);
                for (int iDepth = 0; iDepth <= iMaxLoopDepth; iDepth++)
                {
                    foreach(strLoops objLoop in lstLoops.FindAll(x => x.iMaxDepth == iDepth))
                    {
                        //foreach ()
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
        */
        private void CalculateLoopDepths()
        {
            try
            {
                foreach (ClsConfigReporterFlowDiagram.strLoops objLoop in lstLoops.OrderBy(x => x.iMaxDepth))
                {
                    int iMaxLeft = getSymbolsInLoop(objLoop).Min(x => x.iHorPos);


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

        private void drawShape(strSymbol objSymbol, ref List<string> lstHtml)
        {
            try
            {
                if (lstShapes.Exists(x => x.eSymbol == objSymbol.eType && x.eSize == objSymbol.eSize))
                {
                    string sLine = "";
                    strShape objShape = lstShapes.Find(x => x.eSymbol == objSymbol.eType && x.eSize == objSymbol.eSize);

                    int iOrderMin = 0;
                    int iOrderMax = 0;

                    if (objShape.lstCurvedLine.Count() > 0)
                    {
                        iOrderMin = objShape.lstCurvedLine.Min(x => x.order);
                        iOrderMax = objShape.lstCurvedLine.Max(x => x.order);
                    }

                    if (objShape.lstStraightLine.Count() > 0)
                    {
                        int iOrderStriaghtMin = objShape.lstStraightLine.Min(x => x.order);
                        int iOrderStriaghtMax = objShape.lstStraightLine.Max(x => x.order);

                        if (iOrderMin > iOrderStriaghtMin)
                        { iOrderMin = iOrderStriaghtMin; }

                        if (iOrderMax < iOrderStriaghtMax)
                        { iOrderMax = iOrderStriaghtMax; }
                    }

                    bool bIsContinuing = false;
                    int iPreviousHor = 0;
                    int iPreviousVert = 0;

                    for (int iOrderCounter = iOrderMin; iOrderCounter <= iOrderMax; iOrderCounter++)
                    {

                        sLine = "";
                        foreach (strCurvedLine objCurvedLine in objShape.lstCurvedLine.FindAll(x => x.order == iOrderCounter))
                        {
                            sLine = "";
                            double dHorScale = 1;
                            double dVertScale = 1;
                            double dHorUnScale = 1;
                            double dVertUnScale = 1;

                            /* Scale */
                            if (objCurvedLine.radiusHor != objCurvedLine.radiusVert)
                            {
                                dHorScale = 0;
                                dVertScale = 0;

                                if (objCurvedLine.radiusHor > objCurvedLine.radiusVert)
                                {
                                    dHorScale = 1.0;
                                    dVertScale = (double)objCurvedLine.radiusVert / (double)objCurvedLine.radiusHor;
                                    dHorUnScale = 1.0;
                                    dVertUnScale = (double)objCurvedLine.radiusHor / (double)objCurvedLine.radiusVert;
                                }
                                else
                                {
                                    dHorScale = (double)objCurvedLine.radiusHor / (double)objCurvedLine.radiusVert;
                                    dVertScale = 1.0;
                                    dHorUnScale = (double)objCurvedLine.radiusVert / (double)objCurvedLine.radiusHor;
                                    dVertUnScale = 1.0;
                                }

                                sLine += "ctx.scale(" + dHorScale.ToString() + ", " + dVertScale.ToString() + ");\n";
                            }

                            if (objCurvedLine.freshStart)
                            {
                                int iStartHor = (int)((dHorUnScale * ((double)objCurvedLine.CentreOfRotation.iHor + (double)objSymbol.iHorPos)) + (double)(Math.Cos(objCurvedLine.startAngle) * (double)(objCurvedLine.radiusHor + objCurvedLine.radiusVert) / 2.0));
                                int iStartVert = (int)((dVertUnScale * ((double)objCurvedLine.CentreOfRotation.iVert + (double)objSymbol.iVertPos)) + (double)(Math.Sin(objCurvedLine.startAngle) * (double)(objCurvedLine.radiusHor + objCurvedLine.radiusVert) / 2.0));

                                sLine += "ctx.moveTo(" + iStartHor.ToString() + ", " + iStartVert.ToString() + ");\n";
                            }

                            /*arc*/
                            sLine += "ctx.arc(" + (dHorUnScale * (objCurvedLine.CentreOfRotation.iHor + objSymbol.iHorPos)).ToString() + ", " + (dVertUnScale * (objCurvedLine.CentreOfRotation.iVert + objSymbol.iVertPos)).ToString() + ",";
                            sLine += ((objCurvedLine.radiusHor + objCurvedLine.radiusVert) / 2).ToString() + ",";
                            sLine += objCurvedLine.startAngle.ToString() + "," + objCurvedLine.endAngle.ToString();

                            if (objCurvedLine.startAngle < objCurvedLine.endAngle)
                            { sLine += ");\n"; }
                            else
                            { sLine += ", true);\n"; }

                            /* Un-scale */
                            if (objCurvedLine.radiusHor != objCurvedLine.radiusVert)
                            {
                                sLine += "ctx.scale(" + dHorUnScale.ToString() + ", " + dVertUnScale.ToString() + ");\n";
                            }

                            lstHtml.Add(sLine);
                        }

                        sLine = "";
                        foreach (strStraightLine objStraightLine in objShape.lstStraightLine.FindAll(x => x.order == iOrderCounter))
                        {
                            sLine = "";
                            if (objStraightLine.freshStart)
                            { sLine += "ctx.moveTo(" + (objStraightLine.start.iHor + objSymbol.iHorPos).ToString() + ", " + (objStraightLine.start.iVert + objSymbol.iVertPos).ToString() + ");\n"; }
                            sLine += "ctx.lineTo(" + (objStraightLine.end.iHor + objSymbol.iHorPos).ToString() + ", " + (objStraightLine.end.iVert + objSymbol.iVertPos).ToString() + ");\n";

                            lstHtml.Add(sLine);
                        }
                    }
                    sLine = "ctx.fillText('" + objSymbol.sCaption + "', " + objSymbol.iHorPos + ", " + objSymbol.iVertPos + ");";
                    lstHtml.Add(sLine);
                    sLine = "ctx.stroke();\n";
                    lstHtml.Add(sLine);
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

        private void fillShapes()
        {
            try
            {
                lstShapes = new List<strShape>();

                fillShapesTerminator();
                fillShapesMagneticDisk();
                fillShapesDocument();
                fillShapesDecision();
                fillShapesVDU();
                fillShapesData();
                fillShapesProcess();
                fillShapesPredefinedProcess();
                fillShapesManualInput();
                fillShapesAlternativeProcess();
                fillShapesMultiDocument();
                fillShapesInternalStorage();
                fillShapesSort();
                fillShapesConnectingFlows();

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


        private void fillShapesTerminator()
        {
            try
            {
                strShape objShape = new strShape();
                strStraightLine objStraightLine = new strStraightLine();
                strCurvedLine objCurvedLine = new strCurvedLine();
                strConnectionNode objConnectionNode = new strConnectionNode();
                int iOrder = 0;

                objShape.lstCurvedLine = new List<strCurvedLine>();
                objShape.lstStraightLine = new List<strStraightLine>();

                /*Start*/
                objShape.eSymbol = enumSymbolType.eSym_Terminator;
                objShape.eSize = enumSize.eLarge;
                objShape.lstConnectionNodes = new List<strConnectionNode>();

                objCurvedLine = new strCurvedLine();
                iOrder++;
                objCurvedLine.order = iOrder;
                objCurvedLine.freshStart = true;
                objCurvedLine.CentreOfRotation.iHor = 40;
                objCurvedLine.CentreOfRotation.iVert = 0;
                objCurvedLine.startAngle = -0.5 * Math.PI;
                objCurvedLine.endAngle = 0.5 * Math.PI;
                objCurvedLine.radiusHor = 15;
                objCurvedLine.radiusVert = 15;
                objShape.lstCurvedLine.Add(objCurvedLine);

                objCurvedLine = new strCurvedLine();
                iOrder++;
                objCurvedLine.order = iOrder;
                objCurvedLine.freshStart = false;
                objCurvedLine.CentreOfRotation.iHor = -40;
                objCurvedLine.CentreOfRotation.iVert = 0;
                objCurvedLine.startAngle = 0.5 * Math.PI;
                objCurvedLine.endAngle = 1.5 * Math.PI;
                objCurvedLine.radiusHor = 15;
                objCurvedLine.radiusVert = 15;
                objShape.lstCurvedLine.Add(objCurvedLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = -40;
                objStraightLine.start.iVert = -15;
                objStraightLine.end.iHor = 40;
                objStraightLine.end.iVert = -15;
                objShape.lstStraightLine.Add(objStraightLine);


                objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = 15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eTop;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = -15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                lstShapes.Add(objShape);
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

        private void fillShapesMagneticDisk()
        {
            try
            {
                strShape objShape = new strShape();
                strStraightLine objStraightLine = new strStraightLine();
                strCurvedLine objCurvedLine = new strCurvedLine();
                strConnectionNode objConnectionNode = new strConnectionNode();
                int iOrder = 0;

                objShape.lstCurvedLine = new List<strCurvedLine>();
                objShape.lstStraightLine = new List<strStraightLine>();
                objShape.lstConnectionNodes = new List<strConnectionNode>();

                /*Start*/
                objShape.eSymbol = enumSymbolType.eSym_MagneticDisk;
                objShape.eSize = enumSize.eLarge;

                objCurvedLine = new strCurvedLine();
                iOrder++;
                objCurvedLine.order = iOrder;
                objCurvedLine.freshStart = true;
                objCurvedLine.CentreOfRotation.iHor = 0;
                objCurvedLine.CentreOfRotation.iVert = -20;
                objCurvedLine.startAngle = 0;
                objCurvedLine.endAngle = 2 * Math.PI;
                objCurvedLine.radiusHor = 64;
                objCurvedLine.radiusVert = 32;
                objShape.lstCurvedLine.Add(objCurvedLine);

                objCurvedLine = new strCurvedLine();
                iOrder++;
                objCurvedLine.order = iOrder;
                objCurvedLine.freshStart = false;
                objCurvedLine.CentreOfRotation.iHor = 0;
                objCurvedLine.CentreOfRotation.iVert = 20;
                objCurvedLine.startAngle = 0;
                objCurvedLine.endAngle = Math.PI;
                objCurvedLine.radiusHor = 64;
                objCurvedLine.radiusVert = 32;
                objShape.lstCurvedLine.Add(objCurvedLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = -48;
                objStraightLine.start.iVert = -20;
                objStraightLine.end.iHor = -48;
                objStraightLine.end.iVert = -20;
                objShape.lstStraightLine.Add(objStraightLine);

                objConnectionNode.ePos = enumConnectionNodePosition.eTop;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = -15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = 15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eRight;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                lstShapes.Add(objShape);

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


        private void fillShapesConnectingFlows()
        {
            try
            {
                strShape objShape = new strShape();
                strStraightLine objStraightLine = new strStraightLine();
                strCurvedLine objCurvedLine = new strCurvedLine();
                strConnectionNode objConnectionNode = new strConnectionNode();
                int iOrder = 0;

                objShape.lstCurvedLine = new List<strCurvedLine>();
                objShape.lstStraightLine = new List<strStraightLine>();
                objShape.lstConnectionNodes = new List<strConnectionNode>();

                /*Start*/
                //objShape.eSymbol = enumSymbolType.eSym_Output_ToDatabase;
                objShape.eSize = enumSize.eLarge;

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = -15;
                objStraightLine.start.iVert = -20;
                objStraightLine.end.iHor = 15;
                objStraightLine.end.iVert = -20;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 15;
                objStraightLine.start.iVert = -20;
                objStraightLine.end.iHor = 0;
                objStraightLine.end.iVert = 20;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 0;
                objStraightLine.start.iVert = 20;
                objStraightLine.end.iHor = -15;
                objStraightLine.end.iVert = -20;
                objShape.lstStraightLine.Add(objStraightLine);


                objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = 20;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 20;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eRight;
                objConnectionNode.iHor = 15;
                objConnectionNode.iVert = 20;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objShape.eSymbol = enumSymbolType.eSym_ConnectingFlows;
                lstShapes.Add(objShape);
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



        private void fillShapesDocument()
        {
            try
            {
                strShape objShape = new strShape();
                strStraightLine objStraightLine = new strStraightLine();
                strCurvedLine objCurvedLine = new strCurvedLine();
                strConnectionNode objConnectionNode = new strConnectionNode();
                int iOrder = 0;

                objShape.lstCurvedLine = new List<strCurvedLine>();
                objShape.lstStraightLine = new List<strStraightLine>();
                objShape.lstConnectionNodes = new List<strConnectionNode>();

                /*Start*/
                //objShape.eSymbol = enumSymbolType.eSym_Output_ToDatabase;
                objShape.eSize = enumSize.eLarge;

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = -40;
                objStraightLine.start.iVert = 32;
                objStraightLine.end.iHor = -40;
                objStraightLine.end.iVert = -32;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = -40;
                objStraightLine.start.iVert = -32;
                objStraightLine.end.iHor = 40;
                objStraightLine.end.iVert = -32;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = 40;
                objStraightLine.start.iVert = -32;
                objStraightLine.end.iHor = 40;
                objStraightLine.end.iVert = 32;
                objShape.lstStraightLine.Add(objStraightLine);

                objCurvedLine = new strCurvedLine();
                iOrder++;
                objCurvedLine.order = iOrder;
                objCurvedLine.freshStart = false;
                objCurvedLine.CentreOfRotation.iHor = 20;
                objCurvedLine.CentreOfRotation.iVert = 32;
                objCurvedLine.startAngle = 0;
                objCurvedLine.endAngle = Math.PI;
                objCurvedLine.radiusHor = 22;
                objCurvedLine.radiusVert = 18;
                objShape.lstCurvedLine.Add(objCurvedLine);

                objCurvedLine = new strCurvedLine();
                iOrder++;
                objCurvedLine.order = iOrder;
                objCurvedLine.freshStart = false;
                objCurvedLine.CentreOfRotation.iHor = -20;
                objCurvedLine.CentreOfRotation.iVert = 32;
                objCurvedLine.startAngle = 0;
                objCurvedLine.endAngle = -1 * Math.PI;
                objCurvedLine.radiusHor = 22;
                objCurvedLine.radiusVert = 18;
                objShape.lstCurvedLine.Add(objCurvedLine);

                objConnectionNode.ePos = enumConnectionNodePosition.eTop;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = -15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = 15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eRight;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objShape.eSymbol = enumSymbolType.eSym_Document;
                lstShapes.Add(objShape);
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

        private void fillShapesMultiDocument()
        {
            try
            {
                strShape objShape = new strShape();
                strStraightLine objStraightLine = new strStraightLine();
                strCurvedLine objCurvedLine = new strCurvedLine();
                strConnectionNode objConnectionNode = new strConnectionNode();
                int iOrder = 0;

                objShape.lstCurvedLine = new List<strCurvedLine>();
                objShape.lstStraightLine = new List<strStraightLine>();
                objShape.lstConnectionNodes = new List<strConnectionNode>();

                /*Start*/
                objShape.eSymbol = enumSymbolType.eSym_MultiDocument;
                objShape.eSize = enumSize.eLarge;

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = -40;
                objStraightLine.start.iVert = 32;
                objStraightLine.end.iHor = -40;
                objStraightLine.end.iVert = -32;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = -40;
                objStraightLine.start.iVert = -32;
                objStraightLine.end.iHor = 40;
                objStraightLine.end.iVert = -32;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = 40;
                objStraightLine.start.iVert = -32;
                objStraightLine.end.iHor = 40;
                objStraightLine.end.iVert = 32;
                objShape.lstStraightLine.Add(objStraightLine);

                objCurvedLine = new strCurvedLine();
                iOrder++;
                objCurvedLine.order = iOrder;
                objCurvedLine.freshStart = false;
                objCurvedLine.CentreOfRotation.iHor = 20;
                objCurvedLine.CentreOfRotation.iVert = 32;
                objCurvedLine.startAngle = 0;
                objCurvedLine.endAngle = Math.PI;
                objCurvedLine.radiusHor = 22;
                objCurvedLine.radiusVert = 18;
                objShape.lstCurvedLine.Add(objCurvedLine);

                objCurvedLine = new strCurvedLine();
                iOrder++;
                objCurvedLine.order = iOrder;
                objCurvedLine.freshStart = false;
                objCurvedLine.CentreOfRotation.iHor = -20;
                objCurvedLine.CentreOfRotation.iVert = 32;
                objCurvedLine.startAngle = 0;
                objCurvedLine.endAngle = -1 * Math.PI;
                objCurvedLine.radiusHor = 22;
                objCurvedLine.radiusVert = 18;
                objShape.lstCurvedLine.Add(objCurvedLine);

                /*shadow 1*/
                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = 40;
                objStraightLine.start.iVert = 27;
                objStraightLine.end.iHor = 45;
                objStraightLine.end.iVert = 27;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 45;
                objStraightLine.start.iVert = 27;
                objStraightLine.end.iHor = 45;
                objStraightLine.end.iVert = -37;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 45;
                objStraightLine.start.iVert = -37;
                objStraightLine.end.iHor = -35;
                objStraightLine.end.iVert = -37;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = -35;
                objStraightLine.start.iVert = -37;
                objStraightLine.end.iHor = -35;
                objStraightLine.end.iVert = -32;
                objShape.lstStraightLine.Add(objStraightLine);

                /*shadow 2*/
                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = 45;
                objStraightLine.start.iVert = 22;
                objStraightLine.end.iHor = 50;
                objStraightLine.end.iVert = 22;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 50;
                objStraightLine.start.iVert = 22;
                objStraightLine.end.iHor = 50;
                objStraightLine.end.iVert = -42;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 50;
                objStraightLine.start.iVert = -42;
                objStraightLine.end.iHor = -30;
                objStraightLine.end.iVert = -42;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = -30;
                objStraightLine.start.iVert = -42;
                objStraightLine.end.iHor = -30;
                objStraightLine.end.iVert = -37;
                objShape.lstStraightLine.Add(objStraightLine);

                objConnectionNode.ePos = enumConnectionNodePosition.eTop;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = -15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = 15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eRight;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                lstShapes.Add(objShape);
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

        private void fillShapesDecision()
        {
            try
            {
                strShape objShape = new strShape();
                strStraightLine objStraightLine = new strStraightLine();
                strCurvedLine objCurvedLine = new strCurvedLine();
                strConnectionNode objConnectionNode = new strConnectionNode();
                int iOrder = 0;

                objShape.lstCurvedLine = new List<strCurvedLine>();
                objShape.lstStraightLine = new List<strStraightLine>();
                objShape.lstConnectionNodes = new List<strConnectionNode>();

                /*Start*/
                //objShape.eSymbol = enumSymbolType.eSym_Output_ToDatabase;
                objShape.eSize = enumSize.eLarge;

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = -40;
                objStraightLine.start.iVert = 0;
                objStraightLine.end.iHor = 0;
                objStraightLine.end.iVert = -40;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 0;
                objStraightLine.start.iVert = -40;
                objStraightLine.end.iHor = 40;
                objStraightLine.end.iVert = 0;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 40;
                objStraightLine.start.iVert = 0;
                objStraightLine.end.iHor = 0;
                objStraightLine.end.iVert = 40;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 0;
                objStraightLine.start.iVert = 40;
                objStraightLine.end.iHor = -40;
                objStraightLine.end.iVert = 0;
                objShape.lstStraightLine.Add(objStraightLine);

                objConnectionNode.ePos = enumConnectionNodePosition.eTop;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = -15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = 15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eRight;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objShape.eSymbol = enumSymbolType.eSym_Decision;
                lstShapes.Add(objShape);

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

        private void fillShapesSort()
        {
            try
            {
                strShape objShape = new strShape();
                strStraightLine objStraightLine = new strStraightLine();
                strCurvedLine objCurvedLine = new strCurvedLine();
                strConnectionNode objConnectionNode = new strConnectionNode();
                int iOrder = 0;

                objShape.lstCurvedLine = new List<strCurvedLine>();
                objShape.lstStraightLine = new List<strStraightLine>();
                objShape.lstConnectionNodes = new List<strConnectionNode>();

                /*Start*/
                objShape.eSymbol = enumSymbolType.eSym_Sort;
                objShape.eSize = enumSize.eLarge;

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = -40;
                objStraightLine.start.iVert = 0;
                objStraightLine.end.iHor = 0;
                objStraightLine.end.iVert = -40;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 0;
                objStraightLine.start.iVert = -40;
                objStraightLine.end.iHor = 40;
                objStraightLine.end.iVert = 0;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 40;
                objStraightLine.start.iVert = 0;
                objStraightLine.end.iHor = 0;
                objStraightLine.end.iVert = 40;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 0;
                objStraightLine.start.iVert = 40;
                objStraightLine.end.iHor = -40;
                objStraightLine.end.iVert = 0;
                objShape.lstStraightLine.Add(objStraightLine);
                /*
                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = 0;
                objStraightLine.start.iVert = 40;
                objStraightLine.end.iHor = 0;
                objStraightLine.end.iVert = -40;
                objShape.lstStraightLine.Add(objStraightLine);
                */
                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = 40;
                objStraightLine.start.iVert = 0;
                objStraightLine.end.iHor = -40;
                objStraightLine.end.iVert = 0;
                objShape.lstStraightLine.Add(objStraightLine);

                objConnectionNode.ePos = enumConnectionNodePosition.eTop;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = -40;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = 40;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
                objConnectionNode.iHor = -40;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eRight;
                objConnectionNode.iHor = -40;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                lstShapes.Add(objShape);

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

        private void fillShapesData()
        {
            try
            {
                strShape objShape = new strShape();
                strStraightLine objStraightLine = new strStraightLine();
                strCurvedLine objCurvedLine = new strCurvedLine();
                strConnectionNode objConnectionNode = new strConnectionNode();
                int iOrder = 0;

                objShape.lstCurvedLine = new List<strCurvedLine>();
                objShape.lstStraightLine = new List<strStraightLine>();
                objShape.lstConnectionNodes = new List<strConnectionNode>();

                /*Start*/
                objShape.eSymbol = enumSymbolType.eSym_Data;
                objShape.eSize = enumSize.eLarge;

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = -50;
                objStraightLine.start.iVert = 25;
                objStraightLine.end.iHor = 30;
                objStraightLine.end.iVert = 25;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 30;
                objStraightLine.start.iVert = 25;
                objStraightLine.end.iHor = 50;
                objStraightLine.end.iVert = -25;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 50;
                objStraightLine.start.iVert = -25;
                objStraightLine.end.iHor = -30;
                objStraightLine.end.iVert = -25;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = -30;
                objStraightLine.start.iVert = -25;
                objStraightLine.end.iHor = -50;
                objStraightLine.end.iVert = 25;
                objShape.lstStraightLine.Add(objStraightLine);

                objConnectionNode.ePos = enumConnectionNodePosition.eTop;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = -15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = 15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eRight;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                lstShapes.Add(objShape);
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

        private void fillShapesProcess()
        {
            try
            {
                strShape objShape = new strShape();
                strStraightLine objStraightLine = new strStraightLine();
                strCurvedLine objCurvedLine = new strCurvedLine();
                strConnectionNode objConnectionNode = new strConnectionNode();
                int iOrder = 0;

                objShape.lstCurvedLine = new List<strCurvedLine>();
                objShape.lstStraightLine = new List<strStraightLine>();
                objShape.lstConnectionNodes = new List<strConnectionNode>();

                /*Start*/
                objShape.eSymbol = enumSymbolType.eSym_Process;
                objShape.eSize = enumSize.eLarge;

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = -50;
                objStraightLine.start.iVert = 25;
                objStraightLine.end.iHor = 50;
                objStraightLine.end.iVert = 25;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 50;
                objStraightLine.start.iVert = 25;
                objStraightLine.end.iHor = 50;
                objStraightLine.end.iVert = -25;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 50;
                objStraightLine.start.iVert = -25;
                objStraightLine.end.iHor = -50;
                objStraightLine.end.iVert = -25;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = -50;
                objStraightLine.start.iVert = -25;
                objStraightLine.end.iHor = -50;
                objStraightLine.end.iVert = 25;
                objShape.lstStraightLine.Add(objStraightLine);

                objConnectionNode.ePos = enumConnectionNodePosition.eTop;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = -15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = 15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eRight;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                lstShapes.Add(objShape);
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


        private void fillShapesInternalStorage()
        {
            try
            {
                strShape objShape = new strShape();
                strStraightLine objStraightLine = new strStraightLine();
                strCurvedLine objCurvedLine = new strCurvedLine();
                strConnectionNode objConnectionNode = new strConnectionNode();
                int iOrder = 0;

                objShape.lstCurvedLine = new List<strCurvedLine>();
                objShape.lstStraightLine = new List<strStraightLine>();
                objShape.lstConnectionNodes = new List<strConnectionNode>();

                /*Start*/
                objShape.eSymbol = enumSymbolType.eSym_InternalStorage;
                objShape.eSize = enumSize.eLarge;

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = -40;
                objStraightLine.start.iVert = 40;
                objStraightLine.end.iHor = 40;
                objStraightLine.end.iVert = 40;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 40;
                objStraightLine.start.iVert = 40;
                objStraightLine.end.iHor = 40;
                objStraightLine.end.iVert = -40;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 40;
                objStraightLine.start.iVert = -40;
                objStraightLine.end.iHor = -40;
                objStraightLine.end.iVert = -40;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = -40;
                objStraightLine.start.iVert = -40;
                objStraightLine.end.iHor = -40;
                objStraightLine.end.iVert = 40;
                objShape.lstStraightLine.Add(objStraightLine);


                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = -35;
                objStraightLine.start.iVert = -40;
                objStraightLine.end.iHor = -35;
                objStraightLine.end.iVert = 40;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = -40;
                objStraightLine.start.iVert = -35;
                objStraightLine.end.iHor = 40;
                objStraightLine.end.iVert = -35;
                objShape.lstStraightLine.Add(objStraightLine);

                objConnectionNode.ePos = enumConnectionNodePosition.eTop;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = -15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = 15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eRight;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                lstShapes.Add(objShape);
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

        private void fillShapesAlternativeProcess()
        {
            try
            {
                strShape objShape = new strShape();
                strStraightLine objStraightLine = new strStraightLine();
                strCurvedLine objCurvedLine = new strCurvedLine();
                strConnectionNode objConnectionNode = new strConnectionNode();
                int iOrder = 0;

                objShape.lstCurvedLine = new List<strCurvedLine>();
                objShape.lstStraightLine = new List<strStraightLine>();
                objShape.lstConnectionNodes = new List<strConnectionNode>();

                /*Start*/
                objShape.eSymbol = enumSymbolType.eSym_AlternativeProcess;
                objShape.eSize = enumSize.eLarge;

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = -40;
                objStraightLine.start.iVert = 25;
                objStraightLine.end.iHor = 40;
                objStraightLine.end.iVert = 25;
                objShape.lstStraightLine.Add(objStraightLine);

                objCurvedLine = new strCurvedLine();
                iOrder++;
                objCurvedLine.order = iOrder;
                objCurvedLine.freshStart = false;
                objCurvedLine.CentreOfRotation.iHor = 40;
                objCurvedLine.CentreOfRotation.iVert = 15;
                objCurvedLine.startAngle = 0.5 * Math.PI;
                objCurvedLine.endAngle = 0 * Math.PI;
                objCurvedLine.radiusHor = 10;
                objCurvedLine.radiusVert = 10;
                objShape.lstCurvedLine.Add(objCurvedLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 50;
                objStraightLine.start.iVert = 15;
                objStraightLine.end.iHor = 50;
                objStraightLine.end.iVert = -15;
                objShape.lstStraightLine.Add(objStraightLine);

                objCurvedLine = new strCurvedLine();
                iOrder++;
                objCurvedLine.order = iOrder;
                objCurvedLine.freshStart = false;
                objCurvedLine.CentreOfRotation.iHor = 40;
                objCurvedLine.CentreOfRotation.iVert = -15;
                objCurvedLine.startAngle = 0 * Math.PI;
                objCurvedLine.endAngle = -0.5 * Math.PI;
                objCurvedLine.radiusHor = 10;
                objCurvedLine.radiusVert = 10;
                objShape.lstCurvedLine.Add(objCurvedLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 40;
                objStraightLine.start.iVert = -25;
                objStraightLine.end.iHor = -40;
                objStraightLine.end.iVert = -25;
                objShape.lstStraightLine.Add(objStraightLine);

                objCurvedLine = new strCurvedLine();
                iOrder++;
                objCurvedLine.order = iOrder;
                objCurvedLine.freshStart = false;
                objCurvedLine.CentreOfRotation.iHor = -40;
                objCurvedLine.CentreOfRotation.iVert = -15;
                objCurvedLine.startAngle = -0.5 * Math.PI;
                objCurvedLine.endAngle = -1 * Math.PI;
                objCurvedLine.radiusHor = 10;
                objCurvedLine.radiusVert = 10;
                objShape.lstCurvedLine.Add(objCurvedLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = -50;
                objStraightLine.start.iVert = -15;
                objStraightLine.end.iHor = -50;
                objStraightLine.end.iVert = 15;
                objShape.lstStraightLine.Add(objStraightLine);

                objCurvedLine = new strCurvedLine();
                iOrder++;
                objCurvedLine.order = iOrder;
                objCurvedLine.freshStart = false;
                objCurvedLine.CentreOfRotation.iHor = -40;
                objCurvedLine.CentreOfRotation.iVert = 15;
                objCurvedLine.startAngle = -1 * Math.PI;
                objCurvedLine.endAngle = -1.5 * Math.PI;
                objCurvedLine.radiusHor = 10;
                objCurvedLine.radiusVert = 10;
                objShape.lstCurvedLine.Add(objCurvedLine);

                objConnectionNode.ePos = enumConnectionNodePosition.eTop;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = -15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = 15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eRight;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                lstShapes.Add(objShape);
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

        private void fillShapesManualInput()
        {
            try
            {
                strShape objShape = new strShape();
                strStraightLine objStraightLine = new strStraightLine();
                strCurvedLine objCurvedLine = new strCurvedLine();
                strConnectionNode objConnectionNode = new strConnectionNode();
                int iOrder = 0;

                objShape.lstCurvedLine = new List<strCurvedLine>();
                objShape.lstStraightLine = new List<strStraightLine>();
                objShape.lstConnectionNodes = new List<strConnectionNode>();

                /*Start*/
                objShape.eSymbol = enumSymbolType.eSym_ManualInput;
                objShape.eSize = enumSize.eLarge;

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = -50;
                objStraightLine.start.iVert = 25;
                objStraightLine.end.iHor = 50;
                objStraightLine.end.iVert = 25;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 50;
                objStraightLine.start.iVert = 25;
                objStraightLine.end.iHor = 50;
                objStraightLine.end.iVert = -25;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 50;
                objStraightLine.start.iVert = -25;
                objStraightLine.end.iHor = -50;
                objStraightLine.end.iVert = -10;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = -50;
                objStraightLine.start.iVert = -10;
                objStraightLine.end.iHor = -50;
                objStraightLine.end.iVert = 25;
                objShape.lstStraightLine.Add(objStraightLine);

                objConnectionNode.ePos = enumConnectionNodePosition.eTop;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = -15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = 15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eRight;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                lstShapes.Add(objShape);
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

        private void fillShapesPredefinedProcess()
        {
            try
            {
                strShape objShape = new strShape();
                strStraightLine objStraightLine = new strStraightLine();
                strCurvedLine objCurvedLine = new strCurvedLine();
                strConnectionNode objConnectionNode = new strConnectionNode();
                int iOrder = 0;

                objShape.lstCurvedLine = new List<strCurvedLine>();
                objShape.lstStraightLine = new List<strStraightLine>();
                objShape.lstConnectionNodes = new List<strConnectionNode>();

                /*Start*/
                objShape.eSymbol = enumSymbolType.eSym_PredefinedProcess;
                objShape.eSize = enumSize.eLarge;

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = -50;
                objStraightLine.start.iVert = 25;
                objStraightLine.end.iHor = 50;
                objStraightLine.end.iVert = 25;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 50;
                objStraightLine.start.iVert = 25;
                objStraightLine.end.iHor = 50;
                objStraightLine.end.iVert = -25;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = 50;
                objStraightLine.start.iVert = -25;
                objStraightLine.end.iHor = -50;
                objStraightLine.end.iVert = -25;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = false;
                objStraightLine.start.iHor = -50;
                objStraightLine.start.iVert = -25;
                objStraightLine.end.iHor = -50;
                objStraightLine.end.iVert = 25;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = -40;
                objStraightLine.start.iVert = -25;
                objStraightLine.end.iHor = -40;
                objStraightLine.end.iVert = 25;
                objShape.lstStraightLine.Add(objStraightLine);

                objStraightLine = new strStraightLine();
                iOrder++;
                objStraightLine.order = iOrder;
                objStraightLine.freshStart = true;
                objStraightLine.start.iHor = 40;
                objStraightLine.start.iVert = -25;
                objStraightLine.end.iHor = 40;
                objStraightLine.end.iVert = 25;
                objShape.lstStraightLine.Add(objStraightLine);

                objConnectionNode.ePos = enumConnectionNodePosition.eTop;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = -15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = 15;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eRight;
                objConnectionNode.iHor = -15;
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                lstShapes.Add(objShape);
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

        private void fillShapesVDU()
        {
            try
            {
                strShape objShape = new strShape();
                strStraightLine objStraightLine = new strStraightLine();
                strCurvedLine objCurvedLine = new strCurvedLine();
                strConnectionNode objConnectionNode = new strConnectionNode();
                int iOrder = 0;

                objShape.lstCurvedLine = new List<strCurvedLine>();
                objShape.lstStraightLine = new List<strStraightLine>();
                objShape.lstConnectionNodes = new List<strConnectionNode>();

                /*Start*/
                //objShape.eSymbol = enumSymbolType.eSym_Output_ToDatabase;
                objShape.eSize = enumSize.eLarge;

                objCurvedLine = new strCurvedLine();
                iOrder++;
                objCurvedLine.order = iOrder;
                objCurvedLine.freshStart = true;
                objCurvedLine.CentreOfRotation.iHor = 0;
                objCurvedLine.CentreOfRotation.iVert = 16;
                objCurvedLine.startAngle = Math.PI + Math.Sin(2.0 / 5.0);
                objCurvedLine.endAngle = 1.5 * Math.PI;
                objCurvedLine.radiusHor = 40;
                objCurvedLine.radiusVert = 40;
                objShape.lstCurvedLine.Add(objCurvedLine);

                objCurvedLine = new strCurvedLine();
                iOrder++;
                objCurvedLine.order = iOrder;
                objCurvedLine.freshStart = false;
                objCurvedLine.CentreOfRotation.iHor = -15;
                objCurvedLine.CentreOfRotation.iVert = 0;
                objCurvedLine.startAngle = -1.0 * Math.Sin(3.0 / 4.0);
                objCurvedLine.endAngle = Math.Sin(3.0 / 4.0);
                objCurvedLine.radiusHor = 40;
                objCurvedLine.radiusVert = 40;
                objShape.lstCurvedLine.Add(objCurvedLine);

                objCurvedLine = new strCurvedLine();
                iOrder++;
                objCurvedLine.order = iOrder;
                objCurvedLine.freshStart = false;
                objCurvedLine.CentreOfRotation.iHor = 0;
                objCurvedLine.CentreOfRotation.iVert = -16;
                objCurvedLine.startAngle = 0.5 * Math.PI;
                objCurvedLine.endAngle = Math.PI - Math.Sin(2.0 / 5.0);
                objCurvedLine.radiusHor = 40;
                objCurvedLine.radiusVert = 40;
                objShape.lstCurvedLine.Add(objCurvedLine);

                objConnectionNode.ePos = enumConnectionNodePosition.eLeft;
                objConnectionNode.iHor = (int)(Math.Sqrt((40 ^ 2) - (16 ^ 2)));
                objConnectionNode.iVert = 0;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eBottom;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = 30;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objConnectionNode.ePos = enumConnectionNodePosition.eTop;
                objConnectionNode.iHor = 0;
                objConnectionNode.iVert = -30;
                objShape.lstConnectionNodes.Add(objConnectionNode);

                objShape.eSymbol = enumSymbolType.eSym_VDU;
                lstShapes.Add(objShape);

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

        public void reposition()
        {
            try
            {
                /*
                 * Go to the this.rootOfAllSquaresIdand then find all the child squares
                 * repeat by finding all the child squares of those until you get to the last generation of child squares
                 * lay out the symbols in the last generation of squares
                 * go back one generation of squares and lay out the symbols and the squares in that generation.
                 * and repeat up the generations.
                 * 
                 */

                int iIndexRootSquare = lstSquares.FindIndex(x => x.iId == this.rootOfAllSquaresId);



                strSquare objSquare = lstSquares[iIndexRootSquare];



                //lstSquares.Max(x => x. )

                //foreach(strSquare objSquare in lstSquares.FindAll()) 

                

                /*
                int iMaxHor = lstSymbols.Max(x => x.iHorPos);
                int iMinHor = lstSymbols.Min(x => x.iHorPos);

                int iMaxVert = lstSymbols.Max(x => x.iVertPos);
                int iMinVert = lstSymbols.Min(x => x.iVertPos);

                for (int iIndex = 0; iIndex < lstSymbols.Count; iIndex++)
                {
                    ClsConfigReporterFlowDiagram.strSymbol objSymbol = lstSymbols[iIndex];
                    objSymbol.iHorPos += ciHorSpacing - iMinHor;
                    objSymbol.iVertPos += ciVertSpacing - iMinVert;
                    lstSymbols[iIndex] = objSymbol;
                }

                iCanvasWidth = iMaxHor - iMinHor + 2 * ciHorSpacing;
                iCanvasHeight = iMaxVert - iMinVert + 2 * ciVertSpacing;
                */
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

        public void reposition(int iSquareId)
        {
            try
            {
                int iIndexSquare = lstSquares.FindIndex(x => x.iId == this.rootOfAllSquaresId);
                strSquare objSquare = lstSquares[iIndexSquare];

                foreach (int iChildSquareId in objSquare.lstSquares)
                { reposition(iChildSquareId); }

                int iIndexSymbol = lstSymbols.FindIndex(x => x.id == objSquare.iStartSymbol);

                strSymbol objSymbol = lstSymbols[iIndexSymbol];

                reposition(ref objSymbol, objSquare.iLeft, objSquare.iTop);
                objSymbol.iHorPos = objSquare.iLeft;
                objSymbol.iVertPos = objSquare.iTop;

                switch (objSymbol.lstSymbolsGoingTo.Count)
                {
                    case 0:
                        //nothing
                        break;
                    case 1:
                        //next symbol is down
                        break;
                    case 2:
                        //one symbol is down 
                        //the other symbol is to the right
                        break;
                    case 3:
                        //one symbol is down 
                        //an other symbol is to the right
                        //the other symbol is to the left
                        break;
                    default:
                        //all symbols are down and spread cross forming a line
                        break;
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

        public void reposition(ref strSymbol objSymbol, int iHor, int iVert)
        {
            try
            {
                int iSymbolId = objSymbol.id;

                if (lstSquares.Exists(x => x.lstSymbols.Exists(y => y == iSymbolId)))
                {
                    //reposition_AddDelta();
                }
                else
                {
                objSymbol.iHorPos = iHor;
                objSymbol.iVertPos = iVert;
                

                switch (objSymbol.lstSymbolsGoingTo.Count)
                {
                    case 0:
                        //nothing
                        break;
                    case 1:
                        //next symbol is down
                        int iNextSymbolIdA = objSymbol.lstSymbolsGoingTo[0];
                        int iIndexNextSymbolA = lstSymbols.FindIndex(x => x.id == iNextSymbolIdA);
                        strSymbol objNextSymbolA = lstSymbols[iIndexNextSymbolA];
                        reposition(ref objNextSymbolA, iHor, iVert + ciVertSpacing);
                        break;
                    case 2:
                        //one symbol is down 
                        //the other symbol is to the right
                        int iNextSymbolIdB0 = objSymbol.lstSymbolsGoingTo[0];
                        int iIndexNextSymbolB0 = lstSymbols.FindIndex(x => x.id == iNextSymbolIdB0);
                        strSymbol objNextSymbolB0 = lstSymbols[iIndexNextSymbolB0];
                        reposition(ref objNextSymbolB0, iHor, iVert + ciVertSpacing);

                        int iNextSymbolIdB1 = objSymbol.lstSymbolsGoingTo[1];
                        int iIndexNextSymbolB1 = lstSymbols.FindIndex(x => x.id == iNextSymbolIdB1);
                        strSymbol objNextSymbolB1 = lstSymbols[iIndexNextSymbolB1];
                        reposition(ref objNextSymbolB1, iHor + ciHorSpacing, iVert);
                        break;
                    case 3:
                        //one symbol is down 
                        //an other symbol is to the right
                        //the other symbol is to the left
                        int iNextSymbolIdC0 = objSymbol.lstSymbolsGoingTo[0];
                        int iIndexNextSymbolC0 = lstSymbols.FindIndex(x => x.id == iNextSymbolIdC0);
                        strSymbol objNextSymbolC0 = lstSymbols[iIndexNextSymbolC0];
                        reposition(ref objNextSymbolC0, iHor, iVert + ciVertSpacing);

                        int iNextSymbolIdC1 = objSymbol.lstSymbolsGoingTo[1];
                        int iIndexNextSymbolC1 = lstSymbols.FindIndex(x => x.id == iNextSymbolIdC1);
                        strSymbol objNextSymbolC1 = lstSymbols[iIndexNextSymbolC1];
                        reposition(ref objNextSymbolC1, iHor + ciHorSpacing, iVert);

                        int iNextSymbolIdC2 = objSymbol.lstSymbolsGoingTo[2];
                        int iIndexNextSymbolC2 = lstSymbols.FindIndex(x => x.id == iNextSymbolIdC2);
                        strSymbol objNextSymbolC2 = lstSymbols[iIndexNextSymbolC2];
                        reposition(ref objNextSymbolC2, iHor - ciHorSpacing, iVert);
                        break;
                    default:
                        //all symbols are down and spread cross forming a line
                        for (int iCounter = 0; iCounter < objSymbol.lstSymbolsGoingTo.Count; iCounter++)
                        {
                            int iNextSymbolIdD = objSymbol.lstSymbolsGoingTo[iCounter];
                            int iIndexNextSymbolD = lstSymbols.FindIndex(x => x.id == iNextSymbolIdD);
                            strSymbol objNextSymbolD = lstSymbols[iIndexNextSymbolD];
                            reposition(ref objNextSymbolD, iHor + (ciHorSpacing * iCounter), iVert + ciVertSpacing);
                        }
                        break;
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


        public void reposition_AddDelta(ref strSquare objSquare, int iDeltaHor, int iDeltaVert)
        {
            try
            {
                foreach (int iSymbolId in objSquare.lstSymbols)
                {
                    int iSymbolIndex = lstSymbols.FindIndex(x => x.id == iSymbolId);
                    strSymbol objSymbol = lstSymbols[iSymbolIndex];

                    objSymbol.iHorPos += iDeltaHor;
                    objSymbol.iVertPos += iDeltaVert;

                    lstSymbols[iSymbolIndex] = objSymbol;
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

        public int addSymbol(strSymbol objSymbol)
        {
            try
            {
                int iId;

                if (lstSymbols.Count == 0)
                { iId = 1; }
                else
                { iId = lstSymbols.Max(x => x.id) + 1; }
                
                objSymbol.id = iId;
                lstSymbols.Add(objSymbol);

                return iId;
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

                return -1;
            }
        }

        public void addLoop(strLoops objLoop)
        {
            try
            {
                int iId;

                if (lstLoops.Count == 0)
                { iId = 1; }
                else
                { iId = lstLoops.Max(x => x.iId) + 1; }

                objLoop.iId = iId;

                lstLoops.Add(objLoop);
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

        public void loopEndIf(int iSymbolStart, int iSymbolEnd)
        {
            try
            {
                int iIndex = lstLoops.FindIndex(x => x.iSymbolStart == iSymbolStart);

                if (iIndex > 0)
                {
                    strLoops objLoop = lstLoops[iIndex];
                    objLoop.iSymbolEnd = iSymbolEnd; 
                    lstLoops[iIndex] = objLoop;
                }
                else
                {
                    strLoops objLoop = new strLoops();
                    int iId;

                    if (lstLoops.Count == 0)
                    { iId = 1; }
                    else
                    { iId = lstLoops.Max(x => x.iId) + 1; }

                    objLoop.iId = iId;
                    objLoop.iSymbolStart = -1; 
                    objLoop.iSymbolEnd = iSymbolEnd; 
                    lstLoops.Add(objLoop);
                }

                
                /*
                objLoop.iSymbolStart 

                int iId;

                if (lstLoops.Count == 0)
                { iId = 1; }
                else
                { iId = lstLoops.Max(x => x.iId) + 1; }

                objLoop.iId = iId;

                lstLoops.Add(objLoop);
                */ 
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
        /*
        public void addSquare(strSquare objSquare)
        {
            try
            {
                int iId;

                if (lstSquares.Count == 0)
                { iId = 1; }
                else
                { iId =lstSquares.Max(x=>x.iId) + 1; }

                objSquare.iId = iId;

                lstSquares.Add(objSquare);
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
        public int addSquare()
        {
            try
            {
                strSquare objSquare = new strSquare();
                int iId;

                if (lstSquares.Count == 0)
                { iId = 1; }
                else
                { iId = lstSquares.Max(x => x.iId) + 1; }

                objSquare.iId = iId;
                objSquare.iHeight = 0;
                objSquare.iLeft = 0;
                objSquare.iTop = 0;
                objSquare.iWidth = 0;
                objSquare.lstLoops = new List<int>();
                objSquare.lstSquares = new List<int>();
                objSquare.lstSymbols = new List<int>();

                lstSquares.Add(objSquare);

                return iId;
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

                return -1;
            }
        }

        public void squareAddSymbol(int iSquareId, int iSymbolId)
        {
            try
            {
                int iIndex = lstSquares.FindIndex(x => x.iId == iSquareId);

                if (iIndex >= 0)
                {
                    strSquare objSquare = lstSquares[iIndex];

                    objSquare.lstSymbols.Add(iSymbolId);

                    lstSquares[iIndex] = objSquare;
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

        public void squareAddSquare(int iSquareId, int iSquareChildId)
        {
            try
            {
                int iIndex = lstSquares.FindIndex(x => x.iId == iSquareId);

                if (iIndex >= 0)
                {
                    strSquare objSquare = lstSquares[iIndex];

                    objSquare.lstSquares.Add(iSquareChildId);

                    lstSquares[iIndex] = objSquare;
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


        private List<ClsConfigReporterFlowDiagram.strSymbol> getSymbolsInLoop(ClsConfigReporterFlowDiagram.strLoops objLoop)
        {
            try
            {
                List<ClsConfigReporterFlowDiagram.strSymbol> lstResult = new List<ClsConfigReporterFlowDiagram.strSymbol>();
                bool bDone = false;
                List<int> lstIdToAdd = new List<int>();
                int iPerviousCount = 0;

                lstIdToAdd.Add(objLoop.iSymbolStart);

                while (!bDone)
                {
                    foreach (ClsConfigReporterFlowDiagram.strSymbol objSymbol in lstSymbols.FindAll(x => lstIdToAdd.Contains(x.id) && x.id != objLoop.iSymbolEnd))
                    { lstResult.Add(objSymbol); }

                    lstResult = lstResult.Distinct().ToList<ClsConfigReporterFlowDiagram.strSymbol>();

                    foreach (ClsConfigReporterFlowDiagram.strSymbol objSymbol in lstResult)
                    {
                        foreach (int iGoingTo in objSymbol.lstSymbolsGoingTo)
                        { lstIdToAdd.Add(iGoingTo); }
                    }

                    lstIdToAdd = lstIdToAdd.Distinct().ToList<int>();

                    if (iPerviousCount == lstIdToAdd.Count)
                    { bDone = true; }
                }

                return lstResult;
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

                return new List<ClsConfigReporterFlowDiagram.strSymbol>();
            }
        }
    }
}
