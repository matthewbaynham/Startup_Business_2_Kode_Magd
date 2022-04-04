using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Windows.Forms;

namespace KodeMagd.Documenter
{
    class ClsObjectMapper
    {
        /*
        <!DOCTYPE HTML>
        <html>
        <head>
        <style>
        #div1 {left:0px;top:0px;width:100%;height:1000px;padding:0px;border:1px solid #aaaaaa;}
        #drag1 {position:absolute;left:100px;top:100px;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
        #drag2 {position:absolute;left:200px;top:200px;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}

        </style>
        <script>

        function allowDrop(ev)
        {
            ev.preventDefault();
        }

        function drag(ev)
        {
            ev.dataTransfer.setData("Text",ev.target.id);
        }

        function drop(ev)
        {
            ev.preventDefault();
            var data=ev.dataTransfer.getData("Text");
            ev.target.appendChild(document.getElementById(data));


            document.getElementById("id_name").innerHTML = document.getElementById(data).id;
            document.getElementById("id_height").innerHTML = ev.clientY;
            document.getElementById("id_width").innerHTML = ev.clientX;

            document.getElementById("div1").style.backgroundColor = "red";
            document.getElementById("drag1").style.backgroundColor = "blue";
            document.getElementById("drag2").style.backgroundColor = "yellow";

            document.getElementById(data).style.left = ev.clientX + "px";
            document.getElementById(data).style.top = ev.clientY + "px";

        }

        function init()
        {
            document.getElementById("div1").style.backgroundColor = "red";
            document.getElementById("drag1").style.backgroundColor = "blue";
            document.getElementById("drag2").style.backgroundColor = "yellow";
        }

        </script>
        </head>
        <body onLoad='init()'>
            <div id="div1" ondrop="drop(event)" ondragover="allowDrop(event)">
                 <div id="drag1" draggable="true" ondragstart="drag(event)">
                     <table>
                        <tr>
                            <td>Name</td>
                            <td><div id="id_name"></div></td>
                        </tr>
                        <tr>
                            <td>Height</td>
                            <td><div id="id_height"></div></td>
                        </tr>
                        <tr>
                            <td>Width</td>
                            <td><div id="id_width"></div></td>
                        </tr>
                     </table>

                 </div>
                 <div id="drag2" draggable="true" ondragstart="drag(event)">
                     <table>
                        <tr>
                            <td>Name</td>
                            <td><div id="id_name"></div></td>
                        </tr>
                        <tr>
                            <td>Height</td>
                            <td><div id="id_height"></div></td>
                        </tr>
                        <tr>
                            <td>Width</td>
                            <td><div id="id_width"></div></td>
                        </tr>
                     </table>

                 </div>
            </div>
        </body>
        </html>
         */

        public ClsObjectMapper(ref ClsCodeMapperWrk cCodeMapperWrk)
        {
            try
            {
                mapWrk(ref cCodeMapperWrk);
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

        public void mapWrk(ref ClsCodeMapperWrk cCodeMapperWrk)
        {
            try
            {
                foreach (ClsCodeMapper.strModuleDetails objModuleDetails in cCodeMapperWrk.getLstModuleDetails())
                {
                    ClsCodeMapper cCodeMapper = cCodeMapperWrk.getCodeMapper(objModuleDetails.sName);

                    foreach (ClsCodeMapper.strFunctions objFunction in cCodeMapper.getLstFunctions())
                    {

                    }

                    foreach (ClsCodeMapper.strVariables objVariable in cCodeMapper.lstVariables())
                    {

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

        public string generateHtml() {
            try
            {
                string sResults = "";

                /*
<!DOCTYPE HTML>
<html>
<head>
<title>Test 012</title>
<style>


#div1 {left:0px;top:0px;width:100%;height:1000px;padding:0px;border:1px solid #aaaaaa;}
#drag1 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag2 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag3 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag4 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag5 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag6 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag7 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag8 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag9 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag10 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag11 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag12 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag13 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag14 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag15 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag16 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag17 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag18 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag19 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}
#drag20 {position:absolute;width:100px;height:60px;padding:10px;border:1px solid #aaaaaa;}

</style>
<script>

function allowDrop(ev)
{
    ev.preventDefault();
}

function drag(ev)
{
    ev.dataTransfer.setData("Text",ev.target.id);
}

function drop(ev)
{
    ev.preventDefault();
    var data=ev.dataTransfer.getData("Text");
    ev.target.appendChild(document.getElementById(data));


    document.getElementById("id_name").innerHTML = document.getElementById(data).id;
    document.getElementById("id_height").innerHTML = ev.clientY;
    document.getElementById("id_width").innerHTML = ev.clientX;

    document.getElementById(data).style.left = ev.clientX + "px";
    document.getElementById(data).style.top = ev.clientY + "px";

}

function init()
{

//loop through array of names and position them encase, note take the size of the doc as is because the window could be any size.
var arrNames = new Array("drag1", "drag2", "drag3", "drag4", "drag5", "drag6", "drag7", "drag8", "drag9", "drag10", "drag11", "drag12", "drag13", "drag14", "drag15", "drag16", "drag17", "drag18", "drag19", "drag20");

//var iWidth = parseInt(document.width);
var iHorSpacing = (+150);
var iWidthCounter = (+iHorSpacing);

for (sname in arrNames)
{
    document.getElementById(arrNames[sname]).style.top = iHorSpacing + "px";
    document.getElementById(arrNames[sname]).style.left = iWidthCounter + "px";
    document.getElementById(arrNames[sname]).style.backgroundColor = "red";

    iWidthCounter = (+iWidthCounter) + (+document.getElementById(arrNames[sname]).style.width) + (+iHorSpacing);
}

}

</script>
</head>
<body onLoad='init()'>
    <div id="div1" ondrop="drop(event)" ondragover="allowDrop(event)">
         <div id="drag1" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>

         </div>
         <div id="drag2" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag3" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag4" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag5" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag6" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag7" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag8" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag9" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag10" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag11" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag12" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag13" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag14" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag15" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag16" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag17" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag18" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag19" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
         <div id="drag20" draggable="true" ondragstart="drag(event)">
             <table>
                <tr>
                    <td>Name</td>
                    <td><div id="id_name"></div></td>
                </tr>
                <tr>
                    <td>Height</td>
                    <td><div id="id_height"></div></td>
                </tr>
                <tr>
                    <td>Width</td>
                    <td><div id="id_width"></div></td>
                </tr>
             </table>
         </div>
    </div>
</body>
</html>
                 */


                return sResults;
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

                return "";
            }
        }

    }
}
