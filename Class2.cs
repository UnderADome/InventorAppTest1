using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Inventor;

/*
 * 二次开发入门的几篇简单博客
 https://blog.csdn.net/qq_43006346/category_9762721.html

 */

namespace InventorAppTest1
{
    class Class2
    {

        #region 读取和修改Inventor模型的属性

        static void main(string[] args)
        {
            Inventor.Application inventorApp = null;
            try
            {
                inventorApp = Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            }
            catch
            {
                var inventorType = Type.GetTypeFromProgID("Inventor.Application");
                inventorApp = Activator.CreateInstance(inventorType) as Inventor.Application;
                inventorApp.Visible = true;
            }

            AssemblyDocument asmDoc = (AssemblyDocument)inventorApp.Documents.Open(@"Path", true);

            AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;

            foreach (ComponentOccurrence oOcc in asmDef.Occurrences)
            {
                Document oDoc = (Document)oOcc.Definition.Document;
                // using GUID to find the "Summary Information"
                PropertySet oPS = oDoc.PropertySets["{F29F85E0-4FF9-1068-AB91-08002B27B3D9}"];
                // using proID to find TITLE
                Inventor.Property oP = oPS.ItemByPropId[(int)PropertiesForSummaryInformationEnum.kTitleSummaryInformation];
                oP.Value = oP.Value + "new";
            }
        }/* https://blog.csdn.net/beihuanlihe130/article/details/107352288 */
        #endregion
        
        




    }
}
