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

        public static void Main(string[] args)
        {
            GetDrawingDimension();
            //防止闪退
            System.Console.ReadKey();
            System.Console.ReadKey();
            System.Console.ReadKey();
        }

        #region 读取和修改Inventor模型的属性
        static void ReadAndReviseModelProperties()
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

        private static Inventor.Application inventorApp = null;

        #region 获取工程图中的标注、键号等等
        private static void GetDrawingDimension()
        {
            try
            {
                inventorApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
                Console.WriteLine("查找到可用的实例");
            }
            catch { Console.WriteLine("未打开Inventor"); return; }
            DrawingDocument drawingDocument = (DrawingDocument)inventorApp.ActiveDocument;

            //在Inventor当前正在显示的工程图不一样的时候，ActiveSheet也会发生变化
            Console.WriteLine("打开的图纸："+ drawingDocument.ActiveSheet.Name + " " + drawingDocument.FullFileName);
            DrawingView drawingView = drawingDocument.ActiveSheet.DrawingViews[1];
            
            //特别注明：该类及其方法仅针对模型和草图文件
            GeneralDimensionsEnumerator generalDimensionsEnumerator = 
                drawingDocument.ActiveSheet.DrawingDimensions.GeneralDimensions.Retrieve(drawingView);
            Console.WriteLine("generalDimensionsEnumerator.Count = " + generalDimensionsEnumerator.Count);
            if (generalDimensionsEnumerator.Count != 0)
            {
                for (int i = 1; i <= generalDimensionsEnumerator.Count; i++)
                {
                    Console.WriteLine(generalDimensionsEnumerator[i].Text);
                }
            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////
            BaselineDimensionSets baselineDimensionSets = drawingDocument.ActiveSheet.DrawingDimensions.BaselineDimensionSets;
            Console.WriteLine("baselineDimensionSets.Count = " + baselineDimensionSets.Count);
            if (baselineDimensionSets.Count != 0)
            {
                for (int i=1; i<=baselineDimensionSets.Count; i++)
                {
                    BaselineDimensionSet baselineDimensionSet = baselineDimensionSets[i];
                    Console.WriteLine("baselineDimensionSet.Members = " + baselineDimensionSet.Members);
                    Console.WriteLine("baselineDimensionSet.DimensionType" + baselineDimensionSet.DimensionType);
                }
            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////
            Balloons ballons = drawingDocument.ActiveSheet.Balloons;
            Console.WriteLine("ballons.Count = " + ballons.Count);
            Balloon balloon = null;
            if (ballons.Count != 0)
            {
                for (int i=1; i<=ballons.Count; i++)
                {
                    Console.WriteLine("\n------------------------ballons[" + i+ "]------------------------");
                    balloon = ballons[i];
                    //Console.WriteLine("balloon.Leader.RootNode = " + balloon.Leader.RootNode);  //打印出 System.__ComObject
                    //Console.WriteLine("balloon.Position = " + balloon.Position);  //打印出 System.__ComObject
                    AttributeSets attributeSets = balloon.AttributeSets;
                    Console.WriteLine("attributeSets.Count = " + attributeSets.Count);
                    for (int j=1; j<=attributeSets.Count; j++)
                    {
                        AttributeSet attributeSet = attributeSets[j];
                        Console.WriteLine("attributeSet.Name = " + attributeSet.Name);
                    }

                    BalloonValueSets balloonValueSets = balloon.BalloonValueSets;
                    for (int j = 1; j <= balloonValueSets.Count; j++)
                    {
                        BalloonValueSet balloonValueSet = balloonValueSets[j];
                        Console.WriteLine("balloonValueSet.ItemNumber = "+balloonValueSet.ItemNumber);
                        Console.WriteLine("balloonValueSet.Value = " + balloonValueSet.Value);
                        Console.WriteLine("balloonValueSet.OverrideValue = " + balloonValueSet.OverrideValue);
                        //Console.WriteLine("balloonValueSet.ReferencedFiles = " + balloonValueSet.ReferencedFiles);
                        Console.WriteLine("balloonValueSet.Type = " + balloonValueSet.Type);  
                    }

                    Leader leader = balloon.Leader;
                    Console.WriteLine("leader.ArrowheadType = " + leader.ArrowheadType);
                    Console.WriteLine("leader.Type = "+leader.Type);
                    AttributeSets attributeSets_leader = leader.AttributeSets;
                    Console.WriteLine("attributeSets_leader.Count = " + attributeSets_leader.Count);
                    for (int j=0; j<attributeSets_leader.Count; j++)
                    {
                        AttributeSet attributeSet = attributeSets[j];
                        Console.WriteLine("attributeSet_leader.Name = " + attributeSet.Name);
                    }
                    
                    Console.WriteLine("END------------------------ballons[" + i + "]------------------------\n");
                }

            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////
            //DrawingViews views = drawingDocument.ActiveSheet.DrawingViews;
            //Console.WriteLine("views.count = " + views.Count);

            Console.WriteLine("drawingDocument.SelectSet.Count = " + drawingDocument.SelectSet.Count);
            SelectSet selectSet = null;
            DrawingCurveSegment drawingCurveSegment = null;
            if (drawingDocument.SelectSet.Count == 0)
            {
                Console.WriteLine("Select a drawing view");
                DrawingView view = inventorApp.CommandManager.Pick(SelectionFilterEnum.kDrawingViewFilter, "Select a drawing view");
                //selectSet = inventorApp.CommandManager.Pick(SelectionFilterEnum.kDrawingSheetFilter, "Select drawing sheet!");
                drawingCurveSegment = inventorApp.CommandManager.Pick(SelectionFilterEnum.kDrawingCurveSegmentFilter, "Select drawing segment filter");
            }
            else
            {
                selectSet = drawingDocument.SelectSet;
            }
             
            //DrawingCurveSegment drawingCurveSegment = selectSet[1];//drawingDocument.SelectSet[1];
            DrawingCurve drawingCurve = drawingCurveSegment.Parent;

            //Get the mid point of the selected curve assuming that the selection curve is linear
            Point2d MidPoint = drawingCurve.MidPoint;

            //Set a reference to the TransientGeometry object.
            TransientGeometry TG = inventorApp.TransientGeometry;
            Console.WriteLine("TG : " + (TG == null));
            ObjectCollection LeaderPoints = inventorApp.TransientObjects.CreateObjectCollection();
            Console.WriteLine("LeaderPoints : " + (LeaderPoints == null));

            LeaderPoints.Add(TG.CreatePoint2d(MidPoint.X + 10, MidPoint.Y + 10));
            LeaderPoints.Add(TG.CreatePoint2d(MidPoint.X + 10, MidPoint.Y + 5));

            //Add the GeometryIntent to the leader points collection.
            //This is the geometry that the balloon will attach to.
            GeometryIntent geometryIntent = drawingDocument.ActiveSheet.CreateGeometryIntent(drawingCurve);
            LeaderPoints.Add(geometryIntent);

            //Set a reference to the parent drawing view of the selected curve
            //DrawingView 
            drawingView = drawingCurve.Parent;

            //Set a reference to the referenced model document
            Document ModelDoc = drawingView.ReferencedDocumentDescriptor.ReferencedDocument;
            Console.WriteLine(ModelDoc.Type);
            //PartDocument ModelDoc = drawingView.ReferencedDocumentDescriptor.ReferencedDocument;
            //AssemblyDocument ModelDoc = drawingView.ReferencedDocumentDescriptor.ReferencedDocument;

            //Check if a partslist or a balloon has already been created for thie model
            Boolean IsDrawingBOMDefined = drawingDocument.DrawingBOMs.IsDrawingBOMDefined(ModelDoc.FullFileName);

            // Balloon balloon;
            
            if (IsDrawingBOMDefined)

            {   //当DrawingBOM已经被定义了
                //Just create the balloon with the leader points. All other arguments can be ignored
                Console.WriteLine("当DrawingBOM已经被定义了\n创建气泡标注");
                balloon = drawingDocument.ActiveSheet.Balloons.Add(LeaderPoints);
            }
            else
            {
                //当DrawingBOM没有被定义
                AssemblyDocument assemblyDocument = (AssemblyDocument)ModelDoc;
                AssemblyComponentDefinition assemblyComponentDefinition = assemblyDocument.ComponentDefinition;

                ///* 
                //First check if the 'structured' BOM view has been enabled in the model
                //Set a reference to the model's BOM object
                //BOM bom = ModelDoc.ComponentDefinition.BOM;
                BOM bom = assemblyComponentDefinition.BOM;

                if (bom.StructuredViewEnabled)
                {
                    //Level needs to be specifieed. Numbering options jave already been defined.
                    //Get the Level('All levels' of 'First level only') from the model BOM view - must use the same here
                    PartsListLevelEnum Level;
                    if (bom.StructuredViewFirstLevelOnly)
                        Level = PartsListLevelEnum.kStructured;
                    else
                        Level = PartsListLevelEnum.kStructuredAllLevels;
                }
                else
                {
                    //Level and numbering options must be specifieed. 
                    //The corresponding model BOM view will automatically be enabled
                    NameValueMap NumberingScheme = inventorApp.TransientObjects.CreateNameValueMap();
                    //Add the option for a comma delimiter
                    NumberingScheme.Add("Delimeter", ",");
                    //Create the balloon by specifying the level and numbering scheme
                    balloon = drawingDocument.ActiveSheet.Balloons.Add(LeaderPoints, PartsListLevelEnum.kStructuredAllLevels, NumberingScheme);
                }
        //*/
            }

        }

        #endregion



    }
}







//示例代码有C#的...而且与VBA的一一对应...总搞忘记了
