using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;  //交互性操作
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Inventor;
using Application = Inventor.Application;

namespace InventroAppTest1
{
    /**
     * 可以从一个正在运行的Inventor实例中获取应用程序对象
     * 也可以选择启动Inventor并从中获取应用程序对象
     * */
    class Program
    {
        /*
        static void Main(string[] args)
        {
            //GetExtrudeFeature();
            //ShowExtrudeFeature();
            //TestFunction();
            //SuppressOff();

            //GetInventorApp();
            //AFunction();
            //EditDrawingDemensions();
            //防止闪退
            System.Console.ReadKey();
        }
        */

        //拉伸特征、突出的特征
        private static void GetExtrudeFeature()
        {
            Inventor.Application inventorApp = null;
            try
            {
                //获取一个Inventor的参考
                inventorApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
            }
            catch
            {
                MessageBox.Show("没有正常连接到Inventor");
                return;
            }

            PartDocument partDoc = null;
            partDoc = (PartDocument)inventorApp.ActiveDocument;

            ExtrudeFeature extrude = null;
            //最后调用ExtrudeFeatures对象的Item属性，返回具有指定名称的ExtrudeFeature对象
            extrude = partDoc.ComponentDefinition.Features.ExtrudeFeatures[1];  

            MessageBox.Show("Extrusion " + extrude.Name + " is suppressed:" + extrude.Suppressed);

        }

        private static void ShowExtrudeFeature()
        {
            //Inventor.Application inventorApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
            //PartDocument partDoc = (PartDocument)inventorApp.ActiveDocument;
            Inventor.Application inventorApp = null;
            try
            {
                //获取一个Inventor的参考
                inventorApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
            }
            catch
            {
                MessageBox.Show("没有正常连接到Inventor");
                return;
            }

            PartDocument partDoc = null;
            partDoc = (PartDocument)inventorApp.ActiveDocument;
            ExtrudeFeatures extrudeFeatures = partDoc.ComponentDefinition.Features.ExtrudeFeatures;
            for (int i=0; i < extrudeFeatures.Count; i++)
            {
                Console.WriteLine(extrudeFeatures[1].Name);
            }
        }

        private static void TestFunction()
        {
            Console.WriteLine("TestFunction");
            Inventor.Application application = Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            PartDocument partDoc = application.Documents.Add(DocumentTypeEnum.kPartDocumentObject,
                    application.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject,
                                                     SystemOfMeasureEnum.kDefaultSystemOfMeasure,
                                                     DraftingStandardEnum.kDefault_DraftingStandard, null),
                                                     true) as PartDocument;
            /*
             * PartComponentDefinition.Sketches Property
             * 二维草图的基本方法：
             * 
             * Add 方法增加一个二维图纸
             * Count 返回该集合里的草图个数
             * Item 允许通过迭代的方式访问每一个成员
             * 
             * Delete 删除
             * Edit 使草图处于编辑状态
             * ExitEdit 草图退出编辑状态
             * name 名称
             * visible 可见性
             */
            ExtrudeFeatures extrudeFeatures = partDoc.ComponentDefinition.Features.ExtrudeFeatures;
            Console.WriteLine(extrudeFeatures.Count);
            for (int i = 0; i < extrudeFeatures.Count; i++)
            {
                Console.WriteLine(extrudeFeatures[i].Name);
            }
            Console.WriteLine("out");
        }

        private static void SuppressOff()
        {
            Inventor.Application application = Marshal.GetActiveObject("Inventor.Application") as Inventor.Application;
            PartDocument partDoc = application.Documents.Add(DocumentTypeEnum.kPartDocumentObject,
                    application.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject,
                                                     SystemOfMeasureEnum.kDefaultSystemOfMeasure,
                                                     DraftingStandardEnum.kDefault_DraftingStandard, null),
                                                     true) as PartDocument;
            PartFeature partFeature = null;
            for (int i=0; i<partDoc.ComponentDefinition.Features.Count; i++)
            {
                if (partFeature.Suppressed)
                {
                    partFeature.Suppressed = false;
                }
            }
        }

        private static Application GetInventorApp()
        {
            Console.WriteLine("in---------------");
            Console.WriteLine("获取全局Application");
            Application inventorApp = null;
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

            Console.WriteLine("创建零件文档");
            PartDocument partDoc = inventorApp.Documents.Add
                (DocumentTypeEnum.kPartDocumentObject, 
                 inventorApp.FileManager.GetTemplateFile
                        (DocumentTypeEnum.kPartDocumentObject, 
                         SystemOfMeasureEnum.kDefaultSystemOfMeasure, 
                         DraftingStandardEnum.kDefault_DraftingStandard, 
                         null),
                 true) as PartDocument;
            //Console.WriteLine("打开零件文档");
            //PartDocument partDoc = (PartDocument)inventorApp.Documents.Open(filename, true);

            Console.WriteLine("创建和打开部件文档");
            AssemblyDocument asmDoc = inventorApp.Documents.Add
                (DocumentTypeEnum.kAssemblyDocumentObject,
                 inventorApp.FileManager.GetTemplateFile
                        (DocumentTypeEnum.kAssemblyDocumentObject,
                        SystemOfMeasureEnum.kDefaultSystemOfMeasure,
                        DraftingStandardEnum.kDefault_DraftingStandard,
                        null),
                 true) as AssemblyDocument;
            //Console.WriteLine("打开部件文档");
            //AssemblyDocument asmDoc = (AssemblyDocument)inventorApp.Documents.Open(fileName, true);

            return inventorApp;
        }

        /**
         * C#读取和修改Inventor模型的属性
         */
        private static void AFunction()
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
        }

        /*编辑工程图，修改标注尺寸*/
        private static void EditDrawingDemensions()
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

            //DrawingDocument drawingDocument = (DrawingDocument)inventorApp.ActiveDocument;
            DrawingDocument drawingDocument = inventorApp.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject,
                                                        inventorApp.FileManager.GetTemplateFile(DocumentTypeEnum.kDrawingDocumentObject,
                                                            SystemOfMeasureEnum.kDefaultSystemOfMeasure,
                                                            DraftingStandardEnum.kDefault_DraftingStandard,
                                                            null),
                                                        true) as DrawingDocument;
            Sheet sheet = drawingDocument.ActiveSheet;
            long counter = 1;
            foreach (DrawingDimension drawingDimension in sheet.DrawingDimensions)
            {
                drawingDimension.Text.FormattedText = " (Metric)";
                counter = counter + 1;
            }

            //设置对集合中第一个常规维度的引用
            GeneralDimension generalDimension = sheet.DrawingDimensions.GeneralDimensions[1];  //在VBA中是GeneralDimensions.Item[1]

            //设置对该维度的维度样式的引用
            DimensionStyle dimStyle = generalDimension.Style;

            //修改维度样式的一些属性
            //这将修改所有使用此样式的尺寸
            dimStyle.LinearPrecision = LinearPrecisionEnum.kFourDecimalPlacesLinearPrecision;  //精度（长度）：小数点后四位
            dimStyle.AngularPrecision = AngularPrecisionEnum.kFourDecimalPlacesAngularPrecision;  //精度（角度）：小数点后四位
            dimStyle.LeadingZeroDisplay = false;
            dimStyle.Tolerance.SetToSymmetric(0.02);

        }

      

    }
}

/*
Add-In四个放置位置：

All Users, Version Independent
Windows 7 - %ALLUSERSPROFILE%\Autodesk\Inventor Addins\
Windows XP - %ALLUSERSPROFILE%\Application Data\Autodesk\Inventor Addins\

All Users, Version Dependent
Windows 7 - %ALLUSERSPROFILE%\Autodesk\Inventor 2013\Addins\
Windows XP - %ALLUSERSPROFILE%\Application Data\Autodesk\Inventor 2013\Addins\

Per User, Version Dependent
Both Window 7 and XP - %APPDATA%\Autodesk\Inventor 2013\Addins\

Per User, Version Independent
Both Window 7 and XP - %APPDATA%\Autodesk\ApplicationPlugins

在决定将外接程序放置在何处时，有几件事情需要考虑。如果您选择的位置对所有用户都可用，则需要管理员权限才能安装外接程序。
在大多数情况下，计算机很少在多个用户之间共享，因此每个用户安装通常就足够了。

如果你计划为每个Inventor版本积极更新你的外接程序，那么让它依赖于版本会很好，这样用户就只能访问为Inventor版本编写并测试过的外接程序。
由于为使API向上兼容付出了大量努力，所以应该能够使用Inventor的新版本运行较老的插件。
正因为如此，您可以提供和外接程序，而不必将其绑定到特定的版本，假设它将在Inventor的新版本发布时继续运行。

此外，因为你可以在.addin文件中指定插件兼容的版本，所以你仍然可以使用独立于版本的.addin位置，并通过.addin文件控制版本。
通过Autodesk Exchange Store提供的应用程序就利用了这一点，并安装在“每个用户，独立版本”的位置。
 */


