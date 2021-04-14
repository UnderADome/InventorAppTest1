﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;  //交互性操作
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Inventor;
using Application = Inventor.Application;
using System.Diagnostics;
using System.IO;
using System.Collections;

namespace InventroAppTest1
{
    class Class1
    {

        static void Main(string[] args)
        {

            GetInventorApplication();

            OpenDrawingDocuments();

            Console.WriteLine("即将关闭Inventor");
            if (inventorApp != null)
                inventorApp.Quit();

            //防止闪退
            System.Console.ReadKey();
        }

        #region 获取Inventor实例
        private static Inventor.Application inventorApp = null;
        private static void GetInventorApplication()
        {
            
            try
            {
                //Marshal.GetActiveObject 从运行对象表（ROT）获取指定对象的运行实例

                //获取一个Inventor的参考
                inventorApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
                Console.WriteLine("查找到可用的实例");
            }
            catch
            {
                try
                {
                    Console.WriteLine("没有正常连接到Inventor");
                    //创建新实例
                    ///在线认证需要输入172.20.133.35，如果不能正常访问外网，则Inventor打不开，会导致创建Inventor实例失败
                    Type inventorAppType = Type.GetTypeFromProgID("Inventor.Application");
                    Console.WriteLine(inventorAppType.GUID.ToString());
                    Console.WriteLine("重新创建一个Inventor实例");
                    inventorApp = Activator.CreateInstance(inventorAppType) as Application;
                    Console.WriteLine("创建新Inventor实例完毕");
                }
                catch
                {
                    Console.WriteLine("创建新实例失败");
                    Process.GetCurrentProcess().Close();    //Diagnostics.Process 获取新的Process组件并将其与当前活动的进程关联
                }
            }
            finally
            {
                
                if (inventorApp != null)
                {
                    Console.WriteLine("生成了Inventor实例并显示Inventor");
                    inventorApp.ApplicationEvents.OnQuit += ApplicationEvents_OnQuit;
                    //inventorApp.WindowState = WindowsSizeEnum.kMaximize;  //将Inventor窗口大小设置为最大窗口
                    inventorApp.WindowState = WindowsSizeEnum.kNormalWindow;
                    inventorApp.Visible = false;
                    inventorApp.SilentOperation = true;
                }
            }
        }

        private static void ApplicationEvents_OnQuit(EventTimingEnum BeforeOrAfter, NameValueMap context, out HandlingCodeEnum HandlingCode)
        {
            Console.WriteLine("触发OnQuit事件");
            //HandlingCodeEnum 从某些事件中返回的代码

            //OnQuit 当Inventor被关闭时通知client 
            inventorApp.ApplicationEvents.OnQuit -= ApplicationEvents_OnQuit;
            HandlingCode = HandlingCodeEnum.kEventHandled;  //kEventHandled绕过本地行为  /更多详见后面的备注
            //inventorApp = null;  //仅仅用这个不能关闭Inventor进程，必须要用杀进程的方法
            ///inventorApp.Quit();
            // Process.GetCurrentProcess().Kill();
            Console.WriteLine("关闭Inventor，结束操作");
        }

        #endregion


        #region 操作工程图

        //打开文件夹中的工程图
        private static void OpenDrawingDocuments()
        {
            ArrayList fileInfoArray = new ArrayList();
            Console.WriteLine("打开工程图文件夹");
            string DrawingSheetPath = @"C:\Users\14530\Desktop\DrawingSheets";
            try
            {
                
                //得到文件夹中的文件
                string[] files = Directory.GetFiles(DrawingSheetPath);
                foreach(string file in files)
                {
                    string exname = file.Substring(file.LastIndexOf(".") + 1); //得到后缀名
                    if (".dwg".IndexOf(file.Substring(file.LastIndexOf(".")+1)) > -1) //".dwg|.dfg"
                    {
                        Console.WriteLine("file文件名："+file);
                        FileInfo fileInfo = new FileInfo(file);  //https://blog.csdn.net/liubai123/article/details/9858725 操作文件\文件夹
                        fileInfoArray.Add(fileInfo);
                        Console.WriteLine("Name文件名："+fileInfo.Name);
                        Console.WriteLine("FullName完整目录：" + fileInfo.FullName);
                        Console.WriteLine("DirectoryName目录的完整路径" + fileInfo.DirectoryName);

                        //对工程图进行操作，使用工程图API读取工程图中的信息：
                        //图号、文件名称、自然张数、A1张数
                        //按照公司的图纸目录格式列出汇总栏。统计出自然张数和A1张数  //对于inventor文件操作 https://blog.csdn.net/qq_43006346/article/details/104596572
        
                            Console.WriteLine("读取工程图");
                            DrawingDocument drawingDocument = inventorApp.Documents.Open(file, false) as DrawingDocument; //这里需要注意的是前面要用上inventorApp
                            if (drawingDocument.IsInventorDWG == true)
                            {
                                Console.WriteLine("是Inventor的工程图");
                                Console.WriteLine("DrawingBOMs.Count: " + drawingDocument.DrawingBOMs.Count.ToString());
                                //Console.WriteLine("DrawingBOMs.Item.1: " + drawingDocument.DrawingBOMs[1].ToString());
                                Console.WriteLine("Open: " + drawingDocument.Open.ToString());
                            //PropertySet DesignInfo = drawingDocument.PropertySets["Design Tracking Properties"];
                            //Property property = DesignInfo["Part Number"];

                            //["Property Type list"] ["Design Tracking Properties"] ["Inventor User Defined Properties"] ["Edit Property Fields"]
                            Console.WriteLine("PropertySets.Count = " + drawingDocument.PropertySets.Count);
                            Console.WriteLine(drawingDocument.PropertySets["{F29F85E0-4FF9-1068-AB91-08002B27B3D9}"] == null);  //必须要用GUID来表示 https://blog.csdn.net/beihuanlihe130/article/details/107352288
                            /*
                            for (int i=0; i<drawingDocument.PropertySets.Count; i++)  //行不通，这里能够读到count的值，但是不能通过下标的方式去具体读到其中的属性
                            {
                                Console.Write("打印第"+i+"个Set : ");
                                Console.WriteLine(drawingDocument.PropertySets[i] == null);
                                Console.WriteLine(drawingDocument.PropertySets[i].Name);
                            }
                            */

                            PropertySet DesignInfo1 = drawingDocument.PropertySets["{F29F85E0-4FF9-1068-AB91-08002B27B3D9}"];  //这里存在问题。是按照名称，还是按照Item序号？可能需要根据工程图的具体情况来判断
                                Console.WriteLine("DesignInfo1: "+(DesignInfo1 == null));
                                Console.WriteLine("DesignInfo1.Title: " +DesignInfo1.ToString());
                            Property property = DesignInfo1.ItemByPropId[(int)PropertiesForSummaryInformationEnum.kTitleSummaryInformation];
                            Console.WriteLine("Title : "+property.Value);

                                //Property property_0 = DesignInfo1["Sheet number"];
                                //Console.WriteLine("PropertySet.Property[0].value: "+property_0.Value);

                            }
                       /*
                        catch
                        {
                            Console.WriteLine("读工程图失败");
                        }
                        */



                    }
                }
            }
            catch
            {
                Console.WriteLine("读取文件夹失败");
            }
            


        }
        
        //判断Sheet的Application与本次设置的全局Application是否是一样的
        //需要用到Sheet.Application == inventorApp进行判断
        private static void IsTheSameApplication()
        {
            
        }


        #endregion
    }
}


/**
 * 
 * HandlingCodeEnum.kEventHandled / kEventCanceled / kEventNotHandled
 * kEventCanceled 513    使用Inventor原生的“取消”和“失败”，则返回该代码
 * kEventHandled  514    Inventor绕过其本地行为，则返回该代码
 * kEventNotHandled 515  如果Inventor继续其本机行为，则返回该代码 
 * 
 * 
 * Inventor 工程图标题栏中“比例”和“重量”解决方法
 * http://blog.sina.com.cn/s/blog_721ff0820100q8ou.html
 * Inventor文件中保存自定义数据 - 1
 * http://www.voidcn.com/article/p-xjkdclhs-pv.html
 * 
 * Summary Information, {F29F85E0-4FF9-1068-AB91-08002B27B3D9}
 * Document Summary Information, {D5CDD502-2E9C-101B-9397-08002B2CF9AE}
 * Design Tracking Properties, {32853F0F-3444-11D1-9E93-0060B03C1CA6}
 * User Defined Properties, {D5CDD505-2E9C-101B-9397-08002B2CF9AE}
 */