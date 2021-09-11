using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Security.Policy;
using System.Windows.Media.Imaging;
using System.IO;
using System.Drawing;
using AutoDimPlugin.Properties;

namespace AutoDimPlugin
{
    class Plugin:IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            //Create Tab
            application.CreateRibbonTab("AutoDim Tools");

            //Create Panel
            RibbonPanel panelAxesCols =  application.CreateRibbonPanel("AutoDim Tools", "Axes & Columns");

            //Create Button
            string path = Assembly.GetExecutingAssembly().Location;
            PushButtonData btnDataAxesDim = new PushButtonData("btnAxesCol", "Auto Dimention", path, "AutoDimPlugin.AxesColumnDimensions");

            //
            PushButton pushBtn1 =  panelAxesCols.AddItem(btnDataAxesDim) as PushButton;
            //Uri uri1 = new Uri(@"C:\Users\obayo\OneDrive\Assignments\Revit\AutoDimPlugin\icon1.jpg");
            //BitmapImage img1 = new BitmapImage(uri1);
            BitmapImage img1 = NewBitmapImage(Assembly.GetExecutingAssembly(), "icon.png");
            pushBtn1.LargeImage = img1;
            return Result.Succeeded;
        }
        private BitmapImage NewBitmapImage(Assembly a,string imageName)
        {
            //to read from an external file:
            //return new BitmapImage(new Uri(
            //  Path.Combine(_imageFolder, imageName)));

            System.IO.Stream s = a.GetManifestResourceStream(
                "AutoDimPlugin." + imageName);

            BitmapImage img = new BitmapImage();

            img.BeginInit();
            img.StreamSource = s;
            img.EndInit();

            return img;
        }
    }
}
