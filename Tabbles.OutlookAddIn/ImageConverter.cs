using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using stdole;
using System.Drawing;

namespace Tabbles.OutlookAddIn
{
    class ImageConverter : AxHost
    {
        public ImageConverter() : base(string.Empty)
        {
        }

        public static IPictureDisp GetPictureDisp(Image image)
        {
            return (IPictureDisp)GetIPictureDispFromPicture(image);
        }
    }
}
