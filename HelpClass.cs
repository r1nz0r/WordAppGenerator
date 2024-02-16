using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace AppendixGenConsole
{
    internal static class HelpClass
    {
        public static List<string> SortListByElevations(List<string> unsortedList)
        {
            

            List<string> strlPos = new List<string>();
            List<string> strlNeg = new List<string>();

            foreach (var s in unsortedList)
            {
                if (s.Contains("-")) strlNeg.Add(s);
                else strlPos.Add(s);
            }

            strlPos.Sort(new NaturalStringComparer());
            strlNeg.Sort(new NaturalStringComparer());            
            //strlNeg.Reverse();

            List<string> sortedList = new List<string>();
            sortedList.AddRange(strlNeg);
            sortedList.AddRange(strlPos);

            return sortedList;
        }
        public static void SortListByElevationsFRS(ref List<string> unsortedList)
        {
            List<string> strlPos = new List<string>();
            List<string> strlNeg = new List<string>();


            foreach (var s in unsortedList)
            {
                if (s.Contains("-")) strlNeg.Add(s);
                else strlPos.Add(s);
            }

            strlPos.Sort(new NaturalStringComparer());
            strlNeg.Sort(new NaturalStringComparer());
            strlNeg.Reverse();
            int numOfDataPerElevation = 3; // X, Y, Z оси


            for (int i = 0; i < strlNeg.Count / numOfDataPerElevation; i++)
            {
                Swap<string>(strlNeg, i * numOfDataPerElevation, i * numOfDataPerElevation + 2);
            }

            unsortedList = new List<string>();
            unsortedList.AddRange(strlNeg);
            unsortedList.AddRange(strlPos);
        }
        public static void Swap<T>(IList<T> list, int index1, int index2)
        {
            T temp = list[index1];
            list[index1] = list[index2];
            list[index2] = temp;
        }
        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);
            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBilinear;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }

            }

            return destImage;
        }
        public static Bitmap ResizeImage(Image image, float downScaleFactor)
        {
            int width = (int)(image.Width * downScaleFactor);
            int height = (int)(image.Height * downScaleFactor);
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);
            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }

            }

            return destImage;
        }
        public static void RemoveExtentionFromName(ref List<string> _listOfNames)
        {
            for (int index = 0; index < _listOfNames.Count; index++)
            {
                _listOfNames[index] = _listOfNames[index].Remove(_listOfNames[index].LastIndexOf("."));
            }
        }
    }

    internal static class NativeMethods
    {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
        public static extern int StrCmpLogicalW(string psz1, string psz2);
    }

    public sealed class NaturalStringComparer : IComparer<string>
    {
        public int Compare(string a, string b)
        {
            return NativeMethods.StrCmpLogicalW(a, b);
        }
    }
}
