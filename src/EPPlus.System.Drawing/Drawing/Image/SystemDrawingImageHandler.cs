﻿using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Image;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Drawing;

namespace OfficeOpenXml.SystemDrawing.Image
{
    /// <summary>
    /// 
    /// </summary>
    public class SystemDrawingImageHandler : IImageHandler
    {
        /// <summary>
        /// 
        /// </summary>
        public SystemDrawingImageHandler()
        {
            if(IsWindows())
            {
                SupportedTypes= new HashSet<ePictureType>() { ePictureType.Bmp, ePictureType.Jpg, ePictureType.Gif, ePictureType.Png, ePictureType.Tif, ePictureType.Emf, ePictureType.Wmf };
            }
            else
            {
                SupportedTypes = new HashSet<ePictureType>() { ePictureType.Bmp, ePictureType.Jpg, ePictureType.Gif, ePictureType.Png, ePictureType.Tif };
            }
        }

        private bool IsWindows()
        {
            if(Environment.OSVersion.Platform == PlatformID.Unix ||
#if NET
               Environment.OSVersion.Platform == PlatformID.Other ||
#endif
               Environment.OSVersion.Platform == PlatformID.MacOSX)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        public HashSet<ePictureType> SupportedTypes
        {
            get;
        } 
        /// <summary>
        /// 
        /// </summary>
        public Exception LastException { get; private set; }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="image"></param>
        /// <param name="type"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <param name="horizontalResolution"></param>
        /// <param name="verticalResolution"></param>
        /// <returns></returns>

        public bool GetImageBounds(MemoryStream image, ePictureType type, out double width, out double height, out double horizontalResolution, out double verticalResolution)
        {
            try
            {
                var bmp = new Bitmap(image);
                width = bmp.Width;
                height = bmp.Height;
                horizontalResolution = bmp.HorizontalResolution;
                verticalResolution = bmp.VerticalResolution;
                return true;
            }
            catch(Exception ex)
            {
                width = 0;
                height = 0;
                horizontalResolution = 0;
                verticalResolution = 0;
                LastException = ex;
                return false;
            }
        }
        bool? _validForEnvironment = null;
      
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool ValidForEnvironment()
        {
            if (_validForEnvironment.HasValue == false)
            {
                try
                {
                    var g = Graphics.FromHwnd(IntPtr.Zero); //Fails if no gdi.
                    _validForEnvironment = true;
                }
                catch
                {
                    _validForEnvironment = false;
                }
            }
            return _validForEnvironment.Value;
        }
    }
}
