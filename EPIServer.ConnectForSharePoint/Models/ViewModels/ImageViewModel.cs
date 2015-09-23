using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPiServer.ConnectForSharePoint.Models.ViewModels
{
    public class ImageViewModel
    {
        /// <summary>
        /// Gets or sets the URL to the image.
        /// </summary>
        public string Url { get; set; }

        /// <summary>
        /// Gets or sets the name of the image.
        /// </summary>
        public string Name { get; set; }
    }
}