using EPiServer.Framework.Blobs;
using System;

namespace EPiServer.ConnectForSharePoint
{
    public class SharePointBlobProvider : BlobProvider
    {
        public override Blob CreateBlob(Uri id, string extension)
        {
            return new SharePointBlob(id);
        }

        public override void Delete(Uri id)
        {
            throw new NotImplementedException();
        }

        public override Blob GetBlob(Uri id)
        {
            return new SharePointBlob(id);
        }
    }
}
