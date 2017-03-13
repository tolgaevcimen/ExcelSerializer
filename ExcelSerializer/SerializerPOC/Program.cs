using ExcelSerializer;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SerializerPOC
{
    class Program
    {
        static void Main ( string[] args )
        {
            var x = new DownloadLink
            {
                LinkName = "testlink",
                Recipe = "dafgf",
                CreateDate = null,
                EndDate = DateTime.Now,
                AnotherLink = new DownloadLink
                {
                    LinkName = "testlink2",
                    Recipe = "123asdfdafgf",
                    CreateDate = DateTime.Now.AddDays(8),
                    EndDate = DateTime.Now,
                },
                AnotherLink2 = new DownloadLink
                {
                    LinkName = "testlink2",
                    Recipe = "123asdfdafgf",
                    CreateDate = DateTime.Now.AddDays(8),
                    EndDate = DateTime.Now,
                },
                //Listy = new List<string>
                //{
                //    "1", "2", "3"
                //}
            };

            var ex = new ExcelSerializer.ExcelSerializer();
            ex.Serialize(x, "text.xlsx");

            Process.Start("text.xlsx");
        }

        public partial class DownloadLink
        {
            public DownloadLink ()
            {
                this.DownloadLink_logs = new HashSet<DownloadLink_log>();
                Listy = new List<string>();
            }

            public System.Guid Id { get; set; }
            public string LinkName { get; set; }
            public string Recipe { get; set; }
            public Nullable<System.DateTime> CreateDate { get; set; }
            public Nullable<System.DateTime> EndDate { get; set; }
            public string EncryptionKey { get; set; }
            public Nullable<System.Guid> Salt { get; set; }
            public string PermittedIpList { get; set; }
            public string OwnerLogonName { get; set; }
            public byte Status { get; set; }
            public DownloadLink AnotherLink { get; set; }
            public DownloadLink AnotherLink2 { get; set; }
            public string SentUsersByDesktopAgent { get; set; }
            public List<string> Listy { get; set; }
            public virtual ICollection<DownloadLink_log> DownloadLink_logs { get; set; }
        }
        public partial class DownloadLink_log
        {
            public long Id { get; set; }
            public System.Guid LinkId { get; set; }
            public string Path { get; set; }
            public System.DateTime Time { get; set; }
            public byte Type { get; set; }
            public string ClientIp { get; set; }
            public Nullable<long> FileSize { get; set; }

            public virtual DownloadLink DownloadLinks { get; set; }
        }
    }
}
