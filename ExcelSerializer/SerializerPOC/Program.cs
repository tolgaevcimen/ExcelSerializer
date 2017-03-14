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
                Listy = new List<DownloadLink>
                {
                    new DownloadLink
                    {
                        LinkName = "1",
                        Recipe = "11",
                        CreateDate = DateTime.Now.AddDays(8),
                        EndDate = DateTime.Now,
                    },
                    new DownloadLink
                    {
                        LinkName = "2",
                        Recipe = "22",
                        CreateDate = DateTime.Now.AddDays(8),
                        EndDate = DateTime.Now,
                    },
                    new DownloadLink
                    {
                        LinkName = "3",
                        Recipe = "33",
                        CreateDate = DateTime.Now.AddDays(8),
                        EndDate = DateTime.Now,
                    }
                },
                Listy2 = new List<DownloadLink>
                {
                    new DownloadLink
                    {
                        LinkName = "111",
                        Recipe = "1111",
                        CreateDate = DateTime.Now.AddDays(8),
                        EndDate = DateTime.Now,
                    },
                    new DownloadLink
                    {
                        LinkName = "222",
                        Recipe = "2222",
                        CreateDate = DateTime.Now.AddDays(8),
                        EndDate = DateTime.Now,
                    },
                    new DownloadLink
                    {
                        LinkName = "333",
                        Recipe = "3333",
                        CreateDate = DateTime.Now.AddDays(8),
                        EndDate = DateTime.Now,
                    }
                }
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
                Listy = new List<DownloadLink>();
                Listy2 = new List<DownloadLink>();
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
            public List<DownloadLink> Listy { get; set; }
            public List<DownloadLink> Listy2 { get; set; }
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
