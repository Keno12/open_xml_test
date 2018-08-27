using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test1
{
    public class TOffice
    {
        public String Name { get; set; }
        public String NameEn { get; set; }
        public String Telephone { get; set; }
        public String Fax { get; set; }
        public String Address { get; set; }
        public String AddressEn { get; set; }

        public List<TOffice> GetListTOffice()
        {
            List<TOffice> list = new List<TOffice>()
            {
                new TOffice { Name="武汉办公室" ,NameEn="Wuhan Office",Address="武汉市江汉区建设大道568号新世界国贸大厦I座50层",AddressEn="50/F, New World International Trade Tower, 568 Jianshe Road,Wuhan 430022, P.R.China", Telephone="86-571-5692 1222", Fax="86-571-5692 1222"},
                new TOffice { Name="武汉办公室" ,NameEn="Wuhan Office",Address="武汉市江汉区建设大道568号新世界国贸大厦I座50层",AddressEn="50/F, New World International Trade Tower, 568 Jianshe Road,Wuhan 430022, P.R.China", Telephone="86-571-5692 1222", Fax="86-571-5692 1222"},
                new TOffice { Name="武汉办公室" ,NameEn="Wuhan Office",Address="武汉市江汉区建设大道568号新世界国贸大厦I座50层",AddressEn="50/F, New World International Trade Tower, 568 Jianshe Road,Wuhan 430022, P.R.China", Telephone="86-571-5692 1222", Fax="86-571-5692 1222"},
                new TOffice { Name="武汉办公室" ,NameEn="Wuhan Office",Address="武汉市江汉区建设大道568号新世界国贸大厦I座50层",AddressEn="50/F, New World International Trade Tower, 568 Jianshe Road,Wuhan 430022, P.R.China", Telephone="86-571-5692 1222", Fax="86-571-5692 1222"},
                new TOffice { Name="武汉办公室" ,NameEn="Wuhan Office",Address="武汉市江汉区建设大道568号新世界国贸大厦I座50层",AddressEn="50/F, New World International Trade Tower, 568 Jianshe Road,Wuhan 430022, P.R.China", Telephone="86-571-5692 1222", Fax="86-571-5692 1222"},
                new TOffice { Name="武汉办公室" ,NameEn="Wuhan Office",Address="武汉市江汉区建设大道568号新世界国贸大厦I座50层",AddressEn="50/F, New World International Trade Tower, 568 Jianshe Road,Wuhan 430022, P.R.China", Telephone="86-571-5692 1222", Fax="86-571-5692 1222"},
                new TOffice { Name="武汉办公室" ,NameEn="Wuhan Office",Address="武汉市江汉区建设大道568号新世界国贸大厦I座50层",AddressEn="50/F, New World International Trade Tower, 568 Jianshe Road,Wuhan 430022, P.R.China", Telephone="86-571-5692 1222", Fax="86-571-5692 1222"},
                new TOffice { Name="武汉办公室" ,NameEn="Wuhan Office",Address="武汉市江汉区建设大道568号新世界国贸大厦I座50层",AddressEn="50/F, New World International Trade Tower, 568 Jianshe Road,Wuhan 430022, P.R.China", Telephone="86-571-5692 1222", Fax="86-571-5692 1222"},
                new TOffice { Name="武汉办公室" ,NameEn="Wuhan Office",Address="武汉市江汉区建设大道568号新世界国贸大厦I座50层",AddressEn="50/F, New World International Trade Tower, 568 Jianshe Road,Wuhan 430022, P.R.China", Telephone="86-571-5692 1222", Fax="86-571-5692 1222"},

                new TOffice { Name="武汉办公室" ,NameEn="Wuhan Office",Address="武汉市江汉区建设大道568号新世界国贸大厦I座50层",AddressEn="50/F, New World International Trade Tower, 568 Jianshe Road,Wuhan 430022, P.R.China", Telephone="86-571-5692 1222", Fax="86-571-5692 1222"},
                new TOffice { Name="武汉办公室" ,NameEn="Wuhan Office",Address="武汉市江汉区建设大道568号新世界国贸大厦I座50层",AddressEn="50/F, New World International Trade Tower, 568 Jianshe Road,Wuhan 430022, P.R.China", Telephone="86-571-5692 1222", Fax="86-571-5692 1222"},
                new TOffice { Name="武汉办公室" ,NameEn="Wuhan Office",Address="武汉市江汉区建设大道568号新世界国贸大厦I座50层",AddressEn="50/F, New World International Trade Tower, 568 Jianshe Road,Wuhan 430022, P.R.China", Telephone="86-571-5692 1222", Fax="86-571-5692 1222"},
                new TOffice { Name="武汉办公室" ,NameEn="Wuhan Office",Address="武汉市江汉区建设大道568号新世界国贸大厦I座50层",AddressEn="50/F, New World International Trade Tower, 568 Jianshe Road,Wuhan 430022, P.R.China", Telephone="86-571-5692 1222", Fax="86-571-5692 1222"},
                new TOffice { Name="武汉办公室" ,NameEn="Wuhan Office",Address="武汉市江汉区建设大道568号新世界国贸大厦I座50层",AddressEn="50/F, New World International Trade Tower, 568 Jianshe Road,Wuhan 430022, P.R.China", Telephone="86-571-5692 1222", Fax="86-571-5692 1222"},
                new TOffice { Name="武汉办公室" ,NameEn="Wuhan Office",Address="武汉市江汉区建设大道568号新世界国贸大厦I座50层",AddressEn="50/F, New World International Trade Tower, 568 Jianshe Road,Wuhan 430022, P.R.China", Telephone="86-571-5692 1222", Fax="86-571-5692 1222"},
                new TOffice { Name="武汉办公室" ,NameEn="Wuhan Office",Address="武汉市江汉区建设大道568号新世界国贸大厦I座50层",AddressEn="50/F, New World International Trade Tower, 568 Jianshe Road,Wuhan 430022, P.R.China", Telephone="86-571-5692 1222", Fax="86-571-5692 1222"},
               
            };
            return list;
        }
    }
}
