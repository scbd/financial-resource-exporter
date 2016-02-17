using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace financial_reporting
{
    class Program
    {
        static void Main(string[] args)
        {
            string templatePath = args.Length>0 ? args[0] : Path.Combine(Environment.CurrentDirectory, "template.xlsm");
            
            if(!File.Exists(templatePath))
            {
                Console.Error.WriteLine("File not found: {0}", templatePath);
                return;
            }

            getTerm("ca");//dummy just for console output;

            var reports = getIndexedRecords().OrderBy(o=>o.name.ToLower()).ToArray();
            
			Console.WriteLine("{0} records found", reports.Count());

			foreach(var r in reports)
				Console.WriteLine("{0} - {1}", r.government, r.name);

            Excel.Application xlApp;
            Excel.Workbook    xlWorkBook;

            Console.WriteLine();
			Console.WriteLine("Loading Excel template {0}", Path.GetFileName(templatePath));

            using(new XLDisposable(xlApp = new Excel.Application()))
            using(new XLDisposable(xlWorkBook = (Excel.Workbook)xlApp.Workbooks.Open(templatePath, ReadOnly:true)))
            {
                var oSheetNames = xlWorkBook.Worksheets.OfType<Excel.Worksheet>().Select(o=>o.Name).ToList();

                Excel.Worksheet xlWorkSheetTemplate = (Excel.Worksheet)xlWorkBook.Worksheets["{{template}}"];

                Excel.Worksheet xlWorkSheetMenu = (Excel.Worksheet)xlWorkBook.Worksheets["MENU"];

				var bindings = getBindings(xlWorkSheetTemplate);
                int row = 3;

                foreach(var report in reports)
				{
                    Console.WriteLine();

                    if(oSheetNames.Contains(report.name)) {
                        Console.WriteLine("SKIP: Sheet already exists: {0}", report.name);
                        continue;
                    }

                    oSheetNames.Add(report.name);

					xlWorkSheetTemplate.Copy(Before:xlWorkSheetTemplate);

                    while(!string.IsNullOrWhiteSpace((string)xlWorkSheetMenu.Cells[row, 2].Value2))
                        row++; // next empty row

                    xlWorkSheetMenu.Cells[row, 2] = report.name;

					Excel.Worksheet xlWorkSheet = xlWorkBook.Worksheets[xlWorkSheetTemplate.Index-1];

					var values = mapValues(report.record);

					xlWorkSheet.Name = report.name;

                    populate(xlWorkSheet, values, bindings);
				}

                xlWorkSheetTemplate.Move(After:(Excel.Worksheet)xlWorkBook.Worksheets[xlWorkBook.Worksheets.Count]);

                xlWorkSheetMenu.Activate();

                var outPath = Path.Combine(Path.GetDirectoryName(templatePath), string.Format("out-{0:yyyy-MM-dd-HH-mm-ss}", DateTime.Now)+Path.GetExtension(templatePath));

                xlWorkBook.SaveAs(outPath);
            }
        }

        //==============================
        //
        //==============================
		private static IEnumerable<Binding> getBindings(Excel.Worksheet sheet)
		{
			List<Binding> bindings = new List<Binding>();

			Regex placeHolder = new Regex(@"^\s*{{(.*?)\}\}\s*$");
            Regex pathTarget  = new Regex(@"^(.+?)=(.+?)$");

            var i=0;
            var count = sheet.UsedRange.Rows.Count * sheet.UsedRange.Columns.Count;

            foreach (Excel.Range cell in sheet.UsedRange)
            {

    			Console.Write("\rAnalyzing bindings {0}%    ", ++i*100/count);

                string cellText = string.Format("{0}", cell.Value2 ?? "");

                if (!placeHolder.IsMatch(cellText))
					continue;

                string binding   = placeHolder.Replace(cellText, "$1");
                string condition = null;

                if(pathTarget.IsMatch(binding)) {
                    condition = pathTarget.Replace(binding, "$2");
                    binding   = pathTarget.Replace(binding, "$1");
                }

				bindings.Add(new Binding() {
					col       = cell.Column,
					row       = cell.Row,
					binding   = binding,
					condition = condition
				});
			}

    		Console.WriteLine();

			return bindings;
		}

		class Binding
		{
			public int    col;
			public int    row;
			public string binding;
			public string condition;
		}

        //==============================
        //
        //==============================
        private static string getReportName(JObject report) {
        
            var name   = (string)report["government"]["title"];

            return truncate(name, 30).Trim();
        }

        //==============================
        //
        //==============================
        private static void populate(Excel.Worksheet sheet, IDictionary<string, object> values, IEnumerable<Binding> bindings)
        {
            var i=0;
            var count = bindings.Count();

            foreach (var binding in bindings)
            {
    			Console.Write("\rPopulating {0} {1}%   ", sheet.Name, (++i*100)/count);

				object value = "";

				if(values.ContainsKey(binding.binding))
					value = values[binding.binding];

				if(binding.condition!=null)
				{
					var valueText = string.Format("{0}", value);

					value = StringComparer.InvariantCultureIgnoreCase.Compare(binding.condition, valueText)==0 ? (object)1 : "";
				}

				sheet.Cells[binding.row, binding.col].Value2 = value;
            }

    		Console.WriteLine();
        }

        //==============================
        //
        //==============================
		private static IDictionary<string, object> mapValues(JObject record) 
		{
			JTokenType [] types = new [] {
				//JTokenType.None,
				JTokenType.Object,
				JTokenType.Array,
				//JTokenType.Constructor,
				//JTokenType.Property,
				//JTokenType.Comment,
				JTokenType.Integer,
				JTokenType.Float,
				JTokenType.String,
				JTokenType.Boolean,
				JTokenType.Null,
				JTokenType.Undefined,
				JTokenType.Date,
				//JTokenType.Raw,
				//JTokenType.Bytes,
				//JTokenType.Guid,
				//JTokenType.Uri,
				//JTokenType.TimeSpan
			};

			Dictionary<string, object> map = new Dictionary<string,object>();

			var qValues = record.Descendants().Where(o => types.Contains(o.Type)).ToArray();

            var i=0;

			foreach(var jValue in qValues){
                Console.Write("\rMapping report values {0}%   ", ++i*100/qValues.Length);
				map.Add(jValue.Path, toXlValue(jValue));
            }

            Console.WriteLine();

			return map;
		}

        //==============================
        //
        //==============================
        private static object toXlValue(JToken jValue)
        {
                 if(jValue==null) return "";
            else if(jValue.Type == JTokenType.Integer)   return jValue.Value<int>();
            else if(jValue.Type == JTokenType.String)    return jValue.Value<string>();
            else if(jValue.Type == JTokenType.Float)     return jValue.Value<double>();
            else if(jValue.Type == JTokenType.Date)      return jValue.Value<DateTime>();
            else if(jValue.Type == JTokenType.Array)     return ((JArray)jValue).Count;
            else if(jValue.Type == JTokenType.Object)    return 1;
            else if(jValue.Type == JTokenType.Null)      return "";
            else if(jValue.Type == JTokenType.Undefined) return "";
            else if(jValue.Type == JTokenType.Boolean)
            { 
                if(jValue.Value<bool>())
                    return 1; 
                return 0; 
            }

            throw new ArgumentOutOfRangeException("Unsuported value type:"+jValue.Type.ToString());
        }


        //==============================
        //
        //==============================
        static JObject normalizeRecord(JObject record) {

            // Replace term by term data;

			normalizeTerms(record);

            var baselineData             = record["internationalResources"]["baselineData"];
            var progressData             = record["internationalResources"]["progressData"];
            var domesticExpendituresData = record["domesticExpendituresData"];
            var fundingNeedsData         = record["fundingNeedsData"];

            baselineData            ["baselineFlows"  ] = __keyBy(baselineData            ["baselineFlows"  ], "year");
            progressData            ["progressFlows"  ] = __keyBy(progressData            ["progressFlows"  ], "year");
            domesticExpendituresData["expenditures"   ] = __keyBy(domesticExpendituresData["expenditures"   ], "year");
            fundingNeedsData        ["annualEstimates"] = __keyBy(fundingNeedsData        ["annualEstimates"], "year");

            baselineData            ["odaCategories"] = __keyBy(baselineData["odaCategories"], "identifier");
            baselineData            ["odaoofActions"] = __keyBy(baselineData["odaoofActions"], "identifier");
            baselineData            ["otherActions" ] = __keyBy(baselineData["otherActions" ], "identifier");
            
            
            // Aggregate sources


            var allDomesticSources      = record["nationalPlansData"]["domesticSources"] ?? new JArray();
            var allInternationalSources = record["nationalPlansData"]["internationalSources"] ?? new JArray();

            var domesticSources      = (record["nationalPlansData"]["domesticSources"]      = new JObject());
            var internationalSources = (record["nationalPlansData"]["internationalSources"] = new JObject());

            domesticSources     ["sources"] = allDomesticSources;
            internationalSources["sources"] = allInternationalSources;

            domesticSources["amount2014"] = (decimal)0;
            domesticSources["amount2015"] = (decimal)0;
            domesticSources["amount2016"] = (decimal)0;
            domesticSources["amount2017"] = (decimal)0;
            domesticSources["amount2018"] = (decimal)0;
            domesticSources["amount2019"] = (decimal)0;
            domesticSources["amount2020"] = (decimal)0;

            internationalSources["amount2014"] = (decimal)0;
            internationalSources["amount2015"] = (decimal)0;
            internationalSources["amount2016"] = (decimal)0;
            internationalSources["amount2017"] = (decimal)0;
            internationalSources["amount2018"] = (decimal)0;
            internationalSources["amount2019"] = (decimal)0;
            internationalSources["amount2020"] = (decimal)0;

            foreach(var domesticSource in allDomesticSources) {
                if(domesticSource["amount2014"]!=null) domesticSources["amount2014"] = (decimal)domesticSources["amount2014"] + (decimal)domesticSource["amount2014"];
                if(domesticSource["amount2015"]!=null) domesticSources["amount2015"] = (decimal)domesticSources["amount2015"] + (decimal)domesticSource["amount2015"];
                if(domesticSource["amount2016"]!=null) domesticSources["amount2016"] = (decimal)domesticSources["amount2016"] + (decimal)domesticSource["amount2016"];
                if(domesticSource["amount2017"]!=null) domesticSources["amount2017"] = (decimal)domesticSources["amount2017"] + (decimal)domesticSource["amount2017"];
                if(domesticSource["amount2018"]!=null) domesticSources["amount2018"] = (decimal)domesticSources["amount2018"] + (decimal)domesticSource["amount2018"];
                if(domesticSource["amount2019"]!=null) domesticSources["amount2019"] = (decimal)domesticSources["amount2019"] + (decimal)domesticSource["amount2019"];
                if(domesticSource["amount2020"]!=null) domesticSources["amount2020"] = (decimal)domesticSources["amount2020"] + (decimal)domesticSource["amount2020"];
            }

            foreach(var internationalSource in allInternationalSources) {
                if(internationalSource["amount2014"]!=null) internationalSources["amount2014"] = (decimal)internationalSources["amount2014"] + (decimal)internationalSource["amount2014"];
                if(internationalSource["amount2015"]!=null) internationalSources["amount2015"] = (decimal)internationalSources["amount2015"] + (decimal)internationalSource["amount2015"];
                if(internationalSource["amount2016"]!=null) internationalSources["amount2016"] = (decimal)internationalSources["amount2016"] + (decimal)internationalSource["amount2016"];
                if(internationalSource["amount2017"]!=null) internationalSources["amount2017"] = (decimal)internationalSources["amount2017"] + (decimal)internationalSource["amount2017"];
                if(internationalSource["amount2018"]!=null) internationalSources["amount2018"] = (decimal)internationalSources["amount2018"] + (decimal)internationalSource["amount2018"];
                if(internationalSource["amount2019"]!=null) internationalSources["amount2019"] = (decimal)internationalSources["amount2019"] + (decimal)internationalSource["amount2019"];
                if(internationalSource["amount2020"]!=null) internationalSources["amount2020"] = (decimal)internationalSources["amount2020"] + (decimal)internationalSource["amount2020"];
            }

			//foreach(string p in record.Descendants().OrderBy(o=>o.Path).Select(o=>o.Path).Distinct())
			//	Console.WriteLine("{{{{{0}}}}}",p);

			//Console.WriteLine(record);

            return record;
        }

        //==============================
        //
        //==============================
		private static JObject normalizeTerms(JObject record)
		{
            var eTerms = record.Descendants().Where(o=>o.Type==JTokenType.Property && ((JProperty)o).Name=="identifier").Select(o=>o.Parent).ToArray();

            foreach(var eTerm in eTerms) {

                if(eTerm["identifier"].Type != JTokenType.String)
                {
                    if(eTerm["identifier"].Type != JTokenType.Null && eTerm["identifier"].Type != JTokenType.Undefined) {
                        Console.WriteLine("WARNING(normalizeTerms): Invalid identifier type {0} of `({2})` report\n{1}", eTerm["identifier"].Type, eTerm.Parent, getReportName(record));
                    }

                    continue;
                }

                var identifier = (string)eTerm["identifier"];
                var term = getTerm(identifier??"");

                if(term!=null)
                    eTerm["title"] = term.title;

                if(!string.IsNullOrWhiteSpace(identifier))
                    eTerm[identifier] =  eTerm.DeepClone(); // add identifier 'value' property 
            }

			return record;
		}

        //==============================
        //
        //==============================
        static recordInfo [] getIndexedRecords()
        {
			Console.WriteLine("Loading records...");

            WebClient wc = new WebClient();

            UriBuilder url = new UriBuilder("https://api.cbd.int/api/v2013/index");

            url.Query += string.Format("&q={0}",    Uri.EscapeUriString("schema_s:resourceMobilisation AND _state_s:public AND realm_ss:chm"));
            url.Query += string.Format("&fl={0}",   Uri.EscapeUriString("identifier_s,government_s"));
            url.Query += string.Format("&rows={0}", Uri.EscapeUriString("2000"));

            url.Query = url.Query.TrimStart('&', '?');
            
            var result = JObject.Parse(wc.DownloadString(url.ToString()));

            return result["response"]["docs"].Select(o=> new recordInfo() {
                identifier = (string)o["identifier_s"],
                government = (string)o["government_s"]
            }).ToArray();
        }

        class recordInfo {

            public string identifier;
            public string government;

            Term    mGovernment = null;
            JObject mRecord = null;

            public string name
            {
                get { 

                    if(mGovernment==null)
                        mGovernment = getTerm(government);

                    return truncate(mGovernment.title, 30);
                }
            }

            public JObject record
            {
                get { 

                    if(mRecord==null) {

                        Console.WriteLine("Loading report: {0}...", getTerm(government).title);

                        WebClient wc = new WebClient();
            
                        mRecord = normalizeRecord(JObject.Parse(wc.DownloadString("https://api.cbd.int/api/v2013/documents/"+identifier)));
                    }

                    return mRecord;
                }
            }
        }

        //==============================
        //
        //==============================
        static string truncate(string text, int len) 
        {
            return text.Substring(0, Math.Min(text.Length, len));
        }

        static SortedList<string, Term> termsCache;

        //==============================
        //
        //==============================
        static Term getTerm(string code) 
        {
            if(termsCache==null)
            {
				Console.WriteLine("Loading terms...");

				termsCache = new SortedList<string,Term>();

                foreach (var term in loadDomainTerms("countries")) termsCache.Add(term.identifier, MapCountry(term.identifier)); //TODO USE Special Countries mapping
                foreach(var term in loadDomainTerms("ISO-4217"))  termsCache.Add(term.identifier, term); //TODO USE Special Currencies mapping
                foreach(var term in loadDomainTerms("AB782477-9942-4C6B-B9F0-79A82915A069")) termsCache.Add(term.identifier, term);
                foreach(var term in loadDomainTerms("1FBEF0A8-EE94-4E6B-8547-8EDFCB1E2301")) termsCache.Add(term.identifier, term);
                foreach(var term in loadDomainTerms("33D62DA5-D4A9-48A6-AAE0-3EEAA23D5EB0")) termsCache.Add(term.identifier, term);
                foreach(var term in loadDomainTerms("6BDB1F2A-FDD8-4922-BB40-D67C22236581")) termsCache.Add(term.identifier, term);
                foreach(var term in loadDomainTerms("A9AB3215-353C-4077-8E8C-AF1BF0A89645")) termsCache.Add(term.identifier, term);
            }

            return termsCache.ContainsKey(code) ? termsCache[code] : null;
        }

        //==============================
        //
        //==============================
        static IEnumerable<Term> loadDomainTerms(string domain)
        {
            WebClient wc = new WebClient();

            wc.Headers[HttpRequestHeader.Accept] = "application/json";

            UriBuilder url = new UriBuilder("https://api.cbd.int/api/v2013/thesaurus/domains/"+domain+"/terms");

            var result = wc.DownloadString(url.ToString());

            JArray terms = JArray.Parse(result);

            return terms.Select(o=> new Term() { 
                identifier = (string)o["identifier"], 
                title      = (string)o["name"]
            });
        }

        class Term
        {
            public string identifier;
            public string title;
        }

        //==============================
        //
        //==============================
        static JObject __keyBy(JToken array, string key, Func<JToken,JToken> transform=null) 
        {

            JObject ret = new JObject();

			if(array!=null) {
				foreach(JToken item in array)
				{
                    if(item[key]==null) {
                        Console.WriteLine("WARNING(__keyBy): `{0}` is not specified {1}", key, getReportName((JObject)array.Root));
                        continue;
                    }


					string slot  = item[key].ToString();
					JToken value = item;

					if(transform!=null)
						value = transform(item);

					ret[slot] = value;
				}
			}

            return ret;
        }


        static SortedList<string, string> termOverrides = new SortedList<string, string>();

        private static Term MapCountry(string countryCode)
        {
            termOverrides = new SortedList<string, string>();

            termOverrides["af"] = "Afghanistan";
            termOverrides["al"] = "Albania";
            termOverrides["dz"] = "Algeria";
            termOverrides["ad"] = "Andorra";
            termOverrides["ao"] = "Angola";
            termOverrides["ag"] = "Antigua and Barbuda";
            termOverrides["ar"] = "Argentina";
            termOverrides["am"] = "Armenia";
            termOverrides["au"] = "Australia";
            termOverrides["at"] = "Austria";
            termOverrides["az"] = "Azerbaijan";
            termOverrides["bs"] = "Bahamas";
            termOverrides["bh"] = "Bahrain";
            termOverrides["bd"] = "Bangladesh";
            termOverrides["bb"] = "Barbados";
            termOverrides["by"] = "Belarus";
            termOverrides["be"] = "Belgium";
            termOverrides["bz"] = "Belize";
            termOverrides["bj"] = "Benin";
            termOverrides["bt"] = "Bhutan";
            termOverrides["bo"] = "Bolivia";
            termOverrides["ba"] = "Bosnia and Herzegovina";
            termOverrides["bw"] = "Botswana";
            termOverrides["br"] = "Brazil";
            termOverrides["bn"] = "Brunei Darussalam";
            termOverrides["bg"] = "Bulgaria";
            termOverrides["bf"] = "Burkina Faso";
            termOverrides["bi"] = "Burundi";
            termOverrides["cv"] = "Cape Verde";
            termOverrides["kh"] = "Cambodia";
            termOverrides["cm"] = "Cameroon";
            termOverrides["ca"] = "Canada";
            termOverrides["cf"] = "Central African Republic";
            termOverrides["td"] = "Chad";
            termOverrides["cl"] = "Chile";
            termOverrides["cn"] = "China";
            termOverrides["co"] = "Colombia";
            termOverrides["km"] = "Comoros";
            termOverrides["cg"] = "Congo";
            termOverrides["ck"] = "Cook Islands";
            termOverrides["cr"] = "Costa Rica";
            termOverrides["hr"] = "Croatia";
            termOverrides["cu"] = "Cuba";
            termOverrides["cy"] = "Cyprus";
            termOverrides["cz"] = "Czech Republic";
            termOverrides["ci"] = "Côte d'Ivoire";
            termOverrides["kp"] = "Korea, Democratic People's Republic of";
            termOverrides["cd"] = "Congo, Democratic Republic of the";
            termOverrides["dk"] = "Denmark";
            termOverrides["dj"] = "Djibouti";
            termOverrides["dm"] = "Dominica";
            termOverrides["do"] = "Dominican Republic";
            termOverrides["ec"] = "Ecuador";
            termOverrides["eg"] = "Egypt";
            termOverrides["sv"] = "El Salvador";
            termOverrides["gq"] = "Equatorial Guinea";
            termOverrides["er"] = "Eritrea";
            termOverrides["ee"] = "Estonia";
            termOverrides["et"] = "Ethiopia";
            termOverrides["eu"] = "European Union";
            termOverrides["fj"] = "Fiji";
            termOverrides["fi"] = "Finland";
            termOverrides["fr"] = "France";
            termOverrides["ga"] = "Gabon";
            termOverrides["gm"] = "Gambia";
            termOverrides["ge"] = "Georgia";
            termOverrides["de"] = "Germany";
            termOverrides["gh"] = "Ghana";
            termOverrides["gr"] = "Greece";
            termOverrides["gd"] = "Grenada";
            termOverrides["gt"] = "Guatemala";
            termOverrides["gn"] = "Guinea";
            termOverrides["gw"] = "Guinea-Bissau";
            termOverrides["gy"] = "Guyana";
            termOverrides["ht"] = "Haiti";
            termOverrides["va"] = "Holy See";
            termOverrides["hn"] = "Honduras";
            termOverrides["hu"] = "Hungary";
            termOverrides["is"] = "Iceland";
            termOverrides["in"] = "India";
            termOverrides["id"] = "Indonesia";
            termOverrides["ir"] = "Iran, Islamic Republic of";
            termOverrides["iq"] = "Iraq";
            termOverrides["ie"] = "Ireland";
            termOverrides["il"] = "Israel";
            termOverrides["it"] = "Italy";
            termOverrides["jm"] = "Jamaica";
            termOverrides["jp"] = "Japan";
            termOverrides["jo"] = "Jordan";
            termOverrides["kz"] = "Kazakhstan";
            termOverrides["ke"] = "Kenya";
            termOverrides["ki"] = "Kiribati";
            termOverrides["kw"] = "Kuwait";
            termOverrides["kg"] = "Kyrgyzstan";
            termOverrides["la"] = "Lao People's Democratic Republic";
            termOverrides["lv"] = "Latvia";
            termOverrides["lb"] = "Lebanon";
            termOverrides["ls"] = "Lesotho";
            termOverrides["lr"] = "Liberia";
            termOverrides["ly"] = "Libya";
            termOverrides["li"] = "Liechtenstein";
            termOverrides["lt"] = "Lithuania";
            termOverrides["lu"] = "Luxembourg";
            termOverrides["mg"] = "Madagascar";
            termOverrides["mw"] = "Malawi";
            termOverrides["my"] = "Malaysia";
            termOverrides["mv"] = "Maldives";
            termOverrides["ml"] = "Mali";
            termOverrides["mt"] = "Malta";
            termOverrides["mh"] = "Marshall Islands";
            termOverrides["mr"] = "Mauritania";
            termOverrides["mu"] = "Mauritius";
            termOverrides["mx"] = "Mexico";
            termOverrides["fm"] = "Micronesia, Federated States of";
            termOverrides["mc"] = "Monaco";
            termOverrides["mn"] = "Mongolia";
            termOverrides["me"] = "Montenegro";
            termOverrides["ma"] = "Morocco";
            termOverrides["mz"] = "Mozambique";
            termOverrides["mm"] = "Myanmar";
            termOverrides["na"] = "Namibia";
            termOverrides["nr"] = "Nauru";
            termOverrides["np"] = "Nepal";
            termOverrides["nl"] = "Netherlands";
            termOverrides["nz"] = "New Zealand";
            termOverrides["ni"] = "Nicaragua";
            termOverrides["ne"] = "Niger";
            termOverrides["ng"] = "Nigeria";
            termOverrides["nu"] = "Niue";
            termOverrides["no"] = "Norway";
            termOverrides["om"] = "Oman";
            termOverrides["pk"] = "Pakistan";
            termOverrides["pw"] = "Palau";
            termOverrides["pa"] = "Panama";
            termOverrides["pg"] = "Papua New Guinea";
            termOverrides["py"] = "Paraguay";
            termOverrides["pe"] = "Peru";
            termOverrides["ph"] = "Philippines";
            termOverrides["pl"] = "Poland";
            termOverrides["pt"] = "Portugal";
            termOverrides["qa"] = "Qatar";
            termOverrides["kr"] = "Korea, Republic of";
            termOverrides["md"] = "Moldova, Republic of";
            termOverrides["ro"] = "Romania";
            termOverrides["ru"] = "Russian Federation";
            termOverrides["rw"] = "Rwanda";
            termOverrides["kn"] = "Saint Kitts and Nevis";
            termOverrides["lc"] = "Saint Lucia";
            termOverrides["vc"] = "Saint Vincent and the Grenadines";
            termOverrides["ws"] = "Samoa";
            termOverrides["sm"] = "San Marino";
            termOverrides["st"] = "Sao Tome and Principe";
            termOverrides["sa"] = "Saudi Arabia";
            termOverrides["sn"] = "Senegal";
            termOverrides["rs"] = "Serbia";
            termOverrides["sc"] = "Seychelles";
            termOverrides["sl"] = "Sierra Leone";
            termOverrides["sg"] = "Singapore";
            termOverrides["sk"] = "Slovakia";
            termOverrides["si"] = "Slovenia";
            termOverrides["sb"] = "Solomon Islands";
            termOverrides["so"] = "Somalia";
            termOverrides["za"] = "South Africa";
            termOverrides["ss"] = "Sout Sudan";
            termOverrides["es"] = "Spain";
            termOverrides["lk"] = "Sri Lanka";
            termOverrides["ps"] = "State of Palestine";
            termOverrides["sd"] = "Sudan";
            termOverrides["sr"] = "Suriname";
            termOverrides["sz"] = "Swaziland";
            termOverrides["se"] = "Sweden";
            termOverrides["ch"] = "Switzerland";
            termOverrides["sy"] = "Syrian Arab Republic";
            termOverrides["tj"] = "Tajikistan";
            termOverrides["th"] = "Thailand";
            termOverrides["mk"] = "Macedonia, The Former Yugoslav Republic of";
            termOverrides["tl"] = "Timor-Leste";
            termOverrides["tg"] = "Togo";
            termOverrides["to"] = "Tonga";
            termOverrides["tt"] = "Trinidad and Tobago";
            termOverrides["tn"] = "Tunisia";
            termOverrides["tr"] = "Turkey";
            termOverrides["tm"] = "Turkmenistan";
            termOverrides["tv"] = "Tuvalu";
            termOverrides["ug"] = "Uganda";
            termOverrides["ua"] = "Ukraine";
            termOverrides["ae"] = "United Arab Emirates";
            termOverrides["gb"] = "United Kingdom of Great Britain and Northern Ireland";
            termOverrides["tz"] = "Tanzania, United Republic of";
            termOverrides["us"] = "United States of America";
            termOverrides["uy"] = "Uruguay";
            termOverrides["uz"] = "Uzbekistan";
            termOverrides["vu"] = "Vanuatu";
            termOverrides["ve"] = "Venezuela";
            termOverrides["vn"] = "Viet Nam";
            termOverrides["ye"] = "Yemen";
            termOverrides["zm"] = "Zambia";
            termOverrides["zw"] = "Zimbabwe";

            Term country = new Term();

            country.identifier = countryCode;
            country.title = termOverrides.ContainsKey(countryCode) ? termOverrides[countryCode] : "";

            return country;
        }

        
    }
}
