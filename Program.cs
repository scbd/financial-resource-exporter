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
            var records = getRecords().OrderBy(o=>((string)o["government"]["title"]).ToLower());

			Console.WriteLine("{0} records found", records.Count());

			foreach(dynamic r in records)
				Console.WriteLine(r.government.title);

            Excel.Application xlApp;
            Excel.Workbook    xlWorkBook;

            using(new XLDisposable(xlApp = new Excel.Application()))
            using(new XLDisposable(xlWorkBook = (Excel.Workbook)xlApp.Workbooks.Open(Environment.CurrentDirectory+"\\"+args[0], ReadOnly:true)))
            {
                Excel.Worksheet xlWorkSheetTemplate = (Excel.Worksheet)xlWorkBook.Worksheets["{{template}}"];

				var bindings = getBindings(xlWorkSheetTemplate);

                foreach(var record in records)
				{
					xlWorkSheetTemplate.Copy(Before:xlWorkSheetTemplate);

					Excel.Worksheet xlWorkSheet = xlWorkBook.Worksheets[xlWorkSheetTemplate.Index-1];

					var values = mapValues(record);
					var name   = (string)values["government.title"];

					xlWorkSheet.Name = name.Substring(0, Math.Min(name.Length, 30));

                    populate(xlWorkSheet, values, bindings);
				}

                xlWorkSheetTemplate.Delete();

                xlWorkBook.SaveAs(Environment.CurrentDirectory+string.Format("\\out-{1:yyyy-MM-dd-HH-mm-ss}{0}", Path.GetExtension(args[0]), DateTime.Now));
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

            foreach (Excel.Range cell in sheet.UsedRange)
            {
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
        private static void populate(Excel.Worksheet sheet, IDictionary<string, object> values, IEnumerable<Binding> bindings)
        {
			Console.WriteLine("Compiling {0}...", sheet.Name);

            foreach (var binding in bindings)
            {
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

			var qValues = record.Descendants().Where(o => types.Contains(o.Type));

			foreach(var jValue in qValues)
				map.Add(jValue.Path, toXlValue(jValue));

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
                return ""; 
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
        static IEnumerable<JObject> getRecords()
        {
			Console.WriteLine("Loading records...");

            WebClient wc = new WebClient();

            UriBuilder url = new UriBuilder("https://api.cbd.int/api/v2013/index");

            url.Query += string.Format("&q={0}",    Uri.EscapeUriString("schema_s:resourceMobilisation AND _state_s:public AND realm_ss:chm"));
            url.Query += string.Format("&fl={0}",   Uri.EscapeUriString("identifier_s"));
            url.Query += string.Format("&rows={0}", Uri.EscapeUriString("2000"));

            url.Query = url.Query.TrimStart('&', '?');
            
            var result = JObject.Parse(wc.DownloadString(url.ToString()));

            var identifiers = result["response"]["docs"].Select(o=>(string)o["identifier_s"]);

            var jRecords = identifiers.Select(id=> JObject.Parse(wc.DownloadString("https://api.cbd.int/api/v2013/documents/"+id)));

            return jRecords.Select(o=>normalizeRecord((JObject)o)).ToArray();
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

                foreach(var term in loadDomainTerms("countries")) termsCache.Add(term.identifier, term); //TODO USE Special Countries mapping
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
					string slot  = item[key].ToString();
					JToken value = item;

					if(transform!=null)
						value = transform(item);

					ret[slot] = value;
				}
			}

            return ret;
        }
    }
}
