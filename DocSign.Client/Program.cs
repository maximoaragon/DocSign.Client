using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace DocSign.Client
{
    class Program
    {
        //static string _docSignApiURL = "http://js-ml-dev/docsign/";
        static string _docSignApiURL = "http://showcase.equivant.com/docsign/";
        static string _docSignSharePath = @"\\js-ml-fs1\ShowCase\Images\DocSign\";

        static void Main(string[] args)
        {
            Console.WriteLine("*****::DocSign.ClientDemo::*****");

            //string templateDocument = Environment.CurrentDirectory + @"\SampleDocuments\MultipleSigDoc.docx";
            string templateDocument = Environment.CurrentDirectory + @"\SampleDocuments\DemoDoc.docx";

            string mergedDocument = MergeDoc(templateDocument);

            SendSignatureRequest(mergedDocument).Wait();

            Console.WriteLine("*****::DocSign.END::*****");
            Console.ReadLine();
        }

        private static string MergeDoc(string templateDocument)
        {
            string unsignedDocument = templateDocument;

            string mergedDocument = unsignedDocument.Replace(".docx", "-merged.docx");

            //create a copy that we can merge. 
            File.Copy(unsignedDocument, mergedDocument, true);

            System.Threading.Thread.Sleep(1000);//Copy and overwrite returns too quick sometimes.

            //Get app fields 
            Console.WriteLine(">>>>::Fetching Field Values");
            List<AppField> appFields = new List<AppField>()
            {
                new AppField() { FieldID = "f8731_0001", Value = "SampleDocuments\\ShowCase_DocSign.bmp", SignatureTag = "SC_Login_4979_DS"}, //Clerk
                new AppField() { FieldID = "f7651_0002", Value = "SampleDocuments\\ShowCase_DocSign.bmp", SignatureTag ="SC_Party_7315604_DS"} //Defendant
            };            

            using (var d = new DocxProcessor(mergedDocument))
            {
                Console.WriteLine(">>>>::Merging");
                d.MergeFields(appFields);
            }

            return mergedDocument;
        }

        static async Task SendSignatureRequest(string testDocument)
        {
            Console.WriteLine(">>>>::SendSignatureRequest");

            string DocSignAPIUrl = _docSignApiURL;

            //string srString = @"{'Documents':[{'ExternalID':56875042,'FilePath':'\\js-ml-fs1\ShowCase\Images\DocSign\ee74d6dd-3ce0-4efc-919c-0ecdf5f5e64f.docx','MetaData':null,'Name':'CRIMINAL-NOTICE OF HEARING - PB','Pages':1,'Signatures':[{'MediatorID':4979,'SignatureTemplate':'SC_Party_7315604_DS','SigneeID':0,'SigneeName':'Party #1 DEFENDANT\u000d\u000aBROWNING, JEFFERY LEON JR','Type':1}],'Type':null}],'MeetingRoom':'KD (Gun Club) ','PackageName':'50-2013-CF-000458-AXXX-WB'}";
            //var signatureReuquest = ShowCaseUtil.SerializationHelper.DeSerializeJson<SignaturePackage>(srString); ;

            string signatureReuquest = CreateSignaturePackage(testDocument);

            try
            {
                //var sm = new SignInManager();
                //sm.Authenticate("max", "testing", false);

                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri( _docSignApiURL);
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    var content = new StringContent(signatureReuquest, Encoding.UTF8, "application/json");

                    var response = await client.PostAsync("api/signature", content);
                    
                    if (response.IsSuccessStatusCode)
                    {
                        //var r = response.Content.ReadAsStringAsync().Result;
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("SignatureRequest: 200 - OK :-)");
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Something went wrong :-(");
                    }
                    Console.ForegroundColor = ConsoleColor.Gray;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error Sending request:" + ex.Message);
            }
        }

        private static string CreateSignaturePackage(string testDocument)
        {
            string docInDocSignPath =  _docSignSharePath + Guid.NewGuid().ToString() + ".docx"  ;
            string jsonDocSignSharePath = docInDocSignPath.Replace(@"\", @"\\");

            if (!File.Exists(testDocument))
            {
                Console.WriteLine("File " + testDocument + " was not found");
                return "";
            }

            string base64Doc = Convert.ToBase64String(File.ReadAllBytes( testDocument));

            File.Copy(testDocument, docInDocSignPath, true);

            string jsonData = "{ " +
                "\"Documents\":[" +
                "{" +
                "   \"Name\":\"CRIMINAL-NOTICE OF HEARING\"," +
                "   \"FilePath\":," +
                "   \"ExternalID\":56884988," + //ContainerID
                "   \"Pages\":1," +
                "   \"Type\":null," +
                "   \"FileData\": \"" + base64Doc + "\"," +
                "   \"MetaData\":" +
                "       \"{'ExternalID': 56884988}\"," +
                "   \"Signatures\":[" +
                "   {" +
                "       \"SigneeID\":5434," +
                "       \"SigneeName\":\"Clerk Aragones, Maxuel\"," +
                "       \"MediatorID\":5434," +
                "       \"SignatureTemplate\":\"SC_Login_4979_DS\"," +
                "       \"Type\":0," +
                "       \"Storage\":0}," +
                "   {" +
                "       \"SigneeID\":0," +
                "       \"SigneeName\":\"Party #1 DEFENDANT BROWNING, JEFFERY LEON JR\"," +
                "       \"MediatorID\":4979," +
                "       \"SignatureTemplate\":\"SC_Party_7315604_DS\"," +
                "       \"Type\":1," +
                "       \"Storage\":0}]" +
                "   }" +
                "]," +
                "   \"SignatureRoom\":\"#5 (South Branch)\"," +
                "   \"PackageName\":\"Unit test package\"" +
                "}";


            return jsonData;
        }
    }
}
