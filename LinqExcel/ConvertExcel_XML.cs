using LinqToExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace LinqExcel
{
    class ConvertExcel_XML
    {
        public static XDocument ConversorExcel_XML(string pathToExcelFile)
        {
            //string pathToExcelFile = "" + @"C:\Users\LGMONEZI\Desktop\testlink\Cadastro de Proposta.xls";
            string sheetName = "Query1";
            var excelFile = new ExcelQueryFactory(pathToExcelFile);
            var casetest = from a in excelFile.Worksheet<CaseTest>(sheetName) select a;

            //foreach (var a in casetest)
            //{
            //    string Info = " NOMEDOTESTE: {0}; PRECONDICAO: {1}; ETAPA: {2}; DESCRICAO: {3};RESULTADOESPERADO: {4}";
            //    Console.WriteLine(Info, a.NOMEDOTESTE, a.PRECONDICAO, a.ETAPA, a.DESCRICAO, a.RESULTADOESPERADO);
            //}
            // Console.ReadKey();

            //pega a quantidade de caso de test
            //var NOMETESTE = casetest.Select(t => t.NOMEDOTESTE).ToArray().Distinct();
            //var PRECONDICAO = casetest.Select(t => t.PRECONDICAO).ToArray().Distinct();
            

            string[] nomes = casetest.Select(t => t.NOMEDOTESTE).ToArray();
            string[] pre = casetest.Select(t => t.PRECONDICAO).ToArray();
            string[] etapa = casetest.Select(t => t.ETAPA.Replace("Etapa ", "")).ToArray();
            string[] des = casetest.Select(t => t.DESCRICAO).ToArray();
            string[] res = casetest.Select(t => t.RESULTADOESPERADO).ToArray();

            //Cria o elemento raiz do XML
            var xsurvey = new XDocument(new XDeclaration("1.0", "UTF-8", "No"));
            var xtestcases = new XElement("testcases");
            int i;
            int x;
            for ( i = 0; i < nomes.Length; i++)
            {
                string internalid = "";
                var xmltestcase = new XElement("testcase", new XAttribute("internalid", internalid), new XAttribute("name", nomes[i]));
                var node_order = new XElement("node_order", new XCData(""));
                var extid = new XElement("externalid", new XCData(""));
                var version = new XElement("version", new XCData(""));
                var summary = new XElement("summary", new XCData(""));
                var preconditions = new XElement("preconditions", new XCData(pre[i]));
                var execution_type = new XElement("execution_type", new XCData("1"));
                var importance = new XElement("importance", new XCData("2"));
                var steps = new XElement("steps", new XCData(""));
                for ( x = 0; x < nomes.Length; x++)
                {
                    i++;
                    
                    var step_number = new XElement("step_number", new XCData(etapa[x]));
                    var actions = new XElement("actions", new XCData(des[x]));
                    var expectedresults = new XElement("expectedresults", new XCData(des[x]));
                    var executiontype = new XElement("execution_type", new XCData("1"));
                    steps.Add(step_number);
                    steps.Add(actions);
                    steps.Add(expectedresults);
                    steps.Add(executiontype);
                    if ("1".Contains(etapa[x])) break;
                }
                xmltestcase.Add(node_order);
                xmltestcase.Add(extid);
                xmltestcase.Add(version);
                xmltestcase.Add(summary);
                xmltestcase.Add(preconditions);
                xmltestcase.Add(importance);
                xmltestcase.Add(steps);
            }
            xsurvey.Add(xtestcases);
            //xsurvey.Save(Console.Out); Console.WriteLine(); Console.ReadKey();
            // xsurvey.Save(@"C:\Users\LGMONEZI\Desktop\testlink\Caso.xml");
            return xsurvey;
        }
        //private static XElement gera_Linha_XML(string nomes, string pre, string etapa, string des, string res)
       // private static XElement gera_Testcases_XML(string nomes, string pre, string etapa, string des, string res)
       // {
       //     string internalid = "";
       //     var xmltestcase = new XElement("testcase", new XAttribute("internalid", internalid), new XAttribute("name", nomes));
       //     var node_order = new XElement("node_order", new XCData(""));
       //     var extid = new XElement("externalid", new XCData(""));
       //     var version = new XElement("version", new XCData(""));
       //     var summary = new XElement("summary", new XCData(""));
       //     var preconditions = new XElement("preconditions", new XCData(pre));
       //     var execution_type = new XElement("execution_type", new XCData("1"));
       //     var importance = new XElement("importance", new XCData("2"));
       //     var steps = new XElement("steps", new XCData(""));
       //     for (int i = 0; i < etapa.Length; i++)
       //     {
       //         var step_number = new XElement("step_number", new XCData(etapa));
       //         var actions = new XElement("actions", new XCData(des));
       //         var expectedresults = new XElement("expectedresults", new XCData(des));
       //         var executiontype = new XElement("execution_type", new XCData("1"));
       //         steps.Add(step_number);
       //         steps.Add(actions);
       //         steps.Add(expectedresults);
       //         steps.Add(executiontype);
       //     }
       //     xmltestcase.Add(node_order);
       //     xmltestcase.Add(extid);
       //     xmltestcase.Add(version);
       //     xmltestcase.Add(summary);
       //     xmltestcase.Add(preconditions);
       //     xmltestcase.Add(importance);
       //     xmltestcase.Add(steps);
       //     return xmltestcase;
       //}
            
        

            //so com uma linha
            //var xmlLinha = new XElement("testcase", new XAttribute("internalid", internalid), new XAttribute("name", name),
            //                            new XElement("node_order", new XCData("")),
            //                            new XElement("externalid", new XCData("")),
            //                            new XElement("version", new XCData("")),
            //                            new XElement("summary", new XCData("")),
            //                            new XElement("preconditions", new XCData(preconditions)),
            //                            new XElement("execution_type", new XCData("1")),
            //                            new XElement("importance", new XCData("2")),
            //                            new XElement("steps", new XElement("step", new XElement("step_number", new XCData(steps)),
            //                                                                    new XElement("actions", new XCData(actions)),
            //                                                                    new XElement("expectedresults", new XCData(expectedresults)),
            //                                                                    new XElement("execution_type", new XCData("1")))));

            // return xmlLinha;
        


    }
}
