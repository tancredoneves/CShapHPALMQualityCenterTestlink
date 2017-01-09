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

            string sheetName = "Query1";
            var excelFile = new ExcelQueryFactory(pathToExcelFile);
            //mapeia o arquivo CLV na class <CaseTest> 
            var casetest = from a in excelFile.Worksheet<CaseTest>(sheetName) select a;

            //pega a quantidade de caso de test
            string[] nomes = casetest.Select(t => t.NOMEDOTESTE).ToArray();
            string[] pre = casetest.Select(t => t.PRECONDICAO).ToArray();
            string[] etapa = casetest.Select(t => t.ETAPA.Replace("Etapa ", "")).ToArray();
            string[] des = casetest.Select(t => t.DESCRICAO).ToArray();
            string[] res = casetest.Select(t => t.RESULTADOESPERADO).ToArray();

            //Cria o elemento raiz do XML
            var xsurvey = new XDocument(new XDeclaration("1.0", "UTF-8", "No"));
            var xtestcases = new XElement("testcases");

            // cria índice para casos de test e passos 
            int i, x;
            for (i = 0; i < nomes.Length; i++)
            {

                string internalid = "";
                var xmltestcase = new XElement("testcase", new XAttribute("internalid", internalid), new XAttribute("name", nomes[i]));
                //Cria os elementos do  "testcase"
                var node_order = new XElement("node_order", new XCData(""));
                var extid = new XElement("externalid", new XCData(""));
                var version = new XElement("version", new XCData(""));
                var summary = new XElement("summary", new XCData(""));
                var preconditions = new XElement("preconditions", new XCData(pre[i]));
                var execution_type = new XElement("execution_type", new XCData("1"));
                var importance = new XElement("importance", new XCData("2"));
                var steps = new XElement("steps");

                xmltestcase.Add(node_order);
                xmltestcase.Add(extid);
                xmltestcase.Add(version);
                xmltestcase.Add(summary);
                xmltestcase.Add(preconditions);
                xmltestcase.Add(execution_type);
                xmltestcase.Add(importance);

                do
                {
                    //Cria os elementos do  "steps"
                    var step = new XElement("step");
                    var step_number = new XElement("step_number", new XCData(etapa[i]));
                    var actions = new XElement("actions", new XCData(des[i]));
                    var expectedresults = new XElement("expectedresults", new XCData(res[i]));
                    var executiontype = new XElement("execution_type", new XCData("1"));
                    step.Add(step_number);
                    step.Add(actions);
                    step.Add(expectedresults);
                    step.Add(executiontype);
                    steps.Add(step);

                    i++;

                    //limita o contador 
                    if (i == nomes.Length)
                    {
                        break;
                    }

                    x = Convert.ToInt16(etapa[i]);

                    //condição de passos(Steps)
                } while (x != 1);
                i--;

                //Adiciona passos (Steps) e caso de teste(xmltestcase)
                xmltestcase.Add(steps);
                //Adiciona caso de teste(xmltestcase) em lista de cados de testes(xtestcases)
                xtestcases.Add(xmltestcase);
                //limita o contador 
                if (i == nomes.Length) break;

            }
            //Adiciona casos de testes
            xsurvey.Add(xtestcases);
            //xsurvey.Save(Console.Out); Console.WriteLine(); Console.ReadKey();
            // xsurvey.Save(@"C:\Users\LGMONEZI\Desktop\testlink\Caso.xml");
            return xsurvey;
        }
    }
}
