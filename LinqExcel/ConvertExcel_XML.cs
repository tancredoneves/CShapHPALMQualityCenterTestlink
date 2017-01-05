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
            string[] linhas = casetest.Select(t => t.NOMEDOTESTE).ToArray();
            string[] pre = casetest.Select(t => t.NOMEDOTESTE).ToArray();
            string[] etapa = casetest.Select(t => t.ETAPA.Replace("Etapa ", "")).ToArray();
            string[] des = casetest.Select(t => t.DESCRICAO).ToArray();
            string[] res = casetest.Select(t => t.RESULTADOESPERADO).ToArray();

            //Cria o elemento raiz do XML
            var xsurvey = new XDocument(new XDeclaration("1.0", "UTF-8", "No"));
            var xtestcases = new XElement("testcases");
            //percorre cada linha do arquivo casetest
            for (int i = 0; i < linhas.Length; i++)
            {
                //Cria cada linha
                if (i > 0)
                {
                    xtestcases.Add(gera_Linha_XML(linhas[i], linhas[0], pre[i], etapa[i], des[i], res[i]));
                }
            }
            xsurvey.Add(xtestcases);
            //xsurvey.Save(Console.Out); Console.WriteLine(); Console.ReadKey();
            // xsurvey.Save(@"C:\Users\LGMONEZI\Desktop\testlink\Caso.xml");
            return xsurvey;
        }

        private static XElement gera_Linha_XML(string linha, string primeiraLinha, string pre, string etapa, string des, string res)
        {
            //separa o nome dos campos do arquivo CSV
            string nomes = linha;
            string preconditions = pre;
            string steps = etapa;
            string actions = des;
            string expectedresults = res;
            string internalid = "";
            //Cria o elemento var e os atributos com o nome do campo e valor
            var xmlLinha = new XElement("testcase", new XAttribute("internalid", internalid), new XAttribute("name", nomes),
                                        new XElement("node_order", new XCData("")),
                                        new XElement("externalid", new XCData("")),
                                        new XElement("version", new XCData("")),
                                        new XElement("summary", new XCData("")),
                                        new XElement("preconditions", new XCData(preconditions)),
                                        new XElement("execution_type", new XCData("1")),
                                        new XElement("importance", new XCData("2")),
                                        new XElement("steps", new XElement("step", new XElement("step_number", new XCData(steps)),
                                                                                new XElement("actions", new XCData(actions)),
                                                                                new XElement("expectedresults", new XCData(expectedresults)),
                                                                                new XElement("execution_type", new XCData("1")))));

            return xmlLinha;
        }


    }
}
