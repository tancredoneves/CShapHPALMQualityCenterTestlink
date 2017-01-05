using LinqToExcel.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LinqExcel
{
   public class CaseTest
    {
       // [ExcelColumn("FOLDER")]
       // public string FOLDER { get; set; }
       // [ExcelColumn("SUITE")]
       // public string SUITE { get; set; }
        [ExcelColumn("NOME DO CASO DE TESTE")]
        public string NOMEDOTESTE { get; set; }
        [ExcelColumn("PRÉ-CONDIÇÃO")]
        public string PRECONDICAO { get; set; }
        [ExcelColumn("ETAPA")]
        public string ETAPA { get; set; }
        [ExcelColumn("DESCRIÇÃO")]
        public string DESCRICAO { get; set; }
        [ExcelColumn("RESULTADO ESPERADO")]
        public string RESULTADOESPERADO { get; set; }
    }
}
