using System;

namespace ImportarTributacoes
{
    public class TributacaoImportacao
    {
        public string Produto { get; set; } = string.Empty;
        public string Ncm { get; set; } = string.Empty;
        public string Cst { get; set; } = string.Empty;
        public string ClassTrib { get; set; } = string.Empty;
        public decimal AliqIbsUf { get; set; }
        public decimal AliqIbsMun { get; set; }
        public decimal AliqCbs { get; set; }
        public decimal ReducIbs { get; set;}
        public decimal ReducCbs { get; set;}
    }


}
