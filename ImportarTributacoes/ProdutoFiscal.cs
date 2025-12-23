using System;

namespace ImportarTributacoes
{
    public class ProdutoFiscal
    {
        public string Codigo { get; set; }
        public string Ncm { get; set; }
        public decimal Ibs { get; set; }
        public decimal Cbs { get; set; }
        public decimal Aliquota { get; set; }
    }

}
