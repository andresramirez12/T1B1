using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.Nancy.Modules.MediosMagneticos
{
    public class _MediosMagneticosClass
    {
        public void executeMM()
        {

        }

        private void Formato1009()
        {
            List<_Formato1009> listFormato1099 = new List<_Formato1009>();

            string strbasequery = " select JDT1.transid,JDT1.Line_ID,JDt1.Account,OACT.AcctName,case when JDt1.Account != JDT1.ShortName then JDT1.ShortName else null end as 'BaseCardCode',JDT1.shortname,jdt1.RefDate,JDT1.U_BYB_TRBK,0 as 'SI' " +
            " 	,JDT1.Debit as 'Debito',Jdt1.Credit as 'Credito',T0.CardCode as 'OINVCardCode',T1.CardCode as 'ORINCardCode',T2.CardCode as 'ODLNCardCode',T3.CardCode as 'ORDNCardCode',T4.CardCode as 'OPCHCardCode',T5.CardCode as 'ORPCCardCode' " +
            " 	,T6.CardCode as 'OPDNCardCode',T7.CardCode as 'ORPDCardCode',T8.CardCode as 'ORTCCardCode',T9.CardCode as 'OVPMCardCode',T10.CardCode as 'OIPFCardCode',OJDT.TransType " +
            " 	,jdt1.BaseRef from jdt1 inner join OJDT on OJDT.TransId = JDT1.TransId inner join OACT on jdt1.Account = oact.AcctCode   left join OINV T0 on T0.DocNum = jdt1.BaseRef and jdt1.TransType = 13 left join ORIN T1 on T1.DocNum = jdt1.BaseRef and jdt1.TransType = 14 " +
            " left join ODLN T2 on T2.DocNum = jdt1.BaseRef and jdt1.TransType = 15 left join ORDN T3 on T3.DocNum = jdt1.BaseRef and jdt1.TransType = 16  left join OPCH T4 on T4.DocNum = jdt1.BaseRef and jdt1.TransType = 18 left join ORPC T5 on T5.DocNum = jdt1.BaseRef and jdt1.TransType = 19 " +
            " left join OPDN T6 on T6.DocNum = jdt1.BaseRef and jdt1.TransType = 20 left join ORPD T7 on T7.DocNum = jdt1.BaseRef and jdt1.TransType = 21 left join ORCT T8 on T8.DocNum = jdt1.BaseRef and jdt1.TransType = 24  left join OVPM T9 on T9.DocNum = jdt1.BaseRef and jdt1.TransType = 46 " +
             " left join OIPF T10 on T10.DocNum = jdt1.BaseRef and jdt1.TransType = 69 where OJDT.RefDate < '2016-01-01' and substring(jdt1.account, 1, 1) = '2' order by JDT1.Account ";

            SAPbobsCOM.Recordset objRecordSet = T1.Classes.BYBB1MainObject.Instance.B1Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            objRecordSet.DoQuery(strbasequery);
            if (objRecordSet.RecordCount > 0)
            {
                while (!objRecordSet.EoF)
                {




                    objRecordSet.MoveNext();
                }
            }
        }
    }

    internal class _Formato1009
    {
        string Concepto { get; set; }
        string TD { get; set; }
        string Documento { get; set; }
        string DIgitoVerifiacion { get; set; }
        string PrimerApellido { get; set; }
        string SegundoApellido { get; set; }
        string PrimerNombre { get; set; }
        string OtrosNombres { get; set; }
        string RazonSocial { get; set; }
        string Direccion { get; set; }
        string Departamento { get; set; }
        string Municipio { get; set; }
        string Pais { get; set; }
        double Saldo { get; set; }

    }
}
