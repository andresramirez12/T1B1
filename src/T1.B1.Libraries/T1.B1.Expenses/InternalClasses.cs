using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T1.B1.Expenses
{
    public class projectInfo
    {
        public string ProjectCode { get; set; }
        public string ProjectName { get; set; }
    }

    public class costCenterInfo
    {
        public string CostCenterCode { get; set; }
        public string CostCenterName { get; set; }
        public int DimensionCode { get; set; }
    }

    public class dimensionInfo
    {
        public int DimentionCode { get; set; }
        public string DimensionName { get; set; }
        public bool isActive { get; set; }
    }

    public class Concept
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string Account { get; set; }
        public List<ConceptThirdParty> ValidThirdparty { get; set; }
        public List<ConceptWHTax> ValidWHTax { get; set; }
        public List<ConceptVAT> ValidVAT { get; set; }
        public List<string> validExpenseType { get; set; }
        
    }

    public class ConceptLines
    {
        public string ConceptCode { get; set; }
        public string Description { get; set; }
        public double TotalBeforeTaxes { get; set; }
        public DateTime Date { get; set; }
        public double WHTax { get; set; }
        public double VAT { get; set; }
        public double LineTotal { get; set; }
        public string ThirdParty { get; set; }
        public string Project { get; set; }
        public string ProfitCenter { get; set; }
        public string DIM1 { get; set; }
        public string DIM2 { get; set; }
        public string DIM3 { get; set; }
        public string DIM4 { get; set; }
        public string DIM5 { get; set; }

        public Concept Concept { get; set; }
        public int FormLineNum { get; set; }

    }

    public class ConceptThirdParty
    {
        public string Code { get; set; }
        public bool Default { get; set; }
    }

    public class ConceptWHTax
    {
        public string Code { get; set; }
        
    }

    public class ConceptVAT
    {
        public string Code { get; set; }

    }

    
    public class LegalizationFormCache
    {
        public int DocEntry { get; set; }
        public int DocNum { get; set; }
        public int Series { get; set; }
        public Expense expense { get; set; }
        public DateTime PostingDate { get; set; }
        public int JournalEntry { get; set; }
        public bool isPosted { get; set; }
        public double TotalValue { get; set; }
        public double ExpenseValue { get; set; }
        public List<Concept> ConceptList { get; set; }
        public List<ConceptLines> DocumentLines { get; set; }
        public string remarks { get; set; }

    }

    public class Expense
    {
        public int docEntry { get; set; }
        public int docNum { get; set; }
        public int series { get; set; }
        public DateTime startDate { get; set; }
        public DateTime endDate { get; set; }
        public List<ExpenseResponsableThirdParty> expenseResponsableThirdParty { get; set; }
        public ExpenseCostAccounting expenseCostAccounting { get;set;}
        public double expectedValue { get; set; }
        public double legalizedValue { get; set; }
        public string remark { get; set; }
        public string status { get; set; }
        public ExpenseType expnseType { get; set; }


    }

    public class ExpenseResponsableThirdParty
    {
        public string CardCode { get; set; }

    }
    public class ExpenseCostAccounting
    {
        public string Project { get; set; }
        public string CostCenter { get; set; }
        public string DIM1 { get; set; }
        public string DIM2 { get; set; }
        public string DIM3 { get; set; }
        public string DIM4 { get; set; }
        public string DIM5 { get; set; }
    }

    public class ExpenseType
    {
        public string code { get; set; }
        public string name { get; set; }
        public string account { get; set; }
        public string remark { get; set; }
        public string expenseClass { get; set; }
    }

    #region Caja Menor

    public class PCConcept
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public string VATCode { get; set; }
        public string Account { get; set; }
        public List<PCConceptThirdParty> ValidThirdparty { get; set; }
        public List<PCConceptWHTax> ValidWHTax { get; set; }
        

    }

    public class PCConceptThirdParty
    {
        public string Code { get; set; }
        public bool Default { get; set; }
    }

    public class PCConceptWHTax
    {
        public string Code { get; set; }

    }

    
    public class PCLegalizationFormCache
    {
        public int DocEntry { get; set; }
        public int DocNum { get; set; }
        public int Series { get; set; }
        public PC pettyCash { get; set; }
        public bool isPosted { get; set; }
        public double TotalValue { get; set; }
        public double PCValue { get; set; }
        public List<PCConceptLines> DocumentLines { get; set; }
        public List<PCJournalEntries> JournalEntries { get; set; }
        public List<PCExternalDocuments> ExternalDocuments { get; set; }
        public string remarks { get; set; }

    }

    public class PCConceptLines
    {
        public string ConceptCode { get; set; }
        public string Description { get; set; }
        public double TotalBeforeTaxes { get; set; }
        public DateTime Date { get; set; }
        public string ThirdParty { get; set; }
        public double WHTax { get; set; }
        public double VAT { get; set; }
        public double LineTotal { get; set; }
        public string Project { get; set; }
        public string ProfitCenter { get; set; }
        public string DIM1 { get; set; }
        public string DIM2 { get; set; }
        public string DIM3 { get; set; }
        public string DIM4 { get; set; }
        public string DIM5 { get; set; }

        public PCConcept Concept { get; set; }
        public int FormLineNum { get; set; }

    }

    public class PCJournalEntries
    {
        public int JournalEntry { get; set; }
        public double Total { get; set; }
        public DateTime Date { get; set; }
    }

    public class PCExternalDocuments
    {
        public int DocEntry { get; set; }
        public int DocNum { get; set; }
        public double DocTotal { get; set; }
        public string DocType { get; set; }
        public DateTime DocDate { get; set; }
    }

    public class PC
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public string ThirdParty { get; set; }
        public string Remark { get; set; }
        public string PCAccount { get; set; }
        public string ControlAccount { get; set; }
        public double Value { get; set; }
        public double AvailableValue { get; set; }
        public bool isPosted { get; set; }
    }

    #endregion


}
