using DocumentFormat.OpenXml.Wordprocessing;

namespace ValidateApp;

public class Clause
{
    public Clause(int tableIdx,int rowIdx){

    }
    public int TableIdx { get; set; }
    public int StartRowIdx { get; set; }


    
}

public class ScanTable
{
    public ScanTable(int tableIdx, int startRowIdx)
    {
        this.TableIdx = tableIdx;
        this.StartRowIdx = startRowIdx;
    }
    public int TableIdx { get; set; }
    public int StartRowIdx { get; set; }
    public List<Clause> CLause {get;set;}
    public int CurrentRowIdx { get; set;}
    public int currentRowCells{get;set;}
    public void NextRow(){
        this.CurrentRowIdx++;
    }
    public  void DoVerdict(){
        if(currentRowCells==3){
            // record clause no.
            // record row no.
        }
        if(currentRowCells==4){
            
        }
    }
}