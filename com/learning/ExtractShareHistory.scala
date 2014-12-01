package com.learning

import org.ccil.cowan.tagsoup.jaxp.SAXFactoryImpl;
import scala.xml.parsing.NoBindingFactoryAdapter;
import com.norbitltd.spoiwo.natures.xlsx.Model2XlsxConversions._
import com.norbitltd.spoiwo.model.enums.CellFill
import com.norbitltd.spoiwo.model._
import org.joda.time.LocalDate

object ExtractShareHistory { 
    
    var mySheet=Sheet(
    name = "Share History Chart sheet"
  )   
    
  
  def main (args: Array[String]){
   
    println("Hello world1")
    val parserFactory = new org.ccil.cowan.tagsoup.jaxp.SAXFactoryImpl
    val parser = parserFactory.newSAXParser()
    val source = new org.xml.sax.InputSource("http://www.moneycontrol.com/stocks/hist_stock_result.php?ex=B&sc_id=IT&pno=1&hdn=daily&fdt=2000-01-01&todt=2014-01-02")
    val adapter = new scala.xml.parsing.NoBindingFactoryAdapter
    val node=adapter.loadXML(source, parser)

for(tablenodes<- node\\ "table") 
 if((tablenodes \ "@class").toString().equalsIgnoreCase("tblchart"))
  
  {
   for(tablechildnodes<-tablenodes\\ "tr"){
      val shareHistChartHeader =(tablechildnodes\\ "th").map(_.text)
      val shareHistChartData= (tablechildnodes\\ "td").map(_.text)
      println("Before add row")
      println(((shareHistChartHeader.toSeq).toList).toString)
     mySheet= mySheet.addRow(Row().withCellValues((shareHistChartHeader.toSeq).toList))  
     mySheet=mySheet.addRow(Row().withCellValues((shareHistChartData.toSeq).toList))
     println(shareHistChartHeader.mkString(" "))
     println(shareHistChartData.mkString(" "))
  }
   println("for loop")   
   
  }
   
  
    mySheet.saveAsXlsx("D:\\Pgm\\Scala\\MySheet.xlsx") 
    
   
 //println("GOOD BYE WORLD" + node.child(0))
  }
}
