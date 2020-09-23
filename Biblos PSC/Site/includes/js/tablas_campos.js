/*
*** Multiple dynamic combo boxes
*** by Mirko Elviro, 9 Mar 2005
*** Script featured and available on JavaScript Kit (http://www.javascriptkit.com)
***
***Please do not remove this comment
*/

// This script supports an unlimited number of linked combo boxed
// Their id must be "combo_0", "combo_1", "combo_2" etc.
// Here you have to put the data that will fill the combo boxes
// ie. data_2_1 will be the first option in the second combo box
// when the first combo box has the second option selected


// second combo box

data_1_1 = new Option("descripcion", "81");
data_2_1 = new Option("descripcion", "78");
data_3_1 = new Option("email", "68");
data_3_2 = new Option("tel1", "66");
data_3_3 = new Option("nombre", "65");
data_3_4 = new Option("tel2", "67");
data_3_5 = new Option("web", "69");
data_3_6 = new Option("domicilio_calle", "70");
data_3_7 = new Option("domicilio_nro", "71");
data_3_8 = new Option("domicilio_unidad", "73");
data_3_9 = new Option("domicilio_cod_postal", "74");
data_3_10 = new Option("domicilio_piso", "72");
data_4_1 = new Option("titulo", "63");
data_5_1 = new Option("descripcion", "59");
data_6_1 = new Option("autor", "50");
data_6_2 = new Option("titulo", "49");
data_6_3 = new Option("codigo_libro", "48");
data_6_4 = new Option("año", "52");
data_6_5 = new Option("ISBN", "51");
data_7_1 = new Option("descripcion", "45");
data_7_2 = new Option("direccion", "46");
data_8_3 = new Option("descripcion", "42");
data_9_1 = new Option("descripcion", "40");
data_10_1 = new Option("fecha_desde", "33");
data_10_2 = new Option("fecha_hasta", "34");
data_10_3 = new Option("fecha_devolucion", "38");
data_11_1 = new Option("valor", "31");
data_12_1 = new Option("descripcion", "25");
data_12_2 = new Option("fecha_creacion", "26");
data_13_1 = new Option("descripcion", "23");
data_14_1 = new Option("descripcion", "19");
data_14_2 = new Option("titulo", "20");
data_15_1 = new Option("dni", "7");
data_15_2 = new Option("mail", "6");
data_15_3 = new Option("apellido", "5");
data_15_4 = new Option("nombre", "4");
data_15_5 = new Option("password", "3");
data_15_6 = new Option("username", "2");
data_15_7 = new Option("domicilio_nro", "11");
data_15_8 = new Option("matricula", "8");
data_15_9 = new Option("fecha_nacimiento", "9");
data_15_10 = new Option("domicilio_calle", "10");
data_15_11 = new Option("domicilio_piso", "12");
data_15_12 = new Option("domicilio_unidad", "13");
data_15_13 = new Option("domicilio_cod_postal", "14");
data_15_14 = new Option("tel1", "15");
data_15_15 = new Option("tel2", "16");


// other parameters

    displaywhenempty=""
    valuewhenempty="-1"

    displaywhennotempty="-Seleccione-"
    valuewhennotempty="-1"


function change(currentbox) {
	numb = currentbox.id.split("_");
	currentbox = numb[1];

    i=parseInt(currentbox)+1

// I empty all combo boxes following the current one

    while ((eval("typeof(document.getElementById(\"combo_"+i+"\"))!='undefined'")) &&
           (document.getElementById("combo_"+i)!=null)) {
         son = document.getElementById("combo_"+i);
	     // I empty all options except the first one (it isn't allowed)
	     for (m=son.options.length-1;m>0;m--) son.options[m]=null;
	     // I reset the first option
	     son.options[0]=new Option(displaywhenempty,valuewhenempty)
	     i=i+1
    }


// now I create the string with the "base" name ("stringa"), ie. "data_1_0"
// to which I'll add _0,_1,_2,_3 etc to obtain the name of the combo box to fill

    stringa='data'
    i=0
    while ((eval("typeof(document.getElementById(\"combo_"+i+"\"))!='undefined'")) &&
           (document.getElementById("combo_"+i)!=null)) {
           eval("stringa=stringa+'_'+document.getElementById(\"combo_"+i+"\").selectedIndex")
           if (i==currentbox) break;
           i=i+1
    }


// filling the "son" combo (if exists)

    following=parseInt(currentbox)+1

    if ((eval("typeof(document.getElementById(\"combo_"+following+"\"))!='undefined'")) &&
       (document.getElementById("combo_"+following)!=null)) {
       son = document.getElementById("combo_"+following);
       stringa=stringa+"_"
       i=0
       while ((eval("typeof("+stringa+i+")!='undefined'")) || (i==0)) {

       // if there are no options, I empty the first option of the "son" combo
	   // otherwise I put "-select-" in it

	   	  if ((i==0) && eval("typeof("+stringa+"0)=='undefined'"))
	   	      if (eval("typeof("+stringa+"1)=='undefined'"))
	   	         eval("son.options[0]=new Option(displaywhenempty,valuewhenempty)")
	   	      else
	             eval("son.options[0]=new Option(displaywhennotempty,valuewhennotempty)")
	      else
              eval("son.options["+i+"]=new Option("+stringa+i+".text,"+stringa+i+".value)")
	      i=i+1
	   }
       //son.focus()
       i=1
       combostatus=''
       cstatus=stringa.split("_")
       while (cstatus[i]!=null) {
          combostatus=combostatus+cstatus[i]
          i=i+1
          }
       return combostatus;
    }
}
