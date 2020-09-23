var emailexp = /^[a-z][a-z_0-9\-\.]+@[a-z_0-9\.\-]+\.[a-z]{2,3}$/i;
var phoneexp =  /^\d{10}$/;
var dniexp = /^\d{8}$/;

function CRLF () {
	return String.fromCharCode(10) + String.fromCharCode(13);
}

function TAB(howMany) {
	var tempStr;
	for (count = 0; count < howMany; count++) {
		tempStr = tempStr & String.fromCharCode(9);
	}
}

function validateEmail(str) {	
	return emailexp.test(str);
}

function validatePhone(str) {
	return phoneexp.test(str);
}

function validateDNI(str) {
	return dniexp.test(str);
}

function StripChars(ItemsToStrip, str) {
	returnString = "";
	for (i = 0; i < str.length; i++) {  
		var c = str.charAt(i);
        	if (ItemsToStrip.indexOf(c) == -1) returnString += c;       	 	 
	}
	return returnString;
}

function AllSpace(str) {   //Makes String Blank if noting but spaces
	for (i=0; i < str.length; i++) {
		if (str.charAt(i) != " ") {
			return str;
		}
	}
	return "";
}

function SetDec(str, places) { //chops decimal places to max number of places	
	if (isNaN(str)) {
		return str;
	}
	if (str.indexOf(".") != -1) {
	    if (places > 0) {
		str = str.substring(0, eval(str.indexOf(".")) + eval(places) + eval(1));
	    } else {
		str = str.substring(0, str.indexOf("."));
	    }
	}
	return str;
}

function DateFormat(dateVal) {	
	DayVal = dateVal.getDate();
	MonthVal = dateVal.getMonth();
	YearVal = dateVal.getYear();	
	if (YearVal.length <= 2) {
		YearVal = eval(YearVal) + 1900;				
	}
	tempStr = eval(MonthVal + 1) + "/" + DayVal + "/" + YearVal;	
	return tempStr;
}

function stripNonDigits(str) {
	return str.replace(/[^0-9]/g,"")
}

function checkform(form, errColor, startColor, showAlert, showErrors, fontStyle) {
    Error = false;	
    alertStr = "";
    for (x=0; x < form.elements.length; x++ ) {	
	fieldError = false;   
	if (form.elements[x].type == "text" || form.elements[x].type == "select-one" || form.elements[x].type == "password"  || form.elements[x].type == "textarea") {
	    if (showErrors == true) {
		document.all[form.elements[x].name + 'Error'].innerHTML = "";
	    }		
	    form.elements[x].value = AllSpace(form.elements[x].value);
	    if (x+1 < form.length && form.elements[x+1].name.charAt(0) == "@") {
		paramStr = form.elements[x+1].name.substring(1, form.elements[x+1].name.length);
		params = null;
		params = paramStr.split("_");
			
		if (params[7] != null) {
			backColor = params[7];
		} else {
			backColor = startColor;
		} 
		
		if (params[6] != null && AllSpace(params[6]) != "" ) {
			defaultValue = params[6];
		} else {
			defaultValue = "";
		}
		if (params[1] == "NoBlank" && form.elements[x].value == "" && defaultValue == "") {
			alertStr = alertStr + "El campo " + params[2].replace(/°/gi, " ") + " no debe estar vacío." + CRLF();
			if (showErrors == true) {
				document.all[form.elements[x].name + 'Error'].innerHTML = "<font  style='" + fontStyle + "'>El campo " + params[2].replace(/°/gi, " ") + " no debe estar vacío.</font>";
			}
			Error = true;
			fieldError = true;	
		
		} else if (params[1] == "NoBlank" && form.elements[x].value == "" && defaultValue != "") {
			form.elements[x].value = defaultValue;

		} else if (params[0] == "email") {
			if (!validateEmail(form.elements[x].value) && form.elements[x].value != "") {				
				alertStr = alertStr + "El campo " + params[2].replace(/°/gi, " ") + " no contiene una dirección de email válida." + CRLF();
				if (showErrors == true) {
					document.all[form.elements[x].name + 'Error'].innerHTML = "<font  style='" + fontStyle + "'>El campo " + params[2].replace(/°/gi, " ") + " no contiene una dirección de email válida.</font>";
				}
				Error = true;
				fieldError = true;
			} 
		} else if (params[0] == "number" && form.elements[x].value != "") {
			form.elements[x].value =  StripChars("$,%", form.elements[x].value);			
			if (params[3] != null) {
				form.elements[x].value = SetDec(form.elements[x].value, params[3]);	
			}
			if (isNaN(form.elements[x].value)) {
			 	alertStr = alertStr + "El campo " + params[2].replace(/°/gi, " ") + " no contiene un valor numérico válido." + CRLF();
				if (showErrors == true) {
					document.all[form.elements[x].name + 'Error'].innerHTML = "<font  style='" + fontStyle + "'>El campo " + params[2].replace(/°/gi, " ") + " no contiene un valor numérico válido.</font>";
				}
				Error = true;
				fieldError = true;
			} else {
				if (params[4] != null) {
					if (eval(form.elements[x].value) < eval(params[4])) {
						alertStr = alertStr + "El campo " + params[2].replace(/°/gi, " ") + " debe ser mas grande que " + params[4] + CRLF();
						if (showErrors == true) {
							document.all[form.elements[x].name + 'Error'].innerHTML = "<font  style='" + fontStyle + "'>El campo " + params[2].replace(/°/gi, " ") + " debe ser mas grande que " + params[4] + "</font>";
						}
						Error = true;
						fieldError = true;
					}
				}
				if (params[5] != null) {
					if (eval(form.elements[x].value) > eval(params[5])) {
						alertStr = alertStr + "El campo " + params[2].replace(/°/gi, " ") + " debe ser mas menor a " + params[5] + CRLF();
						if (showErrors == true) {
							document.all[form.elements[x].name + 'Error'].innerHTML = "<font  style='" + fontStyle + "'>El campo " + params[2].replace(/°/gi, " ") + " debe ser mas menor a " + params[5] + "</font>";
						}
						Error = true;
						fieldError = true;
					}
				}
			}
			
		} else if (params[0] == "age" && form.elements[x].value != "") {
			
			form.elements[x].value = SetDec(form.elements[x].value,0);
			if (eval(form.elements[x].value) < 0 || eval(form.elements[x].value) > 120) {
				alertStr = alertStr + "El campo " + params[2].replace(/°/gi, " ") + " no parece contener una edad válida." + CRLF();
				if (showErrors == true) {
					document.all[form.elements[x].name + 'Error'].innerHTML = "<font  style='" + fontStyle + "'>El campo " + params[2].replace(/°/gi, " ") + " no parece contener una edad válida.</font>";
				}
				Error = true;
				fieldError = true;
			}			
		} else if (params[0] == "date" && form.elements[x].value != "") {
			dateYear = new String();
			curDate = new Date();
			tempDate = new Date(form.elements[x].value);
			dateYear = dateYear + tempDate.getYear();			
			if (dateYear.length <= 2) {
				dateYear = eval(dateYear) + eval(1900);				
			}			
			if (form.elements[x].value != "") {
			    if (tempDate == "NaN") {
				alertStr = alertStr + "El campo " + params[2].replace(/°/gi, " ") + " no es una fecha válida." + CRLF();
				if (showErrors == true) {
					document.all[form.elements[x].name + 'Error'].innerHTML = "<font  style='" + fontStyle + "'>El campo " + params[2].replace(/°/gi, " ") + " no es una fecha válida.</font>";
				}
				Error = true;
				fieldError = true;			
			    } else if (params[4] != null && dateYear < eval(curDate.getYear()) - eval(params[4])) {
				alertStr = alertStr + "El campo " + params[2].replace(/°/gi, " ") + " es demasiado bajo." + CRLF();
				if (showErrors == true) {
					document.all[form.elements[x].name + 'Error'].innerHTML = "<font  style='" + fontStyle + "'>El campo " + params[2].replace(/°/gi, " ") + " es demasiado bajo.</font>";
				}
				Error = true;
				fieldError = true;			
	
			    } else if (params[5] != null && dateYear > eval(curDate.getYear()) + eval(params[5])) {
				alertStr = alertStr + "El campo " + params[2].replace(/°/gi, " ") + " es demasiado alto." + CRLF();
				if (showErrors == true) {
					document.all[form.elements[x].name + 'Error'].innerHTML = "<font  style='" + fontStyle + "'>El campo " + params[2].replace(/°/gi, " ") + " es demasiado alto.</font>";
				}
				Error = true;
				fieldError = true;			
	
			    } else {
				form.elements[x].value = DateFormat(tempDate);
			    }
			}
		} else if (params[0] == "phone" && form.elements[x].value != "") {
			form.elements[x].value = stripNonDigits(form.elements[x].value)
			if (validatePhone(form.elements[x].value)) {
			    	tempP = form.elements[x].value	
			    	form.elements[x].value = "(" + tempP.substring(0, 3) + ") " + tempP.substring(3,6) + "-" + tempP.substring(6, 10)
			} else {
			    	alertStr = alertStr + "El campo " + params[2].replace(/°/gi, " ") + " es inválido. Por favor ingrese el número telefónico completo, incluyendo el código de área." + CRLF();
				if (showErrors == true) {
					document.all[form.elements[x].name + 'Error'].innerHTML = "<font  style='" + fontStyle + "'>El campo " + params[2].replace(/°/gi, " ") + " es inválido. Por favor ingrese el número telefónico completo, incluyendo el código de área.</font>";
				}
				Error = true;
				fieldError = true;	
			}
		} else if (params[0] == "dni" && form.elements[x].value != "") {
			if (!validateDNI(form.elements[x].value)) {
			    alertStr = alertStr + "El campo " + params[2].replace(/°/gi, " ") + " es inválido. Por favor ingrese el número de DNI correctamente." + CRLF();
				if (showErrors == true) {
					document.all[form.elements[x].name + 'Error'].innerHTML = "<font  style='" + fontStyle + "'>El campo " + params[2].replace(/°/gi, " ") + " es inválido. Por favor ingrese el número de DNI correctamente.</font>";
				}
				Error = true;
				fieldError = true;	
			}
		}
		if (fieldError == true) {
			form.elements[x].style.background = errColor;
	    	} else {
			form.elements[x].style.background = backColor;
	   	}			
	    } 
	}				
    }
    
    if (Error == true) {
	if (showAlert == true) {
		alert (alertStr);
	}
	return false;
    }
    
}

//**********************************************************************************************
//   invierta una fecha dada retornando en formato YYYYMMDD
// dFecIni = Fecha a invertir
//   nTipFormat = Formato en que biene la fecha
//             1 = DD/MM/YYYY 
//             2 = MM/DD/YYYY   
//             3 = YYYY/MM/DD
//             4 = YYYY/DD/MM

function invFecha(nTipFormat,dFecIni){
   var dFecIni = dFecIni.replace(/-/g,"/");               // reemplaza el - por /   
   
   // primera division fecha
   var nPosUno = ponCero(dFecIni.substr(0,dFecIni.indexOf("/")));
   // 2º divicion fecha
   var nPosDos = ponCero(dFecIni.substr(parseInt(dFecIni.indexOf("/")) + 1,parseInt(dFecIni.lastIndexOf("/")) - parseInt(dFecIni.indexOf("/")) - 1));
   // 3º divicion fecha
   var nPosTres = ponCero(dFecIni.substr(parseInt(dFecIni.lastIndexOf("/")) + 1));

   switch(nTipFormat){
      case 1 :   //   DD/MM/YYYY
         dReturnFecha = nPosTres + "" + nPosDos + "" + nPosUno;
         break;

      case 2 :   //   MM/DD/YYYY
         dReturnFecha = nPosTres + "" + nPosUno + "" +nPosDos;
         break;

      case 3 :   //   YYYY/MM/DD
         dReturnFecha = nPosUno + "" + nPosDos + "" +nPosTres;
         break;
   
      case 4 :   //   YYYY/DD/MM
         dReturnFecha = nPosUno + "" + nPosTres + "" +nPosDos;
         break;
   }
   
   return dReturnFecha;   // retorna la fecha    
}

// Agrega un cero delante del strPon cuando tenga solo un caracter
function ponCero(strPon){
   if(parseInt(strPon.length) < 2)
      strPon = "0" + strPon;
   return strPon;
}
