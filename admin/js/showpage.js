var xmlHttp
var xmlHttp2
var localWhereToPut
var localWhereToPut2
var xmlHttp22
var localWhereToPut22
var xmlHttp33
var localWhereToPut33
var tipo
var disparaFunc


function showPage(pageToFetch, whereToPut, getdisparaFunc) {
    disparaFunc = getdisparaFunc;
    xmlHttp = GetXmlHttpObject()
    if (xmlHttp == null) {
        alert("Sorry you cannot run AJAX Applications.")
        return
    }
    var url = pageToFetch
    localWhereToPut = whereToPut
    //url=url+"?q="
    //url=url+"&sid="+Math.random()
    xmlHttp.onreadystatechange = stateChanged
    xmlHttp.open("GET", url, true)
    xmlHttp.send(null)
}

function showPage2(pageToFetch2, whereToPut2, getdisparaFunc2) {
    disparaFunc2 = getdisparaFunc2;
    xmlHttp2 = GetXmlHttpObject()
    if (xmlHttp2 == null) {
        alert("Sorry you cannot run AJAX Applications.")
        return
    }
    var url = pageToFetch2
    localWhereToPut2 = whereToPut2
    //url=url+"?q="
    //url=url+"&sid="+Math.random()
    xmlHttp2.onreadystatechange = stateChanged2
    xmlHttp2.open("GET", url, true)
    xmlHttp2.send(null)
}

function showPage22(pageToFetch22, whereToPut22) {
    //alert('pop')
    xmlHttp22 = GetXmlHttpObject22()
    if (xmlHttp22 == null) {
        alert("Sorry you cannot run AJAX Applications.")
        return
    }
    var url = pageToFetch22
    localWhereToPut22 = whereToPut22
    xmlHttp22.onreadystatechange = stateChanged22
    xmlHttp22.open("GET", url, true)
    xmlHttp22.send(null)
}

function showPage33(pageToFetch33, whereToPut33, tipoxx) {
    tipo = tipoxx;
    xmlHttp33 = GetXmlHttpObject33()
    if (xmlHttp33 == null) {
        alert("Sorry you cannot run AJAX Applications.")
        return
    }
    var url = pageToFetch33
    localWhereToPut33 = whereToPut33
    xmlHttp33.onreadystatechange = stateChanged33
    xmlHttp33.open("GET", url, true)
    xmlHttp33.send(null)
}


function stateChanged() {
    if (xmlHttp.readyState == 4 || xmlHttp.readyState == "complete") {
        document.getElementById(localWhereToPut).innerHTML = xmlHttp.responseText
        localWhereToPut = ""
        
        if (disparaFunc == "1") {
           $('#example').dataTable();


 if(document.getElementById("busca_outro")){
 $('#example').DataTable().search(
 
  document.getElementById("search").value = document.getElementById("busca_outro").value,
        $('#global_regex').prop('checked'),
        $('#global_smart').prop('checked')
    ).draw();
        }
		}
        
    }
}


function stateChanged2() {
    if (xmlHttp2.readyState == 4 || xmlHttp2.readyState == "complete") {
        document.getElementById(localWhereToPut2).innerHTML = xmlHttp2.responseText
        localWhereToPut2 = ""
    }


    if (disparaFunc2 == "9") {
        $("a[rel^='pop']").prettyPopin({ width: 710, followScroll: false });
    }
}

function stateChanged22() {
//alert('pop')
    if (xmlHttp22.readyState == 4 || xmlHttp22.readyState == "complete") {
       // document.getElementById('data_texto').innerHTML = ""

        document.getElementById(localWhereToPut22).innerHTML = xmlHttp22.responseText
        localWhereToPut22 = ""
        //dataxxx = document.fol.dataxxx.value;
        //selecaoxxx = document.fol.selecaoxxx.value;
        //document.fol.data.value = dataxxx;
        //document.fol.data.value = data
    }
}
 

function stateChanged33() {
    if (xmlHttp33.readyState == 4 || xmlHttp33.readyState == "complete") {
        document.getElementById(localWhereToPut33).innerHTML = xmlHttp33.responseText
        localWhereToPut33 = ""
    }

    if (tipo == 2) {

        mostrar_carrinho('1');
        mostra_div_open();
        atualiza_carrinho_header();
        
    }
}

function GetXmlHttpObject() {
    var objXMLHttp = null
    if (window.XMLHttpRequest) {
        objXMLHttp = new XMLHttpRequest()
    }
    else if (window.ActiveXObject) {
        objXMLHttp = new ActiveXObject("Microsoft.XMLHTTP")
    }
    return objXMLHttp
}
function GetXmlHttpObject22() {
    var objXMLHttp22 = null
    if (window.XMLHttpRequest) {
        objXMLHttp22 = new XMLHttpRequest()
    }
    else if (window.ActiveXObject) {
        objXMLHttp22 = new ActiveXObject("Microsoft.XMLHTTP")
    }
    return objXMLHttp22
}
function GetXmlHttpObject33() {
    var objXMLHttp33 = null
    if (window.XMLHttpRequest) {
        objXMLHttp33 = new XMLHttpRequest()
    }
    else if (window.ActiveXObject) {
        objXMLHttp33 = new ActiveXObject("Microsoft.XMLHTTP")
    }
    return objXMLHttp33
}
 