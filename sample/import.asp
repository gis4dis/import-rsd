<%@LANGUAGE=JScript%>
<%
var ErrorMessage = ""
var varBody;
var radek = 0;
var cislo = -1;
var filename_poradi = "D:\\DopravniInfo\\poradi.txt";
var output_dir = "D:\\DopravniInfo\\importXML\\";
 
try       {
		WriteLog("Test:zápisu do logu:"+output_dir,output_dir+"\\log.txt")
            while(cislo < 0) {
                        cislo = get_max(filename_poradi)
            }
 
            radek = 2
            var data = Request.BinaryRead(Request.TotalBytes);
            radek = 3
            var BinaryStream = Server.CreateObject("ADODB.Stream")
            radek = 4
            BinaryStream.Type = 1
 
            BinaryStream.Open()
            radek = 5
            BinaryStream.Write(data)
            radek = 6
            BinaryStream.Position = 0
            BinaryStream.Type = 2
            //Specify charset For the source text (unicode) data.
            BinaryStream.CharSet = "utf-8";
            radek = 7
            //Open the stream And get binary data from the object
            var Stream_BinaryToString = BinaryStream.ReadText
            radek = 8
            var objFSO = Server.CreateObject("Scripting.FileSystemObject");
            //filePath = Server.MapPath("/");
            radek = 9
            while (true) {
                        //filePath = Request.ServerVariables("APPL_PHYSICAL_PATH")
                        dnes = new Date();
                        fileName = (dnes.getFullYear()+"") + doplnNulu((dnes.getMonth()+1)+"") + doplnNulu(dnes.getDate()+"") + "_" + cislo
                        filePath = output_dir + fileName + ".xml"
                        radek = 10
                        if (!objFSO.FileExists(filePath)) break;
                        radek = 11
            }
            var xmlDoc = Server.CreateObject("Microsoft.XMLDOM");
            xmlDoc.async=false;
            xmlDoc.loadXML(Stream_BinaryToString);
            xmlDoc.save (filePath)
}
catch(e) {
            ErrorMessage = e.description;
            var fs = Server.CreateObject("Scripting.FileSystemObject");
            //tempFilename = Request.ServerVariables("APPL_PHYSICAL_PATH") + "chyba.txt"
            tempFilename = output_dir + "chyba.txt"
            ts = fs.OpenTextFile ( tempFilename, 8, true )
            dnes = new Date();
            ts.WriteLine(dnes + ":" + ErrorMessage + "; radek=" + radek);
            ts.Close()
}
Response.write(ErrorMessage)
function WriteLog(str,Path) 

var fso = Server.CreateObject("Scripting.FileSystemObject");
var MyFile = fso.OpenTextFile(Path, 2, true);
MyFile.WriteLine(str);	
MyFile.Close();	

function randNum (num) {
    var rnd1 = Math.round( (num-1) * Math.random() + 1 )
    return rnd1;
}
 
function doplnNulu(ret) {
            ret = ret + ""
            if (ret.length == 1)
                        return "0" + ret;
            else
                        return ret;
}
 
function get_max(filename) {
            var cislo = -1;
            try {
                        var fs = Server.CreateObject("Scripting.FileSystemObject");
                        ts = fs.OpenTextFile ( filename, 1, true )
                        if (!ts.AtEndOfStream)
                                   cislo = (ts.ReadLine() * 1);
                        else cislo = 0;
                        ts.Close();
 
                        cislo = cislo+1;
 
                        var fs = Server.CreateObject("Scripting.FileSystemObject");
                        ts = fs.OpenTextFile ( filename, 2, true )
                        ts.WriteLine(cislo);
                        ts.Close();
 
                        return cislo
            }
            catch (e) {
                        cislo = -1;
                        return cislo;
            }
}
 
%>
