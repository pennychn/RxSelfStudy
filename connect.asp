<%
Response.expires = 0
set conn = server.createobject("ADODB.Connection")

' Check DB is exist or not
Set FileObject = Server.CreateObject("Scripting.FileSystemObject")
if Not FileObject.FileExists(Server.MapPath("./db/dbPrescriptionDrug.mdb")) then
	Response.Write("Sorry! System is under construction!")
	Response.End()
end if

params = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("./db/dbPrescriptionDrug.mdb") & ";Jet OLEDB:Database Password="
conn.Open params
' conn.Open "lesson"
' onn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & server.mappath("adv.mdb")
%>
