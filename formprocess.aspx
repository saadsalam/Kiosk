<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Text"%>

<html>

<body>


<script language="vbscript" runat="server">

Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
 
        Dim FullName,Address,PhoneNumber,NoOfStudents,OtherReason,Appointment,AppointmentDateTime As String 
	Dim Enrollment,ChangeOfAddress,Counselor,McKinneyProgram,IntraApplication,LunchApplication,BusApplication,YesOther as Integer
	

                
        FullName= Request.Form("Name") 
        
        Address= Request.Form("Address") 
        PhoneNumber= Request.Form("PhoneNumber") 
	
	Enrollment= Request.Form("Enrollment") 

		
	
	NoOfStudents= Request.Form("NoOfStudents") 
	If NoOfStudents="" Then
	NoOfStudents= " null "
		
	End If
	
	
	ChangeOfAddress= Request.Form("ChangeOfAddress")
	Counselor= Request.Form("Counselor") &""
	McKinneyProgram= Request.Form("McKinneyProgram")
	IntraApplication= Request.Form("IntraApplication") 
	LunchApplication= Request.Form("LunchApplication") 
	BusApplication= Request.Form("BusApplication") 
	YesOther= Request.Form("YesOther") &""
	OtherReason= Request.Form("OtherReason") 
	Appointment= Request.Form("Appointment") 
	
	AppointmentDateTime= Request.Form("AppointmentDateTime") 
	
	If AppointmentDateTime="" Then
	AppointmentDateTime= " null "
			
	End If
	
	
        
	response.write(FullName)
	
	
		 Dim sConnectionString As String 
		 sConnectionString= "User ID=DaiUser;Password=Daiuser;Initial Catalog=DaidbVPCPayrolltest;Data Source=Daitracker2"
		 Dim sqlConnection1 As New SqlConnection(sConnectionString)
	         Dim cmd As New SqlCommand
	        
		sqlConnection1.Open()
		
		Dim sql1 as string
	
	sql1="INSERT INTO TestXML (Name,Address,PhoneNumber,Enrollment,NoOfStudents,ChangeOfAddress,Counselor,McKinneyProgram,IntraApplication,LunchApplication,BusApplication,YesOther,OtherReason,Appointment,AppointmentDateTime) VALUES ('"& FullName & "','"& Address &"','"& PhoneNumber &"',"& Enrollment &","& NoOfStudents &","& ChangeOfAddress &","& Counselor &","
	sql1= sql1 + " "& McKinneyProgram &","& IntraApplication &","& LunchApplication &","& BusApplication &","& YesOther &",'" & OtherReason &"','"& Appointment &"','"& AppointmentDateTime &" ')"
	cmd = New SqlCommand(Sql1, sqlConnection1)
		
	cmd.ExecuteNonQuery()

	
 
End Sub



</script>


</body>
</html>