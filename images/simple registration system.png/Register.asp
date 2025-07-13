<%
    ' Capture form data
    Dim name, email, phone, dob, course
    name = Request.Form("name")
    email = Request.Form("email")
    phone = Request.Form("phone")
    dob = Request.Form("dob")
    course = Request.Form("course")

    ' Database connection
    Dim conn
Set conn = Server.CreateObject("ADODB.Connection")

'Path to your Access database (.accdb or .mdb)
Dim dbPath
dbPath = Server.MapPath("school_registration.accdb") 
'Adjust the path as needed

'Connection string using Microsoft Access Database Engine
Dim connString
connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="& dbPath & ";Persist Security Info=False;"

'Open the database connection
conn.Open connString

    ' Insert data into the database
    sql = "INSERT INTO Students (Name, Email, Phone, DOB, Course) VALUES ('" & name & "', '" & email & "', '" & phone & "', '" & dob & "', '" & course & "')"
    conn.Execute(sql)

    ' Close the connection
    conn.Close
    Set conn = Nothing

    ' Display success message
    Response.Write("Registration Successful!")
    %>





