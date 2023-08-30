
<%

'****** FOR DIFFERENT PAGES IN SITES A COOKIE OR SESSION VARIABLE CAN HOLD YOUR ADLOC
'****** QUITE NICELY AND WILL HELP CUSTOMIZE BANNER CONTENT IN DIFFERENT AREAS OF YOUR SITE  
'****** ADLOC IS A FIELD IN THE DATABASE BTW.

  
 

'****** ESTABLISH VARIABLES - aryRS HOLDS BANNER IDS, FOR MORE BANNERS ADD 0,'s TO THE ARRAY

	
        count = 0
	
        aryRS = array(0,0,0,0,0,0)
	
 
'****** ESTABLISH DATA-CONNECTION AND RETRIEVE RECORDS - MODIFY FOR YOUR SITE
    
  ' Response.Write "SELECT * FROM Banner WHERE AShow = YES and ALocation = '" & sar & "'   "

  'set RS = CreateObject("ADODB.Recordset")
  'RS.Open "SELECT * FROM Banner WHERE AShow = YES and ALocation = '" & sar & "'   ", "DSN=STAREC1" , 1, 4

	' 11/9/15: Using the connection string method instead of the DSN method above.
	Set ConnBanner = Server.CreateObject("ADODB.Connection")
	ConnBanner.Open Session("ConnectionString")
	SQLstring = "SELECT * FROM Banner WHERE AShow = YES and ALocation = '" & sar & "'   "
	Set RS = Server.CreateObject("ADODB.Recordset")
	'RS.Open SQLstring, ConnBanner, adOpenStatic, adLockOptimistic, adCmdText 
	RS.Open SQLstring, ConnBanner, 3, 3, &H0001 

'*** Note that the code in this section needs changed to properly format the output for your site! ***

  needrec = 1
  if RS.BOF and RS.EOF then
    'There are no advertisements to show
    'Response.Write("")
  else
    'There are advertisements to show
    count = 0


'****** THE FOLLOWING CODE CHOOSES RANDOM BANNERS FROM YOUR BANNER DATABASE
'****** AND WILL NOT CHOOSE THE SAME ONE TWICE ( IN THIS APPLICATION THE DATABASE
'****** WOULD NEED A MINIMUM OF 5 BANNERS - I RECOMMEND 10-15 MINIMUM )
'****** ON ERROR AT THE END, IT WILL DISPLAY THE FIRST RECORD.
   
'****** CHANGE VALUE OF DO UNTIL COUNT= FOR THE NUMBER OF BANNERS YOU WANT

  RS.MoveLast
  
  numrec = RS.Recordcount
  if numrec > 1 then
     numrec = 1
     needrec = 0
  elseif numrec < 1 then
     needrec = 1 - numrec
  end if

  do until count = numrec
      IF not error then
        Randomize
        rnumber = (Cint(Rnd * RS.RecordCount))
        if count <> 0 then
          
          badflag = 1
          while badflag > 0
            RS.MoveFirst
            
'****** CHANGE THE VALUE 5 IN THE FOLLOWING FOR LOOPS TO REFLECT THE NUMBER OF BANNERS
            
            for i = 0 to numrec
           
	    for z = 0 to numrec
	    
'****** CHECK ARRAY FOR DUPLICATE BANNERS ON DUPLICATE CHOOSE ANOTHER NUMBER ELSE EXIT VIA BADFLAG=0

              if rnumber = aryRS(z) then
                
                rnumber = (Int(Rnd * RS.RecordCount))
		badflag = 1
              end if
	                badflag = 0
            next 
	  next           
          wend
        else
          aryRS(0) = rnumber
        end if
        


      END IF 


'****** SET THE COUNT ON THE LOOP      
      
	count = (count + 1)
        aryRS(count) = rnumber

'****** ERROR HANDLING ROUTINE - ALL THIS DOES IS RE-DISPLAY THE FIRST BANNER ON ERROR

	If (RS.Move(Cint(rnumber)) <> (RS.EOF)) Then
       	RS.MoveFirst
	End if
	
'****** WRITE BANNERS TO SCREEN - MODIFY FOR YOUR SITE 
'****** THIS RESPONSE DOES NOT NOT REFLECT ALL THE AVAILABLE FIELDS! LOOK AT THE DATABSE OR ADVERTISEMENT1.INC FOR COMPLETE LIST!

	Response.Write("<a href=""" & RS("AURL") & """><img src=""" & RS("AImage") & """ alt=""Advertisement"" width =""350"" height=""60""  border=""0""></a>") 
      '  Response.Write("<br>")
      
    loop
  end if

'****** CLEAN UP THE CONNECTION

  RS.Close
  set RS = nothing

%>