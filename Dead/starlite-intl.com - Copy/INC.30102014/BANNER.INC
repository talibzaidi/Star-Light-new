
<%
on error resume next

'****** FOR DIFFERENT PAGES IN SITES A COOKIE OR SESSION VARIABLE CAN HOLD YOUR ADLOC
'****** QUITE NICELY AND WILL HELP CUSTOMIZE BANNER CONTENT IN DIFFERENT AREAS OF YOUR SITE  
'****** ADLOC IS A FIELD IN THE DATABASE BTW.

  
 

'****** ESTABLISH VARIABLES - aryRS HOLDS BANNER IDS, FOR MORE BANNERS ADD 0,'s TO THE ARRAY

	
	count = 0
	
        aryRS = array(0,0,0,0,0,0)
	
 
'****** ESTABLISH DATA-CONNECTION AND RETRIEVE RECORDS - MODIFY FOR YOUR SITE
'****** REMEMBER TO CHANGE THE DSN, THIS EXAMPLE ALSO GRABS GLOBAL ALocations AS
'****** SEEN IN THE OR STATEMENT, THIS IS USEFUL IF YOU HAVE BANNER CONTENT THAT IS BEING
'****** DISPLAYED FOR A PARTICULAR AREA OF YOUR SITE AS WELL AS SITEWIDE BANNERS IN THIS
'****** CASE CITYWIDE, THEN ANOTHER INCLUDE WOULD CONTAIN A THIRD OR STATEMENT FOR
'****** STATE/PROVINCE    
    
  set RS = CreateObject("ADODB.Recordset")
  RS.Open "SELECT * FROM Banner WHERE AShow = YES and ALocation = '" & sar & "'   ", "DSN=STAREC1" , 1, 4

'*** Note that the code in this section needs changed to properly format the output for your site! ***
'
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
	      'Response.Write("BOO!")
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
        'DEBUG' Response.Write("+")
	RS.MoveFirst
	End if
	
'****** WRITE BANNERS TO SCREEN - MODIFY FOR YOUR SITE 
'****** THIS RESPONSE DOES NOT NOT REFLECT ALL THE AVAILABLE FIELDS! LOOK AT THE DATABSE OR ADVERTISEMENT1.INC FOR COMPLETE LIST!
if RS("AURL") = "" then
	Response.Write("<img src=""" & RS("AImage") & """ alt=""Advertisement"" width =""350"" height=""60""  border=""0""><font face=tahoma size=1><br></font>") 
        Response.Write("<br>")
else
Response.Write("<a href=""" & RS("AURL") & """><img src=""" & RS("AImage") & """ alt=""Advertisement"" width =""350"" height=""60""  border=""0""><font face=tahoma size=1><br></font></a>") 
        Response.Write("<br>")

end if

  
    loop
  end if

'****** CLEAN UP THE CONNECTION

  RS.Close
  set RS = nothing

%>